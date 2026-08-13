/**
 * PR #17 承認済み新規追加の安全判定テスト。
 * API呼び出し・元データ変更は行わない。
 */

function runHotelDbV2ApprovedNewFacilityInsertTests() {
  const failures = [];
  let success = 0;

  function check(label, condition, detail) {
    if (condition) {
      success++;
    } else {
      failures.push(label + (detail ? ': ' + detail : ''));
    }
  }

  function candidate(overrides) {
    const base = {
      state: '承認',
      sourceSheetName: '鳥取のコピー',
      sourceSheetId: '12345',
      targetMunicipality: '鳥取県鳥取市',
      searchTypes: 'ホテル',
      placeId: 'TEST_PR17_NEW_001',
      name: 'PR17テストホテル',
      address: '鳥取県鳥取市戎町455',
      postalCode: '680-0055',
      candidateMunicipality: '鳥取県鳥取市',
      businessStatus: '営業中',
      recommendation: '新規候補有力'
    };
    Object.keys(overrides || {}).forEach(function(key) {
      base[key] = overrides[key];
    });
    return base;
  }

  function place(overrides) {
    const base = {
      id: 'TEST_PR17_NEW_001',
      displayName: { text: 'PR17テストホテル' },
      formattedAddress: '鳥取県鳥取市戎町455',
      businessStatus: 'OPERATIONAL',
      addressComponents: [
        { longText: '鳥取県', types: ['administrative_area_level_1'] },
        { longText: '鳥取市', types: ['locality'] },
        { longText: '戎町', types: ['sublocality_level_2'] },
        { longText: '455', types: ['premise'] },
        { longText: '680-0055', types: ['postal_code'] }
      ],
      nationalPhoneNumber: '0857-00-0000',
      websiteUri: 'https://example.invalid/pr17',
      googleMapsUri: 'https://maps.google.com/?cid=17',
      rating: 4.2,
      userRatingCount: 17,
      location: { latitude: 35.5, longitude: 134.2 },
      types: ['hotel', 'lodging']
    };
    Object.keys(overrides || {}).forEach(function(key) {
      base[key] = overrides[key];
    });
    return base;
  }

  let result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({}));
  check('1 承認済み正常候補', result.ok === true, result.reason);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({ state: '未確認' }));
  check('2 未承認は対象外', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({ sourceSheetId: '' }));
  check('3 探索元シート情報不足を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({ targetMunicipality: '' }));
  check('4 対象自治体なしを拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({ placeId: '' }));
  check('5 Place IDなしを拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({ name: '' }));
  check('6 施設名なしを拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({ postalCode: '680-55' }));
  check('7 不正郵便番号を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({ businessStatus: '一時休業' }));
  check('8 営業中以外を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityPrecheck_(candidate({
    candidateMunicipality: '大阪市旭区',
    address: '大阪府大阪市旭区赤川1-1'
  }));
  check('9 候補作成時の自治体違いを拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityValidateLive_(candidate({}), place({}));
  check('10 追加直前の正常再確認', result.ok === true, result.reason);

  result = hotelDbV2ApprovedNewFacilityValidateLive_(candidate({}), place({ id: 'OTHER_ID' }));
  check('11 Place ID変化を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityValidateLive_(candidate({}), place({
    businessStatus: 'CLOSED_PERMANENTLY'
  }));
  check('12 閉業へ変化した候補を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityValidateLive_(candidate({}), place({
    formattedAddress: '大阪府大阪市旭区赤川1-1',
    addressComponents: [
      { longText: '大阪府', types: ['administrative_area_level_1'] },
      { longText: '大阪市', types: ['locality'] },
      { longText: '旭区', types: ['sublocality_level_1'] },
      { longText: '535-0005', types: ['postal_code'] }
    ]
  }));
  check('13 自治体外へ変化した候補を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityValidateLive_(candidate({}), place({
    displayName: { text: 'PR17別ホテル' }
  }));
  check('14 施設名変更を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityValidateLive_(candidate({}), place({
    formattedAddress: '鳥取県鳥取市戎町999'
  }));
  check('15 住所変更を拒否', result.ok === false);

  result = hotelDbV2ApprovedNewFacilityValidateLive_(candidate({}), place({
    addressComponents: [
      { longText: '鳥取県', types: ['administrative_area_level_1'] },
      { longText: '鳥取市', types: ['locality'] },
      { longText: '戎町', types: ['sublocality_level_2'] },
      { longText: '455', types: ['premise'] },
      { longText: '680-0056', types: ['postal_code'] }
    ]
  }));
  check('16 郵便番号変更を拒否', result.ok === false);

  const insertData = hotelDbV2ApprovedNewFacilityBuildInsertData_(candidate({}), place({}));
  check('17 元DB用住所は自治体部分を除去', insertData.address === '戎町455', insertData.address);
  check('18 宿泊分類を自動入力しない', insertData.category === '', insertData.category);
  check(
    '19 備考に宿泊分類要確認を残す',
    insertData.notes.indexOf('宿泊分類要確認') !== -1,
    insertData.notes
  );
  check('20 Google営業状態は営業中で保存', insertData.businessStatus === '営業中');

  if (failures.length) {
    throw new Error([
      '承認済み新規追加 安全判定テスト失敗',
      '',
      '成功件数: ' + success + '件',
      '失敗件数: ' + failures.length + '件',
      '',
      failures.join('\n')
    ].join('\n'));
  }

  SpreadsheetApp.getUi().alert([
    '承認済み新規追加 安全判定テスト 成功',
    '',
    '成功件数: 20件',
    '失敗件数: 0件',
    '未承認・閉業・自治体違いの自動追加: なし',
    '候補作成後の名称・住所・郵便番号変更の自動追加: なし',
    '宿泊分類の自動入力: なし',
    '元データの変更: なし'
  ].join('\n'));

  return { success: 20, failure: 0, sourceAutoChange: false };
}

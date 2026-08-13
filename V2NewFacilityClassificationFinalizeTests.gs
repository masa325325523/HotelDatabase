/**
 * PR #18 自己診断。
 * API呼び出し・Spreadsheet書き換えは行わず、安全判定の純粋ロジックだけを確認する。
 */
function runHotelDbV2NewFacilityClassificationFinalizeTests() {
  const failures = [];
  let success = 0;

  function check(label, condition, detail) {
    if (condition) success++;
    else failures.push(label + (detail ? ': ' + detail : ''));
  }

  function reference(types) {
    return hotelDbV2NewFacilityClassificationRecommend_(types).reference;
  }

  check('1 japanese_innは旅館系参考', reference('japanese_inn,lodging') === '旅館系（要確認）', reference('japanese_inn,lodging'));
  check('2 budget_japanese_innは簡易宿所系参考', reference('budget_japanese_inn,lodging') === '簡易宿所系（要確認）', reference('budget_japanese_inn,lodging'));
  check('3 guest_houseは簡易宿所系参考', reference('guest_house,lodging') === '簡易宿所系（要確認）');
  check('4 hostelは簡易宿所系参考', reference('hostel,lodging') === '簡易宿所系（要確認）');
  check('5 private_guest_roomは民泊・簡易宿所系参考', reference('private_guest_room,lodging') === '住宅宿泊事業・簡易宿所系（要確認）');
  check('6 hotelはホテル系参考', reference('hotel,lodging') === 'ホテル系（要確認）');
  check('7 lodgingだけなら広い参考', reference('lodging') === '旅館・簡易宿所系（要確認）');
  check('8 Googleタイプなしは要確認', reference('') === '要確認');

  const valid = {
    sheetName: 'テスト', sheetId: 123, sourceRow: 2,
    placeId: 'PLACE_1', name: 'テストホテル', address: '東京都千代田区丸の内1-1',
    currentCategory: '', currentNotes: 'PR17承認新規追加／宿泊分類要確認／Google探索種別:ホテル',
    finalCategory: 'ホテル営業', finalNotes: '郵便番号一致'
  };
  check('9 正常な承認入力は事前確認通過', hotelDbV2NewFacilityClassificationPrecheck_(valid).ok);
  check('10 確定宿泊分類なしは停止', !hotelDbV2NewFacilityClassificationPrecheck_(Object.assign({}, valid, { finalCategory: '' })).ok);
  check('11 元シート情報なしは停止', !hotelDbV2NewFacilityClassificationPrecheck_(Object.assign({}, valid, { sheetId: 0 })).ok);
  check('12 Place IDなしは停止', !hotelDbV2NewFacilityClassificationPrecheck_(Object.assign({}, valid, { placeId: '' })).ok);
  check('13 長すぎる分類は停止', !hotelDbV2NewFacilityClassificationPrecheck_(Object.assign({}, valid, { finalCategory: new Array(102).join('あ') })).ok);

  const facility = {
    placeId: 'PLACE_1', name: 'テストホテル', address: '東京都千代田区丸の内1-1',
    category: '', notes: valid.currentNotes
  };
  check('14 Place ID・施設名・住所一致は本人確認通過', hotelDbV2NewFacilityClassificationIdentityMatches_(valid, facility).ok);
  check('15 Place ID違いは停止', !hotelDbV2NewFacilityClassificationIdentityMatches_(valid, Object.assign({}, facility, { placeId: 'PLACE_2' })).ok);
  check('16 施設名違いは停止', !hotelDbV2NewFacilityClassificationIdentityMatches_(valid, Object.assign({}, facility, { name: '別ホテル' })).ok);
  check('17 住所違いは停止', !hotelDbV2NewFacilityClassificationIdentityMatches_(valid, Object.assign({}, facility, { address: '東京都千代田区丸の内1-2' })).ok);

  check('18 正常スナップショットは反映可能', hotelDbV2NewFacilityClassificationSnapshotCheck_(valid, facility).ok);
  check('19 既存分類ありは上書き禁止', !hotelDbV2NewFacilityClassificationSnapshotCheck_(valid, Object.assign({}, facility, { category: '旅館営業' })).ok);
  check('20 備考変更後は停止', !hotelDbV2NewFacilityClassificationSnapshotCheck_(valid, Object.assign({}, facility, { notes: facility.notes + '／手修正' })).ok);
  check('21 要確認マーカー消失は停止', !hotelDbV2NewFacilityClassificationSnapshotCheck_(Object.assign({}, valid, { currentNotes: 'PR17承認新規追加' }), Object.assign({}, facility, { notes: 'PR17承認新規追加' })).ok);

  const merged1 = hotelDbV2NewFacilityClassificationMergeNotes_(valid.currentNotes, '郵便番号一致');
  check('22 宿泊分類要確認だけ除去', merged1.indexOf('宿泊分類要確認') === -1, merged1);
  check('23 既存備考を保持して確認済みを追加', merged1.indexOf('PR17承認新規追加') !== -1 && merged1.indexOf('Google探索種別:ホテル') !== -1 && merged1.indexOf('宿泊分類確認済') !== -1, merged1);
  check('24 人が入力した確定備考を追記', merged1.indexOf('郵便番号一致') !== -1, merged1);

  if (failures.length) {
    throw new Error([
      '新規施設分類・備考 安全判定テスト失敗',
      '',
      '成功件数: ' + success + '件',
      '失敗件数: ' + failures.length + '件',
      '',
      failures.join('\n')
    ].join('\n'));
  }

  SpreadsheetApp.getUi().alert([
    '新規施設分類・備考 安全判定テスト 成功',
    '',
    '成功件数: 24件',
    '失敗件数: 0件',
    'Googleタイプからの自動確定: なし',
    '既存宿泊分類の上書き: なし',
    '施設名・住所・Place ID変更時の反映: なし',
    '備考変更時の反映: なし',
    '宿泊分類要確認以外の既存備考の削除: なし',
    '元データの変更: なし'
  ].join('\n'));

  return { success: true, successCount: 24, failureCount: 0, sourceAutoChange: false };
}

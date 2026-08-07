/**
 * PR #7: 郵便番号正規化を本体へ組み込んだことを確認する回帰テスト。
 * シートやB列「状態」は変更しない。
 */
function runHotelDbV2CorePostalNormalizationTests() {
  const failures = [];
  let passed = 0;

  function check(label, actual, expected) {
    if (actual !== expected) {
      failures.push(
        label + ': ' + String(actual) + '（期待=' + String(expected) + '）'
      );
    } else {
      passed++;
    }
  }

  const normalizationCases = [
    ['ASCII標準', '680-0907', '680-0907'],
    ['末尾だけ全角', '680-090７', '680-0907'],
    ['すべて全角', '６８０－０９０７', '680-0907'],
    ['全角ハイフン', '680－0907', '680-0907'],
    ['〒と空白付き', '〒 ６８０－０９０７', '680-0907'],
    ['ハイフンなし', '6800907', '680-0907'],
    ['空白区切り', '680 0907', '680-0907'],
    ['桁不足は無効', '680-090', '']
  ];

  normalizationCases.forEach(function(testCase) {
    check(
      '正規化「' + testCase[0] + '」',
      hotelDbV2NormalizePostalCode_(testCase[1]),
      testCase[2]
    );
  });

  const facility = {
    postalCode: '680-090７',
    municipality: '鳥取県鳥取市',
    address: '賀露町北1-5-36',
    name: '味覚のお宿 山田屋'
  };

  const query = hotelDbV2BuildSearchQuery_(facility);
  check(
    '検索クエリに正規化済み郵便番号を含む',
    query.indexOf('680-0907') !== -1,
    true
  );

  const place = {
    id: 'test-place-6800907',
    displayName: { text: '味覚のお宿 山田屋' },
    formattedAddress: '鳥取県鳥取市賀露町北1-5-36',
    addressComponents: [
      { longText: '680-0907', types: ['postal_code'] },
      { longText: '鳥取県', types: ['administrative_area_level_1'] },
      { longText: '鳥取市', types: ['locality'] }
    ],
    businessStatus: 'OPERATIONAL',
    googleMapsUri: 'https://maps.google.com/?q=test'
  };

  const score = hotelDbV2CalculateMatchScore_(facility, place);
  check('全角郵便番号でも郵便番号点を含めて100点', score, 100);

  const googleData = hotelDbV2BuildGoogleData_(
    place,
    score,
    facility,
    'テスト'
  );
  const differences = hotelDbV2CompareFacility_(facility, googleData);
  check(
    '全角半角だけでは郵便番号差分を作らない',
    differences.indexOf('郵便番号') === -1,
    true
  );
  check('完全一致ケースは差分0件', differences.length, 0);

  const differentPostalData = Object.assign({}, googleData, {
    proposedPostalCode: '680-0908'
  });
  const differentPostalDifferences = hotelDbV2CompareFacility_(
    facility,
    differentPostalData
  );
  check(
    '実際に違う郵便番号は差分として残す',
    differentPostalDifferences.indexOf('郵便番号') !== -1,
    true
  );

  const nameDifferenceData = Object.assign({}, googleData, {
    proposedName: '味覚のお宿 山田屋 別館'
  });
  const nameDifferences = hotelDbV2CompareFacility_(
    facility,
    nameDifferenceData
  );
  check(
    '郵便番号表記差を消しても施設名差分は残す',
    nameDifferences.join('・'),
    '施設名'
  );

  const message = [
    failures.length
      ? '本体郵便番号正規化テスト 要確認'
      : '本体郵便番号正規化テスト 成功',
    '',
    '成功件数: ' + passed + '件',
    '失敗件数: ' + failures.length + '件',
    '状態列の自動変更: なし'
  ];

  if (failures.length) {
    message.push('', failures.join('\n'));
  }

  SpreadsheetApp.getUi().alert(message.join('\n'));

  if (failures.length) {
    throw new Error(failures.join(' / '));
  }

  return {
    passed: passed,
    failed: failures.length,
    stateChanged: false
  };
}

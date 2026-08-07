/**
 * Ver2.0 Google住所末尾の施設内部表記ノイズ対策。
 *
 * 「ロビー」「フロント」「受付」等は、削除後の住所が元住所と一致する場合に限り
 * 住所差分として無視する。番地・建物名・部屋番号等の実差分は保持する。
 */

const HOTEL_DB_V2_INTERNAL_LOCATION_SUFFIXES = Object.freeze([
  'フロントデスク',
  '受付カウンター',
  'レセプション',
  'ロビー',
  'フロント',
  '受付'
]);

function hotelDbV2StripTrailingInternalLocationSuffix_(address) {
  const original = hotelDbV2Clean_(address).normalize('NFKC');
  if (!original) return original;

  for (let i = 0; i < HOTEL_DB_V2_INTERNAL_LOCATION_SUFFIXES.length; i++) {
    const pattern = hotelDbV2FlexibleSuffixPattern_(
      HOTEL_DB_V2_INTERNAL_LOCATION_SUFFIXES[i]
    );
    if (!pattern) continue;

    const stripped = original.replace(
      new RegExp(
        '[\\s　・･,，、/\\\\()（）\\[\\]【】]*' +
        pattern +
        '[\\s　・･,，、/\\\\()（）\\[\\]【】]*$',
        'iu'
      ),
      ''
    ).trim();

    if (stripped !== original) return stripped;
  }

  return original;
}

// V2ConfigApi.gs の既存処理を保持したうえで、施設内部表記だけを追加で処理する。
const hotelDbV2CleanAddressForSourceBeforeInternalNoise_ =
  hotelDbV2CleanAddressForSource_;

hotelDbV2CleanAddressForSource_ = function(
  candidateAddress,
  sourceAddress,
  sourceName,
  googleName
) {
  const cleaned = hotelDbV2CleanAddressForSourceBeforeInternalNoise_(
    candidateAddress,
    sourceAddress,
    sourceName,
    googleName
  );
  const stripped = hotelDbV2StripTrailingInternalLocationSuffix_(cleaned);

  // 末尾に対象語がない、または元住所がない場合は従来結果をそのまま使う。
  if (stripped === cleaned || !sourceAddress) return cleaned;

  const sourceComparable = hotelDbV2NormalizeAddressForComparison_(sourceAddress);
  const strippedComparable = hotelDbV2NormalizeAddressForComparison_(stripped);

  // 対象語だけを除けば元住所と同一になる場合に限って元住所を採用する。
  // 番地・建物名・部屋番号等が異なる場合は cleaned を保持し、差分判定に残す。
  return sourceComparable && sourceComparable === strippedComparable
    ? hotelDbV2Clean_(sourceAddress)
    : cleaned;
};

/**
 * 施設内部表記ノイズ対策の回帰テスト。
 * Apps Script の関数一覧から runHotelDbV2InternalAddressNoiseTests を実行する。
 */
function runHotelDbV2InternalAddressNoiseTests() {
  const cases = [
    {
      name: '末尾ロビーは元住所と一致する場合だけ無視',
      sourceAddress: '永楽温泉町651',
      googleAddress: '永楽温泉町651 ロビー',
      expectedCleaned: '永楽温泉町651',
      expectedAddressDifference: false
    },
    {
      name: '末尾フロントを無視',
      sourceAddress: '永楽温泉町651',
      googleAddress: '永楽温泉町651 フロント',
      expectedCleaned: '永楽温泉町651',
      expectedAddressDifference: false
    },
    {
      name: '末尾受付を無視',
      sourceAddress: '永楽温泉町651',
      googleAddress: '永楽温泉町651 受付',
      expectedCleaned: '永楽温泉町651',
      expectedAddressDifference: false
    },
    {
      name: '末尾レセプションを無視',
      sourceAddress: '永楽温泉町651',
      googleAddress: '永楽温泉町651 レセプション',
      expectedCleaned: '永楽温泉町651',
      expectedAddressDifference: false
    },
    {
      name: '番地が違えばロビー付きでも住所差分を保持',
      sourceAddress: '永楽温泉町651',
      googleAddress: '永楽温泉町652 ロビー',
      expectedCleaned: '永楽温泉町652 ロビー',
      expectedAddressDifference: true
    },
    {
      name: '部屋番号が違えばロビー付きでも住所差分を保持',
      sourceAddress: '春岡一丁目4番21号 Fuchsia901',
      googleAddress: '春岡1-4-21 Fuchsia301 ロビー',
      expectedCleaned: '春岡1-4-21 Fuchsia301 ロビー',
      expectedAddressDifference: true
    }
  ];

  const failures = [];
  const results = [];

  cases.forEach(function(testCase, index) {
    const cleanedAddress = hotelDbV2CleanAddressForSource_(
      testCase.googleAddress,
      testCase.sourceAddress,
      'テスト旅館',
      'テスト旅館'
    );

    const differences = hotelDbV2CompareFacility_(
      {
        postalCode: '',
        municipality: '',
        address: testCase.sourceAddress,
        name: 'テスト旅館'
      },
      {
        proposedPostalCode: '',
        proposedMunicipality: '',
        proposedAddress: cleanedAddress,
        proposedName: 'テスト旅館'
      }
    );

    const addressDifference = differences.indexOf('住所') !== -1;
    const passed =
      cleanedAddress === testCase.expectedCleaned &&
      addressDifference === testCase.expectedAddressDifference;

    results.push({
      number: index + 1,
      name: testCase.name,
      passed: passed,
      cleanedAddress: cleanedAddress,
      addressDifference: addressDifference
    });

    if (!passed) {
      failures.push(
        '例' + (index + 1) + '「' + testCase.name + '」' +
        ': cleanedAddress=' + cleanedAddress +
        '（期待=' + testCase.expectedCleaned + '）' +
        ', addressDifference=' + addressDifference +
        '（期待=' + testCase.expectedAddressDifference + '）'
      );
    }
  });

  if (failures.length) {
    const message = [
      '施設内部表記ノイズ対策テストに失敗しました。',
      '',
      failures.join('\n')
    ].join('\n');

    console.error(message);
    throw new Error(message);
  }

  const successMessage = [
    '施設内部表記ノイズ対策テスト成功',
    '',
    '成功件数: ' + results.length + '件',
    '失敗件数: 0件'
  ].join('\n');

  console.log(successMessage);

  try {
    SpreadsheetApp.getUi().alert(successMessage);
  } catch (error) {
    // スプレッドシートに紐付いていない実行環境ではログだけを使用する。
  }

  return {
    success: true,
    passed: results.length,
    failed: 0,
    results: results
  };
}

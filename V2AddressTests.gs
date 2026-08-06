/**
 * Ver2.0 住所正規化の自己診断テスト。
 * Apps Scriptの関数一覧から runHotelDbV2AddressNormalizationTests を実行する。
 */
function runHotelDbV2AddressNormalizationTests() {
  const cases = [
    {
      name: '施設名・施設種別・同一番地の重複を除く',
      sourceAddress: '戎町455',
      googleAddress: '戎町４５５ 旅館 とらや455',
      sourceName: 'とらや旅館',
      googleName: 'とらや旅館',
      equivalent: true,
      cleanedAddress: '戎町455'
    },
    {
      name: '全角数字と半角数字を同一視する',
      sourceAddress: '栄町230',
      googleAddress: '栄町２３０',
      sourceName: '常天',
      googleName: '寿司 旅館 常天',
      equivalent: true,
      cleanedAddress: '栄町230'
    },
    {
      name: '各種ハイフンを同一視する',
      sourceAddress: '商栄町108-1',
      googleAddress: '商栄町１０８−１',
      sourceName: '満月荘',
      googleName: '満月荘',
      equivalent: true,
      cleanedAddress: '商栄町108-1'
    },
    {
      name: '異なる番地は住所差分にする',
      sourceAddress: '商栄町108-1',
      googleAddress: '商栄町108-2',
      sourceName: '満月荘',
      googleName: '満月荘',
      equivalent: false,
      cleanedAddress: '商栄町108-2'
    },
    {
      name: '建物名と部屋番号の違いを残す',
      sourceAddress: '春岡一丁目4番21号 Fuchsia901',
      googleAddress: '春岡1-4-21 Fuchsia301',
      sourceName: 'Japan Hinata 1',
      googleName: 'Japan Hinata 1',
      equivalent: false,
      cleanedAddress: '春岡1-4-21 Fuchsia301'
    },
    {
      name: '末尾の施設種別だけを除く',
      sourceAddress: '戎町455',
      googleAddress: '戎町455 ホテル',
      sourceName: 'とらや旅館',
      googleName: 'とらや旅館',
      equivalent: true,
      cleanedAddress: '戎町455'
    }
  ];

  const failures = [];
  const results = [];

  cases.forEach(function(testCase, index) {
    const equivalent = hotelDbV2AddressesEquivalent_(
      testCase.sourceAddress,
      testCase.googleAddress,
      testCase.sourceName,
      testCase.googleName
    );

    const cleanedAddress = hotelDbV2CleanAddressForSource_(
      testCase.googleAddress,
      testCase.sourceAddress,
      testCase.sourceName,
      testCase.googleName
    );

    const differences = hotelDbV2CompareFacility_(
      {
        postalCode: '',
        municipality: '',
        address: testCase.sourceAddress,
        name: testCase.sourceName
      },
      {
        proposedPostalCode: '',
        proposedMunicipality: '',
        proposedAddress: cleanedAddress,
        proposedName: testCase.sourceName
      }
    );

    const addressDifference = differences.indexOf('住所') !== -1;
    const expectedAddressDifference = !testCase.equivalent;
    const passed =
      equivalent === testCase.equivalent &&
      cleanedAddress === testCase.cleanedAddress &&
      addressDifference === expectedAddressDifference;

    results.push({
      number: index + 1,
      name: testCase.name,
      passed: passed,
      equivalent: equivalent,
      cleanedAddress: cleanedAddress,
      addressDifference: addressDifference
    });

    if (!passed) {
      failures.push(
        '例' + (index + 1) + '「' + testCase.name + '」' +
        ': equivalent=' + equivalent +
        '（期待=' + testCase.equivalent + '）' +
        ', cleanedAddress=' + cleanedAddress +
        '（期待=' + testCase.cleanedAddress + '）' +
        ', addressDifference=' + addressDifference +
        '（期待=' + expectedAddressDifference + '）'
      );
    }
  });

  if (failures.length) {
    const message = [
      '住所正規化テストに失敗しました。',
      '',
      failures.join('\n')
    ].join('\n');

    console.error(message);
    throw new Error(message);
  }

  const successMessage = [
    '住所正規化テスト成功',
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

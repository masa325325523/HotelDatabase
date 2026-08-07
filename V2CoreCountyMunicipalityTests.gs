/**
 * PR #8: 郡・町村の列分け差を本体の差分判定へ組み込んだことを確認する回帰テスト。
 * シートやB列「状態」は変更しない。
 */
function runHotelDbV2CoreCountyMunicipalityTests() {
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

  function makeFacility(overrides) {
    return Object.assign({
      postalCode: '681-0024',
      municipality: '鳥取県岩美郡',
      address: '岩美町岩井536',
      name: '明石家'
    }, overrides || {});
  }

  function makeGoogleData(overrides) {
    return Object.assign({
      placeId: 'test-place-id',
      proposedPostalCode: '681-0024',
      proposedMunicipality: '鳥取県岩美町',
      proposedAddress: '岩美郡岩美町岩井536',
      proposedName: '明石家',
      matchScore: 93,
      businessStatus: '営業中'
    }, overrides || {});
  }

  function diff(facility, googleData) {
    return hotelDbV2CompareFacility_(facility, googleData).join('・');
  }

  check(
    '岩美郡と岩美町の列分け差は本体で差分0件',
    diff(makeFacility(), makeGoogleData()),
    ''
  );

  check(
    '八頭郡と八頭町の列分け差は本体で差分0件',
    diff(
      makeFacility({
        postalCode: '680-0611',
        municipality: '鳥取県八頭郡',
        address: '八頭町富枝38',
        name: '古民家 太田邸'
      }),
      makeGoogleData({
        proposedPostalCode: '680-0611',
        proposedMunicipality: '鳥取県八頭町',
        proposedAddress: '八頭郡八頭町富枝38',
        proposedName: '古民家 太田邸'
      })
    ),
    ''
  );

  check(
    '東伯郡と三朝町の列分け差は本体で差分0件',
    diff(
      makeFacility({
        postalCode: '682-0123',
        municipality: '鳥取県東伯郡',
        address: '三朝町三朝365-1',
        name: '依山楼岩崎'
      }),
      makeGoogleData({
        proposedPostalCode: '682-0123',
        proposedMunicipality: '鳥取県三朝町',
        proposedAddress: '東伯郡三朝町三朝365-1',
        proposedName: '依山楼岩崎'
      })
    ),
    ''
  );

  check(
    '日野郡江府町の大字省略も差分0件',
    diff(
      makeFacility({
        postalCode: '689-4401',
        municipality: '鳥取県日野郡',
        address: '江府町大字江尾2064',
        name: '門脇旅館'
      }),
      makeGoogleData({
        proposedPostalCode: '689-4401',
        proposedMunicipality: '鳥取県江府町',
        proposedAddress: '日野郡江府町江尾2064',
        proposedName: '門脇旅館'
      })
    ),
    ''
  );

  check(
    'PR7の全角郵便番号正規化と併用しても差分0件',
    diff(
      makeFacility({ postalCode: '681-002４' }),
      makeGoogleData({ proposedPostalCode: '681-0024' })
    ),
    ''
  );

  check(
    '番地違いは市区町村・住所差分を残す',
    diff(
      makeFacility(),
      makeGoogleData({ proposedAddress: '岩美郡岩美町岩井544' })
    ),
    '市区町村・住所'
  );

  check(
    '町が違えば市区町村・住所差分を残す',
    diff(
      makeFacility(),
      makeGoogleData({
        proposedMunicipality: '鳥取県八頭町',
        proposedAddress: '岩美郡八頭町岩井536'
      })
    ),
    '市区町村・住所'
  );

  check(
    '郵便番号が本当に違えば郵便番号差分も残す',
    diff(
      makeFacility(),
      makeGoogleData({ proposedPostalCode: '681-0003' })
    ),
    '郵便番号・市区町村・住所'
  );

  check(
    '施設名が違えば施設名差分も残す',
    diff(
      makeFacility(),
      makeGoogleData({ proposedName: '別の旅館' })
    ),
    '市区町村・住所・施設名'
  );

  check(
    '一致スコア89は同一扱いしない',
    diff(
      makeFacility(),
      makeGoogleData({ matchScore: 89 })
    ),
    '市区町村・住所'
  );

  check(
    '営業中以外は同一扱いしない',
    diff(
      makeFacility(),
      makeGoogleData({ businessStatus: '閉業' })
    ),
    '市区町村・住所'
  );

  check(
    'Place IDなしは同一扱いしない',
    diff(
      makeFacility(),
      makeGoogleData({ placeId: '' })
    ),
    '市区町村・住所'
  );

  check(
    '通常の完全一致は従来どおり差分0件',
    diff(
      makeFacility({
        postalCode: '680-0831',
        municipality: '鳥取県鳥取市',
        address: '栄町230',
        name: '常天'
      }),
      makeGoogleData({
        proposedPostalCode: '680-0831',
        proposedMunicipality: '鳥取県鳥取市',
        proposedAddress: '栄町230',
        proposedName: '常天',
        matchScore: 100
      })
    ),
    ''
  );

  check(
    '郡形式でない通常の市区町村差・住所差は残す',
    diff(
      makeFacility({
        postalCode: '680-0001',
        municipality: '鳥取県鳥取市',
        address: '本町1-1',
        name: 'テスト旅館'
      }),
      makeGoogleData({
        proposedPostalCode: '680-0001',
        proposedMunicipality: '鳥取県米子市',
        proposedAddress: '本町1-1',
        proposedName: 'テスト旅館',
        matchScore: 93
      })
    ),
    '市区町村'
  );

  const message = [
    failures.length
      ? '本体 郡・町村表記差テスト 要確認'
      : '本体 郡・町村表記差テスト 成功',
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

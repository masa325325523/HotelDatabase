/**
 * PR #11: 本体「大字」住所正規化の回帰テスト。
 * 元データや状態列は変更しない。
 */
function runHotelDbV2CoreOazaAddressTests() {
  const failures = [];
  let passed = 0;

  function check(name, actual, expected) {
    const ok = JSON.stringify(actual) === JSON.stringify(expected);
    if (ok) {
      passed++;
      return;
    }
    failures.push(
      name + ': actual=' + JSON.stringify(actual) +
      ', expected=' + JSON.stringify(expected)
    );
  }

  check(
    '大字を比較用住所から除く',
    hotelDbV2NormalizeAddressForComparison_('江府町大字江尾2064'),
    '江府町江尾2064'
  );

  check(
    '全角の大字表記もNFKC後に同一化',
    hotelDbV2NormalizeAddressForComparison_('江府町 大字江尾２０６４'),
    '江府町江尾2064'
  );

  check(
    '大字の有無だけなら住所同一',
    hotelDbV2AddressesEquivalent_(
      '江府町大字江尾2064',
      '江府町江尾2064',
      '門脇旅館',
      '門脇旅館'
    ),
    true
  );

  check(
    '大字があっても番地違いは別住所',
    hotelDbV2AddressesEquivalent_(
      '江府町大字江尾2064',
      '江府町江尾2065',
      '門脇旅館',
      '門脇旅館'
    ),
    false
  );

  check(
    '階追加は大字正規化で消さない',
    hotelDbV2AddressesEquivalent_(
      '江府町大字江尾2064',
      '江府町江尾2064 1階',
      '門脇旅館',
      '門脇旅館'
    ),
    false
  );

  const facility = {
    postalCode: '689-4401',
    municipality: '鳥取県日野郡',
    address: '江府町大字江尾2064',
    name: '門脇旅館'
  };

  const place = {
    id: 'test-kadowaki-place',
    displayName: { text: '門脇旅館' },
    formattedAddress: '鳥取県日野郡江府町江尾2064',
    addressComponents: [
      { longText: '689-4401', types: ['postal_code'] },
      { longText: '鳥取県', types: ['administrative_area_level_1'] },
      { longText: '日野郡', types: ['administrative_area_level_2'] },
      { longText: '江府町', types: ['locality'] }
    ],
    businessStatus: 'OPERATIONAL'
  };

  check(
    '門脇旅館型の本体一致スコアは93',
    hotelDbV2CalculateMatchScore_(facility, place),
    93
  );

  const googleData93 = {
    proposedPostalCode: '689-4401',
    proposedMunicipality: '鳥取県江府町',
    proposedAddress: '日野郡江府町江尾2064',
    proposedName: '門脇旅館',
    placeId: 'test-kadowaki-place',
    matchScore: 93,
    businessStatus: '営業中'
  };

  check(
    '93点なら郡町村＋大字差を本体で差分なしにする',
    hotelDbV2CompareFacility_(facility, googleData93),
    []
  );

  const googleData78 = Object.assign({}, googleData93, { matchScore: 78 });
  check(
    '既存78点候補も厳しい大字条件なら差分なしにできる',
    hotelDbV2CompareFacility_(facility, googleData78),
    []
  );

  const googleData74 = Object.assign({}, googleData93, { matchScore: 74 });
  check(
    '75点未満は大字差でも自動同一化しない',
    hotelDbV2CompareFacility_(facility, googleData74),
    ['市区町村', '住所']
  );

  check(
    '郵便番号違いは大字差でも残す',
    hotelDbV2CompareFacility_(
      facility,
      Object.assign({}, googleData93, { proposedPostalCode: '689-4402' })
    ),
    ['郵便番号', '市区町村', '住所']
  );

  check(
    '施設名違いは大字差でも残す',
    hotelDbV2CompareFacility_(
      facility,
      Object.assign({}, googleData93, { proposedName: '別の旅館' })
    ),
    ['市区町村', '住所', '施設名']
  );

  check(
    '番地違いは93点でも残す',
    hotelDbV2CompareFacility_(
      facility,
      Object.assign({}, googleData93, {
        proposedAddress: '日野郡江府町江尾2065'
      })
    ),
    ['市区町村', '住所']
  );

  if (failures.length) {
    throw new Error(
      '本体 大字住所正規化テスト失敗\n\n' + failures.join('\n')
    );
  }

  const message = [
    '本体 大字住所正規化テスト 成功',
    '',
    '成功件数: ' + passed + '件',
    '失敗件数: 0件',
    '状態列の自動変更: なし'
  ].join('\n');

  console.log(message);
  SpreadsheetApp.getUi().alert(message);

  return {
    success: passed,
    failure: 0,
    stateAutoChange: false
  };
}

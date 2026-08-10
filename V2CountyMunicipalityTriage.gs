/**
 * Ver2.0 郡・町村の列分け差を安全に再判定する。
 *
 * 例:
 * 元:     市区町村=鳥取県岩美郡 / 住所=岩美町岩井536
 * Google: 市区町村=鳥取県岩美町 / 住所=岩美郡岩美町岩井536
 *
 * 結合すると同一所在地になる場合、修正候補の推奨判定だけを
 * 「却下候補」に更新する。B列「状態」と元データは変更しない。
 */

const HOTEL_DB_V2_COUNTY_TRIAGE_MIN_SCORE = 90;

function runHotelDbV2CountyMunicipalityTriage() {
  return withHotelDbV2Lock_('郡・町村表記差の再判定', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '郡・町村表記差を再判定',
      '「修正候補」のうち、郡と町村の列分けが違うだけの候補を再判定します。\n\n' +
      'B列「状態」は変更しません。\n' +
      '元データも変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2RefineCountyMunicipalityTriage_();

    ui.alert([
      '郡・町村表記差の再判定完了',
      '',
      '確認件数: ' + result.scanned,
      '却下候補へ変更: ' + result.refined,
      '変更なし: ' + result.unchanged,
      '',
      'B列「状態」は変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2RefineCountyMunicipalityTriage_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(
    HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS
  );

  if (!sheet || sheet.getLastRow() < 2) {
    return { scanned: 0, refined: 0, unchanged: 0 };
  }

  const headerMap = hotelDbV2TriageEnsureHeaders_(sheet);
  hotelDbV2TriageValidateHeaders_(headerMap);

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getDisplayValues();

  const recommendationColumn = headerMap['推奨判定'];
  const reasonColumn = headerMap['自動判定理由'];
  const confidenceColumn = headerMap['信頼度'];

  const result = {
    scanned: values.length,
    refined: 0,
    unchanged: 0
  };

  values.forEach(function(row, offset) {
    const currentRecommendation = hotelDbV2Clean_(
      row[recommendationColumn - 1]
    );

    if (
      currentRecommendation &&
      currentRecommendation !== '要人確認'
    ) {
      result.unchanged++;
      return;
    }

    const input = hotelDbV2TriageInputFromRow_(row, headerMap);
    const decision = hotelDbV2CountyMunicipalityDecision_(input);

    if (!decision.equivalent) {
      result.unchanged++;
      return;
    }

    const rowNumber = offset + 2;
    sheet.getRange(rowNumber, recommendationColumn).setValue('却下候補');
    sheet.getRange(rowNumber, reasonColumn).setValue(decision.reason);
    sheet.getRange(rowNumber, confidenceColumn).setValue(decision.confidence);
    result.refined++;
  });

  return result;
}

function hotelDbV2CountyMunicipalityDecision_(input) {
  const data = input || {};
  const differences = hotelDbV2Clean_(data.differences)
    .split('・')
    .map(hotelDbV2Clean_)
    .filter(Boolean)
    .sort();

  if (hotelDbV2Clean_(data.state) === '反映済み') {
    return hotelDbV2CountyMunicipalityNo_('反映済み');
  }

  if (!hotelDbV2Clean_(data.placeId)) {
    return hotelDbV2CountyMunicipalityNo_('Place IDなし');
  }

  if (hotelDbV2Clean_(data.businessStatus) !== '営業中') {
    return hotelDbV2CountyMunicipalityNo_('営業中以外');
  }

  if (
    differences.length !== 2 ||
    differences[0] !== '住所' ||
    differences[1] !== '市区町村'
  ) {
    return hotelDbV2CountyMunicipalityNo_('対象差分ではない');
  }

  const sourcePostal = hotelDbV2NormalizePostalCode_(data.sourcePostalCode);
  const proposedPostal = hotelDbV2NormalizePostalCode_(data.proposedPostalCode);

  if (!sourcePostal || !proposedPostal || sourcePostal !== proposedPostal) {
    return hotelDbV2CountyMunicipalityNo_('郵便番号不一致');
  }

  if (
    hotelDbV2NormalizeText_(data.sourceName) !==
    hotelDbV2NormalizeText_(data.proposedName)
  ) {
    return hotelDbV2CountyMunicipalityNo_('施設名不一致');
  }

  const sourceMunicipality = hotelDbV2Clean_(data.sourceMunicipality)
    .normalize('NFKC')
    .replace(/\s+/g, '');
  const proposedMunicipality = hotelDbV2Clean_(data.proposedMunicipality)
    .normalize('NFKC')
    .replace(/\s+/g, '');
  const sourceAddress = hotelDbV2Clean_(data.sourceAddress)
    .normalize('NFKC')
    .replace(/^\s+|\s+$/g, '');
  const proposedAddress = hotelDbV2Clean_(data.proposedAddress)
    .normalize('NFKC')
    .replace(/^\s+|\s+$/g, '');

  const countyMatch = sourceMunicipality.match(/^(.+?[都道府県])(.+?郡)$/u);
  if (!countyMatch) {
    return hotelDbV2CountyMunicipalityNo_('元市区町村が郡形式ではない');
  }

  const prefecture = countyMatch[1];
  const county = countyMatch[2];
  const townMatch = sourceAddress.match(/^(.+?[町村])(.*)$/u);

  if (!townMatch) {
    return hotelDbV2CountyMunicipalityNo_('元住所から町村を取得できない');
  }

  const town = townMatch[1];
  const expectedGoogleMunicipality = prefecture + town;

  if (
    hotelDbV2NormalizeText_(proposedMunicipality) !==
    hotelDbV2NormalizeText_(expectedGoogleMunicipality)
  ) {
    return hotelDbV2CountyMunicipalityNo_('Google市区町村が想定町村と違う');
  }

  const sourceFull = hotelDbV2CountyMunicipalityCanonicalAddress_(
    sourceMunicipality + sourceAddress
  );
  const proposedFull = hotelDbV2CountyMunicipalityCanonicalAddress_(
    prefecture + proposedAddress
  );

  if (!sourceFull || sourceFull !== proposedFull) {
    return hotelDbV2CountyMunicipalityNo_('結合住所が一致しない');
  }

  const proposedAddressNormalized = hotelDbV2NormalizeText_(proposedAddress);
  const expectedPrefix = hotelDbV2NormalizeText_(county + town);

  if (proposedAddressNormalized.indexOf(expectedPrefix) !== 0) {
    return hotelDbV2CountyMunicipalityNo_('Google住所に郡町村の並びがない');
  }

  const score = Number(data.matchScore || 0);
  const hasOazaDifference =
    /大字/u.test(sourceAddress) !== /大字/u.test(proposedAddress);
  const safeOazaLowScore =
    hasOazaDifference &&
    score >= HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE;

  if (
    score < HOTEL_DB_V2_COUNTY_TRIAGE_MIN_SCORE &&
    !safeOazaLowScore
  ) {
    return hotelDbV2CountyMunicipalityNo_('一致スコア不足');
  }

  return {
    equivalent: true,
    confidence: safeOazaLowScore ? 97 : 98,
    reason: safeOazaLowScore
      ? '郵便番号・施設名・結合住所が一致し、低い一致スコアの原因は郡と町村の列分けおよび「大字」の有無による表記差です。番地は一致しているため、元データ維持を推奨します。'
      : '郡と町村の列分けだけが異なります。元データは郡を市区町村列・町村を住所列に保持し、Googleは町村を市区町村列・郡を住所側に保持していますが、結合住所は同一です。元データ維持を推奨します。'
  };
}

function hotelDbV2CountyMunicipalityCanonicalAddress_(value) {
  return hotelDbV2NormalizeAddress_(value)
    .replace(/大字/g, '')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

function hotelDbV2CountyMunicipalityNo_(reason) {
  return {
    equivalent: false,
    confidence: 0,
    reason: reason || ''
  };
}

/**
 * 郡・町村の列分け差に関する自己診断テスト。
 */
function runHotelDbV2CountyMunicipalityTriageTests() {
  const base = {
    state: '未確認',
    sourcePostalCode: '681-0024',
    proposedPostalCode: '681-0024',
    sourceMunicipality: '鳥取県岩美郡',
    proposedMunicipality: '鳥取県岩美町',
    sourceAddress: '岩美町岩井536',
    proposedAddress: '岩美郡岩美町岩井536',
    sourceName: '明石家',
    proposedName: '明石家',
    placeId: 'test-place-id',
    matchScore: 93,
    businessStatus: '営業中',
    differences: '市区町村・住所'
  };

  const cases = [
    {
      name: '岩美郡と岩美町の列分け差は同一所在地',
      input: Object.assign({}, base),
      expected: true
    },
    {
      name: '八頭郡と八頭町の列分け差は同一所在地',
      input: Object.assign({}, base, {
        sourcePostalCode: '680-0611',
        proposedPostalCode: '680-0611',
        sourceMunicipality: '鳥取県八頭郡',
        proposedMunicipality: '鳥取県八頭町',
        sourceAddress: '八頭町富枝38',
        proposedAddress: '八頭郡八頭町富枝38',
        sourceName: '古民家 太田邸',
        proposedName: '古民家 太田邸'
      }),
      expected: true
    },
    {
      name: '東伯郡と三朝町の列分け差は同一所在地',
      input: Object.assign({}, base, {
        sourcePostalCode: '682-0123',
        proposedPostalCode: '682-0123',
        sourceMunicipality: '鳥取県東伯郡',
        proposedMunicipality: '鳥取県三朝町',
        sourceAddress: '三朝町三朝365-1',
        proposedAddress: '東伯郡三朝町三朝365-1',
        sourceName: '依山楼岩崎',
        proposedName: '依山楼岩崎'
      }),
      expected: true
    },
    {
      name: '大字の有無だけなら同一所在地',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-4401',
        proposedPostalCode: '689-4401',
        sourceMunicipality: '鳥取県日野郡',
        proposedMunicipality: '鳥取県江府町',
        sourceAddress: '江府町大字江尾2064',
        proposedAddress: '日野郡江府町江尾2064',
        sourceName: '門脇旅館',
        proposedName: '門脇旅館'
      }),
      expected: true
    },
    {
      name: '大字差ならスコア78でも安全条件を満たせば同一所在地',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-4401',
        proposedPostalCode: '689-4401',
        sourceMunicipality: '鳥取県日野郡',
        proposedMunicipality: '鳥取県江府町',
        sourceAddress: '江府町大字江尾2064',
        proposedAddress: '日野郡江府町江尾2064',
        sourceName: '門脇旅館',
        proposedName: '門脇旅館',
        matchScore: 78
      }),
      expected: true
    },
    {
      name: '大字差でも自動受付基準未満は同一扱いしない',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-4401',
        proposedPostalCode: '689-4401',
        sourceMunicipality: '鳥取県日野郡',
        proposedMunicipality: '鳥取県江府町',
        sourceAddress: '江府町大字江尾2064',
        proposedAddress: '日野郡江府町江尾2064',
        sourceName: '門脇旅館',
        proposedName: '門脇旅館',
        matchScore: 74
      }),
      expected: false
    },
    {
      name: '大字差があっても番地が違えば同一扱いしない',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-4401',
        proposedPostalCode: '689-4401',
        sourceMunicipality: '鳥取県日野郡',
        proposedMunicipality: '鳥取県江府町',
        sourceAddress: '江府町大字江尾2064',
        proposedAddress: '日野郡江府町江尾2065',
        sourceName: '門脇旅館',
        proposedName: '門脇旅館',
        matchScore: 78
      }),
      expected: false
    },
    {
      name: '番地が違えば同一扱いしない',
      input: Object.assign({}, base, {
        proposedAddress: '岩美郡岩美町岩井544'
      }),
      expected: false
    },
    {
      name: '町が違えば同一扱いしない',
      input: Object.assign({}, base, {
        proposedMunicipality: '鳥取県八頭町',
        proposedAddress: '岩美郡八頭町岩井536'
      }),
      expected: false
    },
    {
      name: '郵便番号が違えば同一扱いしない',
      input: Object.assign({}, base, {
        proposedPostalCode: '681-0003'
      }),
      expected: false
    },
    {
      name: '施設名が違えば同一扱いしない',
      input: Object.assign({}, base, {
        proposedName: '別の旅館'
      }),
      expected: false
    },
    {
      name: '大字差がないスコア90未満は同一扱いしない',
      input: Object.assign({}, base, {
        matchScore: 89
      }),
      expected: false
    },
    {
      name: '施設名差分も含む場合は同一扱いしない',
      input: Object.assign({}, base, {
        differences: '市区町村・住所・施設名'
      }),
      expected: false
    }
  ];

  const failures = [];

  cases.forEach(function(testCase, index) {
    const snapshot = JSON.stringify(testCase.input);
    const result = hotelDbV2CountyMunicipalityDecision_(testCase.input);
    const unchanged = JSON.stringify(testCase.input) === snapshot;

    if (result.equivalent !== testCase.expected || !unchanged) {
      failures.push(
        '例' + (index + 1) + '「' + testCase.name + '」: ' +
        '同一判定=' + result.equivalent +
        '（期待=' + testCase.expected + '）, ' +
        '入力保持=' + unchanged +
        (result.reason ? ', 理由=' + result.reason : '')
      );
    }
  });

  if (failures.length) {
    const message = [
      '郡・町村表記差テストに失敗しました。',
      '',
      failures.join('\n')
    ].join('\n');

    console.error(message);
    throw new Error(message);
  }

  const successMessage = [
    '郡・町村表記差テスト成功',
    '',
    '成功件数: ' + cases.length + '件',
    '失敗件数: 0件',
    '状態列の自動変更: なし'
  ].join('\n');

  console.log(successMessage);

  try {
    SpreadsheetApp.getUi().alert(successMessage);
  } catch (error) {
  }

  return {
    success: true,
    passed: cases.length,
    failed: 0
  };
}

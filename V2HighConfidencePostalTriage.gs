/**
 * PR #9: 郵便番号が異なるが、施設名と郡・町村・番地までが一致する
 * 高信頼ケースだけを「承認候補」へ再分類する。
 *
 * B列「状態」と元データは変更しない。
 */

const HOTEL_DB_V2_HIGH_CONFIDENCE_POSTAL_MIN_SCORE = 83;

function runHotelDbV2HighConfidencePostalTriage() {
  return withHotelDbV2Lock_('高信頼郵便番号修正候補の再判定', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '高信頼郵便番号修正候補を再判定',
      '「修正候補」のうち、郵便番号は違うものの、施設名と郡・町村・番地までが一致する高信頼ケースだけを承認候補へ再分類します。\n\n' +
      'B列「状態」は変更しません。\n' +
      '元データも変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2RefineHighConfidencePostalTriage_();

    ui.alert([
      '高信頼郵便番号修正候補の再判定完了',
      '',
      '確認件数: ' + result.scanned,
      '承認候補へ変更: ' + result.promoted,
      '変更なし: ' + result.unchanged,
      '',
      'B列「状態」は変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2RefineHighConfidencePostalTriage_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(
    HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS
  );

  if (!sheet || sheet.getLastRow() < 2) {
    return { scanned: 0, promoted: 0, unchanged: 0 };
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
    promoted: 0,
    unchanged: 0
  };

  values.forEach(function(row, offset) {
    const currentRecommendation = hotelDbV2Clean_(
      row[recommendationColumn - 1]
    );

    if (currentRecommendation !== '要人確認') {
      result.unchanged++;
      return;
    }

    const input = hotelDbV2TriageInputFromRow_(row, headerMap);
    const decision = hotelDbV2HighConfidencePostalDecision_(input);

    if (!decision.eligible) {
      result.unchanged++;
      return;
    }

    const rowNumber = offset + 2;
    sheet.getRange(rowNumber, recommendationColumn).setValue('承認候補');
    sheet.getRange(rowNumber, reasonColumn).setValue(decision.reason);
    sheet.getRange(rowNumber, confidenceColumn).setValue(decision.confidence);
    result.promoted++;
  });

  return result;
}

function hotelDbV2HighConfidencePostalDecision_(input) {
  const data = input || {};
  const differences = hotelDbV2Clean_(data.differences)
    .split('・')
    .map(hotelDbV2Clean_)
    .filter(Boolean)
    .sort();

  if (hotelDbV2Clean_(data.state) === '反映済み') {
    return hotelDbV2HighConfidencePostalNo_('反映済み');
  }

  if (!hotelDbV2Clean_(data.placeId)) {
    return hotelDbV2HighConfidencePostalNo_('Place IDなし');
  }

  if (hotelDbV2Clean_(data.businessStatus) !== '営業中') {
    return hotelDbV2HighConfidencePostalNo_('営業中以外');
  }

  if (
    Number(data.matchScore || 0) <
    HOTEL_DB_V2_HIGH_CONFIDENCE_POSTAL_MIN_SCORE
  ) {
    return hotelDbV2HighConfidencePostalNo_('一致スコア不足');
  }

  if (
    differences.length !== 3 ||
    differences[0] !== '住所' ||
    differences[1] !== '市区町村' ||
    differences[2] !== '郵便番号'
  ) {
    return hotelDbV2HighConfidencePostalNo_('対象差分ではない');
  }

  const sourcePostal = hotelDbV2NormalizePostalCode_(data.sourcePostalCode);
  const proposedPostal = hotelDbV2NormalizePostalCode_(data.proposedPostalCode);

  if (!sourcePostal || !proposedPostal) {
    return hotelDbV2HighConfidencePostalNo_('有効な7桁郵便番号ではない');
  }

  if (sourcePostal === proposedPostal) {
    return hotelDbV2HighConfidencePostalNo_('郵便番号が同一');
  }

  if (
    hotelDbV2NormalizeText_(data.sourceName) !==
    hotelDbV2NormalizeText_(data.proposedName)
  ) {
    return hotelDbV2HighConfidencePostalNo_('施設名不一致');
  }

  const sourceMunicipality = hotelDbV2Clean_(data.sourceMunicipality)
    .normalize('NFKC')
    .replace(/\s+/g, '');
  const proposedMunicipality = hotelDbV2Clean_(data.proposedMunicipality)
    .normalize('NFKC')
    .replace(/\s+/g, '');
  const sourceAddress = hotelDbV2Clean_(data.sourceAddress)
    .normalize('NFKC')
    .trim();
  const proposedAddress = hotelDbV2Clean_(data.proposedAddress)
    .normalize('NFKC')
    .trim();

  const countyMatch = sourceMunicipality.match(/^(.+?[都道府県])(.+?郡)$/u);
  if (!countyMatch) {
    return hotelDbV2HighConfidencePostalNo_('元市区町村が郡形式ではない');
  }

  const prefecture = countyMatch[1];
  const townMatch = sourceAddress.match(/^(.+?[町村])(.*)$/u);

  if (!townMatch) {
    return hotelDbV2HighConfidencePostalNo_('元住所から町村を取得できない');
  }

  const town = townMatch[1];
  const expectedGoogleMunicipality = prefecture + town;

  if (
    hotelDbV2NormalizeText_(proposedMunicipality) !==
    hotelDbV2NormalizeText_(expectedGoogleMunicipality)
  ) {
    return hotelDbV2HighConfidencePostalNo_('Google市区町村が想定町村と違う');
  }

  const sourceFull = hotelDbV2CountyMunicipalityCanonicalAddress_(
    sourceMunicipality + sourceAddress
  );
  const proposedFull = hotelDbV2CountyMunicipalityCanonicalAddress_(
    prefecture + proposedAddress
  );

  if (!sourceFull || sourceFull !== proposedFull) {
    return hotelDbV2HighConfidencePostalNo_('結合住所が完全一致しない');
  }

  return {
    eligible: true,
    confidence: 98,
    reason:
      '施設名が一致し、郡・町村・番地までを組み直した住所も完全一致しています。Googleの7桁郵便番号だけが元データと異なるため、高信頼の郵便番号修正候補として優先確認します。'
  };
}

function hotelDbV2HighConfidencePostalNo_(reason) {
  return {
    eligible: false,
    confidence: 0,
    reason: reason || ''
  };
}

/**
 * 高信頼郵便番号修正候補判定の自己診断テスト。
 * 入力データやB列「状態」は変更しない。
 */
function runHotelDbV2HighConfidencePostalTriageTests() {
  const base = {
    state: '未確認',
    sourcePostalCode: '680-0400',
    proposedPostalCode: '680-0611',
    sourceMunicipality: '鳥取県八頭郡',
    proposedMunicipality: '鳥取県八頭町',
    sourceAddress: '八頭町富枝453-11',
    proposedAddress: '八頭郡八頭町富枝453-11',
    sourceName: '谷口旅館',
    proposedName: '谷口旅館',
    placeId: 'test-place-id',
    matchScore: 83,
    businessStatus: '営業中',
    differences: '郵便番号・市区町村・住所'
  };

  const cases = [
    {
      name: '谷口旅館型は承認候補',
      input: Object.assign({}, base),
      expected: true
    },
    {
      name: 'アルパインヒュッテ型は承認候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '680-0700',
        proposedPostalCode: '680-0728',
        sourceMunicipality: '鳥取県八頭郡',
        proposedMunicipality: '鳥取県若桜町',
        sourceAddress: '若桜町舂米635-12',
        proposedAddress: '八頭郡若桜町舂米635-12',
        sourceName: 'アルパインヒュッテ',
        proposedName: 'アルパインヒュッテ'
      }),
      expected: true
    },
    {
      name: 'モリス荘型は承認候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '680-0700',
        proposedPostalCode: '680-0728',
        sourceMunicipality: '鳥取県八頭郡',
        proposedMunicipality: '鳥取県若桜町',
        sourceAddress: '若桜町舂米632',
        proposedAddress: '八頭郡若桜町舂米632',
        sourceName: 'モリス荘',
        proposedName: 'モリス荘'
      }),
      expected: true
    },
    {
      name: '斉木別館型は承認候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '682-0123',
        proposedPostalCode: '682-0122',
        sourceMunicipality: '鳥取県東伯郡',
        proposedMunicipality: '鳥取県三朝町',
        sourceAddress: '三朝町山田70',
        proposedAddress: '東伯郡三朝町山田70',
        sourceName: '大江戸温泉物語 Premium 斉木別館',
        proposedName: '大江戸温泉物語Premium 斉木別館'
      }),
      expected: true
    },
    {
      name: 'ホテル東伯イン型は承認候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '682-0123',
        proposedPostalCode: '689-2301',
        sourceMunicipality: '鳥取県東伯郡',
        proposedMunicipality: '鳥取県琴浦町',
        sourceAddress: '琴浦町八橋211',
        proposedAddress: '東伯郡琴浦町八橋211',
        sourceName: 'ホテル東伯イン',
        proposedName: 'ホテル東伯イン'
      }),
      expected: true
    },
    {
      name: '河上旅館型は承認候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '682-0123',
        proposedPostalCode: '689-2501',
        sourceMunicipality: '鳥取県東伯郡',
        proposedMunicipality: '鳥取県琴浦町',
        sourceAddress: '琴浦町赤碕1104-5',
        proposedAddress: '東伯郡琴浦町赤碕1104-5',
        sourceName: '河上旅館',
        proposedName: '河上旅館'
      }),
      expected: true
    },
    {
      name: '北条オートキャンプ場型は承認候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-2301',
        proposedPostalCode: '689-2103',
        sourceMunicipality: '鳥取県東伯郡',
        proposedMunicipality: '鳥取県北栄町',
        sourceAddress: '北栄町田井488-1',
        proposedAddress: '東伯郡北栄町田井488-1',
        sourceName: '北条オートキャンプ場',
        proposedName: '北条オートキャンプ場'
      }),
      expected: true
    },
    {
      name: 'Google住所に追加所在地表記があれば自動承認しない',
      input: Object.assign({}, base, {
        sourcePostalCode: '680-0700',
        proposedPostalCode: '680-0728',
        sourceMunicipality: '鳥取県八頭郡',
        proposedMunicipality: '鳥取県若桜町',
        sourceAddress: '若桜町舂米631-29',
        proposedAddress: '八頭郡若桜町舂米631-29 つくよね631-29',
        sourceName: 'ヒュッテ白樺',
        proposedName: 'ヒュッテ白樺'
      }),
      expected: false
    },
    {
      name: '郵便番号が同じなら対象外',
      input: Object.assign({}, base, {
        proposedPostalCode: '680-0400'
      }),
      expected: false
    },
    {
      name: '施設名が違えば自動承認しない',
      input: Object.assign({}, base, {
        proposedName: '別の旅館'
      }),
      expected: false
    },
    {
      name: '町村が違えば自動承認しない',
      input: Object.assign({}, base, {
        proposedMunicipality: '鳥取県若桜町',
        proposedAddress: '八頭郡若桜町富枝453-11'
      }),
      expected: false
    },
    {
      name: '一致スコア83未満は自動承認しない',
      input: Object.assign({}, base, {
        matchScore: 82
      }),
      expected: false
    },
    {
      name: '営業中以外は自動承認しない',
      input: Object.assign({}, base, {
        businessStatus: '閉業'
      }),
      expected: false
    },
    {
      name: 'Place IDなしは自動承認しない',
      input: Object.assign({}, base, {
        placeId: ''
      }),
      expected: false
    }
  ];

  const failures = [];
  let passed = 0;

  cases.forEach(function(testCase, index) {
    const snapshot = JSON.stringify(testCase.input);
    const decision = hotelDbV2HighConfidencePostalDecision_(testCase.input);
    const unchanged = JSON.stringify(testCase.input) === snapshot;

    if (decision.eligible === testCase.expected && unchanged) {
      passed++;
      return;
    }

    failures.push(
      '例' + (index + 1) + '「' + testCase.name + '」: ' +
      '承認候補=' + decision.eligible +
      '（期待=' + testCase.expected + '）, ' +
      '入力保持=' + unchanged +
      (decision.reason ? ', 理由=' + decision.reason : '')
    );
  });

  const message = [
    failures.length
      ? '高信頼郵便番号修正候補テスト 要確認'
      : '高信頼郵便番号修正候補テスト 成功',
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

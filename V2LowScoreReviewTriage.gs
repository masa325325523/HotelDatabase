/**
 * Ver2.0 低スコア要確認の安全な自動仕分け。
 * 「要確認」シートの状態・元データ・Place IDは変更せず、
 * 推奨判定・自動判定理由・信頼度だけを追加する。
 */

const HOTEL_DB_V2_LOW_SCORE_REVIEW_HEADERS = Object.freeze([
  '推奨判定',
  '自動判定理由',
  '信頼度'
]);

const HOTEL_DB_V2_LOW_SCORE_REASONS = Object.freeze([
  '一致スコア不足',
  '一致スコア要確認',
  '詳細取得後の一致スコア要確認'
]);

const HOTEL_DB_V2_LOW_SCORE_NAME_RISK_TOKENS = Object.freeze([
  '別館', '別邸', '新館', '本館', '離れ', '支店', '駅前',
  'annex', 'branch', 'tower', 'wing',
  'east', 'west', 'north', 'south',
  'cafe', 'カフェ', 'restaurant', 'レストラン'
]);

const HOTEL_DB_V2_LOW_SCORE_RULES = Object.freeze({
  STRONG_NAME_SIMILARITY: 0.55,
  WEAK_NAME_SIMILARITY: 0.25,
  SAME_ADDRESS_MIN_SCORE: 50,
  WRONG_CANDIDATE_MAX_SCORE: 45
});

function runHotelDbV2LowScoreReviewTriage() {
  return withHotelDbV2Lock_('低スコア要確認自動仕分け', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '低スコア要確認を自動仕分け',
      '「要確認」の低スコア行を、安全側に仕分けします。\n\n' +
      '本体の自動採用基準75点は変更しません。\n' +
      '状態・元データ・Place IDも変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2TriageLowScoreReviews_();

    ui.alert([
      '低スコア要確認の自動仕分け完了',
      '',
      '対象件数: ' + result.total,
      '同一施設有力: ' + result.sameLikely,
      '誤候補有力: ' + result.wrongLikely,
      '要人確認: ' + result.humanReview,
      '対象外: ' + result.outOfScope,
      '',
      '状態・元データ・Place IDは変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2TriageLowScoreReviews_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      total: 0,
      sameLikely: 0,
      wrongLikely: 0,
      humanReview: 0,
      outOfScope: 0
    };
  }

  let headerMap = hotelDbV2LowScoreEnsureHeaders_(sheet);
  hotelDbV2LowScoreValidateHeaders_(headerMap);

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet
    .getRange(2, 1, rowCount, sheet.getLastColumn())
    .getDisplayValues();

  const outputs = [];
  const result = {
    total: values.length,
    sameLikely: 0,
    wrongLikely: 0,
    humanReview: 0,
    outOfScope: 0
  };

  values.forEach(function(row) {
    const input = hotelDbV2LowScoreInputFromRow_(row, headerMap);
    const decision = hotelDbV2ClassifyLowScoreReview_(input);
    outputs.push(decision);

    if (decision.recommendation === '同一施設有力') result.sameLikely++;
    else if (decision.recommendation === '誤候補有力') result.wrongLikely++;
    else if (decision.recommendation === '要人確認') result.humanReview++;
    else result.outOfScope++;
  });

  HOTEL_DB_V2_LOW_SCORE_REVIEW_HEADERS.forEach(function(header) {
    const column = headerMap[header];
    const columnValues = outputs.map(function(decision) {
      if (header === '推奨判定') return [decision.recommendation];
      if (header === '自動判定理由') return [decision.reason];
      return [decision.confidence];
    });
    sheet.getRange(2, column, rowCount, 1).setValues(columnValues);
  });

  return result;
}

function hotelDbV2LowScoreEnsureHeaders_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues()[0];
  const normalized = headers.map(hotelDbV2NormalizeText_);
  const missing = HOTEL_DB_V2_LOW_SCORE_REVIEW_HEADERS.filter(function(header) {
    return normalized.indexOf(hotelDbV2NormalizeText_(header)) === -1;
  });

  if (missing.length) {
    sheet
      .getRange(1, sheet.getLastColumn() + 1, 1, missing.length)
      .setValues([missing]);
  }

  return hotelDbV2LowScoreHeaderMap_(sheet);
}

function hotelDbV2LowScoreHeaderMap_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues()[0];
  const map = {};

  headers.forEach(function(header, index) {
    const text = hotelDbV2Clean_(header);
    if (text) map[text] = index + 1;
  });

  return map;
}

function hotelDbV2LowScoreValidateHeaders_(map) {
  const required = [
    '状態', '市区町村', '住所', '施設名', '理由',
    '候補施設名', '候補住所', '候補Place ID',
    '一致スコア', '営業状態', '詳細',
    '推奨判定', '自動判定理由', '信頼度'
  ];

  const missing = required.filter(function(header) {
    return !map[header];
  });

  if (missing.length) {
    throw new Error(
      '要確認シートの見出しが不足しています: ' + missing.join(', ')
    );
  }
}

function hotelDbV2LowScoreInputFromRow_(row, map) {
  function value(header) {
    return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : '';
  }

  return {
    state: value('状態'),
    municipality: value('市区町村'),
    sourceAddress: value('住所'),
    sourceName: value('施設名'),
    reason: value('理由'),
    candidateName: value('候補施設名'),
    candidateAddress: value('候補住所'),
    candidatePlaceId: value('候補Place ID'),
    matchScore: Number(value('一致スコア') || 0),
    businessStatus: value('営業状態'),
    detail: value('詳細')
  };
}

function hotelDbV2ClassifyLowScoreReview_(input) {
  const data = input || {};
  const state = hotelDbV2Clean_(data.state);
  const reason = hotelDbV2Clean_(data.reason);
  const score = Number(data.matchScore || 0);

  if (state && state !== '未確認') {
    return hotelDbV2LowScoreDecision_(
      '対象外',
      '状態が「未確認」ではないため、自動仕分け対象外です。',
      100
    );
  }

  if (HOTEL_DB_V2_LOW_SCORE_REASONS.indexOf(reason) === -1) {
    return hotelDbV2LowScoreDecision_(
      '対象外',
      '低スコア理由の行ではないため、この仕分けの対象外です。',
      100
    );
  }

  if (score >= HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE) {
    return hotelDbV2LowScoreDecision_(
      '対象外',
      '一致スコアが本体の自動採用基準以上のため、低スコア仕分け対象外です。',
      100
    );
  }

  if (!hotelDbV2Clean_(data.candidatePlaceId)) {
    return hotelDbV2LowScoreDecision_(
      '要人確認',
      '候補Place IDがないため、候補施設の同一性を自動判定しません。',
      100
    );
  }

  if (hotelDbV2Clean_(data.businessStatus) !== '営業中') {
    return hotelDbV2LowScoreDecision_(
      '要人確認',
      '候補のGoogle営業状態が「営業中」ではないため、人が確認します。',
      100
    );
  }

  if (
    !hotelDbV2Clean_(data.sourceAddress) ||
    !hotelDbV2Clean_(data.candidateAddress) ||
    !hotelDbV2Clean_(data.sourceName) ||
    !hotelDbV2Clean_(data.candidateName)
  ) {
    return hotelDbV2LowScoreDecision_(
      '要人確認',
      '施設名または住所が不足しているため、自動判定しません。',
      100
    );
  }

  const sourceName = hotelDbV2NormalizeText_(data.sourceName);
  const candidateName = hotelDbV2NormalizeText_(data.candidateName);
  const nameSimilarity = hotelDbV2SimilarityRatio_(sourceName, candidateName);
  const nameContains = hotelDbV2LowScoreNameContains_(sourceName, candidateName);
  const nameStrong =
    sourceName === candidateName ||
    nameContains ||
    nameSimilarity >= HOTEL_DB_V2_LOW_SCORE_RULES.STRONG_NAME_SIMILARITY;

  const municipalitySame = hotelDbV2LowScoreMunicipalityMatches_(
    data.municipality,
    data.candidateAddress
  );
  const addressEquivalent = hotelDbV2LowScoreAddressesEquivalent_(
    data.sourceAddress,
    data.candidateAddress,
    data.sourceName,
    data.candidateName
  );
  const addressContains = hotelDbV2LowScoreAddressContains_(
    data.sourceAddress,
    data.candidateAddress
  );
  const strongLocation = addressEquivalent || addressContains;

  if (
    hotelDbV2LowScoreHasSubpremiseRisk_(
      data.sourceAddress,
      data.candidateAddress
    )
  ) {
    return hotelDbV2LowScoreDecision_(
      '要人確認',
      '階・部屋番号などの差である可能性があるため、同一施設とは自動判定しません。',
      98
    );
  }

  if (hotelDbV2LowScoreHasNameRisk_(data.sourceName, data.candidateName)) {
    return hotelDbV2LowScoreDecision_(
      '要人確認',
      '別館・別邸・新館・支店・飲食店表記などの構造差があるため、人が確認します。',
      98
    );
  }

  if (strongLocation && municipalitySame && nameStrong) {
    return hotelDbV2LowScoreDecision_(
      '同一施設有力',
      '住所と市区町村が一致し、施設名も同一・包含・高類似です。75点基準は下げず、同一施設の有力候補として確認します。',
      Math.max(92, Math.min(98, Math.round(90 + nameSimilarity * 8)))
    );
  }

  if (
    strongLocation &&
    municipalitySame &&
    score >= HOTEL_DB_V2_LOW_SCORE_RULES.SAME_ADDRESS_MIN_SCORE
  ) {
    return hotelDbV2LowScoreDecision_(
      '同一施設有力',
      '住所と市区町村が一致しています。施設名の表記差が大きいため自動採用はせず、同一施設の有力候補として人が確認します。',
      86
    );
  }

  if (
    !municipalitySame &&
    !nameStrong &&
    score <= 55
  ) {
    return hotelDbV2LowScoreDecision_(
      '誤候補有力',
      '候補住所が元の市区町村と一致せず、施設名の類似も弱いため、別施設候補の可能性が高いです。',
      96
    );
  }

  if (
    !strongLocation &&
    nameSimilarity < HOTEL_DB_V2_LOW_SCORE_RULES.WEAK_NAME_SIMILARITY &&
    score <= HOTEL_DB_V2_LOW_SCORE_RULES.WRONG_CANDIDATE_MAX_SCORE
  ) {
    return hotelDbV2LowScoreDecision_(
      '誤候補有力',
      '住所一致が確認できず、施設名の類似も弱く、一致スコアも低いため、別施設候補の可能性が高いです。',
      Math.max(88, Math.min(95, 100 - Math.round(score / 4)))
    );
  }

  return hotelDbV2LowScoreDecision_(
    '要人確認',
    '同一施設または誤候補と安全に断定できる条件が不足しているため、人が確認します。',
    Math.max(80, Math.min(95, Math.round(score + 20)))
  );
}

function hotelDbV2LowScoreDecision_(recommendation, reason, confidence) {
  return {
    recommendation: recommendation,
    reason: reason,
    confidence: Math.max(0, Math.min(100, Math.round(Number(confidence) || 0)))
  };
}

function hotelDbV2LowScoreNameContains_(left, right) {
  const a = hotelDbV2Clean_(left);
  const b = hotelDbV2Clean_(right);
  if (!a || !b) return false;
  const shorter = a.length <= b.length ? a : b;
  const longer = a.length <= b.length ? b : a;
  return shorter.length >= 3 && longer.indexOf(shorter) !== -1;
}

function hotelDbV2LowScoreMunicipalityMatches_(municipality, candidateAddress) {
  const city = hotelDbV2NormalizeText_(municipality);
  const address = hotelDbV2NormalizeText_(candidateAddress);
  if (!city || !address) return false;
  return address.indexOf(city) !== -1;
}

function hotelDbV2LowScoreAddressesEquivalent_(
  sourceAddress,
  candidateAddress,
  sourceName,
  candidateName
) {
  if (typeof hotelDbV2AddressesEquivalent_ === 'function') {
    return hotelDbV2AddressesEquivalent_(
      sourceAddress,
      candidateAddress,
      sourceName,
      candidateName
    );
  }

  return hotelDbV2NormalizeAddress_(sourceAddress) ===
    hotelDbV2NormalizeAddress_(candidateAddress);
}

function hotelDbV2LowScoreAddressContains_(sourceAddress, candidateAddress) {
  const left = hotelDbV2NormalizeAddress_(sourceAddress);
  const right = hotelDbV2NormalizeAddress_(candidateAddress);
  if (!left || !right) return false;
  const shorter = left.length <= right.length ? left : right;
  const longer = left.length <= right.length ? right : left;
  return shorter.length >= 8 && longer.indexOf(shorter) !== -1;
}

function hotelDbV2LowScoreHasSubpremiseRisk_(sourceAddress, candidateAddress) {
  const source = hotelDbV2ConvertAddressKanjiNumbers_(
    hotelDbV2Clean_(sourceAddress).normalize('NFKC').toLowerCase()
  );
  const candidate = hotelDbV2ConvertAddressKanjiNumbers_(
    hotelDbV2Clean_(candidateAddress).normalize('NFKC').toLowerCase()
  );

  if (hotelDbV2NormalizeAddress_(source) === hotelDbV2NormalizeAddress_(candidate)) {
    return false;
  }

  const subpremisePattern = /(?:\d+\s*(?:階|f|号室|室)|(?:room|floor)\s*\d+)/i;
  return subpremisePattern.test(source) || subpremisePattern.test(candidate);
}

function hotelDbV2LowScoreHasNameRisk_(sourceName, candidateName) {
  const source = hotelDbV2NormalizeText_(sourceName);
  const candidate = hotelDbV2NormalizeText_(candidateName);

  return HOTEL_DB_V2_LOW_SCORE_NAME_RISK_TOKENS.some(function(token) {
    const normalizedToken = hotelDbV2NormalizeText_(token);
    const sourceHas = source.indexOf(normalizedToken) !== -1;
    const candidateHas = candidate.indexOf(normalizedToken) !== -1;
    return sourceHas !== candidateHas;
  });
}

function runHotelDbV2LowScoreReviewTriageTests() {
  const base = {
    state: '未確認',
    municipality: '大阪市北区',
    sourceAddress: '大阪府大阪市北区芝田2-4-32',
    sourceName: 'サンプルホテル大阪',
    reason: '一致スコア不足',
    candidateName: 'サンプルホテル大阪',
    candidateAddress: '大阪府大阪市北区芝田2丁目4-32',
    candidatePlaceId: 'TEST_PLACE',
    matchScore: 60,
    businessStatus: '営業中',
    detail: '自動採用基準=75'
  };

  function make(overrides) {
    const item = {};
    Object.keys(base).forEach(function(key) { item[key] = base[key]; });
    Object.keys(overrides || {}).forEach(function(key) { item[key] = overrides[key]; });
    return item;
  }

  const cases = [
    { name: '同住所・同名', input: make({}), expected: '同一施設有力' },
    {
      name: '日本語名と英字名でも住所一致',
      input: make({
        sourceName: 'ホープツリー天王寺',
        candidateName: 'HOPETREE 天王寺',
        municipality: '大阪市阿倍野区',
        sourceAddress: '大阪府大阪市阿倍野区松崎町1丁目1-7',
        candidateAddress: '大阪府大阪市阿倍野区松崎町1-1-7',
        matchScore: 57
      }),
      expected: '同一施設有力'
    },
    {
      name: '施設名包含',
      input: make({ candidateName: 'サンプルホテル大阪 by GRANVIA', matchScore: 63 }),
      expected: '同一施設有力'
    },
    {
      name: '別館差分は人確認',
      input: make({ candidateName: 'サンプルホテル大阪 別館', matchScore: 65 }),
      expected: '要人確認'
    },
    {
      name: 'Cafe差分は人確認',
      input: make({ sourceName: 'Drop Inn Tottori', candidateName: 'Cafe Drop Inn Tottori', matchScore: 65 }),
      expected: '要人確認'
    },
    {
      name: '別部屋は人確認',
      input: make({
        municipality: '名古屋市千種区',
        sourceAddress: '愛知県名古屋市千種区春岡一丁目4番21号 Fuchsia901',
        candidateAddress: '愛知県名古屋市千種区春岡1-4-21 Fuchsia301',
        sourceName: 'Japan Hinata',
        candidateName: 'Japan Hinata',
        matchScore: 60
      }),
      expected: '要人確認'
    },
    {
      name: '別自治体・別名は誤候補有力',
      input: make({
        candidateName: '京都サンプル旅館',
        candidateAddress: '京都府京都市下京区東塩小路町1',
        matchScore: 30
      }),
      expected: '誤候補有力'
    },
    {
      name: '同自治体でも低スコア別住所別名',
      input: make({
        candidateName: '全く別の宿',
        candidateAddress: '大阪府大阪市北区梅田3-1-1',
        matchScore: 25
      }),
      expected: '誤候補有力'
    },
    {
      name: '44点・弱い施設名・別住所',
      input: make({
        candidateName: 'XYZ GUEST HOUSE',
        candidateAddress: '大阪府大阪市北区堂島1-1-1',
        matchScore: 44
      }),
      expected: '誤候補有力'
    },
    {
      name: '同名だが別住所は人確認',
      input: make({
        candidateAddress: '大阪府大阪市北区梅田1-1-1',
        matchScore: 63
      }),
      expected: '要人確認'
    },
    {
      name: '同住所でも40点・無関係名は人確認',
      input: make({ candidateName: '別会社ホテル', matchScore: 40 }),
      expected: '要人確認'
    },
    {
      name: '74点でも低スコア対象',
      input: make({ matchScore: 74 }),
      expected: '同一施設有力'
    },
    {
      name: '75点以上は対象外',
      input: make({ matchScore: 75 }),
      expected: '対象外'
    },
    {
      name: '閉業理由は対象外',
      input: make({ reason: '閉業', matchScore: 60 }),
      expected: '対象外'
    },
    {
      name: '確認済み状態は対象外',
      input: make({ state: '確認済み' }),
      expected: '対象外'
    },
    {
      name: '営業中でない候補は人確認',
      input: make({ businessStatus: '閉業' }),
      expected: '要人確認'
    },
    {
      name: '候補Place IDなしは人確認',
      input: make({ candidatePlaceId: '' }),
      expected: '要人確認'
    },
    {
      name: '元住所なしは人確認',
      input: make({ sourceAddress: '' }),
      expected: '要人確認'
    },
    {
      name: '候補住所なしは人確認',
      input: make({ candidateAddress: '' }),
      expected: '要人確認'
    },
    {
      name: 'ANNEX差分は人確認',
      input: make({ candidateName: 'Sample Hotel Osaka ANNEX', matchScore: 65 }),
      expected: '要人確認'
    }
  ];

  const failures = [];

  cases.forEach(function(testCase, index) {
    const actual = hotelDbV2ClassifyLowScoreReview_(testCase.input);
    if (actual.recommendation !== testCase.expected) {
      failures.push(
        '例' + (index + 1) + '「' + testCase.name + '」: ' +
        '実際=' + actual.recommendation + ', 期待=' + testCase.expected +
        ', 理由=' + actual.reason
      );
    }
  });

  if (failures.length) {
    throw new Error(
      '低スコア要確認安全仕分けテスト失敗\n\n' + failures.join('\n')
    );
  }

  SpreadsheetApp.getUi().alert([
    '低スコア要確認安全仕分けテスト 成功',
    '',
    '成功件数: ' + cases.length + '件',
    '失敗件数: 0件',
    '状態列の自動変更: なし',
    '元データ・Place IDの自動変更: なし',
    '本体75点基準の変更: なし'
  ].join('\n'));

  return {
    success: cases.length,
    failure: 0,
    stateAutoChange: false,
    sourceAutoChange: false,
    autoAcceptThresholdChanged: false
  };
}
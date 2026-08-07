/**
 * Ver2.0 修正候補の自動仕分け。
 * B列「状態」と元データは変更せず、推奨判定・理由・信頼度だけを追加する。
 */

const HOTEL_DB_V2_TRIAGE_HEADERS = Object.freeze([
  '推奨判定',
  '自動判定理由',
  '信頼度'
]);

const HOTEL_DB_V2_TRIAGE_GENERIC_NAME_WORDS = Object.freeze([
  '寿司', '料理', '割烹', '温泉', '宿泊', '旅館', 'ホテル',
  '民宿', 'ペンション', 'ゲストハウス', 'ホステル',
  'ロッジ', 'イン', 'お宿', '宿'
]);

function runHotelDbV2TriageCorrections() {
  return withHotelDbV2Lock_('修正候補自動仕分け', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '修正候補を自動仕分け',
      '「修正候補」に推奨判定・理由・信頼度を付けます。\n\n' +
      'B列「状態」は変更しません。\n' +
      '元データも変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2TriageCorrections_();

    ui.alert([
      '修正候補の自動仕分け完了',
      '',
      '対象件数: ' + result.total,
      '承認候補: ' + result.approveCandidate,
      '却下候補: ' + result.rejectCandidate,
      '要人確認: ' + result.humanReview,
      '対象外: ' + result.outOfScope,
      '',
      'B列「状態」は変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2TriageCorrections_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(
    HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS
  );

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      total: 0,
      approveCandidate: 0,
      rejectCandidate: 0,
      humanReview: 0,
      outOfScope: 0
    };
  }

  const headerMap = hotelDbV2TriageEnsureHeaders_(sheet);
  hotelDbV2TriageValidateHeaders_(headerMap);

  const values = sheet
    .getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .getDisplayValues();

  const outputs = [];
  const result = {
    total: values.length,
    approveCandidate: 0,
    rejectCandidate: 0,
    humanReview: 0,
    outOfScope: 0
  };

  values.forEach(function(row) {
    const input = hotelDbV2TriageInputFromRow_(row, headerMap);
    const decision = hotelDbV2ClassifyCorrectionCandidate_(input);

    outputs.push([
      decision.recommendation,
      decision.reason,
      decision.confidence
    ]);

    if (decision.recommendation === '承認候補') result.approveCandidate++;
    else if (decision.recommendation === '却下候補') result.rejectCandidate++;
    else if (decision.recommendation === '要人確認') result.humanReview++;
    else result.outOfScope++;
  });

  sheet
    .getRange(2, headerMap['推奨判定'], outputs.length, 3)
    .setValues(outputs);

  return result;
}

function hotelDbV2TriageEnsureHeaders_(sheet) {
  const headers = sheet
    .getRange(1, 1, 1, Math.max(1, sheet.getLastColumn()))
    .getDisplayValues()[0];
  const normalized = headers.map(hotelDbV2NormalizeText_);
  const missing = HOTEL_DB_V2_TRIAGE_HEADERS.filter(function(header) {
    return normalized.indexOf(hotelDbV2NormalizeText_(header)) === -1;
  });

  if (missing.length) {
    sheet
      .getRange(1, sheet.getLastColumn() + 1, 1, missing.length)
      .setValues([missing]);
  }

  return hotelDbV2TriageHeaderMap_(sheet);
}

function hotelDbV2TriageHeaderMap_(sheet) {
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

function hotelDbV2TriageValidateHeaders_(map) {
  const required = [
    '状態', '元郵便番号', '修正郵便番号',
    '元市区町村', '修正市区町村',
    '元住所', '修正住所',
    '元施設名', '修正施設名',
    'Place ID', '一致スコア', '営業状態', '差分',
    '推奨判定', '自動判定理由', '信頼度'
  ];

  const missing = required.filter(function(header) {
    return !map[header];
  });

  if (missing.length) {
    throw new Error(
      '修正候補シートの見出しが不足しています: ' + missing.join(', ')
    );
  }
}

function hotelDbV2TriageInputFromRow_(row, map) {
  function value(header) {
    return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : '';
  }

  return {
    state: value('状態'),
    sourcePostalCode: value('元郵便番号'),
    proposedPostalCode: value('修正郵便番号'),
    sourceMunicipality: value('元市区町村'),
    proposedMunicipality: value('修正市区町村'),
    sourceAddress: value('元住所'),
    proposedAddress: value('修正住所'),
    sourceName: value('元施設名'),
    proposedName: value('修正施設名'),
    placeId: value('Place ID'),
    matchScore: Number(value('一致スコア') || 0),
    businessStatus: value('営業状態'),
    differences: value('差分')
  };
}

function hotelDbV2ClassifyCorrectionCandidate_(input) {
  const data = input || {};
  const state = hotelDbV2Clean_(data.state);
  const score = Number(data.matchScore || 0);
  const differences = hotelDbV2Clean_(data.differences)
    .split('・')
    .map(hotelDbV2Clean_)
    .filter(Boolean);

  if (state === '反映済み') {
    return hotelDbV2TriageDecision_(
      '対象外',
      'すでに元データへ反映済みです。',
      100
    );
  }

  if (!hotelDbV2Clean_(data.placeId)) {
    return hotelDbV2TriageDecision_(
      '要人確認',
      'Place IDがないため自動判定しません。',
      100
    );
  }

  if (hotelDbV2Clean_(data.businessStatus) !== '営業中') {
    return hotelDbV2TriageDecision_(
      '要人確認',
      'Google営業状態が「営業中」ではありません。',
      100
    );
  }

  if (!score || score < 90) {
    return hotelDbV2TriageDecision_(
      '要人確認',
      '一致スコアが90点未満のため、人が確認します。',
      Math.max(60, Math.min(89, Math.round(score || 60)))
    );
  }

  if (!differences.length) {
    return hotelDbV2TriageDecision_(
      '対象外',
      '差分がないため修正対象ではありません。',
      100
    );
  }

  if (differences.length > 1) {
    return hotelDbV2TriageDecision_(
      '要人確認',
      '複数項目に差分があるため自動承認しません。',
      95
    );
  }

  const difference = differences[0];
  const postalSame =
    hotelDbV2NormalizePostalCode_(data.sourcePostalCode) ===
    hotelDbV2NormalizePostalCode_(data.proposedPostalCode);
  const municipalitySame =
    hotelDbV2NormalizeText_(data.sourceMunicipality) ===
    hotelDbV2NormalizeText_(data.proposedMunicipality);
  const addressSame = hotelDbV2AddressesEquivalent_(
    data.sourceAddress,
    data.proposedAddress,
    data.sourceName,
    data.proposedName
  );
  const nameSame =
    hotelDbV2NormalizeText_(data.sourceName) ===
    hotelDbV2NormalizeText_(data.proposedName);

  if (difference === '施設名') {
    if (
      postalSame &&
      municipalitySame &&
      addressSame &&
      hotelDbV2TriageNamesEquivalentByNoise_(
        data.sourceName,
        data.proposedName,
        data.sourceMunicipality
      )
    ) {
      return hotelDbV2TriageDecision_(
        '却下候補',
        '住所・郵便番号・市区町村が一致し、施設名の差は地域名や業態語などの付加表記と判定しました。',
        Math.min(98, Math.max(90, Math.round(score + 3)))
      );
    }

    return hotelDbV2TriageDecision_(
      '要人確認',
      '施設名差分が単純な付加表記だけとは断定できません。',
      90
    );
  }

  if (difference === '住所') {
    if (
      hotelDbV2TriageHasSubpremiseRisk_(
        data.sourceAddress,
        data.proposedAddress
      )
    ) {
      return hotelDbV2TriageDecision_(
        '要人確認',
        '建物名・階・部屋番号などの差である可能性があるため、人が確認します。',
        98
      );
    }

    if (postalSame && municipalitySame && nameSame) {
      return hotelDbV2TriageDecision_(
        '承認候補',
        '施設名・郵便番号・市区町村が一致し、差分が住所だけです。元データ更新候補として優先確認します。',
        Math.min(96, Math.max(90, Math.round(score)))
      );
    }

    return hotelDbV2TriageDecision_(
      '要人確認',
      '住所以外の一致条件が不足しているため自動承認候補にしません。',
      92
    );
  }

  if (difference === '郵便番号') {
    if (municipalitySame && addressSame && nameSame && score >= 95) {
      return hotelDbV2TriageDecision_(
        '承認候補',
        '施設名・市区町村・住所が一致し、差分が郵便番号だけです。',
        95
      );
    }

    return hotelDbV2TriageDecision_(
      '要人確認',
      '郵便番号差分は誤修正防止のため人が確認します。',
      92
    );
  }

  return hotelDbV2TriageDecision_(
    '要人確認',
    '市区町村など重要項目の差分は人が確認します。',
    95
  );
}

function hotelDbV2TriageDecision_(recommendation, reason, confidence) {
  return {
    recommendation: recommendation,
    reason: reason,
    confidence: Math.max(
      0,
      Math.min(100, Math.round(Number(confidence) || 0))
    )
  };
}

function hotelDbV2TriageNamesEquivalentByNoise_(
  sourceName,
  proposedName,
  municipality
) {
  const sourceCore = hotelDbV2TriageNameCore_(sourceName, municipality);
  const proposedCore = hotelDbV2TriageNameCore_(proposedName, municipality);

  return Boolean(
    sourceCore &&
    proposedCore &&
    sourceCore === proposedCore
  );
}

function hotelDbV2TriageNameCore_(name, municipality) {
  let normalized = hotelDbV2NormalizeText_(name);
  if (!normalized) return '';

  const words = HOTEL_DB_V2_TRIAGE_GENERIC_NAME_WORDS.slice();

  if (typeof HOTEL_DB_V2_FACILITY_TYPES !== 'undefined') {
    HOTEL_DB_V2_FACILITY_TYPES.forEach(function(word) {
      words.push(word);
    });
  }

  hotelDbV2TriageLocationWords_(municipality).forEach(function(word) {
    words.push(word);
  });

  hotelDbV2Unique_(words)
    .map(hotelDbV2NormalizeText_)
    .filter(Boolean)
    .sort(function(a, b) {
      return b.length - a.length;
    })
    .forEach(function(word) {
      normalized = normalized.split(word).join('');
    });

  return normalized;
}

function hotelDbV2TriageLocationWords_(municipality) {
  const text = hotelDbV2Clean_(municipality).normalize('NFKC');
  if (!text) return [];

  const words = [text];
  const prefectureMatch = text.match(/^(.+?[都道府県])/u);

  if (prefectureMatch) {
    words.push(prefectureMatch[1]);
    words.push(prefectureMatch[1].replace(/[都道府県]$/u, ''));
  }

  const municipalityMatch = text.match(/([^都道府県]*?[市区町村])$/u);

  if (municipalityMatch) {
    words.push(municipalityMatch[1]);
    words.push(municipalityMatch[1].replace(/[市区町村]$/u, ''));
  }

  return hotelDbV2Unique_(words);
}

function hotelDbV2TriageHasSubpremiseRisk_(sourceAddress, proposedAddress) {
  const combined = [
    hotelDbV2Clean_(sourceAddress),
    hotelDbV2Clean_(proposedAddress)
  ].join(' ').normalize('NFKC');

  if (
    /(号室|客室|部屋|室|階|フロア|floor|room|unit)/iu.test(combined)
  ) {
    return true;
  }

  return /[a-z][a-z0-9._ -]*\d{2,}\b/iu.test(combined);
}

/**
 * 修正候補自動仕分けの自己診断テスト。
 */
function runHotelDbV2CorrectionTriageTests() {
  const base = {
    state: '未確認',
    sourcePostalCode: '680-1442',
    proposedPostalCode: '680-1442',
    sourceMunicipality: '鳥取県鳥取市',
    proposedMunicipality: '鳥取県鳥取市',
    sourceAddress: '吉岡温泉町271',
    proposedAddress: '吉岡温泉町771',
    sourceName: '藤田旅館',
    proposedName: '藤田旅館',
    placeId: 'test-place-id',
    matchScore: 91,
    businessStatus: '営業中',
    differences: '住所'
  };

  const cases = [
    {
      name: '地域名付加だけなら却下候補',
      input: Object.assign({}, base, {
        sourceAddress: '吉岡温泉町765',
        proposedAddress: '吉岡温泉町765',
        sourceName: '北川旅館',
        proposedName: '鳥取 北川旅館',
        matchScore: 93,
        differences: '施設名'
      }),
      expected: '却下候補'
    },
    {
      name: '業態語付加だけなら却下候補',
      input: Object.assign({}, base, {
        sourcePostalCode: '680-0831',
        proposedPostalCode: '680-0831',
        sourceAddress: '栄町230',
        proposedAddress: '栄町230',
        sourceName: '常天',
        proposedName: '寿司 旅館 常天',
        matchScore: 93,
        differences: '施設名'
      }),
      expected: '却下候補'
    },
    {
      name: '高スコア・名称一致・住所だけ差分なら承認候補',
      input: Object.assign({}, base),
      expected: '承認候補'
    },
    {
      name: '90点未満は要人確認',
      input: Object.assign({}, base, {
        sourceAddress: '吉岡温泉町268',
        proposedAddress: '吉岡温泉町772',
        sourceName: 'たから屋旅館',
        proposedName: 'たから屋旅館',
        matchScore: 84
      }),
      expected: '要人確認'
    },
    {
      name: '部屋番号差は要人確認',
      input: Object.assign({}, base, {
        sourceAddress: '春岡一丁目4番21号 Fuchsia901',
        proposedAddress: '春岡1-4-21 Fuchsia301',
        sourceName: 'Japan Hinata 1',
        proposedName: 'Japan Hinata 1',
        matchScore: 98
      }),
      expected: '要人確認'
    },
    {
      name: '営業中以外は要人確認',
      input: Object.assign({}, base, {
        businessStatus: '閉業',
        matchScore: 99
      }),
      expected: '要人確認'
    },
    {
      name: '複数差分は要人確認',
      input: Object.assign({}, base, {
        proposedName: '別名旅館',
        differences: '住所・施設名',
        matchScore: 99
      }),
      expected: '要人確認'
    },
    {
      name: 'Place IDなしは要人確認',
      input: Object.assign({}, base, {
        placeId: '',
        matchScore: 99
      }),
      expected: '要人確認'
    }
  ];

  const failures = [];

  cases.forEach(function(testCase, index) {
    const snapshot = JSON.stringify(testCase.input);
    const result = hotelDbV2ClassifyCorrectionCandidate_(testCase.input);
    const unchanged = JSON.stringify(testCase.input) === snapshot;

    if (
      result.recommendation !== testCase.expected ||
      !unchanged
    ) {
      failures.push(
        '例' + (index + 1) + '「' + testCase.name + '」: ' +
        '判定=' + result.recommendation +
        '（期待=' + testCase.expected + '）, ' +
        '入力保持=' + unchanged
      );
    }
  });

  if (failures.length) {
    const message = [
      '修正候補自動仕分けテストに失敗しました。',
      '',
      failures.join('\n')
    ].join('\n');

    console.error(message);
    throw new Error(message);
  }

  const successMessage = [
    '修正候補自動仕分けテスト成功',
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

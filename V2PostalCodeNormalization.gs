/**
 * Ver2.0 郵便番号表記差の安全な正規化・再判定。
 *
 * 例:
 *   680-090７ / ６８０－０９０７ / 〒6800907
 * を比較時には 680-0907 として扱う。
 *
 * このモジュールは元データとB列「状態」を変更しない。
 */

function hotelDbV2NormalizePostalCodeNfkc_(value) {
  const text = hotelDbV2Clean_(value).normalize('NFKC');
  const digits = text.replace(/\D/g, '');

  return digits.length === 7
    ? digits.slice(0, 3) + '-' + digits.slice(3)
    : '';
}

function runHotelDbV2PostalCodeNormalizationRefinement() {
  return withHotelDbV2Lock_('郵便番号表記差の再判定', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '郵便番号表記差を再判定',
      '「修正候補」のうち、全角・半角・ハイフン・空白の違いだけの郵便番号差分を再判定します。\n\n' +
      'B列「状態」は変更しません。\n' +
      '元データも変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2RefinePostalCodeTriage_();

    ui.alert([
      '郵便番号表記差の再判定完了',
      '',
      '確認件数: ' + result.scanned,
      '郵便番号表記差を解消: ' + result.postalFormattingResolved,
      '却下候補へ変更: ' + result.rejected,
      '承認候補へ変更: ' + result.approved,
      '要人確認のまま: ' + result.stillHumanReview,
      '変更なし: ' + result.unchanged,
      '',
      'B列「状態」は変更していません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2RefinePostalCodeTriage_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(
    HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS
  );

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      scanned: 0,
      postalFormattingResolved: 0,
      rejected: 0,
      approved: 0,
      stillHumanReview: 0,
      unchanged: 0
    };
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
    postalFormattingResolved: 0,
    rejected: 0,
    approved: 0,
    stillHumanReview: 0,
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
    const decision = hotelDbV2PostalCodeRefinementDecision_(input);

    if (!decision.postalFormattingResolved) {
      result.unchanged++;
      return;
    }

    result.postalFormattingResolved++;

    const rowNumber = offset + 2;
    sheet
      .getRange(rowNumber, recommendationColumn)
      .setValue(decision.recommendation);
    sheet
      .getRange(rowNumber, reasonColumn)
      .setValue(decision.reason);
    sheet
      .getRange(rowNumber, confidenceColumn)
      .setValue(decision.confidence);

    if (decision.recommendation === '却下候補') {
      result.rejected++;
    } else if (decision.recommendation === '承認候補') {
      result.approved++;
    } else {
      result.stillHumanReview++;
    }
  });

  return result;
}

function hotelDbV2PostalCodeRefinementDecision_(input) {
  const data = input || {};
  const differences = hotelDbV2Clean_(data.differences)
    .split('・')
    .map(hotelDbV2Clean_)
    .filter(Boolean);

  if (hotelDbV2Clean_(data.state) === '反映済み') {
    return hotelDbV2PostalCodeRefinementNo_('反映済み');
  }

  if (differences.indexOf('郵便番号') === -1) {
    return hotelDbV2PostalCodeRefinementNo_('郵便番号差分なし');
  }

  const sourcePostal = hotelDbV2NormalizePostalCodeNfkc_(
    data.sourcePostalCode
  );
  const proposedPostal = hotelDbV2NormalizePostalCodeNfkc_(
    data.proposedPostalCode
  );

  if (!sourcePostal || !proposedPostal || sourcePostal !== proposedPostal) {
    return hotelDbV2PostalCodeRefinementNo_('正規化後も郵便番号不一致');
  }

  const residualDifferences = differences.filter(function(difference) {
    return difference !== '郵便番号';
  });

  if (!residualDifferences.length) {
    return {
      postalFormattingResolved: true,
      recommendation: '却下候補',
      reason:
        '郵便番号は全角・半角・ハイフン・空白などの表記差だけで、正規化後は同一です。元データ維持を推奨します。',
      confidence: 100,
      residualDifferences: []
    };
  }

  const normalizedInput = Object.assign({}, data, {
    sourcePostalCode: sourcePostal,
    proposedPostalCode: proposedPostal,
    differences: residualDifferences.join('・')
  });

  const triageDecision = hotelDbV2ClassifyCorrectionCandidate_(
    normalizedInput
  );

  if (triageDecision.recommendation === '要人確認') {
    return {
      postalFormattingResolved: true,
      recommendation: '要人確認',
      reason:
        '郵便番号の差は全角・半角などの表記差だけで解消しました。残る差分（' +
        residualDifferences.join('・') +
        '）は人が確認します。',
      confidence: triageDecision.confidence,
      residualDifferences: residualDifferences
    };
  }

  return {
    postalFormattingResolved: true,
    recommendation: triageDecision.recommendation,
    reason:
      '郵便番号の差は全角・半角などの表記差だけで解消しました。残る差分（' +
      residualDifferences.join('・') +
      '）を既存ルールで再判定しました。' +
      triageDecision.reason,
    confidence: triageDecision.confidence,
    residualDifferences: residualDifferences
  };
}

function hotelDbV2PostalCodeRefinementNo_(reason) {
  return {
    postalFormattingResolved: false,
    recommendation: '要人確認',
    reason: reason || '',
    confidence: 0,
    residualDifferences: []
  };
}

/**
 * 郵便番号正規化の自己診断テスト。
 */
function runHotelDbV2PostalCodeNormalizationTests() {
  const normalizationCases = [
    ['ASCII標準', '680-0907', '680-0907'],
    ['末尾だけ全角', '680-090７', '680-0907'],
    ['すべて全角', '６８０－０９０７', '680-0907'],
    ['〒と空白付き', '〒 ６８０－０９０７', '680-0907'],
    ['ハイフンなし', '6800907', '680-0907'],
    ['空白区切り', '680 0907', '680-0907'],
    ['異なる郵便番号', '680-0908', '680-0908'],
    ['桁不足は無効', '680-090', '']
  ];

  const failures = [];
  let passed = 0;

  normalizationCases.forEach(function(testCase, index) {
    const actual = hotelDbV2NormalizePostalCodeNfkc_(testCase[1]);
    if (actual !== testCase[2]) {
      failures.push(
        '正規化例' + (index + 1) + '「' + testCase[0] + '」: ' +
        actual + '（期待=' + testCase[2] + '）'
      );
    } else {
      passed++;
    }
  });

  const base = {
    state: '未確認',
    sourcePostalCode: '680-090７',
    proposedPostalCode: '680-0907',
    sourceMunicipality: '鳥取県鳥取市',
    proposedMunicipality: '鳥取県鳥取市',
    sourceAddress: '賀露町北1-5-36',
    proposedAddress: '賀露町北1-5-36',
    sourceName: '味覚のお宿 山田屋',
    proposedName: '味覚のお宿 山田屋',
    placeId: 'test-place-id',
    matchScore: 93,
    businessStatus: '営業中',
    differences: '郵便番号'
  };

  const decisionCases = [
    {
      name: '郵便番号表記差だけなら却下候補',
      input: Object.assign({}, base),
      expected: '却下候補'
    },
    {
      name: '郵便番号表記差を除いて施設名差が残る場合は要人確認',
      input: Object.assign({}, base, {
        sourcePostalCode: '689-340１',
        proposedPostalCode: '689-3401',
        sourceMunicipality: '鳥取県米子市',
        proposedMunicipality: '鳥取県米子市',
        sourceAddress: '淀江町今津50-1',
        proposedAddress: '淀江町今津50-1',
        sourceName: '淀江の宿 今津田中家',
        proposedName: '淀江の宿 今津田中家 Yodoe Inn -Imazu Tanakaya-',
        matchScore: 83,
        differences: '郵便番号・施設名'
      }),
      expected: '要人確認'
    }
  ];

  decisionCases.forEach(function(testCase, index) {
    const snapshot = JSON.stringify(testCase.input);
    const decision = hotelDbV2PostalCodeRefinementDecision_(testCase.input);
    const unchanged = JSON.stringify(testCase.input) === snapshot;

    if (
      decision.recommendation !== testCase.expected ||
      !decision.postalFormattingResolved ||
      !unchanged
    ) {
      failures.push(
        '判定例' + (index + 1) + '「' + testCase.name + '」: ' +
        '判定=' + decision.recommendation +
        '（期待=' + testCase.expected + '）, ' +
        '郵便番号表記差解消=' + decision.postalFormattingResolved +
        ', 入力保持=' + unchanged
      );
    } else {
      passed++;
    }
  });

  if (failures.length) {
    const message = [
      '郵便番号正規化テストに失敗しました。',
      '',
      failures.join('\n')
    ].join('\n');

    console.error(message);
    throw new Error(message);
  }

  const successMessage = [
    '郵便番号正規化テスト成功',
    '',
    '成功件数: ' + passed + '件',
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
    passed: passed,
    failed: 0
  };
}

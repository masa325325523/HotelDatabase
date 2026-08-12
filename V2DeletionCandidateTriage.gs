/**
 * PR #15 閉業・未検出施設の安全な削除候補仕分け。
 *
 * 「要確認」シートに以下3列だけを追加する。
 * - 削除推奨判定
 * - 削除判定理由
 * - 削除信頼度
 *
 * 元データ・Place ID・状態列は変更せず、自動削除もしない。
 */

const HOTEL_DB_V2_DELETION_TRIAGE_HEADERS = Object.freeze([
  '削除推奨判定',
  '削除判定理由',
  '削除信頼度'
]);

const HOTEL_DB_V2_DELETION_TARGET_REASONS = Object.freeze([
  '閉業',
  'Google候補なし',
  '一時休業',
  '開業予定'
]);

function runHotelDbV2DeletionCandidateTriage() {
  return withHotelDbV2Lock_('閉業・未検出施設の削除候補仕分け', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '閉業・未検出施設を削除候補に仕分け',
      '「要確認」シートの閉業・未検出等を安全側に仕分けします。\n\n' +
      'Google候補なしだけでは削除候補にしません。\n' +
      '恒久閉業でも自動削除はしません。\n' +
      '元データ・Place ID・状態列も変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2TriageDeletionCandidates_();

    ui.alert([
      '閉業・未検出施設の削除候補仕分け完了',
      '',
      '確認件数: ' + result.total,
      '削除候補有力: ' + result.deletionLikely,
      '削除非推奨: ' + result.doNotDelete,
      '要人確認: ' + result.humanReview,
      '対象外: ' + result.outOfScope,
      '',
      '自動削除: なし',
      '状態・元データ・Place IDの変更: なし'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2TriageDeletionCandidates_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);

  if (!sheet || sheet.getLastRow() < 2) {
    return {
      total: 0,
      deletionLikely: 0,
      doNotDelete: 0,
      humanReview: 0,
      outOfScope: 0
    };
  }

  const map = hotelDbV2DeletionEnsureHeaders_(sheet);
  hotelDbV2DeletionValidateHeaders_(map);

  const rowCount = sheet.getLastRow() - 1;
  const values = sheet
    .getRange(2, 1, rowCount, sheet.getLastColumn())
    .getDisplayValues();

  const outputs = [];
  const result = {
    total: values.length,
    deletionLikely: 0,
    doNotDelete: 0,
    humanReview: 0,
    outOfScope: 0
  };

  values.forEach(function(row) {
    const input = hotelDbV2DeletionInputFromRow_(row, map);
    const decision = hotelDbV2ClassifyDeletionCandidate_(input);
    outputs.push(decision);

    if (decision.recommendation === '削除候補有力') result.deletionLikely++;
    else if (decision.recommendation === '削除非推奨') result.doNotDelete++;
    else if (decision.recommendation === '要人確認') result.humanReview++;
    else result.outOfScope++;
  });

  HOTEL_DB_V2_DELETION_TRIAGE_HEADERS.forEach(function(header) {
    const column = map[header];
    const columnValues = outputs.map(function(decision) {
      if (header === '削除推奨判定') return [decision.recommendation];
      if (header === '削除判定理由') return [decision.reason];
      return [decision.confidence];
    });
    sheet.getRange(2, column, rowCount, 1).setValues(columnValues);
  });

  return result;
}

function hotelDbV2DeletionEnsureHeaders_(sheet) {
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0];
  const normalized = headers.map(hotelDbV2NormalizeText_);
  const missing = HOTEL_DB_V2_DELETION_TRIAGE_HEADERS.filter(function(header) {
    return normalized.indexOf(hotelDbV2NormalizeText_(header)) === -1;
  });

  if (missing.length) {
    sheet
      .getRange(1, lastColumn + 1, 1, missing.length)
      .setValues([missing]);
  }

  return hotelDbV2DeletionHeaderMap_(sheet);
}

function hotelDbV2DeletionHeaderMap_(sheet) {
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

function hotelDbV2DeletionValidateHeaders_(map) {
  const required = [
    '状態', '理由', '候補Place ID', '一致スコア', '営業状態',
    '削除推奨判定', '削除判定理由', '削除信頼度'
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

function hotelDbV2DeletionInputFromRow_(row, map) {
  function value(header) {
    return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : '';
  }

  return {
    state: value('状態'),
    reason: value('理由'),
    candidatePlaceId: value('候補Place ID'),
    matchScore: Number(value('一致スコア') || 0),
    businessStatus: value('営業状態'),
    detail: value('詳細')
  };
}

function hotelDbV2ClassifyDeletionCandidate_(input) {
  const data = input || {};
  const state = hotelDbV2Clean_(data.state);
  const reason = hotelDbV2Clean_(data.reason);
  const placeId = hotelDbV2Clean_(data.candidatePlaceId);
  const businessStatus = hotelDbV2Clean_(data.businessStatus);
  const score = Number(data.matchScore || 0);

  if (state && state !== '未確認') {
    return hotelDbV2DeletionDecision_(
      '対象外',
      '状態が「未確認」ではないため、この自動仕分けの対象外です。',
      0
    );
  }

  if (HOTEL_DB_V2_DELETION_TARGET_REASONS.indexOf(reason) === -1) {
    return hotelDbV2DeletionDecision_(
      '対象外',
      '閉業・未検出・一時休業・開業予定の行ではないため対象外です。',
      0
    );
  }

  if (reason === '一時休業') {
    return hotelDbV2DeletionDecision_(
      '削除非推奨',
      '一時休業は営業再開の可能性があるため、削除候補にはしません。',
      0
    );
  }

  if (reason === '開業予定') {
    return hotelDbV2DeletionDecision_(
      '削除非推奨',
      '開業予定施設は今後営業開始する可能性があるため、削除候補にはしません。',
      0
    );
  }

  if (reason === 'Google候補なし') {
    return hotelDbV2DeletionDecision_(
      '要人確認',
      'Googleで候補が見つからないだけでは施設不存在と断定できません。名称変更・Google未登録・検索漏れ等を人が確認します。',
      20
    );
  }

  if (reason === '閉業') {
    if (!placeId) {
      return hotelDbV2DeletionDecision_(
        '要人確認',
        '閉業判定ですが候補Place IDがないため、同一施設と断定せず人が確認します。',
        25
      );
    }

    if (businessStatus !== '閉業') {
      return hotelDbV2DeletionDecision_(
        '要人確認',
        '理由は閉業ですがGoogle営業状態と一致しないため、人が確認します。',
        20
      );
    }

    if (score >= HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE) {
      return hotelDbV2DeletionDecision_(
        '削除候補有力',
        'Place ID付きで同一施設の自動採用基準を満たし、Google営業状態も恒久閉業です。ただし自動削除はせず、削除候補として人が最終確認します。',
        98
      );
    }

    return hotelDbV2DeletionDecision_(
      '要人確認',
      'Googleは閉業ですが一致スコアが75点未満のため、別施設候補の可能性を除外できません。人が確認します。',
      score >= HOTEL_DB_V2_CONFIG.MIN_MATCH_SCORE ? 55 : 30
    );
  }

  return hotelDbV2DeletionDecision_(
    '対象外',
    '削除候補仕分けの対象条件に該当しません。',
    0
  );
}

function hotelDbV2DeletionDecision_(recommendation, reason, confidence) {
  return {
    recommendation: recommendation,
    reason: reason,
    confidence: Math.max(0, Math.min(100, Math.round(Number(confidence) || 0)))
  };
}

function runHotelDbV2DeletionCandidateTriageTests() {
  const cases = [
    [{ state: '未確認', reason: '閉業', candidatePlaceId: 'P1', matchScore: 98, businessStatus: '閉業' }, '削除候補有力', '高一致の恒久閉業'],
    [{ state: '未確認', reason: '閉業', candidatePlaceId: 'P2', matchScore: 75, businessStatus: '閉業' }, '削除候補有力', '75点ちょうどの恒久閉業'],
    [{ state: '未確認', reason: '閉業', candidatePlaceId: 'P3', matchScore: 74, businessStatus: '閉業' }, '要人確認', '75点未満は人確認'],
    [{ state: '未確認', reason: '閉業', candidatePlaceId: '', matchScore: 98, businessStatus: '閉業' }, '要人確認', 'Place IDなし'],
    [{ state: '未確認', reason: '閉業', candidatePlaceId: 'P4', matchScore: 98, businessStatus: '営業中' }, '要人確認', '営業状態不整合'],
    [{ state: '未確認', reason: 'Google候補なし', candidatePlaceId: '', matchScore: 0, businessStatus: '' }, '要人確認', '未検出だけでは削除しない'],
    [{ state: '未確認', reason: 'Google候補なし', candidatePlaceId: 'OLD', matchScore: 0, businessStatus: '' }, '要人確認', '古い候補が残っていても未検出は人確認'],
    [{ state: '未確認', reason: '一時休業', candidatePlaceId: 'P5', matchScore: 98, businessStatus: '一時休業' }, '削除非推奨', '一時休業'],
    [{ state: '未確認', reason: '開業予定', candidatePlaceId: 'P6', matchScore: 98, businessStatus: '開業予定' }, '削除非推奨', '開業予定'],
    [{ state: '未確認', reason: 'APIエラー', candidatePlaceId: '', matchScore: 0, businessStatus: '' }, '対象外', 'APIエラー'],
    [{ state: '未確認', reason: '一致スコア不足', candidatePlaceId: 'P7', matchScore: 40, businessStatus: '営業中' }, '対象外', '低スコアはPR14側'],
    [{ state: '確認済み', reason: '閉業', candidatePlaceId: 'P8', matchScore: 98, businessStatus: '閉業' }, '対象外', '確認済み状態'],
    [{ state: '承認', reason: '閉業', candidatePlaceId: 'P9', matchScore: 98, businessStatus: '閉業' }, '対象外', '承認済み状態'],
    [{ state: '', reason: '閉業', candidatePlaceId: 'P10', matchScore: 90, businessStatus: '閉業' }, '削除候補有力', '空状態は未確認相当'],
    [{ state: '未確認', reason: '閉業', candidatePlaceId: 'P11', matchScore: 54, businessStatus: '閉業' }, '要人確認', '低一致閉業'],
    [{ state: '未確認', reason: '開業予定', candidatePlaceId: '', matchScore: 0, businessStatus: '' }, '削除非推奨', 'Place IDなしでも開業予定は削除しない']
  ];

  const failures = [];

  cases.forEach(function(testCase, index) {
    const decision = hotelDbV2ClassifyDeletionCandidate_(testCase[0]);
    if (decision.recommendation !== testCase[1]) {
      failures.push(
        '例' + (index + 1) + '「' + testCase[2] + '」: 実際=' +
        decision.recommendation + ', 期待=' + testCase[1]
      );
    }
  });

  if (failures.length) {
    throw new Error(
      '削除候補安全仕分けテスト失敗\n\n' + failures.join('\n')
    );
  }

  SpreadsheetApp.getUi().alert([
    '削除候補安全仕分けテスト 成功',
    '',
    '成功件数: ' + cases.length + '件',
    '失敗件数: 0件',
    'Google候補なしだけで削除候補化: なし',
    '一時休業・開業予定の削除候補化: なし',
    '自動削除: なし',
    '元データ・Place ID・状態列の変更: なし'
  ].join('\n'));

  return {
    success: cases.length,
    failure: 0,
    autoDelete: false,
    sourceAutoChange: false
  };
}

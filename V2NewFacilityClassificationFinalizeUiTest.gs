/**
 * PR #18 ⑮→⑯ コピー版UI実地テスト。
 *
 * 必ず「宿泊施設DB_PR13_⑧反映テスト」のようなコピー版だけで使用する。
 * 通常の「新規追加候補」「修正履歴」「新規施設分類候補」はバックアップへ退避し、
 * 標準名にはテスト専用の合成データだけを置く。
 */
const HOTEL_DB_V2_PR18_UI_TEST = Object.freeze({
  REQUIRED_NAME_PARTS: ['PR13', '⑧反映テスト'],
  SOURCE_SHEET: 'PR18_UI分類元',
  BACKUP_NEW_CANDIDATES: '__PR18_BACKUP_新規追加候補',
  BACKUP_HISTORY: '__PR18_BACKUP_修正履歴',
  BACKUP_CLASSIFICATION: '__PR18_BACKUP_新規施設分類候補',
  TEST_ROW_COUNT: 4,
  PLACE_IDS: Object.freeze({
    NORMAL: 'PR18_UI_PLACE_NORMAL',
    UNAPPROVED: 'PR18_UI_PLACE_UNAPPROVED',
    CATEGORY_CHANGED: 'PR18_UI_PLACE_CATEGORY_CHANGED',
    NOTES_CHANGED: 'PR18_UI_PLACE_NOTES_CHANGED'
  })
});

function hotelDbV2Pr18UiAssertCopy_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_PR18_UI_TEST.REQUIRED_NAME_PARTS.every(function(part) {
    return name.indexOf(part) !== -1;
  });
  if (!ok) {
    throw new Error(
      'PR #18 UIテストはコピー版専用です。スプレッドシート名に「PR13」と「⑧反映テスト」の両方が必要です。'
    );
  }
  return spreadsheet;
}

function hotelDbV2Pr18UiHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    const text = hotelDbV2Clean_(header);
    if (text && !map[text]) map[text] = index + 1;
  });
  return map;
}

function setupHotelDbV2NewFacilityClassificationUiTest() {
  const spreadsheet = hotelDbV2Pr18UiAssertCopy_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  let sourceSheet = null;
  let testNewCandidateSheet = null;
  let testHistorySheet = null;
  let backedUpNewCandidates = false;
  let backedUpHistory = false;
  let backedUpClassification = false;

  try {
    if (
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.SOURCE_SHEET) ||
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.BACKUP_NEW_CANDIDATES) ||
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.BACKUP_HISTORY) ||
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.BACKUP_CLASSIFICATION)
    ) {
      throw new Error(
        'PR #18 UIテストの残骸があります。先に cleanupHotelDbV2NewFacilityClassificationUiTest を実行してください。'
      );
    }

    const realNewCandidates = spreadsheet.getSheetByName(
      HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SOURCE_CANDIDATE_SHEET
    );
    if (!realNewCandidates || realNewCandidates.getLastColumn() < 1) {
      throw new Error('通常の「新規追加候補」が見つかりません。');
    }

    const realHeaders = realNewCandidates
      .getRange(1, 1, 1, realNewCandidates.getLastColumn())
      .getDisplayValues()[0];
    const realMap = hotelDbV2Pr18UiHeaderMap_(realHeaders);

    // PR17の監査列（追加先シート・追加先行など）は、通常シートにはまだ無くても正常。
    // UIテスト用シートだけに不足分を追加し、通常シートは一切変更しない。
    const requiredBaseHeaders = [
      '候補キー', '状態', '探索元シート', '探索元シートID',
      '候補Place ID', '候補施設名', '候補住所', '候補郵便番号',
      '候補市区町村', '営業状態', 'Googleタイプ'
    ];
    const missingBase = requiredBaseHeaders.filter(function(header) {
      return !realMap[header];
    });
    if (missingBase.length) {
      throw new Error(
        '通常の「新規追加候補」にPR18テストで必要な基本列がありません: ' +
        missingBase.join('、')
      );
    }

    const testHeaders = realHeaders.slice();
    HOTEL_DB_V2_APPROVED_NEW_FACILITY.AUDIT_HEADERS.forEach(function(header) {
      if (testHeaders.indexOf(header) === -1) testHeaders.push(header);
    });
    const testMap = hotelDbV2Pr18UiHeaderMap_(testHeaders);

    realNewCandidates.setName(HOTEL_DB_V2_PR18_UI_TEST.BACKUP_NEW_CANDIDATES);
    backedUpNewCandidates = true;

    const realHistory = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    if (realHistory) {
      realHistory.setName(HOTEL_DB_V2_PR18_UI_TEST.BACKUP_HISTORY);
      backedUpHistory = true;
    }

    const realClassification = spreadsheet.getSheetByName(
      HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME
    );
    if (realClassification) {
      realClassification.setName(HOTEL_DB_V2_PR18_UI_TEST.BACKUP_CLASSIFICATION);
      backedUpClassification = true;
    }

    sourceSheet = spreadsheet.insertSheet(HOTEL_DB_V2_PR18_UI_TEST.SOURCE_SHEET);
    const sourceHeaders = [
      '〒', '市区町村', '住所（番地まで）', '施設名', '宿泊分類', '備考欄',
      'Place ID', 'Google施設名', 'Google住所', '電話番号', '公式サイト',
      '評価', '口コミ数', '営業状態', '緯度', '経度', '一致スコア',
      '最終確認日', 'Google Maps URL', '照合判定'
    ];
    sourceSheet.getRange(1, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
    const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);

    const sourceCases = [
      {
        placeId: HOTEL_DB_V2_PR18_UI_TEST.PLACE_IDS.NORMAL,
        postal: '100-0001',
        municipality: '東京都千代田区',
        address: '東京都千代田区千代田1-1',
        name: 'PR18正常ホテル',
        googleType: 'hotel,lodging'
      },
      {
        placeId: HOTEL_DB_V2_PR18_UI_TEST.PLACE_IDS.UNAPPROVED,
        postal: '100-0002',
        municipality: '東京都千代田区',
        address: '東京都千代田区皇居外苑1-1',
        name: 'PR18未承認ゲストハウス',
        googleType: 'guest_house,lodging'
      },
      {
        placeId: HOTEL_DB_V2_PR18_UI_TEST.PLACE_IDS.CATEGORY_CHANGED,
        postal: '100-0003',
        municipality: '東京都千代田区',
        address: '東京都千代田区一ツ橋1-1',
        name: 'PR18分類変更旅館',
        googleType: 'japanese_inn,lodging'
      },
      {
        placeId: HOTEL_DB_V2_PR18_UI_TEST.PLACE_IDS.NOTES_CHANGED,
        postal: '100-0004',
        municipality: '東京都千代田区',
        address: '東京都千代田区大手町1-1',
        name: 'PR18備考変更民泊',
        googleType: 'private_guest_room,lodging'
      }
    ];

    const sourceRows = sourceCases.map(function(item) {
      const row = new Array(sourceHeaders.length).fill('');
      function setSource_(key, value) {
        if (sourceMap[key]) row[sourceMap[key] - 1] = value;
      }
      setSource_('postalCode', item.postal);
      setSource_('municipality', item.municipality);
      setSource_('address', item.address);
      setSource_('facilityName', item.name);
      setSource_('category', '');
      setSource_('notes', 'PR17承認新規追加／宿泊分類要確認／PR18 UIテスト');
      setSource_('placeId', item.placeId);
      setSource_('googleName', item.name);
      setSource_('googleAddress', item.address);
      setSource_('businessStatus', '営業中');
      setSource_('matchDecision', '新規追加（承認）');
      return row;
    });
    sourceSheet
      .getRange(2, 1, sourceRows.length, sourceHeaders.length)
      .setValues(sourceRows);

    testNewCandidateSheet = spreadsheet.insertSheet(
      HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SOURCE_CANDIDATE_SHEET
    );
    testNewCandidateSheet
      .getRange(1, 1, 1, testHeaders.length)
      .setValues([testHeaders]);

    const candidateRows = sourceCases.map(function(item, index) {
      const row = new Array(testHeaders.length).fill('');
      function setCandidate_(header, value) {
        if (testMap[header]) row[testMap[header] - 1] = value;
      }
      const sourceRow = index + 2;
      setCandidate_('候補キー', 'PR18_UI_' + (index + 1));
      setCandidate_('状態', HOTEL_DB_V2_APPROVED_NEW_FACILITY.APPLIED_STATE);
      setCandidate_('探索元シート', sourceSheet.getName());
      setCandidate_('探索元シートID', String(sourceSheet.getSheetId()));
      setCandidate_('候補Place ID', item.placeId);
      setCandidate_('候補施設名', item.name);
      setCandidate_('候補住所', item.address);
      setCandidate_('候補郵便番号', item.postal);
      setCandidate_('候補市区町村', item.municipality);
      setCandidate_('営業状態', '営業中');
      setCandidate_('Googleタイプ', item.googleType);
      setCandidate_('追加先シート', sourceSheet.getName());
      setCandidate_('追加先行', String(sourceRow));
      setCandidate_('追加結果', '追加済み。宿泊分類要確認');
      return row;
    });
    testNewCandidateSheet
      .getRange(2, 1, candidateRows.length, testHeaders.length)
      .setValues(candidateRows);

    testHistorySheet = spreadsheet.insertSheet(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    testHistorySheet
      .getRange(1, 1, 1, HOTEL_DB_V2_HISTORY_HEADERS.length)
      .setValues([HOTEL_DB_V2_HISTORY_HEADERS]);

    sourceSheet.setFrozenRows(1);
    testNewCandidateSheet.setFrozenRows(1);
    testHistorySheet.setFrozenRows(1);
    spreadsheet.setActiveSheet(testNewCandidateSheet);

    SpreadsheetApp.getUi().alert([
      'PR #18 ⑮→⑯UIテスト 準備完了',
      '',
      '追加済みテスト施設: 4件',
      '正常反映1件・未承認1件・分類変更停止1件・備考変更停止1件',
      '通常の新規追加候補・修正履歴・分類候補: バックアップ済み',
      '',
      '次にメニューの⑮「追加済み新規施設の分類候補を作成」を実行してください。'
    ].join('\n'));

    return { ready: true, rows: 4 };
  } catch (error) {
    try {
      const testClassification = spreadsheet.getSheetByName(
        HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME
      );
      if (testClassification) spreadsheet.deleteSheet(testClassification);
      if (testHistorySheet) spreadsheet.deleteSheet(testHistorySheet);
      if (testNewCandidateSheet) spreadsheet.deleteSheet(testNewCandidateSheet);
      if (sourceSheet) spreadsheet.deleteSheet(sourceSheet);

      const backupNewCandidates = spreadsheet.getSheetByName(
        HOTEL_DB_V2_PR18_UI_TEST.BACKUP_NEW_CANDIDATES
      );
      const backupHistory = spreadsheet.getSheetByName(
        HOTEL_DB_V2_PR18_UI_TEST.BACKUP_HISTORY
      );
      const backupClassification = spreadsheet.getSheetByName(
        HOTEL_DB_V2_PR18_UI_TEST.BACKUP_CLASSIFICATION
      );

      if (backedUpNewCandidates && backupNewCandidates) {
        backupNewCandidates.setName(
          HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SOURCE_CANDIDATE_SHEET
        );
      }
      if (backedUpHistory && backupHistory) {
        backupHistory.setName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
      }
      if (backedUpClassification && backupClassification) {
        backupClassification.setName(HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME);
      }
    } catch (rollbackError) {
      throw new Error(error.message + '\n復元にも失敗しました: ' + rollbackError.message);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function prepareHotelDbV2NewFacilityClassificationApprovalUiTest() {
  const spreadsheet = hotelDbV2Pr18UiAssertCopy_();
  const sourceSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.SOURCE_SHEET);
  const candidateSheet = spreadsheet.getSheetByName(
    HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME
  );
  if (!sourceSheet || !candidateSheet) {
    throw new Error(
      'PR #18 UIテスト用シートがありません。先にセットアップ→⑮を実行してください。'
    );
  }

  const map = hotelDbV2NewFacilityClassificationHeaderMap_(candidateSheet);
  const required = [
    '状態', 'Place ID', '参考分類', '確定宿泊分類', '確定備考'
  ];
  const missing = required.filter(function(header) { return !map[header]; });
  if (missing.length) {
    throw new Error(
      '「新規施設分類候補」に必要な列がありません: ' + missing.join('、')
    );
  }

  const values = candidateSheet
    .getRange(2, 1, candidateSheet.getLastRow() - 1, candidateSheet.getLastColumn())
    .getDisplayValues();
  const rowByPlaceId = {};
  values.forEach(function(row, index) {
    const placeId = hotelDbV2Clean_(row[map['Place ID'] - 1]);
    if (placeId) rowByPlaceId[placeId] = index + 2;
  });

  const ids = HOTEL_DB_V2_PR18_UI_TEST.PLACE_IDS;
  [ids.NORMAL, ids.UNAPPROVED, ids.CATEGORY_CHANGED, ids.NOTES_CHANGED]
    .forEach(function(id) {
      if (!rowByPlaceId[id]) {
        throw new Error('⑮で作成されるはずの分類候補が不足しています: ' + id);
      }
    });

  const expectedReferences = {};
  expectedReferences[ids.NORMAL] = 'ホテル系（要確認）';
  expectedReferences[ids.UNAPPROVED] = '簡易宿所系（要確認）';
  expectedReferences[ids.CATEGORY_CHANGED] = '旅館系（要確認）';
  expectedReferences[ids.NOTES_CHANGED] = '住宅宿泊事業・簡易宿所系（要確認）';

  Object.keys(expectedReferences).forEach(function(placeId) {
    const row = rowByPlaceId[placeId];
    const actual = hotelDbV2Clean_(
      candidateSheet.getRange(row, map['参考分類']).getDisplayValue()
    );
    if (actual !== expectedReferences[placeId]) {
      throw new Error(
        '⑮の参考分類が期待値と一致しません: ' +
        placeId + ' / expected=' + expectedReferences[placeId] + ' / actual=' + actual
      );
    }
  });

  function setCandidate_(placeId, state, category, notes) {
    const row = rowByPlaceId[placeId];
    candidateSheet.getRange(row, map['状態']).setValue(state);
    candidateSheet.getRange(row, map['確定宿泊分類']).setValue(category || '');
    candidateSheet.getRange(row, map['確定備考']).setValue(notes || '');
  }

  setCandidate_(ids.NORMAL, '承認', 'ホテル営業', '郵便番号一致');
  setCandidate_(ids.UNAPPROVED, '未確認', '', '');
  setCandidate_(ids.CATEGORY_CHANGED, '承認', '旅館営業', '分類確認テスト');
  setCandidate_(ids.NOTES_CHANGED, '承認', '簡易宿所営業', '備考確認テスト');

  const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
  sourceSheet.getRange(4, sourceMap.category).setValue('旅館営業');
  sourceSheet.getRange(5, sourceMap.notes).setValue(
    'PR17承認新規追加／宿泊分類要確認／PR18 UIテスト／手修正'
  );

  spreadsheet.setActiveSheet(candidateSheet);
  SpreadsheetApp.getUi().alert([
    'PR #18 ⑯UIテスト 承認準備完了',
    '',
    '承認: 3件（正常1件・元分類変更1件・元備考変更1件）',
    '未確認: 1件',
    '⑮のGoogleタイプ参考分類4ケース: 確認済み',
    '元分類変更・元備考変更は、⑯で要再確認になれば正常です。',
    '',
    '次にメニューの⑯「承認済み宿泊分類・備考を安全に反映」を実行してください。'
  ].join('\n'));

  return { ready: true, approved: 3, unapproved: 1 };
}

function verifyHotelDbV2NewFacilityClassificationUiTest() {
  const spreadsheet = hotelDbV2Pr18UiAssertCopy_();
  const sourceSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.SOURCE_SHEET);
  const candidateSheet = spreadsheet.getSheetByName(
    HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME
  );
  const historySheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
  if (!sourceSheet || !candidateSheet || !historySheet) {
    throw new Error(
      'PR #18 UIテスト用シートがありません。セットアップから順番に実行してください。'
    );
  }

  const failures = [];
  let success = 0;
  function check_(label, condition, detail) {
    if (condition) success++;
    else failures.push(label + (detail ? ': ' + detail : ''));
  }

  const map = hotelDbV2NewFacilityClassificationHeaderMap_(candidateSheet);
  const values = candidateSheet
    .getRange(2, 1, candidateSheet.getLastRow() - 1, candidateSheet.getLastColumn())
    .getDisplayValues();
  const rowByPlaceId = {};
  values.forEach(function(row) {
    const placeId = hotelDbV2Clean_(row[map['Place ID'] - 1]);
    if (placeId) rowByPlaceId[placeId] = row;
  });
  function cv_(placeId, header) {
    const row = rowByPlaceId[placeId];
    return row && map[header]
      ? hotelDbV2Clean_(row[map[header] - 1])
      : '';
  }

  const ids = HOTEL_DB_V2_PR18_UI_TEST.PLACE_IDS;
  check_(
    '正常承認は反映済み',
    cv_(ids.NORMAL, '状態') === '反映済み',
    cv_(ids.NORMAL, '状態')
  );
  check_(
    '未承認は未確認のまま',
    cv_(ids.UNAPPROVED, '状態') === '未確認',
    cv_(ids.UNAPPROVED, '状態')
  );
  check_(
    '元分類変更は要再確認',
    cv_(ids.CATEGORY_CHANGED, '状態') === '要再確認',
    cv_(ids.CATEGORY_CHANGED, '状態')
  );
  check_(
    '元備考変更は要再確認',
    cv_(ids.NOTES_CHANGED, '状態') === '要再確認',
    cv_(ids.NOTES_CHANGED, '状態')
  );
  check_('正常承認に反映日時あり', !!cv_(ids.NORMAL, '反映日時'));
  check_(
    '正常承認の反映結果',
    cv_(ids.NORMAL, '反映結果') === '反映済み',
    cv_(ids.NORMAL, '反映結果')
  );
  check_(
    '分類変更は未反映記録',
    cv_(ids.CATEGORY_CHANGED, '反映結果').indexOf('未反映') !== -1,
    cv_(ids.CATEGORY_CHANGED, '反映結果')
  );
  check_(
    '備考変更は未反映記録',
    cv_(ids.NOTES_CHANGED, '反映結果').indexOf('未反映') !== -1,
    cv_(ids.NOTES_CHANGED, '反映結果')
  );

  const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
  const sourceValues = sourceSheet
    .getRange(2, 1, 4, sourceSheet.getLastColumn())
    .getDisplayValues();
  function sv_(index, key) {
    return sourceMap[key]
      ? hotelDbV2Clean_(sourceValues[index][sourceMap[key] - 1])
      : '';
  }

  check_(
    '正常行だけホテル営業へ確定',
    sv_(0, 'category') === 'ホテル営業',
    sv_(0, 'category')
  );
  check_(
    '正常行の要確認マーカー除去',
    sv_(0, 'notes').indexOf('宿泊分類要確認') === -1,
    sv_(0, 'notes')
  );
  check_(
    '正常行の既存備考保持',
    sv_(0, 'notes').indexOf('PR17承認新規追加') !== -1,
    sv_(0, 'notes')
  );
  check_(
    '正常行に確認済み追記',
    sv_(0, 'notes').indexOf('宿泊分類確認済') !== -1,
    sv_(0, 'notes')
  );
  check_(
    '正常行に確定備考追記',
    sv_(0, 'notes').indexOf('郵便番号一致') !== -1,
    sv_(0, 'notes')
  );
  check_('未承認行は分類未入力', sv_(1, 'category') === '', sv_(1, 'category'));
  check_(
    '未承認行は要確認マーカー保持',
    sv_(1, 'notes').indexOf('宿泊分類要確認') !== -1,
    sv_(1, 'notes')
  );
  check_(
    '分類変更行を上書きしない',
    sv_(2, 'category') === '旅館営業',
    sv_(2, 'category')
  );
  check_(
    '備考変更行の分類は未入力',
    sv_(3, 'category') === '',
    sv_(3, 'category')
  );
  check_(
    '備考変更行の手修正を保持',
    sv_(3, 'notes').indexOf('手修正') !== -1,
    sv_(3, 'notes')
  );

  const historyCount = Math.max(0, historySheet.getLastRow() - 1);
  check_('履歴は正常反映1件だけ', historyCount === 1, String(historyCount));
  if (historyCount === 1) {
    const historyMap = hotelDbV2NewFacilityClassificationHeaderMap_(historySheet);
    const historyRow = historySheet
      .getRange(2, 1, 1, historySheet.getLastColumn())
      .getDisplayValues()[0];
    check_(
      '履歴の処理名',
      historyMap['処理'] &&
      hotelDbV2Clean_(historyRow[historyMap['処理'] - 1]) === '新規施設分類確定'
    );
  } else {
    check_('履歴の処理名', false, '履歴件数が1件ではありません');
  }

  check_(
    '通常の新規追加候補バックアップ保持',
    !!spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.BACKUP_NEW_CANDIDATES)
  );
  check_(
    '分類候補はテスト4件のみ',
    Object.keys(rowByPlaceId).length === 4,
    String(Object.keys(rowByPlaceId).length)
  );

  if (failures.length) {
    throw new Error([
      'PR #18 ⑮→⑯UIテスト 失敗',
      '',
      '成功件数: ' + success + '件',
      '失敗件数: ' + failures.length + '件',
      '',
      failures.join('\n')
    ].join('\n'));
  }

  SpreadsheetApp.getUi().alert([
    'PR #18 ⑮→⑯UIテスト 成功',
    '',
    '成功件数: ' + success + '件',
    '失敗件数: 0件',
    '正常承認: 宿泊分類・備考だけ反映',
    '未承認: 未反映',
    '候補作成後の分類変更: 要再確認・未反映',
    '候補作成後の備考変更: 要再確認・未反映',
    '履歴: 正常反映1件だけ記録',
    '',
    '次に cleanupHotelDbV2NewFacilityClassificationUiTest を実行してください。'
  ].join('\n'));

  return { success: true, successCount: success, failureCount: 0 };
}

function cleanupHotelDbV2NewFacilityClassificationUiTest() {
  const spreadsheet = hotelDbV2Pr18UiAssertCopy_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  try {
    const sourceSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_PR18_UI_TEST.SOURCE_SHEET);
    const backupNewCandidates = spreadsheet.getSheetByName(
      HOTEL_DB_V2_PR18_UI_TEST.BACKUP_NEW_CANDIDATES
    );
    const backupHistory = spreadsheet.getSheetByName(
      HOTEL_DB_V2_PR18_UI_TEST.BACKUP_HISTORY
    );
    const backupClassification = spreadsheet.getSheetByName(
      HOTEL_DB_V2_PR18_UI_TEST.BACKUP_CLASSIFICATION
    );

    if (!sourceSheet && !backupNewCandidates && !backupHistory && !backupClassification) {
      throw new Error('PR #18 UIテストの復元対象が見つかりません。');
    }

    const testNewCandidates = spreadsheet.getSheetByName(
      HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SOURCE_CANDIDATE_SHEET
    );
    const testHistory = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    const testClassification = spreadsheet.getSheetByName(
      HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME
    );

    if (backupNewCandidates && testNewCandidates) {
      spreadsheet.deleteSheet(testNewCandidates);
    }
    if (backupHistory && testHistory) {
      spreadsheet.deleteSheet(testHistory);
    }
    if (backupClassification && testClassification) {
      spreadsheet.deleteSheet(testClassification);
    }
    if (!backupClassification && sourceSheet && testClassification) {
      spreadsheet.deleteSheet(testClassification);
    }
    if (!backupHistory && sourceSheet && testHistory) {
      spreadsheet.deleteSheet(testHistory);
    }
    if (sourceSheet) spreadsheet.deleteSheet(sourceSheet);

    if (backupNewCandidates) {
      backupNewCandidates.setName(
        HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SOURCE_CANDIDATE_SHEET
      );
    }
    if (backupHistory) {
      backupHistory.setName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    }
    if (backupClassification) {
      backupClassification.setName(HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME);
    }

    SpreadsheetApp.getUi().alert([
      'PR #18 ⑮→⑯UIテスト 復元完了',
      '',
      'テスト用分類元・新規追加候補・修正履歴・分類候補: 削除済み',
      '通常の新規追加候補: 復元済み',
      '通常の修正履歴: ' + (backupHistory ? '復元済み' : '元々なし'),
      '通常の新規施設分類候補: ' +
        (backupClassification ? '復元済み' : '元々なし')
    ].join('\n'));

    return { restored: true };
  } finally {
    lock.releaseLock();
  }
}

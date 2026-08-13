/**
 * PR #17 ⑭「承認済み新規追加候補を安全に追加」コピー版UI実地テスト。
 *
 * 必ずコピー版スプレッドシートだけで使用する。
 * 通常の「新規追加候補」「修正履歴」はバックアップ名へ退避し、
 * 標準名にはテスト専用データだけを置く。
 */
const HOTEL_DB_V2_PR17_UI_TEST = Object.freeze({
  REQUIRED_NAME_PARTS: ['PR13', '⑧反映テスト'],
  SOURCE_SHEET: 'PR17_UI追加先',
  BACKUP_CANDIDATES: '__PR17_BACKUP_新規追加候補',
  BACKUP_HISTORY: '__PR17_BACKUP_修正履歴',
  TEST_ROW_COUNT: 4
});

function hotelDbV2Pr17UiAssertCopy_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_PR17_UI_TEST.REQUIRED_NAME_PARTS.every(function(part) {
    return name.indexOf(part) !== -1;
  });
  if (!ok) {
    throw new Error(
      'PR #17 UIテストはコピー版専用です。スプレッドシート名に「PR13」と「⑧反映テスト」の両方が必要です。'
    );
  }
  return spreadsheet;
}

function setupHotelDbV2ApprovedNewFacilityUiTest() {
  const spreadsheet = hotelDbV2Pr17UiAssertCopy_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  let sourceSheet = null;
  let testCandidateSheet = null;
  let testHistorySheet = null;
  let backedUpCandidates = false;
  let backedUpHistory = false;

  try {
    if (
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR17_UI_TEST.SOURCE_SHEET) ||
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR17_UI_TEST.BACKUP_CANDIDATES) ||
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR17_UI_TEST.BACKUP_HISTORY)
    ) {
      throw new Error('PR #17 UIテストの残骸があります。先に cleanupHotelDbV2ApprovedNewFacilityUiTest を実行してください。');
    }

    const realCandidates = spreadsheet.getSheetByName(
      HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME
    );
    if (!realCandidates || realCandidates.getLastRow() < 2) {
      throw new Error('通常の「新規追加候補」が見つかりません。先に⑬で候補を作成してください。');
    }

    const candidateHeaderMap = hotelDbV2ApprovedNewFacilityHeaderMap_(realCandidates);
    const required = HOTEL_DB_V2_APPROVED_NEW_FACILITY.REQUIRED_HEADERS.concat(['推奨判定']);
    const missing = required.filter(function(header) {
      return !candidateHeaderMap[header];
    });
    if (missing.length) {
      throw new Error('通常の「新規追加候補」に必要な列がありません: ' + missing.join('、'));
    }

    const candidateValues = realCandidates
      .getRange(2, 1, realCandidates.getLastRow() - 1, realCandidates.getLastColumn())
      .getDisplayValues();

    const eligible = [];
    const seenPlaceIds = {};
    candidateValues.forEach(function(row) {
      const placeId = hotelDbV2Clean_(row[candidateHeaderMap['候補Place ID'] - 1]);
      const recommendation = hotelDbV2Clean_(row[candidateHeaderMap['推奨判定'] - 1]);
      const businessStatus = hotelDbV2Clean_(row[candidateHeaderMap['営業状態'] - 1]);
      const name = hotelDbV2Clean_(row[candidateHeaderMap['候補施設名'] - 1]);
      const address = hotelDbV2Clean_(row[candidateHeaderMap['候補住所'] - 1]);
      const postal = hotelDbV2NormalizePostalCode_(row[candidateHeaderMap['候補郵便番号'] - 1]);
      const targetMunicipality = hotelDbV2Clean_(row[candidateHeaderMap['対象市区町村'] - 1]);
      const candidateMunicipality = hotelDbV2Clean_(row[candidateHeaderMap['候補市区町村'] - 1]);

      if (
        recommendation !== '新規候補有力' ||
        businessStatus !== '営業中' ||
        !placeId || !name || !address || !postal || !targetMunicipality ||
        !candidateMunicipality || seenPlaceIds[placeId]
      ) return;

      seenPlaceIds[placeId] = true;
      eligible.push(row.slice());
    });

    if (eligible.length < 3) {
      throw new Error(
        'PR #17 UIテストには、営業中かつ「新規候補有力」の異なるPlace IDが3件必要です。現在: ' + eligible.length + '件'
      );
    }

    const normalRow = eligible[0].slice();
    const duplicateRow = eligible[1].slice();
    const changedRow = eligible[2].slice();
    const unapprovedRow = eligible[0].slice();

    realCandidates.setName(HOTEL_DB_V2_PR17_UI_TEST.BACKUP_CANDIDATES);
    backedUpCandidates = true;

    const realHistory = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    if (realHistory) {
      realHistory.setName(HOTEL_DB_V2_PR17_UI_TEST.BACKUP_HISTORY);
      backedUpHistory = true;
    }

    sourceSheet = spreadsheet.insertSheet(HOTEL_DB_V2_PR17_UI_TEST.SOURCE_SHEET);
    const sourceHeaders = [
      '〒', '市区町村', '住所（番地まで）', '施設名', '宿泊分類', '備考欄',
      'Place ID', 'Google施設名', 'Google住所', '電話番号', '公式サイト',
      '評価', '口コミ数', '営業状態', '緯度', '経度', '一致スコア',
      '最終確認日', 'Google Maps URL', '照合判定'
    ];
    sourceSheet.getRange(1, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);

    const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
    const duplicateSourceRow = new Array(sourceHeaders.length).fill('');
    function setSource_(key, value) {
      if (sourceMap[key]) duplicateSourceRow[sourceMap[key] - 1] = value;
    }
    setSource_('postalCode', duplicateRow[candidateHeaderMap['候補郵便番号'] - 1]);
    setSource_('municipality', duplicateRow[candidateHeaderMap['候補市区町村'] - 1]);
    setSource_('address', duplicateRow[candidateHeaderMap['候補住所'] - 1]);
    setSource_('facilityName', duplicateRow[candidateHeaderMap['候補施設名'] - 1]);
    setSource_('category', '');
    setSource_('notes', 'PR17 UIテスト用の既存Place ID行');
    setSource_('placeId', duplicateRow[candidateHeaderMap['候補Place ID'] - 1]);
    sourceSheet.getRange(2, 1, 1, duplicateSourceRow.length).setValues([duplicateSourceRow]);

    testCandidateSheet = spreadsheet.insertSheet(
      HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME
    );
    const originalHeaders = spreadsheet
      .getSheetByName(HOTEL_DB_V2_PR17_UI_TEST.BACKUP_CANDIDATES)
      .getRange(1, 1, 1, spreadsheet.getSheetByName(HOTEL_DB_V2_PR17_UI_TEST.BACKUP_CANDIDATES).getLastColumn())
      .getDisplayValues()[0];
    testCandidateSheet.getRange(1, 1, 1, originalHeaders.length).setValues([originalHeaders]);
    const testMap = hotelDbV2ApprovedNewFacilityHeaderMap_(testCandidateSheet);

    function prepareRow_(row, key, state) {
      const next = row.slice();
      next[testMap['候補キー'] - 1] = key;
      next[testMap['状態'] - 1] = state;
      next[testMap['探索元シート'] - 1] = sourceSheet.getName();
      next[testMap['探索元シートID'] - 1] = String(sourceSheet.getSheetId());
      return next;
    }

    normalRow = prepareRow_(normalRow, 'PR17_UI_NORMAL', '承認');
    unapprovedRow = prepareRow_(unapprovedRow, 'PR17_UI_UNAPPROVED', '未確認');
    duplicateRow = prepareRow_(duplicateRow, 'PR17_UI_DUPLICATE', '承認');
    changedRow = prepareRow_(changedRow, 'PR17_UI_CHANGED', '承認');
    changedRow[testMap['候補施設名'] - 1] =
      hotelDbV2Clean_(changedRow[testMap['候補施設名'] - 1]) + ' PR17変更テスト';

    testCandidateSheet
      .getRange(2, 1, HOTEL_DB_V2_PR17_UI_TEST.TEST_ROW_COUNT, originalHeaders.length)
      .setValues([normalRow, unapprovedRow, duplicateRow, changedRow]);

    testHistorySheet = spreadsheet.insertSheet(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    testHistorySheet
      .getRange(1, 1, 1, HOTEL_DB_V2_HISTORY_HEADERS.length)
      .setValues([HOTEL_DB_V2_HISTORY_HEADERS]);

    sourceSheet.setFrozenRows(1);
    testCandidateSheet.setFrozenRows(1);
    testHistorySheet.setFrozenRows(1);
    spreadsheet.setActiveSheet(testCandidateSheet);

    SpreadsheetApp.getUi().alert([
      'PR #17 ⑭UIテスト 準備完了',
      '',
      '承認: 3件（正常1件・既存Place ID重複1件・候補名変更1件）',
      '未確認: 1件',
      '通常の新規追加候補・修正履歴: バックアップ済み',
      '',
      '次にメニューの⑭「承認済み新規追加候補を安全に追加」を実行してください。'
    ].join('\n'));

    return { ready: true, approved: 3, unapproved: 1 };
  } catch (error) {
    try {
      if (testHistorySheet) spreadsheet.deleteSheet(testHistorySheet);
      if (testCandidateSheet) spreadsheet.deleteSheet(testCandidateSheet);
      if (sourceSheet) spreadsheet.deleteSheet(sourceSheet);

      const backupCandidates = spreadsheet.getSheetByName(
        HOTEL_DB_V2_PR17_UI_TEST.BACKUP_CANDIDATES
      );
      const backupHistory = spreadsheet.getSheetByName(
        HOTEL_DB_V2_PR17_UI_TEST.BACKUP_HISTORY
      );
      if (backedUpCandidates && backupCandidates) {
        backupCandidates.setName(HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME);
      }
      if (backedUpHistory && backupHistory) {
        backupHistory.setName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
      }
    } catch (rollbackError) {
      throw new Error(error.message + '\n復元にも失敗しました: ' + rollbackError.message);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function verifyHotelDbV2ApprovedNewFacilityUiTest() {
  const spreadsheet = hotelDbV2Pr17UiAssertCopy_();
  const sourceSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_PR17_UI_TEST.SOURCE_SHEET);
  const candidateSheet = spreadsheet.getSheetByName(
    HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME
  );
  const historySheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);

  if (!sourceSheet || !candidateSheet || !historySheet) {
    throw new Error('PR #17 UIテスト用シートが見つかりません。先にセットアップを実行してください。');
  }

  const failures = [];
  let successCount = 0;
  function check_(label, condition, detail) {
    if (condition) successCount++;
    else failures.push(label + (detail ? ': ' + detail : ''));
  }

  const candidateMap = hotelDbV2ApprovedNewFacilityHeaderMap_(candidateSheet);
  const candidateValues = candidateSheet
    .getRange(2, 1, HOTEL_DB_V2_PR17_UI_TEST.TEST_ROW_COUNT, candidateSheet.getLastColumn())
    .getDisplayValues();

  function cv_(rowIndex, header) {
    return candidateMap[header]
      ? hotelDbV2Clean_(candidateValues[rowIndex][candidateMap[header] - 1])
      : '';
  }

  check_('正常承認は追加済み', cv_(0, '状態') === '追加済み', cv_(0, '状態'));
  check_('未承認は未確認のまま', cv_(1, '状態') === '未確認', cv_(1, '状態'));
  check_('Place ID重複は要再確認', cv_(2, '状態') === '要再確認', cv_(2, '状態'));
  check_('候補名変更は要再確認', cv_(3, '状態') === '要再確認', cv_(3, '状態'));
  check_('正常承認に追加日時あり', !!cv_(0, '追加処理日時'));
  check_('正常承認の追加先シート', cv_(0, '追加先シート') === sourceSheet.getName(), cv_(0, '追加先シート'));
  check_('正常承認の追加先行あり', !!cv_(0, '追加先行'));
  check_('未承認は監査列未変更', !cv_(1, '追加処理日時') && !cv_(1, '追加結果'));
  check_('重複は未追加記録', cv_(2, '追加結果').indexOf('未追加') !== -1, cv_(2, '追加結果'));
  check_('候補変更は未追加記録', cv_(3, '追加結果').indexOf('未追加') !== -1, cv_(3, '追加結果'));

  const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
  const lastDataRow = hotelDbV2ApprovedNewFacilityLastDataRow_(sourceSheet, sourceMap);
  check_('元シートのデータ行は2件だけ', lastDataRow === 3, '最終データ行=' + lastDataRow);

  const added = hotelDbV2ReadFacility_(sourceSheet, 3, sourceMap);
  check_('正常承認のPlace IDだけ追加', added.placeId === cv_(0, '候補Place ID'), added.placeId);
  check_('追加施設名が一致', hotelDbV2NormalizeText_(added.name) === hotelDbV2NormalizeText_(cv_(0, '候補施設名')), added.name);
  check_('宿泊分類は空欄', !hotelDbV2Clean_(added.category), added.category);
  check_('備考に宿泊分類要確認', added.notes.indexOf('宿泊分類要確認') !== -1, added.notes);

  const historyValues = historySheet.getLastRow() >= 2
    ? historySheet
        .getRange(2, 1, historySheet.getLastRow() - 1, HOTEL_DB_V2_HISTORY_HEADERS.length)
        .getDisplayValues()
    : [];
  const historyMap = hotelDbV2HeaderIndex_(HOTEL_DB_V2_HISTORY_HEADERS);
  const results = historyValues.map(function(row) {
    return hotelDbV2Clean_(row[historyMap['結果'] - 1]);
  });
  check_('履歴は3件', historyValues.length === 3, '実際=' + historyValues.length);
  check_('追加済み履歴あり', results.indexOf('追加済み') !== -1);
  check_('要再確認履歴2件', results.filter(function(value) {
    return value === '要再確認・未追加';
  }).length === 2);

  if (failures.length) {
    throw new Error([
      'PR #17 ⑭UIテスト失敗',
      '',
      '成功件数: ' + successCount + '件',
      '失敗件数: ' + failures.length + '件',
      '',
      failures.join('\n')
    ].join('\n'));
  }

  SpreadsheetApp.getUi().alert([
    'PR #17 ⑭UIテスト 成功',
    '',
    '成功件数: 18件',
    '失敗件数: 0件',
    '正常承認: 1件だけ追加',
    '未承認: 未追加',
    '既存Place ID重複: 要再確認・未追加',
    '候補作成後の施設名変更: 要再確認・未追加',
    '宿泊分類: 自動入力なし',
    '履歴: 3件記録',
    '',
    '次に cleanupHotelDbV2ApprovedNewFacilityUiTest を実行してください。'
  ].join('\n'));

  return { success: true, successCount: 18, failureCount: 0 };
}

function cleanupHotelDbV2ApprovedNewFacilityUiTest() {
  const spreadsheet = hotelDbV2Pr17UiAssertCopy_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  try {
    const testCandidates = spreadsheet.getSheetByName(
      HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME
    );
    const testHistory = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    const testSource = spreadsheet.getSheetByName(HOTEL_DB_V2_PR17_UI_TEST.SOURCE_SHEET);
    const backupCandidates = spreadsheet.getSheetByName(
      HOTEL_DB_V2_PR17_UI_TEST.BACKUP_CANDIDATES
    );
    const backupHistory = spreadsheet.getSheetByName(
      HOTEL_DB_V2_PR17_UI_TEST.BACKUP_HISTORY
    );

    if (!backupCandidates) {
      throw new Error('退避した通常の「新規追加候補」が見つかりません。自動削除は行いません。');
    }

    if (testCandidates) spreadsheet.deleteSheet(testCandidates);
    if (testHistory) spreadsheet.deleteSheet(testHistory);
    if (testSource) spreadsheet.deleteSheet(testSource);

    backupCandidates.setName(HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME);
    if (backupHistory) backupHistory.setName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);

    spreadsheet.setActiveSheet(
      spreadsheet.getSheetByName(HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME)
    );

    SpreadsheetApp.getUi().alert([
      'PR #17 ⑭UIテスト 復元完了',
      '',
      'テスト用シート: 削除済み',
      '通常の新規追加候補: 復元済み',
      '退避した修正履歴: ' + (backupHistory ? '復元済み' : '元々なし')
    ].join('\n'));

    return { cleaned: true };
  } finally {
    lock.releaseLock();
  }
}

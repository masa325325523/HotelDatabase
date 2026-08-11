/**
 * PR #13
 * ⑧「承認済み修正候補を反映」のUI実地テスト補助。
 *
 * 必ず「コピー版スプレッドシート」でだけ使用する。
 * セットアップ時に通常の修正候補・修正履歴をバックアップ名へ退避し、
 * ⑧が触る標準名にはテスト専用データだけを配置する。
 */

const HOTEL_DB_V2_PR13_UI_TEST = Object.freeze({
  REQUIRED_NAME_PARTS: ['PR13', '⑧反映テスト'],
  SOURCE_SHEET: 'PR13_UI反映元',
  BACKUP_CORRECTIONS: '__PR13_BACKUP_修正候補',
  BACKUP_HISTORY: '__PR13_BACKUP_修正履歴'
});

function hotelDbV2Pr13UiAssertCopy_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_PR13_UI_TEST.REQUIRED_NAME_PARTS.every(function(part) {
    return name.indexOf(part) !== -1;
  });
  if (!ok) {
    throw new Error(
      'PR #13 UIテストはコピー版専用です。スプレッドシート名に「PR13」と「⑧反映テスト」の両方が必要です。'
    );
  }
  return spreadsheet;
}

function setupHotelDbV2ApprovedCorrectionUiTest() {
  const spreadsheet = hotelDbV2Pr13UiAssertCopy_();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  let sourceSheet = null;
  let correctionSheet = null;
  let historySheet = null;
  let backedUpCorrections = false;
  let backedUpHistory = false;

  try {
    if (
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.SOURCE_SHEET) ||
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_CORRECTIONS) ||
      spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_HISTORY)
    ) {
      throw new Error('PR #13 UIテストの残骸があります。先に cleanupHotelDbV2ApprovedCorrectionUiTest を実行してください。');
    }

    const realCorrections = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
    const realHistory = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);

    if (realCorrections) {
      realCorrections.setName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_CORRECTIONS);
      backedUpCorrections = true;
    }
    if (realHistory) {
      realHistory.setName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_HISTORY);
      backedUpHistory = true;
    }

    sourceSheet = spreadsheet.insertSheet(HOTEL_DB_V2_PR13_UI_TEST.SOURCE_SHEET);
    correctionSheet = spreadsheet.insertSheet(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
    historySheet = spreadsheet.insertSheet(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);

    const sourceHeaders = ['〒', '市区町村', '住所（番地まで）', '施設名', '宿泊分類'];
    sourceSheet.getRange(1, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
    sourceSheet.getRange(2, 1, 3, sourceHeaders.length).setValues([
      ['680-1001', '鳥取県鳥取市', '本町1-100', 'PR13テスト旅館A', '旅館'],
      ['680-1002', '鳥取県鳥取市', '本町2-200', 'PR13テスト旅館B', '旅館'],
      ['680-1003', '鳥取県鳥取市', '本町3-999', 'PR13テスト旅館C', '旅館']
    ]);

    correctionSheet
      .getRange(1, 1, 1, HOTEL_DB_V2_CORRECTION_HEADERS.length)
      .setValues([HOTEL_DB_V2_CORRECTION_HEADERS]);
    historySheet
      .getRange(1, 1, 1, HOTEL_DB_V2_HISTORY_HEADERS.length)
      .setValues([HOTEL_DB_V2_HISTORY_HEADERS]);

    const sourceSheetId = sourceSheet.getSheetId();
    const rows = [
      hotelDbV2RowFromObject_(HOTEL_DB_V2_CORRECTION_HEADERS, {
        '候補キー': sourceSheetId + '|2',
        '状態': '承認',
        '元シート': sourceSheet.getName(),
        '元シートID': sourceSheetId,
        '元行': 2,
        '元郵便番号': '680-1001',
        '修正郵便番号': '680-1001',
        '元市区町村': '鳥取県鳥取市',
        '修正市区町村': '鳥取県鳥取市',
        '元住所': '本町1-100',
        '修正住所': '本町1-100',
        '元施設名': 'PR13テスト旅館A',
        '修正施設名': 'PR13テスト旅館A 改',
        'Place ID': 'PR13_UI_APPLY',
        '一致スコア': 99,
        '営業状態': '営業中',
        'Google Maps URL': '',
        '差分': '施設名',
        '確認日': hotelDbV2Today_(),
        '反映日時': ''
      }),
      hotelDbV2RowFromObject_(HOTEL_DB_V2_CORRECTION_HEADERS, {
        '候補キー': sourceSheetId + '|3',
        '状態': '未確認',
        '元シート': sourceSheet.getName(),
        '元シートID': sourceSheetId,
        '元行': 3,
        '元郵便番号': '680-1002',
        '修正郵便番号': '680-1002',
        '元市区町村': '鳥取県鳥取市',
        '修正市区町村': '鳥取県鳥取市',
        '元住所': '本町2-200',
        '修正住所': '本町2-200',
        '元施設名': 'PR13テスト旅館B',
        '修正施設名': 'PR13テスト旅館B 改',
        'Place ID': 'PR13_UI_UNAPPROVED',
        '一致スコア': 99,
        '営業状態': '営業中',
        'Google Maps URL': '',
        '差分': '施設名',
        '確認日': hotelDbV2Today_(),
        '反映日時': ''
      }),
      hotelDbV2RowFromObject_(HOTEL_DB_V2_CORRECTION_HEADERS, {
        '候補キー': sourceSheetId + '|4',
        '状態': '承認',
        '元シート': sourceSheet.getName(),
        '元シートID': sourceSheetId,
        '元行': 4,
        '元郵便番号': '680-1003',
        '修正郵便番号': '680-1003',
        '元市区町村': '鳥取県鳥取市',
        '修正市区町村': '鳥取県鳥取市',
        '元住所': '本町3-300',
        '修正住所': '本町3-301',
        '元施設名': 'PR13テスト旅館C',
        '修正施設名': 'PR13テスト旅館C',
        'Place ID': 'PR13_UI_CONFLICT',
        '一致スコア': 95,
        '営業状態': '営業中',
        'Google Maps URL': '',
        '差分': '住所',
        '確認日': hotelDbV2Today_(),
        '反映日時': ''
      })
    ];

    correctionSheet
      .getRange(2, 1, rows.length, HOTEL_DB_V2_CORRECTION_HEADERS.length)
      .setValues(rows);

    correctionSheet.setFrozenRows(1);
    historySheet.setFrozenRows(1);
    sourceSheet.setFrozenRows(1);
    spreadsheet.setActiveSheet(correctionSheet);

    SpreadsheetApp.getUi().alert([
      'PR #13 ⑧UIテスト準備完了',
      '',
      '承認: 2件（正常1件・競合1件）',
      '未確認: 1件',
      '通常の修正候補・修正履歴: バックアップ済み',
      '',
      '次にメニューの⑧「承認済み修正候補を反映」を実行してください。'
    ].join('\n'));

    return { ready: true, approved: 2, unapproved: 1 };
  } catch (error) {
    try {
      if (historySheet) spreadsheet.deleteSheet(historySheet);
      if (correctionSheet) spreadsheet.deleteSheet(correctionSheet);
      if (sourceSheet) spreadsheet.deleteSheet(sourceSheet);
      const backupCorrections = spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_CORRECTIONS);
      const backupHistory = spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_HISTORY);
      if (backedUpCorrections && backupCorrections) backupCorrections.setName(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
      if (backedUpHistory && backupHistory) backupHistory.setName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
    } catch (rollbackError) {
      throw new Error(error.message + '\n復元にも失敗しました: ' + rollbackError.message);
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function verifyHotelDbV2ApprovedCorrectionUiTest() {
  const spreadsheet = hotelDbV2Pr13UiAssertCopy_();
  const sourceSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.SOURCE_SHEET);
  const correctionSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
  const historySheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
  if (!sourceSheet || !correctionSheet || !historySheet) {
    throw new Error('PR #13 UIテスト用シートが見つかりません。先にセットアップを実行してください。');
  }

  const failures = [];
  let successCount = 0;
  function check_(label, condition, detail) {
    if (condition) {
      successCount++;
    } else {
      failures.push(label + (detail ? ': ' + detail : ''));
    }
  }

  const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
  const row2 = hotelDbV2ReadFacility_(sourceSheet, 2, sourceMap);
  const row3 = hotelDbV2ReadFacility_(sourceSheet, 3, sourceMap);
  const row4 = hotelDbV2ReadFacility_(sourceSheet, 4, sourceMap);
  const correctionMap = hotelDbV2HeaderIndex_(HOTEL_DB_V2_CORRECTION_HEADERS);
  const correctionValues = correctionSheet
    .getRange(2, 1, 3, HOTEL_DB_V2_CORRECTION_HEADERS.length)
    .getDisplayValues();
  const historyMap = hotelDbV2HeaderIndex_(HOTEL_DB_V2_HISTORY_HEADERS);
  const historyValues = historySheet.getLastRow() >= 2
    ? historySheet
        .getRange(2, 1, historySheet.getLastRow() - 1, HOTEL_DB_V2_HISTORY_HEADERS.length)
        .getDisplayValues()
    : [];
  const historyResults = historyValues.map(function(row) {
    return hotelDbV2Clean_(row[historyMap['結果'] - 1]);
  });

  check_('正常承認1件を反映', row2.name === 'PR13テスト旅館A 改', row2.name);
  check_('未承認行は未反映', row3.name === 'PR13テスト旅館B', row3.name);
  check_('競合行は未反映', row4.address === '本町3-999', row4.address);
  check_('正常承認の状態', hotelDbV2Clean_(correctionValues[0][correctionMap['状態'] - 1]) === '反映済み');
  check_('正常承認の反映日時', !!hotelDbV2Clean_(correctionValues[0][correctionMap['反映日時'] - 1]));
  check_('未承認の状態維持', hotelDbV2Clean_(correctionValues[1][correctionMap['状態'] - 1]) === '未確認');
  check_('競合を要再確認へ', hotelDbV2Clean_(correctionValues[2][correctionMap['状態'] - 1]) === '要再確認');
  check_('履歴2件', historyValues.length === 2, '実際' + historyValues.length);
  check_('反映済み履歴', historyResults.indexOf('反映済み') !== -1);
  check_('競合・未反映履歴', historyResults.indexOf('競合・未反映') !== -1);
  check_('修正候補はテスト3件のみ', correctionSheet.getLastRow() === 4, '最終行' + correctionSheet.getLastRow());
  check_('修正履歴はテスト2件のみ', historySheet.getLastRow() === 3, '最終行' + historySheet.getLastRow());

  if (failures.length) {
    throw new Error([
      'PR #13 ⑧UIテスト失敗',
      '',
      '成功件数: ' + successCount + '件',
      '失敗件数: ' + failures.length + '件',
      '',
      failures.join('\n')
    ].join('\n'));
  }

  SpreadsheetApp.getUi().alert([
    'PR #13 ⑧UIテスト 成功',
    '',
    '成功件数: 12件',
    '失敗件数: 0件',
    '正常承認: 1件だけ反映',
    '未承認: 未反映',
    '競合: 要再確認・未反映',
    '履歴: 2件記録'
  ].join('\n'));

  return { success: 12, failure: 0 };
}

function cleanupHotelDbV2ApprovedCorrectionUiTest() {
  const spreadsheet = hotelDbV2Pr13UiAssertCopy_();
  const sourceSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.SOURCE_SHEET);
  if (!sourceSheet) {
    throw new Error('PR #13 UIテスト元シートがありません。誤削除防止のため処理を中止しました。');
  }

  const correctionSheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
  const historySheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
  if (correctionSheet) spreadsheet.deleteSheet(correctionSheet);
  if (historySheet) spreadsheet.deleteSheet(historySheet);
  spreadsheet.deleteSheet(sourceSheet);

  const backupCorrections = spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_CORRECTIONS);
  const backupHistory = spreadsheet.getSheetByName(HOTEL_DB_V2_PR13_UI_TEST.BACKUP_HISTORY);
  if (backupCorrections) backupCorrections.setName(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
  if (backupHistory) backupHistory.setName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);

  SpreadsheetApp.getUi().alert([
    'PR #13 ⑧UIテスト 復元完了',
    '',
    'テスト用シート: 削除済み',
    '退避した修正候補・修正履歴: 復元済み'
  ].join('\n'));

  return { cleaned: true };
}

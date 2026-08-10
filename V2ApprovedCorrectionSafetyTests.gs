/**
 * PR #13
 * ⑧「承認済み修正候補を反映」の隔離安全テスト。
 *
 * 本物の「修正候補」「修正履歴」は使用しない。
 * 一時シートだけで本番と同じ中核処理を実行し、終了後に削除する。
 */

function runHotelDbV2ApprovedCorrectionSafetyTests() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  const failures = [];
  let successCount = 0;
  let sourceSheet = null;
  let correctionSheet = null;
  let historySheet = null;

  function snapshotSheet_(name) {
    const sheet = spreadsheet.getSheetByName(name);
    return {
      exists: !!sheet,
      lastRow: sheet ? sheet.getLastRow() : 0
    };
  }

  const realCorrectionsBefore = snapshotSheet_(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
  const realHistoryBefore = snapshotSheet_(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);

  function check_(label, condition, detail) {
    if (condition) {
      successCount++;
      return;
    }
    failures.push(label + (detail ? ': ' + detail : ''));
  }

  try {
    const suffix = String(new Date().getTime());
    sourceSheet = spreadsheet.insertSheet('__PR13反映元_' + suffix);
    correctionSheet = spreadsheet.insertSheet('__PR13修正候補_' + suffix);
    historySheet = spreadsheet.insertSheet('__PR13修正履歴_' + suffix);

    const sourceHeaders = [
      '〒', '市区町村', '住所（番地まで）', '施設名', '宿泊分類'
    ];
    sourceSheet.getRange(1, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
    sourceSheet.getRange(2, 1, 3, sourceHeaders.length).setValues([
      ['680-0001', '鳥取県鳥取市', '本町1-100', 'テスト旅館A', '旅館'],
      ['680-0002', '鳥取県鳥取市', '本町2-200', 'テスト旅館B', '旅館'],
      ['680-0003', '鳥取県鳥取市', '本町3-999', 'テスト旅館C', '旅館']
    ]);

    correctionSheet
      .getRange(1, 1, 1, HOTEL_DB_V2_CORRECTION_HEADERS.length)
      .setValues([HOTEL_DB_V2_CORRECTION_HEADERS]);
    historySheet
      .getRange(1, 1, 1, HOTEL_DB_V2_HISTORY_HEADERS.length)
      .setValues([HOTEL_DB_V2_HISTORY_HEADERS]);

    const sourceSheetId = sourceSheet.getSheetId();
    const correctionRows = [
      hotelDbV2RowFromObject_(HOTEL_DB_V2_CORRECTION_HEADERS, {
        '候補キー': sourceSheetId + '|2',
        '状態': '承認',
        '元シート': sourceSheet.getName(),
        '元シートID': sourceSheetId,
        '元行': 2,
        '元郵便番号': '680-0001',
        '修正郵便番号': '680-0001',
        '元市区町村': '鳥取県鳥取市',
        '修正市区町村': '鳥取県鳥取市',
        '元住所': '本町1-100',
        '修正住所': '本町1-100',
        '元施設名': 'テスト旅館A',
        '修正施設名': 'テスト旅館A 改',
        'Place ID': 'PR13_APPLY',
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
        '元郵便番号': '680-0002',
        '修正郵便番号': '680-0002',
        '元市区町村': '鳥取県鳥取市',
        '修正市区町村': '鳥取県鳥取市',
        '元住所': '本町2-200',
        '修正住所': '本町2-200',
        '元施設名': 'テスト旅館B',
        '修正施設名': 'テスト旅館B 改',
        'Place ID': 'PR13_UNAPPROVED',
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
        '元郵便番号': '680-0003',
        '修正郵便番号': '680-0003',
        '元市区町村': '鳥取県鳥取市',
        '修正市区町村': '鳥取県鳥取市',
        '元住所': '本町3-300',
        '修正住所': '本町3-301',
        '元施設名': 'テスト旅館C',
        '修正施設名': 'テスト旅館C',
        'Place ID': 'PR13_CONFLICT',
        '一致スコア': 95,
        '営業状態': '営業中',
        'Google Maps URL': '',
        '差分': '住所',
        '確認日': hotelDbV2Today_(),
        '反映日時': ''
      })
    ];

    correctionSheet
      .getRange(2, 1, correctionRows.length, HOTEL_DB_V2_CORRECTION_HEADERS.length)
      .setValues(correctionRows);

    const result = hotelDbV2ApplyApprovedCorrectionsWithContext_({
      spreadsheet: spreadsheet,
      correctionSheet: correctionSheet,
      historySheet: historySheet
    });

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

    const realCorrectionsAfter = snapshotSheet_(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS);
    const realHistoryAfter = snapshotSheet_(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);

    check_('承認対象件数', result.approved === 2, '期待2 / 実際' + result.approved);
    check_('反映済み件数', result.applied === 1, '期待1 / 実際' + result.applied);
    check_('競合件数', result.conflicts === 1, '期待1 / 実際' + result.conflicts);
    check_('エラー件数', result.errors === 0, '期待0 / 実際' + result.errors);
    check_('承認行だけ元データへ反映', row2.name === 'テスト旅館A 改', row2.name);
    check_('未承認行は元データ維持', row3.name === 'テスト旅館B', row3.name);
    check_('競合行は元データ維持', row4.address === '本町3-999', row4.address);
    check_(
      '反映済み状態',
      hotelDbV2Clean_(correctionValues[0][correctionMap['状態'] - 1]) === '反映済み'
    );
    check_(
      '反映日時を記録',
      !!hotelDbV2Clean_(correctionValues[0][correctionMap['反映日時'] - 1])
    );
    check_(
      '未承認状態を維持',
      hotelDbV2Clean_(correctionValues[1][correctionMap['状態'] - 1]) === '未確認'
    );
    check_(
      '競合を要再確認へ変更',
      hotelDbV2Clean_(correctionValues[2][correctionMap['状態'] - 1]) === '要再確認'
    );
    check_('履歴件数', historyValues.length === 2, '期待2 / 実際' + historyValues.length);
    check_('反映済み履歴', historyResults.indexOf('反映済み') !== -1);
    check_('競合・未反映履歴', historyResults.indexOf('競合・未反映') !== -1);
    check_(
      '本物の修正候補シートを未変更',
      realCorrectionsBefore.exists === realCorrectionsAfter.exists &&
        realCorrectionsBefore.lastRow === realCorrectionsAfter.lastRow
    );
    check_(
      '本物の修正履歴シートを未変更',
      realHistoryBefore.exists === realHistoryAfter.exists &&
        realHistoryBefore.lastRow === realHistoryAfter.lastRow
    );
  } catch (error) {
    failures.push('テスト実行エラー: ' + (error && error.message ? error.message : error));
  } finally {
    [historySheet, correctionSheet, sourceSheet].forEach(function(sheet) {
      if (!sheet) return;
      try {
        spreadsheet.deleteSheet(sheet);
      } catch (cleanupError) {
        failures.push('一時シート削除エラー: ' + cleanupError.message);
      }
    });
    lock.releaseLock();
  }

  if (failures.length) {
    throw new Error([
      '承認済み修正反映 安全テスト失敗',
      '',
      '成功件数: ' + successCount + '件',
      '失敗件数: ' + failures.length + '件',
      '',
      failures.join('\n')
    ].join('\n'));
  }

  SpreadsheetApp.getUi().alert([
    '承認済み修正反映 安全テスト 成功',
    '',
    '成功件数: 16件',
    '失敗件数: 0件',
    '本物の修正候補・修正履歴: 変更なし',
    '一時テストシート: 削除済み'
  ].join('\n'));

  return {
    success: 16,
    failure: 0,
    realSheetsChanged: false,
    temporarySheetsDeleted: true
  };
}

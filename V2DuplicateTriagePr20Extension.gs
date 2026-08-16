/**
 * PR #20: 既存⑩の判定ロジックを変えず、重複濃厚候補に行スナップショットを追加する。
 * 既存PR #12の純粋判定関数・自己診断はそのまま保持する。
 */
function runHotelDbV2DuplicateTriageWithPr20Snapshots() {
  return withHotelDbV2Lock_('重複候補の自動仕分け＋整理スナップショット', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '重複候補を自動仕分け',
      '「重複候補」シートを従来どおり安全条件で分類し、\n' +
      '「重複濃厚」の候補だけ行1/行2の安全スナップショットを作成します。\n\n' +
      '元データは削除・変更しません。\n' +
      '承認済み候補でスナップショットが新規・変更になった場合だけ、安全のため「要再確認」に戻します。\n' +
      '続行しますか？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2DuplicateTriageWithPr20Snapshots_();
    ui.alert([
      '重複候補の自動仕分け完了', '',
      '確認件数: ' + result.scanned,
      '重複濃厚: ' + result.strongDuplicate,
      '要人確認: ' + result.needReview,
      '整理スナップショット作成: ' + result.snapshotCreated,
      'スナップショット要再確認: ' + result.snapshotReview,
      '承認解除・再確認へ戻した件数: ' + result.approvalReset, '',
      '自動削除: 0',
      '元データの変更: なし'
    ].join('\n'));
    return result;
  });
}

function hotelDbV2DuplicateTriageWithPr20Snapshots_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES);
  if (!sheet || sheet.getLastRow() < 2) {
    return {
      scanned: 0, strongDuplicate: 0, needReview: 0,
      snapshotCreated: 0, snapshotReview: 0, approvalReset: 0
    };
  }

  hotelDbV2DuplicateTriageEnsureHeaders_(sheet);
  const headerMap = hotelDbV2Pr20EnsureDuplicateHeaders_(sheet);
  hotelDbV2DuplicateTriageValidateHeaders_(headerMap);

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const result = {
    scanned: 0, strongDuplicate: 0, needReview: 0,
    snapshotCreated: 0, snapshotReview: 0, approvalReset: 0
  };

  values.forEach(function(row, offset) {
    if (!hotelDbV2DuplicateTriageHasData_(row, headerMap)) return;
    const input = {
      existingDecision: row[headerMap['判定'] - 1],
      name1: row[headerMap['施設名1'] - 1], address1: row[headerMap['住所1'] - 1],
      name2: row[headerMap['施設名2'] - 1], address2: row[headerMap['住所2'] - 1],
      placeId: row[headerMap['Place ID'] - 1], similarity: row[headerMap['類似度'] - 1],
      state: row[headerMap['状態'] - 1]
    };
    const decision = hotelDbV2DuplicateTriageDecision_(input);
    const rowNumber = offset + 2;

    sheet.getRange(rowNumber, headerMap['推奨判定']).setValue(decision.recommendation);
    sheet.getRange(rowNumber, headerMap['自動判定理由']).setValue(decision.reason);
    sheet.getRange(rowNumber, headerMap['信頼度']).setValue(decision.confidence);

    const snapshot = hotelDbV2Pr20PrepareTriageSnapshot_(
      spreadsheet, sheet, rowNumber, headerMap, decision
    );
    result.snapshotCreated += snapshot.created;
    result.snapshotReview += snapshot.review;
    result.approvalReset += snapshot.approvalReset;
    result.scanned++;
    if (decision.recommendation === HOTEL_DB_V2_PR20.STRONG) result.strongDuplicate++;
    else result.needReview++;
  });

  return result;
}

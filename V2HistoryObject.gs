/**
 * Shared history object helper.
 * PR #19 source-sheet guard refactor keeps this helper available unchanged.
 */
function hotelDbV2HistoryObject_(
  sheet,
  facility,
  action,
  result,
  placeId,
  score,
  status,
  detail
) {
  return {
    '日時': hotelDbV2Timestamp_(),
    '元シート': sheet.getName(),
    '元シートID': sheet.getSheetId(),
    '元行': facility.rowNumber,
    '施設名': facility.name,
    '処理': action,
    '結果': result,
    'Place ID': placeId,
    '一致スコア': score,
    '営業状態': status,
    '詳細': detail
  };
}
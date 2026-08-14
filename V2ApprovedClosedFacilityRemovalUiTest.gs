/**
 * PR #19 コピー版UIテスト。
 * 宿泊施設DB_PR13_⑧反映テスト 専用。通常シートは退避し、cleanupで復元する。
 */
const HOTEL_DB_V2_PR19_UI_TEST = Object.freeze({
  REQUIRED_NAME_PARTS: ['PR13', '⑧反映テスト'],
  SOURCE: 'PR19_閉業除外テスト元DB',
  BACKUP_REVIEW: 'PR19_要確認退避',
  BACKUP_ARCHIVE: 'PR19_閉業除外履歴退避',
  BACKUP_HISTORY: 'PR19_修正履歴退避'
});

function hotelDbV2Pr19AssertCopy_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_PR19_UI_TEST.REQUIRED_NAME_PARTS.every(function(part) { return name.indexOf(part) !== -1; });
  if (!ok) throw new Error('PR #19 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」専用です。');
  return spreadsheet;
}

function setupHotelDbV2ApprovedClosedFacilityRemovalUiTest() {
  const ss = hotelDbV2Pr19AssertCopy_();
  [HOTEL_DB_V2_PR19_UI_TEST.BACKUP_REVIEW, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_ARCHIVE, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_HISTORY, HOTEL_DB_V2_PR19_UI_TEST.SOURCE]
    .forEach(function(name) { if (ss.getSheetByName(name)) throw new Error('PR19テスト残骸があります: ' + name + '。先にcleanupしてください。'); });

  hotelDbV2Pr19BackupSheet_(ss, HOTEL_DB_V2_CONFIG.SHEETS.REVIEW, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_REVIEW);
  hotelDbV2Pr19BackupSheet_(ss, HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_ARCHIVE);
  hotelDbV2Pr19BackupSheet_(ss, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_HISTORY);

  const sourceHeaders = ['郵便番号','市区町村名','住所（番地まで）','施設名','宿泊分類','備考','Place ID'];
  const source = ss.insertSheet(HOTEL_DB_V2_PR19_UI_TEST.SOURCE);
  source.getRange(1,1,1,sourceHeaders.length).setValues([sourceHeaders]);
  const sourceRows = [
    ['100-0001','東京都千代田区','千代田1-1','閉業テストホテルA','ホテル営業','PR19正常','TEST_CLOSED_98'],
    ['100-0002','東京都千代田区','皇居外苑1-1','未承認テストホテル','ホテル営業','未承認','TEST_UNAPPROVED'],
    ['100-0003','東京都千代田区','一ツ橋1-1','一時休業テストホテル','ホテル営業','一時休業','TEST_TEMP'],
    ['100-0004','東京都千代田区','大手町1-1','元データ変更テスト','ホテル営業','変更前','TEST_CHANGED'],
    ['100-0005','東京都千代田区','丸の内1-1','営業再開テストホテル','ホテル営業','営業再開','TEST_REOPENED'],
    ['100-0006','東京都千代田区','有楽町1-1','APIエラーテストホテル','ホテル営業','APIエラー','TEST_API_ERROR'],
    ['100-0007','東京都千代田区','内幸町1-1','下行番号保持テストホテル','ホテル営業','営業中','TEST_BELOW']
  ];
  source.getRange(2,1,sourceRows.length,sourceHeaders.length).setValues(sourceRows);

  const review = ss.insertSheet(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);
  const headers = HOTEL_DB_V2_REVIEW_HEADERS.concat(HOTEL_DB_V2_DELETION_TRIAGE_HEADERS).concat(HOTEL_DB_V2_PR19_AUDIT_HEADERS);
  review.getRange(1,1,1,headers.length).setValues([headers]);
  const map = hotelDbV2HeaderIndex_(headers);
  const candidateDefs = [
    [2,'閉業','閉業テストホテルA','東京都千代田区千代田1-1','TEST_CLOSED_98',98,'閉業','未確認'],
    [3,'閉業','未承認テストホテル','東京都千代田区皇居外苑1-1','TEST_UNAPPROVED',98,'閉業','未確認'],
    [4,'一時休業','一時休業テストホテル','東京都千代田区一ツ橋1-1','TEST_TEMP',98,'一時休業','未確認'],
    [5,'閉業','元データ変更テスト','東京都千代田区大手町1-1','TEST_CHANGED',98,'閉業','未確認'],
    [6,'閉業','営業再開テストホテル','東京都千代田区丸の内1-1','TEST_REOPENED',98,'閉業','未確認'],
    [7,'閉業','APIエラーテストホテル','東京都千代田区有楽町1-1','TEST_API_ERROR',98,'閉業','未確認']
  ];
  const rows = candidateDefs.map(function(d,index) {
    return hotelDbV2RowFromObject_(headers, {
      '確認キー':'PR19_TEST|' + (index+1), '状態':d[7], '元シート':source.getName(), '元シートID':source.getSheetId(), '元行':d[0],
      '郵便番号':sourceRows[d[0]-2][0], '市区町村':sourceRows[d[0]-2][1], '住所':sourceRows[d[0]-2][2], '施設名':sourceRows[d[0]-2][3], '宿泊分類':sourceRows[d[0]-2][4],
      '理由':d[1], '候補施設名':d[2], '候補住所':d[3], '候補Place ID':d[4], '一致スコア':d[5], '営業状態':d[6], 'Google Maps URL':'', '確認日':hotelDbV2Today_(), '詳細':'PR19 UI test'
    });
  });
  review.getRange(2,1,rows.length,headers.length).setValues(rows);
  review.activate();

  SpreadsheetApp.getUi().alert([
    'PR #19 UIテスト準備完了','',
    '1) ⑫を実行して削除候補・ハッシュを作成','2) 正常/変更/営業再開/APIエラーの対象を承認','3) 元データ変更テストの備考を「変更後」に変える','4) testHotelDbV2ApprovedClosedFacilityRemovalUiTest() を実行','5) cleanupで復元'
  ].join('\n'));
  return { prepared:true, sourceSheetId:source.getSheetId(), reviewRows:rows.length, belowRow:8 };
}

function testHotelDbV2ApprovedClosedFacilityRemovalUiTest() {
  const ss = hotelDbV2Pr19AssertCopy_();
  const source = ss.getSheetByName(HOTEL_DB_V2_PR19_UI_TEST.SOURCE);
  const review = ss.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);
  if (!source || !review) throw new Error('setupを先に実行してください。');
  const beforeBelowName = source.getRange(8,4).getDisplayValue();
  const result = hotelDbV2ApplyApprovedClosedFacilityRemovals_({
    spreadsheet:ss, reviewSheet:review,
    placeDetailsProvider:function(placeId) {
      if (placeId === 'TEST_API_ERROR') throw new Error('PR19 UIテスト用APIエラー');
      const rowMap = {
        TEST_CLOSED_98:['閉業テストホテルA','東京都千代田区千代田1-1','100-0001','CLOSED_PERMANENTLY'],
        TEST_CHANGED:['元データ変更テスト','東京都千代田区大手町1-1','100-0004','CLOSED_PERMANENTLY'],
        TEST_REOPENED:['営業再開テストホテル','東京都千代田区丸の内1-1','100-0005','OPERATIONAL'],
        TEST_UNAPPROVED:['未承認テストホテル','東京都千代田区皇居外苑1-1','100-0002','CLOSED_PERMANENTLY']
      };
      const d=rowMap[placeId]; if(!d) return null;
      return { id:placeId, displayName:{text:d[0]}, formattedAddress:d[1], businessStatus:d[3], addressComponents:[{longText:d[2],types:['postal_code']}] };
    },
    scoreProvider:function(){ return 98; }
  });
  const afterBelowName = source.getRange(8,4).getDisplayValue();
  if (beforeBelowName !== afterBelowName || afterBelowName !== '下行番号保持テストホテル') throw new Error('下の行番号が変化しました。');
  SpreadsheetApp.getUi().alert(['PR #19 UIテスト実行完了','', '承認対象:'+result.approved, '除外済み:'+result.applied, '要再確認:'+result.conflicts, 'エラー:'+result.errors, '下行番号保持: 成功'].join('\n'));
  return result;
}

function cleanupHotelDbV2ApprovedClosedFacilityRemovalUiTest() {
  const ss = hotelDbV2Pr19AssertCopy_();
  [HOTEL_DB_V2_CONFIG.SHEETS.REVIEW, HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY, HOTEL_DB_V2_PR19_UI_TEST.SOURCE]
    .forEach(function(name) { const s=ss.getSheetByName(name); if(s) ss.deleteSheet(s); });
  hotelDbV2Pr19RestoreBackup_(ss, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_REVIEW, HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);
  hotelDbV2Pr19RestoreBackup_(ss, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_ARCHIVE, HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME);
  hotelDbV2Pr19RestoreBackup_(ss, HOTEL_DB_V2_PR19_UI_TEST.BACKUP_HISTORY, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
  SpreadsheetApp.getUi().alert('PR #19 UIテスト復元完了\n\n通常シートを復元しました。');
  return { cleaned:true };
}

function hotelDbV2Pr19BackupSheet_(ss, original, backup) {
  const sheet=ss.getSheetByName(original); if(sheet) sheet.setName(backup);
}
function hotelDbV2Pr19RestoreBackup_(ss, backup, original) {
  const sheet=ss.getSheetByName(backup); if(sheet) sheet.setName(original);
}
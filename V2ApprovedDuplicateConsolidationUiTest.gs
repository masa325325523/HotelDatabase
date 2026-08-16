/**
 * PR #20 コピー版UIテスト。
 * 宿泊施設DB_PR13_⑧反映テスト 専用。通常の重複候補・履歴は退避し、cleanupで復元する。
 */
const HOTEL_DB_V2_PR20_UI_TEST = Object.freeze({
  REQUIRED_NAME_PARTS: ['PR13', '⑧反映テスト'],
  SOURCE: 'PR20_重複整理テスト元DB',
  BACKUP_DUPLICATES: 'PR20_重複候補退避',
  BACKUP_ARCHIVE: 'PR20_重複整理履歴退避',
  BACKUP_HISTORY: 'PR20_修正履歴退避'
});

function hotelDbV2Pr20AssertCopy_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_PR20_UI_TEST.REQUIRED_NAME_PARTS.every(function(part) {
    return name.indexOf(part) !== -1;
  });
  if (!ok) throw new Error('PR #20 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」専用です。');
  return spreadsheet;
}

function setupHotelDbV2ApprovedDuplicateConsolidationUiTest() {
  const ss = hotelDbV2Pr20AssertCopy_();
  [
    HOTEL_DB_V2_PR20_UI_TEST.BACKUP_DUPLICATES,
    HOTEL_DB_V2_PR20_UI_TEST.BACKUP_ARCHIVE,
    HOTEL_DB_V2_PR20_UI_TEST.BACKUP_HISTORY,
    HOTEL_DB_V2_PR20_UI_TEST.SOURCE
  ].forEach(function(name) {
    if (ss.getSheetByName(name)) throw new Error('PR20テスト残骸があります: ' + name + '。先にcleanupしてください。');
  });

  hotelDbV2Pr20BackupSheet_(ss, HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES, HOTEL_DB_V2_PR20_UI_TEST.BACKUP_DUPLICATES);
  hotelDbV2Pr20BackupSheet_(ss, HOTEL_DB_V2_PR20_ARCHIVE_SHEET_NAME, HOTEL_DB_V2_PR20_UI_TEST.BACKUP_ARCHIVE);
  hotelDbV2Pr20BackupSheet_(ss, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY, HOTEL_DB_V2_PR20_UI_TEST.BACKUP_HISTORY);

  const headers = ['郵便番号','市区町村名','住所（番地まで）','施設名','宿泊分類','備考','Place ID'];
  const source = ss.insertSheet(HOTEL_DB_V2_PR20_UI_TEST.SOURCE);
  source.getRange(1,1,1,headers.length).setValues([headers]);
  const rows = [
    ['100-0001','東京都千代田区','千代田1-1','同一ホテルA','ホテル営業','確認済み','P1'],
    ['100-0001','東京都千代田区','千代田1-1','同一ホテルA','ホテル営業','','P1'],
    ['100-0001','東京都千代田区','千代田1-1','同一ホテルA','ホテル営業','','P1'],
    ['100-0001','東京都千代田区','千代田1-1','同一ホテルA','ホテル営業','除外行だけの重要備考','P1'],
    ['100-0001','東京都千代田区','千代田1-1','同一ホテルA','ホテル営業','','P1'],
    ['100-0002','東京都千代田区','大手町1-1 101号室','同建物テスト','簡易宿所営業','','ROOM101'],
    ['100-0002','東京都千代田区','大手町1-1 102号室','同建物テスト','簡易宿所営業','','ROOM102'],
    ['100-0003','東京都千代田区','丸の内1-1','下行番号保持テストホテル','ホテル営業','営業中','BELOW']
  ];
  source.getRange(2,1,rows.length,headers.length).setValues(rows);

  const duplicateHeaders = HOTEL_DB_V2_DUPLICATE_HEADERS
    .concat(HOTEL_DB_V2_DUPLICATE_TRIAGE_HEADERS)
    .concat(HOTEL_DB_V2_PR20_AUDIT_HEADERS);
  const duplicates = ss.insertSheet(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES);
  duplicates.getRange(1,1,1,duplicateHeaders.length).setValues([duplicateHeaders]);
  const candidates = [
    [2,3,'同一ホテルA','東京都千代田区千代田1-1','同一ホテルA','東京都千代田区千代田1-1','P1',100],
    [2,4,'同一ホテルA','東京都千代田区千代田1-1','同一ホテルA','東京都千代田区千代田1-1','P1',100],
    [2,5,'同一ホテルA','東京都千代田区千代田1-1','同一ホテルA','東京都千代田区千代田1-1','P1',100],
    [2,6,'同一ホテルA','東京都千代田区千代田1-1','同一ホテルA','東京都千代田区千代田1-1','P1',100],
    [7,8,'同建物テスト','東京都千代田区大手町1-1 101号室','同建物テスト','東京都千代田区大手町1-1 102号室','',95]
  ].map(function(d,index) {
    return hotelDbV2RowFromObject_(duplicateHeaders, {
      '重複キー':'PR20_TEST|' + (index+1), '元シート':source.getName(), '元シートID':source.getSheetId(),
      '判定': index === 4 ? '類似・要確認' : '施設名＋住所一致',
      '行1':d[0], '施設名1':d[2], '住所1':d[3],
      '行2':d[1], '施設名2':d[4], '住所2':d[5],
      'Place ID':d[6], '類似度':d[7], '確認日':hotelDbV2Today_(), '状態':'未確認'
    });
  });
  duplicates.getRange(2,1,candidates.length,duplicateHeaders.length).setValues(candidates);
  duplicates.activate();

  SpreadsheetApp.getUi().alert([
    'PR #20 UIテスト準備完了','',
    '次にメニュー⑩「重複候補を自動仕分け」を実行してください。','',
    'その後 Apps Script で testHotelDbV2ApprovedDuplicateConsolidationUiTest() を実行します。',
    'テスト関数が承認・残す行指定・元データ変更を安全なテストデータ内だけで作成します。'
  ].join('\n'));
  return { prepared:true, sourceSheetId:source.getSheetId(), duplicateRows:candidates.length, belowRow:9 };
}

function testHotelDbV2ApprovedDuplicateConsolidationUiTest() {
  const ss = hotelDbV2Pr20AssertCopy_();
  const source = ss.getSheetByName(HOTEL_DB_V2_PR20_UI_TEST.SOURCE);
  const duplicates = ss.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES);
  if (!source || !duplicates) throw new Error('setupを先に実行してください。');

  const map = hotelDbV2Pr20EnsureDuplicateHeaders_(duplicates);
  const snapshotStates = duplicates.getRange(2, map['整理スナップショット状態'], 5, 1).getDisplayValues();
  if (snapshotStates[0][0] !== '作成済み' || snapshotStates[1][0] !== '作成済み' ||
      snapshotStates[2][0] !== '作成済み' || snapshotStates[3][0] !== '作成済み' ||
      snapshotStates[4][0] !== '対象外') {
    throw new Error('⑩のスナップショット結果が期待値と一致しません。');
  }

  // 正常、情報損失、元データ変更、別部屋の4件を意図的に承認。
  [2,4,5,6].forEach(function(rowNumber) {
    duplicates.getRange(rowNumber, map['状態']).setValue('承認');
  });
  duplicates.getRange(2, map['残す行']).setValue(2);
  duplicates.getRange(4, map['残す行']).setValue(2);
  duplicates.getRange(5, map['残す行']).setValue(2);
  duplicates.getRange(6, map['残す行']).setValue(7);

  // 候補作成後変更を再現。行6（候補4の行2）の備考だけ変更。
  source.getRange(6,6).setValue('候補作成後に変更');

  const beforeBelow = source.getRange(9,4).getDisplayValue();
  const result = hotelDbV2ApplyApprovedDuplicateConsolidations_({
    spreadsheet:ss, duplicateSheet:duplicates
  });
  const afterBelow = source.getRange(9,4).getDisplayValue();

  if (result.approved !== 4 || result.applied !== 1 || result.conflicts !== 3 || result.errors !== 0) {
    throw new Error('処理件数が期待値と一致しません: ' + JSON.stringify(result));
  }
  if (!hotelDbV2Pr19RowIsEmpty_(source, 3)) throw new Error('正常重複の除外行3が空になっていません。');
  if (!source.getRange(2,4).getDisplayValue()) throw new Error('残す行2が消えています。');
  if (!source.getRange(5,6).getDisplayValue()) throw new Error('情報損失ケースの行5が誤って除外されています。');
  if (source.getRange(6,6).getDisplayValue() !== '候補作成後に変更') throw new Error('変更後行6が誤って除外されています。');
  if (!source.getRange(7,4).getDisplayValue() || !source.getRange(8,4).getDisplayValue()) throw new Error('別部屋行が誤って除外されています。');
  if (beforeBelow !== afterBelow || afterBelow !== '下行番号保持テストホテル') throw new Error('下の行番号が変化しました。');

  const archive = ss.getSheetByName(HOTEL_DB_V2_PR20_ARCHIVE_SHEET_NAME);
  if (!archive || archive.getLastRow() !== 2) throw new Error('重複整理履歴が正常1件だけ作成されていません。');

  SpreadsheetApp.getUi().alert([
    'PR #20 UIテスト実行完了','',
    '承認対象: ' + result.approved,
    '整理済み: ' + result.applied,
    '要再確認: ' + result.conflicts,
    'エラー: ' + result.errors,
    '情報損失ケース保護: 成功',
    '別部屋保護: 成功',
    '下行番号保持: 成功'
  ].join('\n'));
  return result;
}

function cleanupHotelDbV2ApprovedDuplicateConsolidationUiTest() {
  const ss = hotelDbV2Pr20AssertCopy_();
  [
    HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES,
    HOTEL_DB_V2_PR20_ARCHIVE_SHEET_NAME,
    HOTEL_DB_V2_CONFIG.SHEETS.HISTORY,
    HOTEL_DB_V2_PR20_UI_TEST.SOURCE
  ].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });
  hotelDbV2Pr20RestoreBackup_(ss, HOTEL_DB_V2_PR20_UI_TEST.BACKUP_DUPLICATES, HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES);
  hotelDbV2Pr20RestoreBackup_(ss, HOTEL_DB_V2_PR20_UI_TEST.BACKUP_ARCHIVE, HOTEL_DB_V2_PR20_ARCHIVE_SHEET_NAME);
  hotelDbV2Pr20RestoreBackup_(ss, HOTEL_DB_V2_PR20_UI_TEST.BACKUP_HISTORY, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
  SpreadsheetApp.getUi().alert('PR #20 UIテスト復元完了\n\n通常シートを復元しました。');
  return { cleaned:true };
}

function hotelDbV2Pr20BackupSheet_(ss, original, backup) {
  const sheet = ss.getSheetByName(original);
  if (sheet) sheet.setName(backup);
}

function hotelDbV2Pr20RestoreBackup_(ss, backup, original) {
  const sheet = ss.getSheetByName(backup);
  if (sheet) sheet.setName(original);
}

/**
 * PR #24 コピー版UIテスト。
 * 対象: 宿泊施設DB_PR13_⑧反映テスト
 *
 * 専用テストフォルダ内に実バックアップ・復元候補を作成して検証する。
 * 元スプレッドシートは変更しない。cleanupではテストで作ったDrive項目だけをゴミ箱へ移す。
 */

const HOTEL_DB_V2_PR24_UI_TEST = Object.freeze({
  STATE_KEY: 'HOTEL_DB_V2_PR24_BACKUP_UI_STATE',
  FOLDER_PREFIX: 'PR24_バックアップ復元テスト_'
});

function setupHotelDbV2BackupRestoreUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2BackupUiAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  if (store.getProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY)) {
    throw new Error('前回のPR #24 UIテスト状態が残っています。先に cleanupHotelDbV2BackupRestoreUiTest() を実行してください。');
  }

  const protectedSnapshot = hotelDbV2HealthCheckSnapshot_(ss);
  const sourceFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(ss);
  const testFolder = DriveApp.createFolder(HOTEL_DB_V2_PR24_UI_TEST.FOLDER_PREFIX + hotelDbV2BackupTimestamp_(new Date()));
  let backup = null;
  try {
    backup = hotelDbV2BackupCreateInFolder_(ss, testFolder, {nameSuffix:'PR24_TEST'});
  } catch (error) {
    try { testFolder.setTrashed(true); } catch (ignore) {}
    throw error;
  }

  store.setProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY, JSON.stringify({
    spreadsheetId:ss.getId(), folderId:testFolder.getId(), backupId:backup.id, recoveryId:'',
    protectedSnapshot:protectedSnapshot, sourceFingerprint:sourceFingerprint.fingerprint,
    createdAt:new Date().toISOString()
  }));

  SpreadsheetApp.getUi().alert([
    'PR #24 UIテスト 準備完了', '',
    '対象: ' + ss.getName(),
    'テスト専用Driveフォルダ: 作成済み',
    '検証済みバックアップ: 1件作成済み',
    '元スプレッドシートの変更: なし',
    'APIキー本体の保存: なし', '',
    '次に testHotelDbV2BackupRestoreUiTest() を実行してください。'
  ].join('\n'));
  return {ready:true, backupId:backup.id, folderId:testFolder.getId()};
}

function testHotelDbV2BackupRestoreUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2BackupUiAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const raw = store.getProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY);
  if (!raw) throw new Error('PR #24 UIテスト状態がありません。先に setupHotelDbV2BackupRestoreUiTest() を実行してください。');
  const state = JSON.parse(raw);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。元のコピー版で続行してください。');
  if (state.recoveryId) throw new Error('復元候補はすでに作成済みです。先にcleanupしてから再実行してください。');

  const failures = [];
  function check(label, condition) { if (!condition) failures.push(label); }
  const folder = DriveApp.getFolderById(state.folderId);
  const backupFile = hotelDbV2BackupFindFileInFolder_(folder, state.backupId);
  check('テストバックアップが見つかりません。', Boolean(backupFile));

  const metadata = backupFile ? hotelDbV2BackupReadMetadata_(backupFile) : null;
  const validation = hotelDbV2BackupValidateMetadata_(metadata, ss.getId());
  check('バックアップメタデータが不正です。', validation.valid === true);

  const backupFingerprint = backupFile ? hotelDbV2BackupFingerprintSpreadsheet_(SpreadsheetApp.openById(backupFile.getId())) : null;
  check('バックアップの指紋がメタデータと一致しません。', Boolean(backupFingerprint && metadata && backupFingerprint.fingerprint === metadata.fingerprint));

  const recovery = backupFile ? hotelDbV2BackupCreateRecoveryFromFile_(ss, backupFile, folder, {nameSuffix:'PR24_TEST'}) : null;
  check('復元候補が作成されません。', Boolean(recovery && recovery.verified));
  check('復元候補が本番上書きを示しています。', Boolean(recovery && recovery.currentFileOverwritten === false));
  if (recovery) {
    state.recoveryId = recovery.id;
    store.setProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY, JSON.stringify(state));
  }

  const recoveryFingerprint = recovery ? hotelDbV2BackupFingerprintSpreadsheet_(SpreadsheetApp.openById(recovery.id)) : null;
  check('復元候補の指紋がバックアップと一致しません。', Boolean(recoveryFingerprint && metadata && recoveryFingerprint.fingerprint === metadata.fingerprint));

  const currentFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(ss);
  check('元スプレッドシートの内容がsetup時から変化しています。', currentFingerprint.fingerprint === state.sourceFingerprint);
  const currentProtected = hotelDbV2HealthCheckSnapshot_(ss);
  check('運用シートまたはAPIキー指紋がsetup時から変化しています。', hotelDbV2HealthCheckSnapshotsEqual_(state.protectedSnapshot, currentProtected));

  const realKey = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  const safePayload = JSON.stringify({metadata:metadata, recovery:recovery});
  check('APIキー本体がバックアップ管理情報へ露出しています。', !realKey || safePayload.indexOf(realKey) === -1);

  const html = HtmlService.createHtmlOutputFromFile(HOTEL_DB_V2_BACKUP.DIALOG_FILE).getContent();
  check('バックアップ・復元画面を読み込めません。', html.indexOf('バックアップ・復元') !== -1);
  check('上書きしない説明がありません。', html.indexOf('上書きしません') !== -1 || html.indexOf('上書きせず') !== -1);
  check('バックアップ作成ボタンがありません。', html.indexOf('バックアップを作成') !== -1);
  check('復元候補作成ボタンがありません。', html.indexOf('復元候補を作成') !== -1);
  check('Google Places APIを呼ぶコードがHTMLにあります。', html.indexOf('UrlFetchApp') === -1);

  const entries = hotelDbV2HealthCheckEntryPoints_();
  const regressions = hotelDbV2HealthCheckRegressionSuites_();
  check('ヘルスチェック機能入口22件になっていません。', entries.length === 22);
  check('ヘルスチェック安全回帰8件になっていません。', regressions.length === 8);

  if (failures.length) {
    SpreadsheetApp.getUi().alert('PR #24 UIテスト 失敗\n\n' + failures.join('\n'));
    throw new Error('PR #24 UIテストで失敗があります: ' + failures.join(' / '));
  }

  SpreadsheetApp.getUi().alert([
    'PR #24 UIテスト 成功', '',
    '実バックアップ作成: 成功',
    'バックアップ指紋検証: 成功',
    '復元候補作成: 成功',
    '復元候補指紋検証: 成功',
    '現在の本番ファイル上書き: なし',
    '元スプレッドシート内容: 変更なし',
    '運用シート内容・数式: 変更なし',
    'Script Properties APIキー: 変更なし',
    'APIキー本体の管理情報への露出: なし',
    'Google Places API呼出: なし', '',
    '次にメニュー「💾 バックアップ・復元」を開いて画面を目視確認してください。'
  ].join('\n'));
  return {success:true, backupId:state.backupId, recoveryId:state.recoveryId};
}

function cleanupHotelDbV2BackupRestoreUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2BackupUiAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const raw = store.getProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY);
  if (!raw) {
    SpreadsheetApp.getUi().alert('PR #24 UIテスト cleanup\n\n一時状態は残っていません。');
    return {cleaned:false, alreadyClean:true};
  }
  const state = JSON.parse(raw);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。元のコピー版でcleanupしてください。');

  const currentFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(ss);
  if (currentFingerprint.fingerprint !== state.sourceFingerprint) {
    throw new Error('cleanup前確認で元スプレッドシートの内容変化を検出しました。テスト項目は削除せず停止します。');
  }
  if (!hotelDbV2HealthCheckSnapshotsEqual_(state.protectedSnapshot, hotelDbV2HealthCheckSnapshot_(ss))) {
    throw new Error('cleanup前確認で運用シートまたはAPIキーの状態変化を検出しました。テスト項目は削除せず停止します。');
  }

  const cleanupErrors = [];
  [state.recoveryId, state.backupId].filter(Boolean).forEach(function(id) {
    try { DriveApp.getFileById(id).setTrashed(true); }
    catch (error) { cleanupErrors.push('file ' + id + ': ' + error.message); }
  });
  try { DriveApp.getFolderById(state.folderId).setTrashed(true); }
  catch (error) { cleanupErrors.push('folder ' + state.folderId + ': ' + error.message); }

  if (cleanupErrors.length) throw new Error('テストDrive項目のcleanupに失敗しました。一時状態は保持します: ' + cleanupErrors.join(' / '));

  store.deleteProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY);
  SpreadsheetApp.getUi().alert([
    'PR #24 UIテスト復元完了', '',
    'テスト用バックアップ・復元候補・フォルダをゴミ箱へ移しました。',
    '既存の元スプレッドシートは変更していません。',
    '既存運用シートは変更していません。',
    'Script PropertiesのAPIキーも変更していません。'
  ].join('\n'));
  return {cleaned:true};
}

function hotelDbV2BackupUiAssertCopy_(ss) {
  const name = ss && ss.getName ? String(ss.getName()) : '';
  if (name.indexOf('PR13') === -1 || name.indexOf('⑧反映テスト') === -1) {
    throw new Error('PR #24 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」専用です。現在: ' + name);
  }
}

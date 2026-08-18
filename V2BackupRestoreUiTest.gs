/**
 * PR #24 コピー版UIテスト。
 * 対象: 宿泊施設DB_PR13_⑧反映テスト
 *
 * 専用テストフォルダ内に実バックアップ・復元候補を作成して検証する。
 * 元スプレッドシートは変更しない。cleanupではテストで作ったDrive項目だけをゴミ箱へ移す。
 *
 * タイムアウト安全性:
 * - テストフォルダ作成直後にUser PropertiesへfolderIdを先行保存する。
 * - setup途中で6分上限などにより停止しても、次回cleanupでテストフォルダを特定できる。
 * - setupでは元DBの重複フィンガープリント走査を行わない。
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

  console.log('PR24 UI setup 1/3: test folder creating.');
  const testFolder = DriveApp.createFolder(
    HOTEL_DB_V2_PR24_UI_TEST.FOLDER_PREFIX + hotelDbV2BackupTimestamp_(new Date())
  );
  const state = {
    spreadsheetId: ss.getId(),
    folderId: testFolder.getId(),
    backupId: '',
    recoveryId: '',
    sourceFingerprint: '',
    apiKeyFingerprint: hotelDbV2BackupUiApiKeyFingerprint_(),
    createdAt: new Date().toISOString(),
    stage: 'folder-created'
  };

  // 重要: Driveコピーより先に保存する。以後タイムアウトしてもcleanupがfolderIdを追跡できる。
  store.setProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY, JSON.stringify(state));
  console.log('PR24 UI setup 1/3: test folder checkpoint saved.');

  try {
    // createInFolder_ 自身が「元DB 1回 + バックアップ 1回」の指紋比較を行う。
    // setup側で元DBを事前にもう1回走査しないことで実行時間を削減する。
    console.log('PR24 UI setup 2/3: backup copy and fingerprint verification started.');
    const backup = hotelDbV2BackupCreateInFolder_(ss, testFolder, {nameSuffix:'PR24_TEST'});
    const backupFile = hotelDbV2BackupFindFileInFolder_(testFolder, backup.id);
    const metadata = backupFile ? hotelDbV2BackupReadMetadata_(backupFile) : null;
    const validation = hotelDbV2BackupValidateMetadata_(metadata, ss.getId());
    if (!validation.valid) throw new Error('作成済みテストバックアップのメタデータ検証に失敗しました: ' + validation.message);

    state.backupId = backup.id;
    state.sourceFingerprint = metadata.fingerprint;
    state.stage = 'backup-verified';
    store.setProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY, JSON.stringify(state));
    console.log('PR24 UI setup 3/3: verified backup checkpoint saved.');

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
  } catch (error) {
    // 状態は削除しない。folderIdを残し、cleanupで安全に片付けられるようにする。
    state.stage = 'setup-error';
    state.lastError = hotelDbV2BackupSafeErrorMessage_(error);
    try { store.setProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY, JSON.stringify(state)); } catch (ignore) {}
    throw new Error(
      'PR #24 UIテスト setup が完了しませんでした。テスト状態はcleanup用に保持しています。' +
      '先に cleanupHotelDbV2BackupRestoreUiTest() を実行してから再試行してください。\n' +
      hotelDbV2BackupSafeErrorMessage_(error)
    );
  }
}

function testHotelDbV2BackupRestoreUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2BackupUiAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const raw = store.getProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY);
  if (!raw) throw new Error('PR #24 UIテスト状態がありません。先に setupHotelDbV2BackupRestoreUiTest() を実行してください。');
  const state = JSON.parse(raw);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。元のコピー版で続行してください。');
  if (!state.backupId || !state.sourceFingerprint) {
    throw new Error('setupがバックアップ検証まで完了していません。先に cleanupHotelDbV2BackupRestoreUiTest() を実行してからsetupをやり直してください。');
  }
  if (state.recoveryId) throw new Error('復元候補はすでに作成済みです。先にcleanupしてから再実行してください。');

  const failures = [];
  function check(label, condition) { if (!condition) failures.push(label); }
  const folder = DriveApp.getFolderById(state.folderId);
  const backupFile = hotelDbV2BackupFindFileInFolder_(folder, state.backupId);
  check('テストバックアップが見つかりません。', Boolean(backupFile));

  const metadata = backupFile ? hotelDbV2BackupReadMetadata_(backupFile) : null;
  const validation = hotelDbV2BackupValidateMetadata_(metadata, ss.getId());
  check('バックアップメタデータが不正です。', validation.valid === true);
  check('setup時の元DB指紋とバックアップ記録が一致しません。', Boolean(metadata && metadata.fingerprint === state.sourceFingerprint));

  // createRecoveryFromFile_ 内部でバックアップを再指紋検証し、
  // 復元候補作成後にも復元候補を再指紋検証する。UIテスト側で同じ2走査を重複しない。
  console.log('PR24 UI test 1/2: recovery candidate copy and verification started.');
  const recovery = backupFile
    ? hotelDbV2BackupCreateRecoveryFromFile_(ss, backupFile, folder, {nameSuffix:'PR24_TEST'})
    : null;
  check('復元候補が作成されません。', Boolean(recovery && recovery.verified));
  check('復元候補が本番上書きを示しています。', Boolean(recovery && recovery.currentFileOverwritten === false));
  if (recovery) {
    state.recoveryId = recovery.id;
    state.stage = 'recovery-verified';
    store.setProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY, JSON.stringify(state));
    console.log('PR24 UI test 2/2: recovery checkpoint saved.');
  }

  check(
    'Script Properties APIキーがsetup時から変化しています。',
    hotelDbV2BackupUiApiKeyFingerprint_() === state.apiKeyFingerprint
  );

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
  check('既存ヘルスチェック機能入口21件が維持されていません。', entries.length === 21);
  check('既存ヘルスチェック安全回帰7件が維持されていません。', regressions.length === 7);

  if (failures.length) {
    SpreadsheetApp.getUi().alert('PR #24 UIテスト 失敗\n\n' + failures.join('\n'));
    throw new Error('PR #24 UIテストで失敗があります: ' + failures.join(' / '));
  }

  SpreadsheetApp.getUi().alert([
    'PR #24 UIテスト 成功', '',
    '実バックアップ作成: 成功',
    'バックアップ再検証: 成功',
    '復元候補作成: 成功',
    '復元候補指紋検証: 成功',
    '現在の本番ファイル上書き: なし',
    '元スプレッドシート内容: cleanupで最終確認予定',
    'Script Properties APIキー: 変更なし',
    'APIキー本体の管理情報への露出: なし',
    'Google Places API呼出: なし', '',
    '次にメニュー「💾 バックアップ・復元」を開いて画面を目視確認してください。',
    '目視確認後、cleanupで元スプレッドシート全体の不変を最終確認します。'
  ].join('\n'));
  return {success:true, backupId:state.backupId, recoveryId:state.recoveryId};
}

function cleanupHotelDbV2BackupRestoreUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2BackupUiAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const raw = store.getProperty(HOTEL_DB_V2_PR24_UI_TEST.STATE_KEY);
  if (!raw) {
    SpreadsheetApp.getUi().alert([
      'PR #24 UIテスト cleanup', '',
      '追跡可能な一時状態は残っていません。',
      '旧版setupがタイムアウトして作った未追跡フォルダは、この関数では自動削除しません。'
    ].join('\n'));
    return {cleaned:false, alreadyClean:true};
  }
  const state = JSON.parse(raw);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。元のコピー版でcleanupしてください。');

  // setupがバックアップ検証まで到達した場合のみ、全元DBの不変を最終確認する。
  if (state.sourceFingerprint) {
    console.log('PR24 UI cleanup 1/2: source fingerprint verification started.');
    const currentFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(ss);
    if (currentFingerprint.fingerprint !== state.sourceFingerprint) {
      throw new Error('cleanup前確認で元スプレッドシートの内容変化を検出しました。テスト項目は削除せず停止します。');
    }
  }
  if (state.apiKeyFingerprint && hotelDbV2BackupUiApiKeyFingerprint_() !== state.apiKeyFingerprint) {
    throw new Error('cleanup前確認でScript PropertiesのAPIキー状態変化を検出しました。テスト項目は削除せず停止します。');
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
  console.log('PR24 UI cleanup 2/2: tracked Drive items trashed and state cleared.');
  SpreadsheetApp.getUi().alert([
    'PR #24 UIテスト復元完了', '',
    'テスト用バックアップ・復元候補・フォルダをゴミ箱へ移しました。',
    state.sourceFingerprint ? '元スプレッドシート全体: setup時から変更なし' : '元スプレッドシート全体: setup未完了のため全体比較は省略',
    '既存運用シートは変更していません。',
    'Script PropertiesのAPIキーも変更していません。'
  ].join('\n'));
  return {cleaned:true, sourceVerified:Boolean(state.sourceFingerprint)};
}

function hotelDbV2BackupUiApiKeyFingerprint_() {
  const key = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  return hotelDbV2HealthCheckDigest_(String(key));
}

function hotelDbV2BackupUiAssertCopy_(ss) {
  const name = ss && ss.getName ? String(ss.getName()) : '';
  if (name.indexOf('PR13') === -1 || name.indexOf('⑧反映テスト') === -1) {
    throw new Error('PR #24 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」専用です。現在: ' + name);
  }
}

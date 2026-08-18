/**
 * PR #23 コピー版UIテスト。
 * 対象: 宿泊施設DB_PR13_⑧反映テスト
 *
 * スプレッドシート / Script Propertiesは変更しない。
 * テスト状態は実行ユーザーのUser Propertiesへ指紋だけ一時保存し、cleanupで削除する。
 */

const HOTEL_DB_V2_PR23_UI_TEST = Object.freeze({
  STATE_KEY: 'HOTEL_DB_V2_PR23_HEALTH_UI_STATE'
});

function setupHotelDbV2HealthCheckUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2HealthCheckAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  if (store.getProperty(HOTEL_DB_V2_PR23_UI_TEST.STATE_KEY)) {
    throw new Error('前回のPR #23 UIテスト状態が残っています。先に cleanupHotelDbV2HealthCheckUiTest() を実行してください。');
  }

  const snapshot = hotelDbV2HealthCheckSnapshot_(ss);
  store.setProperty(HOTEL_DB_V2_PR23_UI_TEST.STATE_KEY, JSON.stringify({
    spreadsheetId: snapshot.spreadsheetId,
    apiKeyFingerprint: snapshot.apiKeyFingerprint,
    sheetHashes: snapshot.sheetHashes,
    createdAt: new Date().toISOString()
  }));

  SpreadsheetApp.getUi().alert([
    'PR #23 UIテスト 準備完了', '',
    '対象: ' + ss.getName(),
    '運用シート・APIキーの指紋をUser Propertiesへ一時保存しました。',
    'APIキー本体は保存していません。', '',
    '次に testHotelDbV2HealthCheckUiTest() を実行してください。'
  ].join('\n'));
  return { ready:true, spreadsheet:ss.getName() };
}

function testHotelDbV2HealthCheckUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2HealthCheckAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const rawState = store.getProperty(HOTEL_DB_V2_PR23_UI_TEST.STATE_KEY);
  if (!rawState) throw new Error('PR #23 UIテスト状態がありません。先に setupHotelDbV2HealthCheckUiTest() を実行してください。');
  const state = JSON.parse(rawState);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。cleanup後、対象コピー版でやり直してください。');

  const failures = [];
  function check(label, condition) { if (!condition) failures.push(label); }

  const report = hotelDbV2HealthCheckBuildReport_(ss);
  const current = hotelDbV2HealthCheckSnapshot_(ss);
  const baseline = {
    spreadsheetId: state.spreadsheetId,
    apiKeyFingerprint: state.apiKeyFingerprint,
    sheetHashes: state.sheetHashes
  };
  const html = HtmlService.createHtmlOutputFromFile(HOTEL_DB_V2_HEALTH.DIALOG_FILE).getContent();
  const realKey = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  const reportText = JSON.stringify(report);

  check('レポートが生成されません。', Boolean(report));
  check('全体状態が不正です。', ['正常','要確認','重大'].indexOf(report.overall) !== -1);
  check('主要セクションが不足しています。', Array.isArray(report.sections) && report.sections.length >= 5);
  check('機能入口の確認件数が不足しています。', report.checks.filter(function(item){ return item.section === '機能入口'; }).length === 22);
  check('安全回帰スイートの確認件数が不足しています。', report.checks.filter(function(item){ return item.section === '安全回帰'; }).length === 8);
  check('外部API呼出フラグがtrueです。', report.externalApiCalled === false);
  check('運用書込みフラグがtrueです。', report.operationalWrites === false);
  check('APIキー値が結果へ含まれています。', report.apiKeyValue === null && report.apiKeyHash === null);
  check('実APIキー文字列が結果へ露出しています。', !realKey || reportText.indexOf(realKey) === -1);
  check('安全マーカーが不正です。', report.safetyMarker === HOTEL_DB_V2_HEALTH.SAFETY_MARKER);
  check('設定画面連携が確認できません。', report.checks.some(function(item){ return item.id === 'SETUP-API'; }));
  check('安全設定75点が確認できません。', report.checks.some(function(item){ return item.id === 'CFG-AUTO-SCORE' && item.severity === 'ok'; }));
  check('バッチ50件が確認できません。', report.checks.some(function(item){ return item.id === 'CFG-BATCH' && item.severity === 'ok'; }));
  check('診断前後の不変確認が成功していません。', report.checks.some(function(item){ return item.id === 'SAFE-UNCHANGED' && item.severity === 'ok'; }));
  check('運用シート/APIキーの指紋がsetup時から変化しています。', hotelDbV2HealthCheckSnapshotsEqual_(baseline, current));
  check('HTML画面を読み込めません。', html.indexOf('製品全体ヘルスチェック') !== -1);
  check('HTMLに再診断ボタンがありません。', html.indexOf('再診断') !== -1);
  check('HTMLが外部APIを直接呼ぶ設計です。', html.indexOf('UrlFetchApp') === -1);

  if (failures.length) {
    SpreadsheetApp.getUi().alert('PR #23 UIテスト 失敗\n\n' + failures.join('\n'));
    throw new Error('PR #23 UIテストで失敗があります: ' + failures.join(' / '));
  }

  SpreadsheetApp.getUi().alert([
    'PR #23 UIテスト 成功', '',
    'レポート生成: 正常',
    '機能入口22件: 正常',
    '安全回帰8件: 正常',
    '安全設定: 正常',
    'APIキー本体の非露出: 成功',
    '運用シート内容・数式: 変更なし',
    'Script Properties APIキー: 変更なし',
    'Google Places API呼出: なし', '',
    '次にメニュー「🩺 製品全体ヘルスチェック」を開いて目視確認してください。'
  ].join('\n'));

  return { success:true, overall:report.overall, checks:report.counts.total, externalApi:false, operationalWrites:false };
}

function cleanupHotelDbV2HealthCheckUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2HealthCheckAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const rawState = store.getProperty(HOTEL_DB_V2_PR23_UI_TEST.STATE_KEY);
  if (!rawState) {
    SpreadsheetApp.getUi().alert('PR #23 UIテスト cleanup\n\n一時状態は残っていません。');
    return { cleaned:false, alreadyClean:true };
  }

  const state = JSON.parse(rawState);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。元のコピー版でcleanupしてください。');

  const current = hotelDbV2HealthCheckSnapshot_(ss);
  const baseline = {
    spreadsheetId: state.spreadsheetId,
    apiKeyFingerprint: state.apiKeyFingerprint,
    sheetHashes: state.sheetHashes
  };
  if (!hotelDbV2HealthCheckSnapshotsEqual_(baseline, current)) {
    throw new Error('cleanup前確認で運用シートまたはAPIキーの状態変化を検出しました。一時状態は削除せず停止します。');
  }

  store.deleteProperty(HOTEL_DB_V2_PR23_UI_TEST.STATE_KEY);
  SpreadsheetApp.getUi().alert([
    'PR #23 UIテスト復元完了', '',
    'テスト用User Propertiesを削除しました。',
    '既存運用シートは変更していません。',
    'Script PropertiesのAPIキーも変更していません。'
  ].join('\n'));
  return { cleaned:true };
}

function hotelDbV2HealthCheckAssertCopy_(ss) {
  const name = ss && ss.getName ? String(ss.getName()) : '';
  if (name.indexOf('PR13') === -1 || name.indexOf('⑧反映テスト') === -1) {
    throw new Error('PR #23 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」専用です。現在: ' + name);
  }
}

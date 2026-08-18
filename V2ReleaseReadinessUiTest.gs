/**
 * PR #25 コピー版UIテスト。
 * 対象: 宿泊施設DB_PR13_⑧反映テスト
 *
 * 読み取り専用。User PropertiesにはAPIキー本体を含まない指紋だけを一時保存する。
 */
const HOTEL_DB_V2_PR25_UI_TEST = Object.freeze({
  STATE_KEY: 'HOTEL_DB_V2_PR25_RELEASE_UI_STATE'
});

function setupHotelDbV2ReleaseReadinessUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2ReleaseAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  if (store.getProperty(HOTEL_DB_V2_PR25_UI_TEST.STATE_KEY)) {
    throw new Error('前回のPR #25 UIテスト状態が残っています。先に cleanupHotelDbV2ReleaseReadinessUiTest() を実行してください。');
  }

  const snapshot = hotelDbV2ReleaseSnapshot_(ss);
  store.setProperty(HOTEL_DB_V2_PR25_UI_TEST.STATE_KEY, JSON.stringify({
    spreadsheetId: snapshot.spreadsheetId,
    activeSheetId: snapshot.activeSheetId,
    activeSheetHash: snapshot.activeSheetHash,
    apiKeyFingerprint: snapshot.apiKeyFingerprint,
    createdAt: new Date().toISOString()
  }));

  SpreadsheetApp.getUi().alert([
    'PR #25 UIテスト 準備完了', '', '対象: ' + ss.getName(),
    '元データ・運用シート・APIキーの指紋をUser Propertiesへ一時保存しました。',
    'APIキー本体は保存していません。', 'Driveへの書込みはありません。', '',
    '次に testHotelDbV2ReleaseReadinessUiTest() を実行してください。'
  ].join('\n'));
  return {ready:true, spreadsheet:ss.getName()};
}

function testHotelDbV2ReleaseReadinessUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2ReleaseAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const raw = store.getProperty(HOTEL_DB_V2_PR25_UI_TEST.STATE_KEY);
  if (!raw) throw new Error('PR #25 UIテスト状態がありません。先に setupHotelDbV2ReleaseReadinessUiTest() を実行してください。');
  const state = JSON.parse(raw);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。cleanup後、対象コピー版でやり直してください。');

  const failures = [];
  function check(label, condition) { if (!condition) failures.push(label); }

  const report = hotelDbV2ReleaseBuildReport_(ss);
  const current = hotelDbV2ReleaseSnapshot_(ss);
  const baseline = {
    spreadsheetId:state.spreadsheetId, activeSheetId:state.activeSheetId,
    activeSheetHash:state.activeSheetHash, apiKeyFingerprint:state.apiKeyFingerprint
  };
  const html = HtmlService.createHtmlOutputFromFile(HOTEL_DB_V2_RELEASE.DIALOG_FILE).getContent();
  const realKey = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  const text = JSON.stringify(report);

  check('レポートを生成できません。', Boolean(report));
  check('コピー版なのにテスト環境判定ではありません。', report.environment && report.environment.kind === 'test');
  check('コピー版なのにリリース可能になっています。', report.releaseReady === false);
  check('全体状態が不正です。', ['リリース可能','要対応','テスト環境'].indexOf(report.overall) !== -1);
  check('チェック項目が不足しています。', Array.isArray(report.checks) && report.checks.length >= 15);
  check('Google Places API呼出フラグがtrueです。', report.externalApiCalled === false);
  check('バックアップ自動作成フラグがtrueです。', report.backupCreated === false);
  check('運用書込みフラグがtrueです。', report.operationalWrites === false);
  check('APIキー値・ハッシュが結果へ含まれています。', report.apiKeyValue === null && report.apiKeyHash === null);
  check('実APIキーが結果へ露出しています。', !realKey || text.indexOf(realKey) === -1);
  check('安全マーカーが不正です。', report.safetyMarker === HOTEL_DB_V2_RELEASE.SAFETY_MARKER);
  check('バックアップ安全方式を確認できません。', report.checks.some(function(item){ return item.id === 'BACKUP-NO-INPLACE' && item.severity === 'ok'; }));
  check('リリース準備入口を確認できません。', report.checks.some(function(item){ return item.id === 'ENTRY-RELEASE' && item.severity === 'ok'; }));
  check('データ不変確認が成功していません。', report.checks.some(function(item){ return item.id === 'SAFE-UNCHANGED' && item.severity === 'ok'; }));
  check('setup時から元データ・運用シート・APIキー状態が変化しています。', hotelDbV2ReleaseSnapshotsEqual_(baseline, current));
  check('HTMLタイトルがありません。', html.indexOf('リリース準備チェック') !== -1);
  check('HTMLに再チェックボタンがありません。', html.indexOf('再チェック') !== -1);
  check('HTMLに検証済みバックアップ案内がありません。', html.indexOf('検証済みバックアップ') !== -1);
  check('HTMLがGoogle Places APIを直接呼ぶ設計です。', html.indexOf('UrlFetchApp') === -1);

  if (failures.length) {
    SpreadsheetApp.getUi().alert('PR #25 UIテスト 失敗\n\n' + failures.join('\n'));
    throw new Error('PR #25 UIテストで失敗があります: ' + failures.join(' / '));
  }

  SpreadsheetApp.getUi().alert([
    'PR #25 UIテスト 成功', '', 'レポート生成: 正常', 'コピー版テスト環境判定: 正常',
    'リリース可能の誤判定: なし', 'バックアップ安全方式: 正常', 'APIキー本体の非露出: 成功',
    '元データ・運用シート内容: 変更なし', 'Script Properties APIキー: 変更なし',
    'Drive書込み: なし', 'Google Places API呼出: なし', '',
    '次にメニュー「🚀 リリース準備チェック」を開いて目視確認してください。'
  ].join('\n'));

  return {success:true, overall:report.overall, blockers:report.blockingCount, environment:report.environment.kind,
    externalApi:false, driveWrites:false, operationalWrites:false};
}

function cleanupHotelDbV2ReleaseReadinessUiTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2ReleaseAssertCopy_(ss);
  const store = PropertiesService.getUserProperties();
  const raw = store.getProperty(HOTEL_DB_V2_PR25_UI_TEST.STATE_KEY);
  if (!raw) {
    SpreadsheetApp.getUi().alert('PR #25 UIテスト cleanup\n\n一時状態は残っていません。');
    return {cleaned:false, alreadyClean:true};
  }
  const state = JSON.parse(raw);
  if (state.spreadsheetId !== ss.getId()) throw new Error('setup時と異なるスプレッドシートです。元のコピー版でcleanupしてください。');

  const current = hotelDbV2ReleaseSnapshot_(ss);
  const baseline = {
    spreadsheetId:state.spreadsheetId, activeSheetId:state.activeSheetId,
    activeSheetHash:state.activeSheetHash, apiKeyFingerprint:state.apiKeyFingerprint
  };
  if (!hotelDbV2ReleaseSnapshotsEqual_(baseline, current)) {
    throw new Error('cleanup前確認で元データ・運用シートまたはAPIキーの状態変化を検出しました。一時状態は削除せず停止します。');
  }

  store.deleteProperty(HOTEL_DB_V2_PR25_UI_TEST.STATE_KEY);
  SpreadsheetApp.getUi().alert([
    'PR #25 UIテスト復元完了', '', 'テスト用User Propertiesを削除しました。',
    '元データ・既存運用シートは変更していません。', 'Script PropertiesのAPIキーも変更していません。',
    'Driveへの書込みはありません。'
  ].join('\n'));
  return {cleaned:true};
}

function hotelDbV2ReleaseAssertCopy_(ss) {
  const name = ss && ss.getName ? String(ss.getName()) : '';
  if (name.indexOf('PR13') === -1 || name.indexOf('⑧反映テスト') === -1) {
    throw new Error('PR #25 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」専用です。現在: ' + name);
  }
}

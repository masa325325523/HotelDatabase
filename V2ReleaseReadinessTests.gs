/**
 * PR #25 リリース準備チェック 自己診断。
 * 外部API・Drive書込み・元DB書込みを行わない。
 */
function runHotelDbV2ReleaseReadinessTests() {
  const failures = [];
  let passed = 0;

  function check(label, condition, detail) {
    if (condition) passed++;
    else failures.push(label + (detail ? ': ' + detail : ''));
  }

  check('CONST-01 version', HOTEL_DB_V2_RELEASE.VERSION === '2.0');
  check('CONST-02 dialog file', HOTEL_DB_V2_RELEASE.DIALOG_FILE === 'V2ReleaseReadinessDialog');
  check('CONST-03 read only', HOTEL_DB_V2_RELEASE.SAFETY_MARKER.indexOf('read-only') !== -1);
  check('CONST-04 no places api', HOTEL_DB_V2_RELEASE.SAFETY_MARKER.indexOf('no-places-api') !== -1);
  check('CONST-05 no backup create', HOTEL_DB_V2_RELEASE.SAFETY_MARKER.indexOf('no-backup-create') !== -1);
  check('CONST-06 api hidden', HOTEL_DB_V2_RELEASE.SAFETY_MARKER.indexOf('api-key-not-returned') !== -1);
  check('CONST-07 release gate', HOTEL_DB_V2_RELEASE.SAFETY_MARKER.indexOf('release-gate') !== -1);

  check('ENV-01 production', hotelDbV2ReleaseEnvironment_('宿泊施設DB 本番').kind === 'production');
  check('ENV-02 pr test', hotelDbV2ReleaseEnvironment_('宿泊施設DB_PR13_⑧反映テスト').kind === 'test');
  check('ENV-03 copy', hotelDbV2ReleaseEnvironment_('宿泊施設DB コピー').kind === 'test');
  check('ENV-04 recovery', hotelDbV2ReleaseEnvironment_('【復元候補】宿泊施設DB').kind === 'test');
  check('ENV-05 backup', hotelDbV2ReleaseEnvironment_('【宿泊施設DBバックアップ】本番').kind === 'test');
  check('ENV-06 label production', hotelDbV2ReleaseEnvironment_('宿泊施設DB 本番').label === '製品候補');
  check('ENV-07 label test', hotelDbV2ReleaseEnvironment_('全国宿泊施設 作業中').label === 'テスト環境');

  let summary = hotelDbV2ReleaseSummarize_([{severity:'ok', blocking:false},{severity:'info', blocking:false}], {kind:'production'});
  check('SUM-01 ready', summary.releaseReady === true);
  check('SUM-02 overall ready', summary.overall === 'リリース可能');
  check('SUM-03 total2', summary.counts.total === 2);
  check('SUM-04 info1', summary.counts.info === 1);
  check('SUM-05 blockers0', summary.blockingCount === 0);

  summary = hotelDbV2ReleaseSummarize_([{severity:'ok', blocking:false},{severity:'critical', blocking:true}], {kind:'production'});
  check('SUM-06 blocked', summary.releaseReady === false);
  check('SUM-07 overall action', summary.overall === '要対応');
  check('SUM-08 critical1', summary.counts.critical === 1);
  check('SUM-09 blockers1', summary.blockingCount === 1);

  summary = hotelDbV2ReleaseSummarize_([{severity:'critical', blocking:true}], {kind:'test'});
  check('SUM-10 test not ready', summary.releaseReady === false);
  check('SUM-11 test overall', summary.overall === 'テスト環境');

  check('ACT-01 setup', hotelDbV2ReleaseActionFor_('SETUP-API','').indexOf('初期セットアップ') !== -1);
  check('ACT-02 health', hotelDbV2ReleaseActionFor_('HEALTH-CRITICAL','').indexOf('ヘルスチェック') !== -1);
  check('ACT-03 backup', hotelDbV2ReleaseActionFor_('BACKUP-EXISTS','').indexOf('バックアップ') !== -1);
  check('ACT-04 residue', hotelDbV2ReleaseActionFor_('TEST-RESIDUE','').indexOf('cleanup') !== -1);
  check('ACT-05 entry', hotelDbV2ReleaseActionFor_('ENTRY-RELEASE','').indexOf('最新main') !== -1);
  check('ACT-06 unchanged', hotelDbV2ReleaseActionFor_('SAFE-UNCHANGED','').indexOf('他の処理') !== -1);
  check('ACT-07 fallback', hotelDbV2ReleaseActionFor_('OTHER','abc') === 'abc');

  const recs = hotelDbV2ReleaseRecommendations_([
    {id:'BACKUP-EXISTS', label:'Backup', severity:'critical', detail:'', blocking:true},
    {id:'HEALTH-WARNING', label:'Warn', severity:'warning', detail:'', blocking:false},
    {id:'OK', label:'OK', severity:'ok', detail:'', blocking:false}
  ], {kind:'production'});
  check('ACT-08 rec count2', recs.length === 2);
  check('ACT-09 rec backup', recs[0].label === 'Backup');
  check('ACT-10 rec warning', recs[1].label === 'Warn');
  const testRecs = hotelDbV2ReleaseRecommendations_([], {kind:'test'});
  check('ACT-11 test recommendation', testRecs.length === 1);
  check('ACT-12 test recommendation label', testRecs[0].label.indexOf('製品候補') !== -1);

  check('ENTRY-01 setup', typeof runHotelDbV2OpenSetup === 'function');
  check('ENTRY-02 health', typeof runHotelDbV2HealthCheck === 'function');
  check('ENTRY-03 backup', typeof runHotelDbV2OpenBackupRestore === 'function');
  check('ENTRY-04 release', typeof runHotelDbV2OpenReleaseReadiness === 'function');
  const healthEntries = hotelDbV2HealthCheckEntryPoints_();
  const healthRegressions = hotelDbV2HealthCheckRegressionSuites_();
  check('HEALTH-01 existing entries remain21', healthEntries.length === 21, 'actual=' + healthEntries.length);
  check('HEALTH-02 existing regressions remain7', healthRegressions.length === 7, 'actual=' + healthRegressions.length);
  check('HEALTH-03 health selftest exists', typeof runHotelDbV2HealthCheckTests === 'function');
  check('BACKUP-01 backup selftest exists', typeof runHotelDbV2BackupRestoreTests === 'function');

  const html = HtmlService.createHtmlOutputFromFile(HOTEL_DB_V2_RELEASE.DIALOG_FILE).getContent();
  check('HTML-01 title', html.indexOf('リリース準備チェック') !== -1);
  check('HTML-02 refresh', html.indexOf('再チェック') !== -1);
  check('HTML-03 backup wording', html.indexOf('検証済みバックアップ') !== -1);
  check('HTML-04 no UrlFetchApp', html.indexOf('UrlFetchApp') === -1);
  check('HTML-05 no raw key placeholder', html.indexOf('GOOGLE_PLACES_API_KEY=') === -1);
  check('HTML-06 no write api', html.indexOf('hotelDbV2BackupCreate(') === -1);
  check('HTML-07 no save api key', html.indexOf('hotelDbV2SetupSaveApiKey') === -1);

  const s1 = {spreadsheetId:'S', activeSheetId:1, activeSheetHash:'A', apiKeyFingerprint:'K'};
  const s2 = {spreadsheetId:'S', activeSheetId:1, activeSheetHash:'A', apiKeyFingerprint:'K'};
  const s3 = {spreadsheetId:'S', activeSheetId:1, activeSheetHash:'B', apiKeyFingerprint:'K'};
  const s4 = {spreadsheetId:'S', activeSheetId:2, activeSheetHash:'A', apiKeyFingerprint:'K'};
  check('SNAP-01 equal', hotelDbV2ReleaseSnapshotsEqual_(s1, s2) === true);
  check('SNAP-02 hash diff', hotelDbV2ReleaseSnapshotsEqual_(s1, s3) === false);
  check('SNAP-03 sheet diff', hotelDbV2ReleaseSnapshotsEqual_(s1, s4) === false);
  check('SNAP-04 null', hotelDbV2ReleaseSnapshotsEqual_(null, s2) === false);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const before = hotelDbV2ReleaseSnapshot_(ss);
  const report = hotelDbV2ReleaseBuildReport_(ss);
  const after = hotelDbV2ReleaseSnapshot_(ss);
  const reportText = JSON.stringify(report);
  const realKey = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';

  check('REAL-01 report exists', Boolean(report));
  check('REAL-02 version', report.version === '2.0');
  check('REAL-03 overall valid', ['リリース可能','要対応','テスト環境'].indexOf(report.overall) !== -1);
  check('REAL-04 checks', Array.isArray(report.checks) && report.checks.length >= 15);
  check('REAL-05 recommendations', Array.isArray(report.recommendations));
  check('REAL-06 no places api', report.externalApiCalled === false);
  check('REAL-07 no backup create', report.backupCreated === false);
  check('REAL-08 no operational writes', report.operationalWrites === false);
  check('REAL-09 api value null', report.apiKeyValue === null);
  check('REAL-10 api hash null', report.apiKeyHash === null);
  check('REAL-11 raw key absent', !realKey || reportText.indexOf(realKey) === -1);
  check('REAL-12 snapshot unchanged', hotelDbV2ReleaseSnapshotsEqual_(before, after) === true);
  check('REAL-13 safety marker', report.safetyMarker === HOTEL_DB_V2_RELEASE.SAFETY_MARKER);
  check('REAL-14 reconciles', report.counts.total === report.checks.length);
  check('REAL-15 environment test on copy', report.environment.kind === 'test');

  const message = [
    failures.length ? 'PR #25 自己診断 失敗' : 'PR #25 自己診断 成功', '',
    '成功件数: ' + passed + '件', '失敗件数: ' + failures.length + '件',
    'Google Places API呼出: なし', 'Drive書込み: なし', 'Script Properties変更: なし', '元DB・候補・履歴の変更: なし'
  ];
  if (failures.length) message.push('', '失敗:', failures.join('\n'));
  try { SpreadsheetApp.getUi().alert(message.join('\n')); } catch (ignore) {}
  if (failures.length) throw new Error('PR #25 自己診断で失敗があります: ' + failures.join(' / '));
  return {success:passed, failed:0, externalApi:false, driveWrites:false, operationalWrites:false};
}

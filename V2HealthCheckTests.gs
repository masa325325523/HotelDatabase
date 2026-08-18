/**
 * PR #23 製品全体ヘルスチェック 自己診断。
 * 外部APIを呼ばず、元DB・候補・履歴・Script Propertiesを書き換えない。
 */
function runHotelDbV2HealthCheckTests() {
  const failures = [];
  let passed = 0;

  function check(label, condition, detail) {
    if (condition) passed++;
    else failures.push(label + (detail ? ': ' + detail : ''));
  }

  // 定数・安全契約
  check('CONST-01 version', HOTEL_DB_V2_HEALTH.VERSION === '2.0');
  check('CONST-02 dialog file', HOTEL_DB_V2_HEALTH.DIALOG_FILE === 'V2HealthCheckDialog');
  check('CONST-03 optional sheets 5', HOTEL_DB_V2_HEALTH.OPTIONAL_SHEETS.length === 5);
  check('CONST-04 no external api marker', HOTEL_DB_V2_HEALTH.SAFETY_MARKER.indexOf('no-external-api') !== -1);
  check('CONST-05 read only marker', HOTEL_DB_V2_HEALTH.SAFETY_MARKER.indexOf('read-only') !== -1);
  check('CONST-06 api key hidden marker', HOTEL_DB_V2_HEALTH.SAFETY_MARKER.indexOf('api-key-not-returned') !== -1);
  check('CONST-07 hash guard marker', HOTEL_DB_V2_HEALTH.SAFETY_MARKER.indexOf('operational-hash-guard') !== -1);

  // 設定の安全基準
  const configChecks = hotelDbV2HealthCheckConfigChecks_();
  check('CFG-01 8 checks', configChecks.length === 8);
  configChecks.forEach(function(item) {
    check('CFG-' + item.id, item.ok === true, item.detail);
  });
  check('CFG-10 api property', HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY === 'GOOGLE_PLACES_API_KEY');
  check('CFG-11 auto accept', HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE === 75);
  check('CFG-12 min match', HOTEL_DB_V2_CONFIG.MIN_MATCH_SCORE === 55);
  check('CFG-13 batch size', HOTEL_DB_V2_CONFIG.BATCH_SIZE === 50);
  check('CFG-14 test rows', HOTEL_DB_V2_CONFIG.TEST_ROWS === 3);
  check('CFG-15 timezone', HOTEL_DB_V2_CONFIG.TIMEZONE === 'Asia/Tokyo');

  // 全メニュー入口
  const entries = hotelDbV2HealthCheckEntryPoints_();
  check('ENTRY-00 total 21', entries.length === 21, 'actual=' + entries.length);
  entries.forEach(function(item) {
    check(item.id + ' exists', item.exists === true, item.label);
  });
  check('ENTRY-unique', new Set(entries.map(function(item){ return item.id; })).size === entries.length);

  // 安全回帰スイート入口
  const regressions = hotelDbV2HealthCheckRegressionSuites_();
  check('REG-00 total 7', regressions.length === 7, 'actual=' + regressions.length);
  regressions.forEach(function(item) {
    check(item.id + ' exists', item.exists === true, item.label);
  });
  check('REG-unique', new Set(regressions.map(function(item){ return item.id; })).size === regressions.length);

  // 集計ロジック
  let summary = hotelDbV2HealthCheckSummarize_([
    {severity:'ok'}, {severity:'ok'}, {severity:'info'}
  ]);
  check('SUM-01 normal', summary.overall === '正常');
  check('SUM-02 ok2', summary.counts.ok === 2);
  check('SUM-03 info1', summary.counts.info === 1);
  check('SUM-04 total3', summary.counts.total === 3);

  summary = hotelDbV2HealthCheckSummarize_([
    {severity:'ok'}, {severity:'warning'}
  ]);
  check('SUM-05 warning overall', summary.overall === '要確認');
  check('SUM-06 warning1', summary.counts.warning === 1);

  summary = hotelDbV2HealthCheckSummarize_([
    {severity:'warning'}, {severity:'critical'}
  ]);
  check('SUM-07 critical overall', summary.overall === '重大');
  check('SUM-08 critical1', summary.counts.critical === 1);

  summary = hotelDbV2HealthCheckSummarize_([{severity:'unknown'}]);
  check('SUM-09 unknown as warning', summary.counts.warning === 1);
  check('SUM-10 unknown total', summary.counts.total === 1);

  // 推奨アクション
  check('ACT-01 api', hotelDbV2HealthCheckActionFor_('SETUP-API','').indexOf('APIキー') !== -1);
  check('ACT-02 core', hotelDbV2HealthCheckActionFor_('CORE-修正候補','').indexOf('コア運用シート') !== -1);
  check('ACT-03 entry', hotelDbV2HealthCheckActionFor_('ENTRY-01','').indexOf('最新main') !== -1);
  check('ACT-04 config', hotelDbV2HealthCheckActionFor_('CFG-BATCH','').indexOf('安全設定') !== -1);
  check('ACT-05 regression', hotelDbV2HealthCheckActionFor_('REG-19','').indexOf('自己診断ファイル') !== -1);
  check('ACT-06 unchanged', hotelDbV2HealthCheckActionFor_('SAFE-UNCHANGED','').indexOf('他処理') !== -1);
  check('ACT-07 fallback', hotelDbV2HealthCheckActionFor_('OTHER','abc') === 'abc');

  const recs = hotelDbV2HealthCheckRecommendations_([
    {severity:'ok', id:'A', label:'A', detail:''},
    {severity:'warning', id:'SETUP-API', label:'API', detail:''},
    {severity:'critical', id:'CFG-BATCH', label:'設定', detail:''}
  ]);
  check('ACT-08 rec count', recs.length === 2);
  check('ACT-09 rec warning first', recs[0].label === 'API');
  check('ACT-10 rec critical second', recs[1].label === '設定');

  // セクション構造
  const sections = hotelDbV2HealthCheckGroupSections_([
    {section:'安全設定', id:'1'},
    {section:'セットアップ', id:'2'},
    {section:'監査・保護', id:'3'}
  ]);
  check('SEC-01 count3', sections.length === 3);
  check('SEC-02 ordered setup', sections[0].name === 'セットアップ');
  check('SEC-03 ordered config', sections[1].name === '安全設定');
  check('SEC-04 ordered audit', sections[2].name === '監査・保護');
  check('SEC-05 one item', sections[0].items.length === 1);

  // ハッシュ・スナップショット比較
  const digestA = hotelDbV2HealthCheckDigest_('abc');
  const digestB = hotelDbV2HealthCheckDigest_('abc');
  const digestC = hotelDbV2HealthCheckDigest_('abcd');
  check('HASH-01 deterministic', digestA === digestB);
  check('HASH-02 64 hex', /^[0-9a-f]{64}$/.test(digestA));
  check('HASH-03 different input', digestA !== digestC);

  const snapA = {spreadsheetId:'S', apiKeyFingerprint:'K', sheetHashes:{A:'1',B:null}};
  const snapB = {spreadsheetId:'S', apiKeyFingerprint:'K', sheetHashes:{B:null,A:'1'}};
  const snapC = {spreadsheetId:'S', apiKeyFingerprint:'X', sheetHashes:{A:'1',B:null}};
  const snapD = {spreadsheetId:'S', apiKeyFingerprint:'K', sheetHashes:{A:'2',B:null}};
  check('HASH-04 equal ignores key order', hotelDbV2HealthCheckSnapshotsEqual_(snapA, snapB) === true);
  check('HASH-05 api diff', hotelDbV2HealthCheckSnapshotsEqual_(snapA, snapC) === false);
  check('HASH-06 sheet diff', hotelDbV2HealthCheckSnapshotsEqual_(snapA, snapD) === false);
  check('HASH-07 null left', hotelDbV2HealthCheckSnapshotsEqual_(null, snapA) === false);

  // 実環境を読み取りだけで診断し、秘密情報と書込みがないことを確認
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const before = hotelDbV2HealthCheckSnapshot_(spreadsheet);
  const report = hotelDbV2HealthCheckBuildReport_(spreadsheet);
  const after = hotelDbV2HealthCheckSnapshot_(spreadsheet);
  const reportText = JSON.stringify(report);
  const realKey = PropertiesService.getScriptProperties()
    .getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';

  check('REAL-01 report exists', Boolean(report));
  check('REAL-02 version', report.version === '2.0');
  check('REAL-03 overall valid', ['正常','要確認','重大'].indexOf(report.overall) !== -1);
  check('REAL-04 sections', Array.isArray(report.sections) && report.sections.length >= 5);
  check('REAL-05 checks', Array.isArray(report.checks) && report.checks.length > 30);
  check('REAL-06 no external api', report.externalApiCalled === false);
  check('REAL-07 no operational writes', report.operationalWrites === false);
  check('REAL-08 api value null', report.apiKeyValue === null);
  check('REAL-09 api hash null', report.apiKeyHash === null);
  check('REAL-10 raw key absent', !realKey || reportText.indexOf(realKey) === -1);
  check('REAL-11 snapshot unchanged', hotelDbV2HealthCheckSnapshotsEqual_(before, after) === true);
  check('REAL-12 safety marker', report.safetyMarker === HOTEL_DB_V2_HEALTH.SAFETY_MARKER);
  check('REAL-13 total reconciles', report.counts.total === report.checks.length);

  const message = [
    failures.length ? 'PR #23 自己診断 失敗' : 'PR #23 自己診断 成功',
    '',
    '成功件数: ' + passed + '件',
    '失敗件数: ' + failures.length + '件',
    'Google Places API呼出: なし',
    'Script Properties変更: なし',
    '元DB・候補・履歴の変更: なし'
  ];
  if (failures.length) message.push('', '失敗:', failures.join('\n'));

  try { SpreadsheetApp.getUi().alert(message.join('\n')); } catch (ignore) {}
  if (failures.length) {
    throw new Error('PR #23 自己診断で失敗があります: ' + failures.join(' / '));
  }
  return { success: passed, failed: failures.length, externalApi:false, operationalWrites:false };
}

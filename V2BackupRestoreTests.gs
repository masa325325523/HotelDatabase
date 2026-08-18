/**
 * PR #24 バックアップ・復元 自己診断。
 * Driveへのコピー作成やGoogle Places API呼出は行わない。
 */
function runHotelDbV2BackupRestoreTests() {
  const failures = [];
  let passed = 0;
  function check(label, condition, detail) {
    if (condition) passed++;
    else failures.push(label + (detail ? ': ' + detail : ''));
  }

  check('CONST-01 version', HOTEL_DB_V2_BACKUP.VERSION === '2.0');
  check('CONST-02 dialog', HOTEL_DB_V2_BACKUP.DIALOG_FILE === 'V2BackupRestoreDialog');
  check('CONST-03 folder property', HOTEL_DB_V2_BACKUP.FOLDER_PROPERTY === 'HOTEL_DB_V2_BACKUP_FOLDER_ID');
  check('CONST-04 max list', HOTEL_DB_V2_BACKUP.MAX_LIST === 20);
  check('CONST-05 block rows', HOTEL_DB_V2_BACKUP.HASH_BLOCK_ROWS === 250);
  check('CONST-06 no overwrite marker', HOTEL_DB_V2_BACKUP.SAFETY_MARKER.indexOf('no-in-place-restore') !== -1);
  check('CONST-07 verified marker', HOTEL_DB_V2_BACKUP.SAFETY_MARKER.indexOf('verified-copy') !== -1);
  check('CONST-08 no api marker', HOTEL_DB_V2_BACKUP.SAFETY_MARKER.indexOf('no-places-api') !== -1);
  check('CONST-09 secret marker', HOTEL_DB_V2_BACKUP.SAFETY_MARKER.indexOf('api-key-not-returned') !== -1);

  const d = new Date('2026-08-18T01:23:45Z');
  const backupName = hotelDbV2BackupBuildBackupName_('宿泊/DB:*?', d, 'TEST');
  const recoveryName = hotelDbV2BackupBuildRecoveryName_('宿泊/DB:*?', d, 'TEST');
  check('NAME-01 backup prefix', backupName.indexOf(HOTEL_DB_V2_BACKUP.BACKUP_PREFIX) === 0);
  check('NAME-02 recovery prefix', recoveryName.indexOf(HOTEL_DB_V2_BACKUP.RECOVERY_PREFIX) === 0);
  check('NAME-03 invalid chars removed', backupName.indexOf('/') === -1 && backupName.indexOf('*') === -1 && backupName.indexOf('?') === -1);
  check('NAME-04 suffix', backupName.slice(-5) === '_TEST');
  check('NAME-05 safe fallback', hotelDbV2BackupSafeName_('') === '宿泊施設DB');
  check('NAME-06 short fingerprint', hotelDbV2BackupShortFingerprint_('1234567890abcdef') === '1234567890ab');
  check('NAME-07 empty short fingerprint', hotelDbV2BackupShortFingerprint_('') === '');

  const digestA = hotelDbV2BackupDigest_('abc');
  const digestB = hotelDbV2BackupDigest_('abc');
  const digestC = hotelDbV2BackupDigest_('abcd');
  check('HASH-01 deterministic', digestA === digestB);
  check('HASH-02 64 hex', /^[0-9a-f]{64}$/.test(digestA));
  check('HASH-03 difference', digestA !== digestC);
  check('HASH-04 serialize number', hotelDbV2BackupSerializeValue_(12.5) === 'number:12.5');
  check('HASH-05 serialize bool', hotelDbV2BackupSerializeValue_(true) === 'boolean:1');
  check('HASH-06 serialize null', hotelDbV2BackupSerializeValue_(null) === 'null:');
  check('HASH-07 serialize string', hotelDbV2BackupSerializeValue_('abc') === 'string:abc');
  check('HASH-08 serialize date', hotelDbV2BackupSerializeValue_(new Date('2026-01-01T00:00:00Z')) === 'date:2026-01-01T00:00:00.000Z');

  const metadata = hotelDbV2BackupBuildMetadata_({
    sourceSpreadsheetId:'SOURCE1', sourceName:'テストDB', createdAt:d,
    fingerprint:digestA, sheetCount:3, usedCells:120
  });
  check('META-01 marker', metadata.marker === HOTEL_DB_V2_BACKUP.META_MARKER);
  check('META-02 version', metadata.version === '2.0');
  check('META-03 source', metadata.sourceSpreadsheetId === 'SOURCE1');
  check('META-04 source name', metadata.sourceName === 'テストDB');
  check('META-05 iso date', metadata.createdAt === d.toISOString());
  check('META-06 fingerprint', metadata.fingerprint === digestA);
  check('META-07 sheet count', metadata.sheetCount === 3);
  check('META-08 cells', metadata.usedCells === 120);
  check('META-09 verified', metadata.verified === true);

  let validation = hotelDbV2BackupValidateMetadata_(metadata, 'SOURCE1');
  check('VALID-01 valid metadata', validation.valid === true, validation.message);
  validation = hotelDbV2BackupValidateMetadata_(metadata, 'OTHER');
  check('VALID-02 source mismatch', validation.valid === false && validation.message.indexOf('このスプレッドシート用') !== -1);
  validation = hotelDbV2BackupValidateMetadata_(Object.assign({}, metadata, {marker:'BAD'}), 'SOURCE1');
  check('VALID-03 marker mismatch', validation.valid === false);
  validation = hotelDbV2BackupValidateMetadata_(Object.assign({}, metadata, {version:'1.0'}), 'SOURCE1');
  check('VALID-04 version mismatch', validation.valid === false);
  validation = hotelDbV2BackupValidateMetadata_(Object.assign({}, metadata, {sourceSpreadsheetId:''}), '');
  check('VALID-05 source missing', validation.valid === false);
  validation = hotelDbV2BackupValidateMetadata_(Object.assign({}, metadata, {fingerprint:'abc'}), 'SOURCE1');
  check('VALID-06 fingerprint invalid', validation.valid === false);
  validation = hotelDbV2BackupValidateMetadata_(Object.assign({}, metadata, {sheetCount:0}), 'SOURCE1');
  check('VALID-07 sheet count invalid', validation.valid === false);
  validation = hotelDbV2BackupValidateMetadata_(Object.assign({}, metadata, {verified:false}), 'SOURCE1');
  check('VALID-08 unverified', validation.valid === false);

  const fakeFile = { getDescription:function(){ return HOTEL_DB_V2_BACKUP.META_PREFIX + JSON.stringify(metadata); } };
  const parsed = hotelDbV2BackupReadMetadata_(fakeFile);
  check('PARSE-01 valid description', parsed && parsed.sourceSpreadsheetId === 'SOURCE1');
  check('PARSE-02 no prefix', hotelDbV2BackupReadMetadata_({getDescription:function(){return 'x';}}) === null);
  check('PARSE-03 invalid json', hotelDbV2BackupReadMetadata_({getDescription:function(){return HOTEL_DB_V2_BACKUP.META_PREFIX + '{';}}) === null);
  check('PARSE-04 null file', hotelDbV2BackupReadMetadata_(null) === null);

  const realSs = SpreadsheetApp.getActiveSpreadsheet();
  const apiKeyBefore = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  const protectedBefore = hotelDbV2HealthCheckSnapshot_(realSs);
  const fingerprint1 = hotelDbV2BackupFingerprintSpreadsheet_(realSs);
  const fingerprint2 = hotelDbV2BackupFingerprintSpreadsheet_(realSs);
  const protectedAfter = hotelDbV2HealthCheckSnapshot_(realSs);
  const apiKeyAfter = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  check('REAL-01 fingerprint present', /^[0-9a-f]{64}$/.test(fingerprint1.fingerprint));
  check('REAL-02 fingerprint deterministic', fingerprint1.fingerprint === fingerprint2.fingerprint);
  check('REAL-03 sheet count', fingerprint1.sheetCount === realSs.getSheets().length);
  check('REAL-04 sheets summary', fingerprint1.sheets.length === fingerprint1.sheetCount);
  check('REAL-05 used cells nonnegative', fingerprint1.usedCells >= 0);
  check('REAL-06 operational unchanged', hotelDbV2HealthCheckSnapshotsEqual_(protectedBefore, protectedAfter));
  check('REAL-07 api key unchanged', apiKeyBefore === apiKeyAfter);

  const html = HtmlService.createHtmlOutputFromFile(HOTEL_DB_V2_BACKUP.DIALOG_FILE).getContent();
  check('HTML-01 title', html.indexOf('バックアップ・復元') !== -1);
  check('HTML-02 no overwrite', html.indexOf('上書きしません') !== -1 || html.indexOf('上書きせず') !== -1);
  check('HTML-03 backup button', html.indexOf('バックアップを作成') !== -1);
  check('HTML-04 recovery button', html.indexOf('復元候補を作成') !== -1);
  check('HTML-05 api key warning', html.indexOf('APIキー本体') !== -1);
  check('HTML-06 no UrlFetchApp', html.indexOf('UrlFetchApp') === -1);
  check('HTML-07 no dynamic apply', html.indexOf('.apply(') === -1);

  const entries = hotelDbV2HealthCheckEntryPoints_();
  const regressions = hotelDbV2HealthCheckRegressionSuites_();
  check('HEALTH-01 existing entries remain 21', entries.length === 21, 'actual=' + entries.length);
  check('HEALTH-02 no health scope change', !entries.some(function(item){ return item.id === 'ENTRY-BACKUP'; }));
  check('HEALTH-03 existing regressions remain 7', regressions.length === 7, 'actual=' + regressions.length);
  check('HEALTH-04 no regression scope change', !regressions.some(function(item){ return item.id === 'REG-24'; }));

  check('SAFETY-01 public create exists', typeof hotelDbV2BackupCreate === 'function');
  check('SAFETY-02 public recovery exists', typeof hotelDbV2BackupCreateRecovery === 'function');
  check('SAFETY-03 status exists', typeof hotelDbV2BackupGetStatus === 'function');
  check('SAFETY-04 menu exists', typeof runHotelDbV2OpenBackupRestore === 'function');
  check('SAFETY-05 no in-place flag', HOTEL_DB_V2_BACKUP.SAFETY_MARKER.indexOf('no-in-place-restore') !== -1);

  // UIテストのタイムアウト安全契約。Driveコピーは実行しない。
  check('UITIME-01 api fingerprint helper exists', typeof hotelDbV2BackupUiApiKeyFingerprint_ === 'function');
  check('UITIME-02 setup exists', typeof setupHotelDbV2BackupRestoreUiTest === 'function');
  check('UITIME-03 cleanup exists', typeof cleanupHotelDbV2BackupRestoreUiTest === 'function');

  const expectedCount = 71;
  const actualCount = passed + failures.length;
  if (actualCount !== expectedCount) {
    failures.push('TEST-COUNT 自己診断定義数が期待値と不一致: actual=' + actualCount + ', expected=' + expectedCount);
  }

  const message = [
    failures.length ? 'PR #24 自己診断 失敗' : 'PR #24 自己診断 成功', '',
    '成功件数: ' + passed + '件', '失敗件数: ' + failures.length + '件',
    'Driveコピー作成: なし', 'Google Places API呼出: なし',
    'Script Properties APIキー変更: なし', '元DB・候補・履歴の変更: なし'
  ];
  if (failures.length) message.push('', '失敗:', failures.join('\n'));
  try { SpreadsheetApp.getUi().alert(message.join('\n')); } catch (ignore) {}
  if (failures.length) throw new Error('PR #24 自己診断で失敗があります: ' + failures.join(' / '));
  return {success:passed, failed:failures.length, driveCopies:0, externalApi:false, operationalWrites:false};
}

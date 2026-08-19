/**
 * PR #25 リリース準備チェック。
 *
 * 読み取り専用で、製品リリース前の必須条件を1画面に集約する。
 * - Google Places APIは呼ばない。
 * - バックアップは作成しない（検証済みバックアップが存在するかだけ確認）。
 * - 元DB・候補・履歴・Script Propertiesを書き換えない。
 * - APIキー本体・ハッシュを結果へ返さない。
 */
const HOTEL_DB_V2_RELEASE = Object.freeze({
  VERSION: '2.0',
  DIALOG_FILE: 'V2ReleaseReadinessDialog',
  DIALOG_TITLE: '宿泊施設DB Ver2.0 リリース準備チェック',
  DIALOG_WIDTH: 900,
  DIALOG_HEIGHT: 780,
  LOCK_TIMEOUT_MS: 10000,
  SAFETY_MARKER: 'read-only; no-places-api; no-backup-create; api-key-not-returned; release-gate',
  TEST_NAME_PATTERN: /(PR\d+|テスト|TEST|test|コピー|copy|復元候補|バックアップ|作業中|開発|DEV|dev)/,
  TEST_PROPERTY_PATTERN: /^(HOTEL_DB_V2_PR\d+_|HOTEL_DB_V2_.*(?:UI_)?TEST_STATE)/i
});

function runHotelDbV2OpenReleaseReadiness() {
  const html = HtmlService
    .createHtmlOutputFromFile(HOTEL_DB_V2_RELEASE.DIALOG_FILE)
    .setWidth(HOTEL_DB_V2_RELEASE.DIALOG_WIDTH)
    .setHeight(HOTEL_DB_V2_RELEASE.DIALOG_HEIGHT);
  SpreadsheetApp.getUi().showModalDialog(html, HOTEL_DB_V2_RELEASE.DIALOG_TITLE);
}

function hotelDbV2ReleaseGetReport() {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(HOTEL_DB_V2_RELEASE.LOCK_TIMEOUT_MS)) {
    throw new Error('別の処理が実行中です。完了後にもう一度リリース準備チェックを実行してください。');
  }
  try {
    return hotelDbV2ReleaseBuildReport_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (error) {
    throw new Error(hotelDbV2ReleaseSafeMessage_(error));
  } finally {
    lock.releaseLock();
  }
}

function hotelDbV2ReleaseBuildReport_(spreadsheet) {
  const startedAt = new Date();
  const before = hotelDbV2ReleaseSnapshot_(spreadsheet);
  const checks = [];

  function add(id, label, severity, detail, blocking) {
    checks.push({
      id: id,
      label: label,
      severity: severity,
      detail: detail || '',
      blocking: Boolean(blocking)
    });
  }

  const setupAvailable = typeof hotelDbV2SetupBuildStatus_ === 'function';
  const healthAvailable = typeof hotelDbV2HealthCheckBuildReport_ === 'function';
  const backupAvailable = typeof hotelDbV2BackupBuildStatus_ === 'function';

  const setup = setupAvailable ? hotelDbV2SetupBuildStatus_(spreadsheet) : null;
  const health = healthAvailable ? hotelDbV2HealthCheckBuildReport_(spreadsheet) : null;
  const backup = backupAvailable ? hotelDbV2BackupBuildStatus_(spreadsheet) : null;
  const environment = hotelDbV2ReleaseEnvironment_(spreadsheet && spreadsheet.getName ? spreadsheet.getName() : '');
  const testResidues = hotelDbV2ReleaseFindTestResidues_();

  add(
    'MODULE-SETUP', '初期セットアップ機能',
    setupAvailable ? 'ok' : 'critical',
    setupAvailable ? '利用可能です。' : 'V2Setup.gs の実行入口を確認できません。',
    !setupAvailable
  );
  add(
    'MODULE-HEALTH', '製品全体ヘルスチェック',
    healthAvailable ? 'ok' : 'critical',
    healthAvailable ? '利用可能です。' : 'V2HealthCheck.gs の診断機能を確認できません。',
    !healthAvailable
  );
  add(
    'MODULE-BACKUP', 'バックアップ・復元',
    backupAvailable ? 'ok' : 'critical',
    backupAvailable ? '利用可能です。' : 'V2BackupRestore.gs の機能を確認できません。',
    !backupAvailable
  );

  add(
    'SETUP-API', 'Google Places APIキー',
    setup && setup.apiKeyConfigured ? 'ok' : 'critical',
    setup && setup.apiKeyConfigured
      ? '設定済みです。キー本体はこの結果へ返しません。'
      : '未設定です。「⚙ 初期セットアップ・設定」から設定してください。',
    !(setup && setup.apiKeyConfigured)
  );
  add(
    'SETUP-SOURCE', '元データシート',
    setup && setup.source && setup.source.ready ? 'ok' : 'critical',
    setup && setup.source && setup.source.ready
      ? '使用可能: ' + setup.source.name
      : 'リリース前に元データシートを開き、列構成を確認してください。',
    !(setup && setup.source && setup.source.ready)
  );
  add(
    'SETUP-CORE', 'コア運用シート',
    setup && setup.coreMissing === 0 && setup.coreNeedsReview === 0 ? 'ok' : 'critical',
    setup
      ? '未作成=' + Number(setup.coreMissing || 0) + ' / 要確認=' + Number(setup.coreNeedsReview || 0)
      : 'セットアップ状態を取得できません。',
    !(setup && setup.coreMissing === 0 && setup.coreNeedsReview === 0)
  );

  const healthCritical = health && health.counts ? Number(health.counts.critical || 0) : null;
  add(
    'HEALTH-CRITICAL', '重大なヘルスチェック項目',
    healthCritical === 0 ? 'ok' : 'critical',
    healthCritical === 0
      ? '重大項目は0件です。'
      : (healthCritical === null ? 'ヘルスチェック結果を取得できません。' : '重大項目が' + healthCritical + '件あります。'),
    healthCritical !== 0
  );
  const healthWarnings = health && health.counts ? Number(health.counts.warning || 0) : null;
  add(
    'HEALTH-WARNING', 'ヘルスチェック警告',
    healthWarnings && healthWarnings > 0 ? 'warning' : 'ok',
    healthWarnings === null ? '取得できません。' : '警告=' + healthWarnings + '件',
    false
  );

  const backupCount = backup ? Number(backup.backupCount || 0) : 0;
  add(
    'BACKUP-EXISTS', '検証済みバックアップ',
    backupCount > 0 ? 'ok' : 'critical',
    backupCount > 0
      ? '利用可能な検証済みバックアップ=' + backupCount + '件'
      : 'リリース前に「💾 バックアップ・復元」から検証済みバックアップを1件以上作成してください。',
    backupCount < 1
  );
  add(
    'BACKUP-NO-INPLACE', '本番ファイル直接上書き禁止',
    backup && backup.inPlaceRestore === false ? 'ok' : 'critical',
    backup && backup.inPlaceRestore === false
      ? '復元は別ファイルの復元候補を作る方式です。'
      : '安全な復元方式を確認できません。',
    !(backup && backup.inPlaceRestore === false)
  );

  add(
    'ENVIRONMENT', '実行環境',
    environment.kind === 'test' ? 'info' : 'ok',
    environment.label + ': ' + environment.name,
    environment.kind === 'test'
  );

  add(
    'TEST-RESIDUE', 'テスト用User Properties',
    testResidues.length === 0 ? 'ok' : 'critical',
    testResidues.length === 0
      ? '残留なし。'
      : '残留キー=' + testResidues.join(', ') + '。対応するcleanupを実行してください。',
    testResidues.length > 0
  );

  add(
    'ENTRY-HEALTH', '🩺 製品全体ヘルスチェック入口',
    typeof runHotelDbV2HealthCheck === 'function' ? 'ok' : 'critical',
    typeof runHotelDbV2HealthCheck === 'function' ? '確認済み。' : '実行入口がありません。',
    typeof runHotelDbV2HealthCheck !== 'function'
  );
  add(
    'ENTRY-BACKUP', '💾 バックアップ・復元入口',
    typeof runHotelDbV2OpenBackupRestore === 'function' ? 'ok' : 'critical',
    typeof runHotelDbV2OpenBackupRestore === 'function' ? '確認済み。' : '実行入口がありません。',
    typeof runHotelDbV2OpenBackupRestore !== 'function'
  );
  add(
    'ENTRY-RELEASE', '🚀 リリース準備チェック入口',
    typeof runHotelDbV2OpenReleaseReadiness === 'function' ? 'ok' : 'critical',
    typeof runHotelDbV2OpenReleaseReadiness === 'function' ? '確認済み。' : '実行入口がありません。',
    typeof runHotelDbV2OpenReleaseReadiness !== 'function'
  );

  add(
    'SAFE-NO-PLACES', 'Google Places API自動呼び出し',
    'ok',
    'このチェックではGoogle Places APIを呼びません。',
    false
  );
  add(
    'SAFE-NO-BACKUP-CREATE', 'バックアップ自動作成',
    'ok',
    '存在確認だけを行い、バックアップは自動作成しません。',
    false
  );
  add(
    'SAFE-API-KEY', 'APIキー非露出',
    'ok',
    '設定済み/未設定だけを確認し、キー本体・ハッシュは結果へ返しません。',
    false
  );

  const after = hotelDbV2ReleaseSnapshot_(spreadsheet);
  const unchanged = hotelDbV2ReleaseSnapshotsEqual_(before, after);
  add(
    'SAFE-UNCHANGED', 'チェック前後のデータ不変',
    unchanged ? 'ok' : 'critical',
    unchanged
      ? '元データシートとScript PropertiesのAPIキー指紋はチェック前後で一致しました。運用シートは内部ヘルスチェック側でも不変確認します。'
      : 'チェック中に元データシートまたはAPIキー状態の変化を検出しました。',
    !unchanged
  );

  const summary = hotelDbV2ReleaseSummarize_(checks, environment);
  return {
    version: HOTEL_DB_V2_RELEASE.VERSION,
    checkedAt: Utilities.formatDate(new Date(), HOTEL_DB_V2_CONFIG.TIMEZONE || 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    durationMs: new Date().getTime() - startedAt.getTime(),
    overall: summary.overall,
    releaseReady: summary.releaseReady,
    environment: environment,
    counts: summary.counts,
    blockingCount: summary.blockingCount,
    backupCount: backupCount,
    healthCritical: healthCritical,
    healthWarnings: healthWarnings,
    sourceName: setup && setup.source ? setup.source.name : '',
    checks: checks,
    recommendations: hotelDbV2ReleaseRecommendations_(checks, environment),
    externalApiCalled: false,
    backupCreated: false,
    operationalWrites: false,
    apiKeyConfigured: Boolean(setup && setup.apiKeyConfigured),
    apiKeyValue: null,
    apiKeyHash: null,
    safetyMarker: HOTEL_DB_V2_RELEASE.SAFETY_MARKER
  };
}

function hotelDbV2ReleaseEnvironment_(name) {
  const cleaned = String(name || '').trim();
  const test = HOTEL_DB_V2_RELEASE.TEST_NAME_PATTERN.test(cleaned);
  return {
    name: cleaned || '名称未取得',
    kind: test ? 'test' : 'production',
    label: test ? 'テスト環境' : '製品候補'
  };
}

function hotelDbV2ReleaseFindTestResidues_() {
  const props = PropertiesService.getUserProperties().getProperties();
  return Object.keys(props).filter(function(key) {
    return HOTEL_DB_V2_RELEASE.TEST_PROPERTY_PATTERN.test(String(key));
  }).sort();
}

function hotelDbV2ReleaseSummarize_(checks, environment) {
  const counts = { ok:0, info:0, warning:0, critical:0, total:0 };
  let blockingCount = 0;
  (checks || []).forEach(function(item) {
    const severity = counts.hasOwnProperty(item.severity) ? item.severity : 'warning';
    counts[severity]++;
    counts.total++;
    if (item.blocking) blockingCount++;
  });
  const isTest = environment && environment.kind === 'test';
  return {
    counts: counts,
    blockingCount: blockingCount,
    releaseReady: !isTest && blockingCount === 0,
    overall: isTest ? 'テスト環境' : (blockingCount === 0 ? 'リリース可能' : '要対応')
  };
}

function hotelDbV2ReleaseRecommendations_(checks, environment) {
  const items = [];
  if (environment && environment.kind === 'test') {
    items.push({
      severity: 'info',
      label: '製品候補ファイルで最終確認',
      action: 'コピー版ではリリース可能判定にしません。製品候補のスプレッドシートで再実行してください。'
    });
  }
  (checks || []).filter(function(item) {
    return item.blocking || item.severity === 'warning';
  }).forEach(function(item) {
    if (items.length >= 8) return;
    items.push({
      severity: item.severity,
      label: item.label,
      action: hotelDbV2ReleaseActionFor_(item.id, item.detail)
    });
  });
  return items.slice(0, 8);
}

function hotelDbV2ReleaseActionFor_(id, detail) {
  if (id === 'SETUP-API' || id === 'SETUP-SOURCE' || id === 'SETUP-CORE') {
    return '「⚙ 初期セットアップ・設定」で状態を確認してください。';
  }
  if (id === 'HEALTH-CRITICAL' || id === 'HEALTH-WARNING') {
    return '「🩺 製品全体ヘルスチェック」で詳細を確認してください。';
  }
  if (id === 'BACKUP-EXISTS' || id === 'BACKUP-NO-INPLACE') {
    return '「💾 バックアップ・復元」で検証済みバックアップを確認してください。';
  }
  if (id === 'TEST-RESIDUE') {
    return '残っているテストのcleanup関数を実行してから再確認してください。';
  }
  if (String(id).indexOf('ENTRY-') === 0 || String(id).indexOf('MODULE-') === 0) {
    return 'Apps Scriptコードが最新mainと一致しているか確認してください。';
  }
  if (id === 'SAFE-UNCHANGED') {
    return '他の処理を止めた状態で再実行し、続く場合はコード差分を確認してください。';
  }
  return detail || '内容を確認してください。';
}

function hotelDbV2ReleaseSnapshot_(spreadsheet) {
  const active = spreadsheet && spreadsheet.getActiveSheet ? spreadsheet.getActiveSheet() : null;
  const activeHash = active && typeof hotelDbV2HealthCheckHashSheet_ === 'function'
    ? hotelDbV2HealthCheckHashSheet_(active)
    : null;
  const key = PropertiesService.getScriptProperties()
    .getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  const keyFingerprint = typeof hotelDbV2HealthCheckDigest_ === 'function'
    ? hotelDbV2HealthCheckDigest_(String(key))
    : String(key).length;
  return {
    spreadsheetId: spreadsheet && spreadsheet.getId ? spreadsheet.getId() : '',
    activeSheetId: active && active.getSheetId ? active.getSheetId() : null,
    activeSheetHash: activeHash,
    apiKeyFingerprint: keyFingerprint
  };
}

function hotelDbV2ReleaseSnapshotsEqual_(left, right) {
  if (!left || !right) return false;
  return left.spreadsheetId === right.spreadsheetId &&
    left.activeSheetId === right.activeSheetId &&
    left.activeSheetHash === right.activeSheetHash &&
    left.apiKeyFingerprint === right.apiKeyFingerprint;
}

function hotelDbV2ReleaseSafeMessage_(error) {
  let message = error && error.message ? String(error.message) : String(error || '不明なエラー');
  try {
    const key = PropertiesService.getScriptProperties()
      .getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY);
    if (key) message = message.split(String(key)).join('[API_KEY]');
  } catch (ignore) {}
  return message;
}

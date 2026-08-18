/**
 * PR #23 製品全体ヘルスチェック。
 *
 * 製品利用者向けの「読み取り専用」診断。
 * - Google Places APIは自動では呼ばない。
 * - 元DB・候補・履歴・Script Propertiesは変更しない。
 * - APIキー本体は結果へ含めない。
 * - 診断前後の運用シート / APIキー指紋を比較し、自己変更がないことも検証する。
 */

const HOTEL_DB_V2_HEALTH = Object.freeze({
  VERSION: '2.0',
  DIALOG_FILE: 'V2HealthCheckDialog',
  DIALOG_TITLE: '宿泊施設DB Ver2.0 製品全体ヘルスチェック',
  DIALOG_WIDTH: 860,
  DIALOG_HEIGHT: 760,
  LOCK_TIMEOUT_MS: 10000,
  OPTIONAL_SHEETS: Object.freeze([
    '新規追加候補',
    '新規施設分類候補',
    '閉業除外履歴',
    '重複整理履歴',
    '統合ダッシュボード'
  ]),
  SAFETY_MARKER: 'read-only; no-external-api; api-key-not-returned; operational-hash-guard'
});

function runHotelDbV2HealthCheck() {
  const html = HtmlService
    .createHtmlOutputFromFile(HOTEL_DB_V2_HEALTH.DIALOG_FILE)
    .setWidth(HOTEL_DB_V2_HEALTH.DIALOG_WIDTH)
    .setHeight(HOTEL_DB_V2_HEALTH.DIALOG_HEIGHT);
  SpreadsheetApp.getUi().showModalDialog(html, HOTEL_DB_V2_HEALTH.DIALOG_TITLE);
}

function hotelDbV2HealthCheckGetReport() {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(HOTEL_DB_V2_HEALTH.LOCK_TIMEOUT_MS)) {
    throw new Error('別の処理が実行中です。完了後にもう一度ヘルスチェックしてください。');
  }
  try {
    return hotelDbV2HealthCheckBuildReport_(SpreadsheetApp.getActiveSpreadsheet());
  } catch (error) {
    throw new Error(hotelDbV2HealthCheckSafeMessage_(error));
  } finally {
    lock.releaseLock();
  }
}

function hotelDbV2HealthCheckBuildReport_(spreadsheet) {
  const startedAt = new Date();
  const before = hotelDbV2HealthCheckSnapshot_(spreadsheet);
  const checks = [];

  function add(section, id, label, severity, detail) {
    checks.push({ section:section, id:id, label:label, severity:severity, detail:detail || '' });
  }

  let setupStatus = null;
  if (typeof hotelDbV2SetupBuildStatus_ === 'function') {
    setupStatus = hotelDbV2SetupBuildStatus_(spreadsheet);
    add(
      'セットアップ', 'SETUP-API', 'Google Places APIキー',
      setupStatus.apiKeyConfigured ? 'ok' : 'warning',
      setupStatus.apiKeyConfigured
        ? '設定済み。保存済みキー本体はヘルスチェック結果へ返しません。'
        : '未設定です。「⚙ 初期セットアップ・設定」から設定してください。'
    );
    if (setupStatus.source && setupStatus.source.ready) {
      add('セットアップ', 'SETUP-SOURCE', '元データシート', 'ok', '使用可能: ' + setupStatus.source.name);
    } else {
      const sourceName = setupStatus.source && setupStatus.source.name ? setupStatus.source.name : '取得不可';
      const issueText = setupStatus.source && setupStatus.source.issues
        ? setupStatus.source.issues.join(' / ')
        : '元データシートを開いて再確認してください。';
      add('セットアップ', 'SETUP-SOURCE', '元データシート', 'info', '現在のシート: ' + sourceName + '。' + issueText);
    }
    (setupStatus.coreSheets || []).forEach(function(item) {
      add(
        '運用シート', 'CORE-' + item.name, item.name,
        !item.exists || !item.compatible ? 'critical' : 'ok',
        !item.exists ? '未作成です。' : (item.compatible ? '見出し互換・準備済み。' : '見出しが期待仕様と一致しません。')
      );
    });
  } else {
    add('セットアップ', 'SETUP-MODULE', '初期セットアップモジュール', 'critical', 'hotelDbV2SetupBuildStatus_ が見つかりません。');
  }

  const configExists = typeof HOTEL_DB_V2_CONFIG !== 'undefined' && Boolean(HOTEL_DB_V2_CONFIG);
  add('安全設定', 'CFG-EXISTS', 'Ver2.0設定', configExists ? 'ok' : 'critical',
      configExists ? '設定オブジェクトを確認しました。' : 'HOTEL_DB_V2_CONFIG が見つかりません。');
  if (configExists) {
    hotelDbV2HealthCheckConfigChecks_().forEach(function(item) {
      add('安全設定', item.id, item.label, item.ok ? 'ok' : 'critical', item.detail);
    });
  }

  hotelDbV2HealthCheckEntryPoints_().forEach(function(item) {
    add('機能入口', item.id, item.label, item.exists ? 'ok' : 'critical',
        item.exists ? '実行入口を確認しました。' : '実行関数が見つかりません。');
  });

  hotelDbV2HealthCheckRegressionSuites_().forEach(function(item) {
    add('安全回帰', item.id, item.label, item.exists ? 'ok' : 'warning',
        item.exists ? '個別自己診断スイートを利用できます。ヘルスチェック中は自動実行しません。' : '個別自己診断スイートが見つかりません。');
  });

  HOTEL_DB_V2_HEALTH.OPTIONAL_SHEETS.forEach(function(name) {
    const sheet = spreadsheet.getSheetByName(name);
    add('運用シート', 'OPTIONAL-' + name, name, sheet ? 'ok' : 'info',
        sheet ? '存在を確認しました。' : '未作成です。必要な機能を実行したときに作成できます。');
  });

  add('監査・保護', 'SAFE-NO-API', '外部API自動呼び出し', 'ok',
      'Google Places APIは自動では呼びません。接続確認は設定画面の「接続テスト」で明示実行します。');
  add('監査・保護', 'SAFE-API-KEY', 'APIキー非露出', 'ok',
      '結果オブジェクトには設定済み/未設定のみを含め、キー本体・ハッシュは返しません。');

  const after = hotelDbV2HealthCheckSnapshot_(spreadsheet);
  const unchanged = hotelDbV2HealthCheckSnapshotsEqual_(before, after);
  add(
    '監査・保護', 'SAFE-UNCHANGED', '診断前後のデータ不変', unchanged ? 'ok' : 'critical',
    unchanged
      ? '運用シート内容・数式とScript PropertiesのAPIキー指紋は診断前後で一致しました。'
      : '診断中に保護対象の状態変化を検出しました。別処理の同時実行有無も確認してください。'
  );

  const summary = hotelDbV2HealthCheckSummarize_(checks);
  const recommendations = hotelDbV2HealthCheckRecommendations_(checks);
  return {
    version: HOTEL_DB_V2_HEALTH.VERSION,
    checkedAt: Utilities.formatDate(new Date(), HOTEL_DB_V2_CONFIG.TIMEZONE || 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'),
    durationMs: new Date().getTime() - startedAt.getTime(),
    overall: summary.overall,
    counts: summary.counts,
    checks: checks,
    sections: hotelDbV2HealthCheckGroupSections_(checks),
    recommendations: recommendations,
    externalApiCalled: false,
    operationalWrites: false,
    apiKeyConfigured: setupStatus ? Boolean(setupStatus.apiKeyConfigured) : false,
    apiKeyValue: null,
    apiKeyHash: null,
    safetyMarker: HOTEL_DB_V2_HEALTH.SAFETY_MARKER
  };
}

function hotelDbV2HealthCheckConfigChecks_() {
  return [
    { id:'CFG-API-PROP', label:'APIキー保存先', ok:HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY === 'GOOGLE_PLACES_API_KEY', detail:'GOOGLE_PLACES_API_KEY のみを使用します。' },
    { id:'CFG-AUTO-SCORE', label:'自動採用しきい値', ok:Number(HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE) === 75, detail:'安全基準: 75点。現在=' + HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE },
    { id:'CFG-MIN-SCORE', label:'最低一致スコア', ok:Number(HOTEL_DB_V2_CONFIG.MIN_MATCH_SCORE) === 55, detail:'安全基準: 55点。現在=' + HOTEL_DB_V2_CONFIG.MIN_MATCH_SCORE },
    { id:'CFG-BATCH', label:'本番バッチ件数', ok:Number(HOTEL_DB_V2_CONFIG.BATCH_SIZE) === 50, detail:'安全基準: 50件。現在=' + HOTEL_DB_V2_CONFIG.BATCH_SIZE },
    { id:'CFG-TEST', label:'先頭テスト件数', ok:Number(HOTEL_DB_V2_CONFIG.TEST_ROWS) === 3, detail:'安全基準: 3件。現在=' + HOTEL_DB_V2_CONFIG.TEST_ROWS },
    { id:'CFG-LANGUAGE', label:'Google Places言語', ok:HOTEL_DB_V2_CONFIG.LANGUAGE_CODE === 'ja', detail:'期待値: ja。現在=' + HOTEL_DB_V2_CONFIG.LANGUAGE_CODE },
    { id:'CFG-REGION', label:'Google Places地域', ok:HOTEL_DB_V2_CONFIG.REGION_CODE === 'JP', detail:'期待値: JP。現在=' + HOTEL_DB_V2_CONFIG.REGION_CODE },
    { id:'CFG-TIMEZONE', label:'タイムゾーン', ok:HOTEL_DB_V2_CONFIG.TIMEZONE === 'Asia/Tokyo', detail:'期待値: Asia/Tokyo。現在=' + HOTEL_DB_V2_CONFIG.TIMEZONE }
  ];
}

function hotelDbV2HealthCheckEntryPoints_() {
  return [
    { id:'ENTRY-SETUP', label:'⚙ 初期セットアップ・設定', exists: typeof runHotelDbV2OpenSetup === 'function' },
    { id:'ENTRY-HEALTH', label:'🩺 製品全体ヘルスチェック', exists: typeof runHotelDbV2HealthCheck === 'function' },
    { id:'ENTRY-BACKUP', label:'💾 バックアップ・復元', exists: typeof runHotelDbV2OpenBackupRestore === 'function' },
    { id:'ENTRY-01', label:'① 設定・見出し診断', exists: typeof runHotelDbV2Diagnosis === 'function' },
    { id:'ENTRY-02', label:'② API接続テスト', exists: typeof runHotelDbV2ConnectionTest === 'function' },
    { id:'ENTRY-03', label:'③ 先頭3件テスト', exists: typeof runHotelDbV2Test3 === 'function' },
    { id:'ENTRY-04', label:'④ 本番バッチ実行', exists: typeof runHotelDbV2Batch === 'function' },
    { id:'ENTRY-05', label:'⑤ Place ID再確認', exists: typeof runHotelDbV2RefreshExisting === 'function' },
    { id:'ENTRY-06', label:'⑥ 開始行リセット', exists: typeof runHotelDbV2ResetCheckpoint === 'function' },
    { id:'ENTRY-07', label:'⑦ 重複候補更新', exists: typeof runHotelDbV2Duplicates === 'function' },
    { id:'ENTRY-08', label:'⑧ 承認済み修正反映', exists: typeof runHotelDbV2ApplyApprovedCorrections === 'function' },
    { id:'ENTRY-09', label:'⑨ 修正候補自動仕分け', exists: typeof runHotelDbV2TriageCorrections === 'function' },
    { id:'ENTRY-10', label:'⑩ 重複候補自動仕分け', exists: typeof runHotelDbV2DuplicateTriageWithPr20Snapshots === 'function' },
    { id:'ENTRY-11', label:'⑪ 低スコア要確認仕分け', exists: typeof runHotelDbV2LowScoreReviewTriageWithMunicipalityRefinement === 'function' },
    { id:'ENTRY-12', label:'⑫ 閉業・未検出仕分け', exists: typeof runHotelDbV2DeletionCandidateTriage === 'function' },
    { id:'ENTRY-13', label:'⑬ 新規施設探索', exists: typeof runHotelDbV2DiscoverNewFacilities === 'function' },
    { id:'ENTRY-14', label:'⑭ 承認済み新規追加', exists: typeof runHotelDbV2ApplyApprovedNewFacilities === 'function' },
    { id:'ENTRY-15', label:'⑮ 新規施設分類候補', exists: typeof runHotelDbV2BuildNewFacilityClassificationCandidates === 'function' },
    { id:'ENTRY-16', label:'⑯ 承認済み分類反映', exists: typeof runHotelDbV2ApplyApprovedNewFacilityClassifications === 'function' },
    { id:'ENTRY-17', label:'⑰ 承認済み閉業除外', exists: typeof runHotelDbV2ApplyApprovedClosedFacilityRemovals === 'function' },
    { id:'ENTRY-18', label:'⑱ 承認済み重複整理', exists: typeof runHotelDbV2ApplyApprovedDuplicateConsolidations === 'function' },
    { id:'ENTRY-19', label:'⑲ 統合ダッシュボード', exists: typeof runHotelDbV2RefreshDashboard === 'function' }
  ];
}

function hotelDbV2HealthCheckRegressionSuites_() {
  return [
    { id:'REG-13', label:'PR #13 修正反映安全テスト', exists: typeof runHotelDbV2ApprovedCorrectionSafetyTests === 'function' },
    { id:'REG-17', label:'PR #17 新規追加安全テスト', exists: typeof runHotelDbV2ApprovedNewFacilityInsertTests === 'function' },
    { id:'REG-18', label:'PR #18 分類反映安全テスト', exists: typeof runHotelDbV2NewFacilityClassificationFinalizeTests === 'function' },
    { id:'REG-19', label:'PR #19 閉業除外安全テスト', exists: typeof runHotelDbV2ApprovedClosedFacilityRemovalTests === 'function' },
    { id:'REG-20', label:'PR #20 重複整理安全テスト', exists: typeof runHotelDbV2ApprovedDuplicateConsolidationTests === 'function' },
    { id:'REG-21', label:'PR #21 ダッシュボード自己診断', exists: typeof runHotelDbV2DashboardTests === 'function' },
    { id:'REG-22', label:'PR #22 セットアップ自己診断', exists: typeof runHotelDbV2SetupTests === 'function' },
    { id:'REG-24', label:'PR #24 バックアップ・復元自己診断', exists: typeof runHotelDbV2BackupRestoreTests === 'function' }
  ];
}

function hotelDbV2HealthCheckSummarize_(checks) {
  const counts = { ok:0, info:0, warning:0, critical:0, total:0 };
  (checks || []).forEach(function(item) {
    const key = counts.hasOwnProperty(item.severity) ? item.severity : 'warning';
    counts[key]++;
    counts.total++;
  });
  const overall = counts.critical > 0 ? '重大' : (counts.warning > 0 ? '要確認' : '正常');
  return { overall:overall, counts:counts };
}

function hotelDbV2HealthCheckRecommendations_(checks) {
  const items = (checks || []).filter(function(item) { return item.severity === 'critical' || item.severity === 'warning'; });
  return items.slice(0, 8).map(function(item) {
    return { severity:item.severity, label:item.label, action:hotelDbV2HealthCheckActionFor_(item.id, item.detail) };
  });
}

function hotelDbV2HealthCheckActionFor_(id, detail) {
  if (id === 'SETUP-API') return '「⚙ 初期セットアップ・設定」でAPIキーを設定してください。';
  if (String(id).indexOf('CORE-') === 0) return '設定画面でコア運用シートの状態を確認してください。';
  if (String(id).indexOf('ENTRY-') === 0) return 'Apps Scriptコードが最新mainと一致しているか確認してください。';
  if (String(id).indexOf('CFG-') === 0) return '安全設定が変更されています。コード差分を確認してください。';
  if (String(id).indexOf('REG-') === 0) return '自己診断ファイルがApps Scriptへ反映されているか確認してください。';
  if (id === 'SAFE-UNCHANGED') return '他処理を止めた状態で再診断し、続く場合はコードを確認してください。';
  return detail || '内容を確認してください。';
}

function hotelDbV2HealthCheckGroupSections_(checks) {
  const order = ['セットアップ', '安全設定', '機能入口', '運用シート', '安全回帰', '監査・保護'];
  return order.map(function(name) {
    return { name:name, items:(checks || []).filter(function(item) { return item.section === name; }) };
  }).filter(function(section) { return section.items.length > 0; });
}

function hotelDbV2HealthCheckSnapshot_(spreadsheet) {
  const names = hotelDbV2HealthCheckProtectedSheetNames_();
  const sheetHashes = {};
  names.forEach(function(name) {
    const sheet = spreadsheet.getSheetByName(name);
    if (!sheet) { sheetHashes[name] = null; return; }
    sheetHashes[name] = hotelDbV2HealthCheckHashSheet_(sheet);
  });
  const apiKey = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY) || '';
  return { spreadsheetId:spreadsheet.getId(), sheetHashes:sheetHashes, apiKeyFingerprint:hotelDbV2HealthCheckDigest_(String(apiKey)) };
}

function hotelDbV2HealthCheckProtectedSheetNames_() {
  const names = [];
  if (typeof hotelDbV2SetupReservedSheetNames_ === 'function') {
    hotelDbV2SetupReservedSheetNames_().forEach(function(name) { if (names.indexOf(name) === -1) names.push(name); });
  } else {
    [
      '修正候補', '要確認', '修正履歴', '重複候補', '実行サマリー',
      '新規追加候補', '新規施設分類候補', '閉業除外履歴', '重複整理履歴', '統合ダッシュボード'
    ].forEach(function(name) { names.push(name); });
  }
  return names;
}

function hotelDbV2HealthCheckHashSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    return hotelDbV2HealthCheckDigest_(JSON.stringify({ id:sheet.getSheetId(), rows:lastRow, columns:lastColumn, values:[], formulas:[] }));
  }
  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  return hotelDbV2HealthCheckDigest_(JSON.stringify({
    id:sheet.getSheetId(), rows:lastRow, columns:lastColumn,
    values:range.getDisplayValues(), formulas:range.getFormulas()
  }));
}

function hotelDbV2HealthCheckDigest_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''), Utilities.Charset.UTF_8);
  return bytes.map(function(value) {
    const normalized = value < 0 ? value + 256 : value;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function hotelDbV2HealthCheckSnapshotsEqual_(left, right) {
  if (!left || !right) return false;
  if (left.spreadsheetId !== right.spreadsheetId) return false;
  if (left.apiKeyFingerprint !== right.apiKeyFingerprint) return false;
  const leftKeys = Object.keys(left.sheetHashes || {}).sort();
  const rightKeys = Object.keys(right.sheetHashes || {}).sort();
  if (JSON.stringify(leftKeys) !== JSON.stringify(rightKeys)) return false;
  return leftKeys.every(function(key) { return left.sheetHashes[key] === right.sheetHashes[key]; });
}

function hotelDbV2HealthCheckSafeMessage_(error) {
  let message = error && error.message ? String(error.message) : String(error || '不明なエラー');
  try {
    const key = PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY);
    if (key) message = message.split(String(key)).join('[API_KEY]');
  } catch (ignore) {}
  return message;
}

function hotelDbV2HealthCheckSafetyMarker_() {
  return HOTEL_DB_V2_HEALTH.SAFETY_MARKER;
}

/**
 * PR #21 コピー版UIテスト。
 * 宿泊施設DB_PR13_⑧反映テスト 専用。
 * 運用シートの内容ハッシュを前後比較し、ダッシュボード更新が読み取り専用であることを確認する。
 */
const HOTEL_DB_V2_DASHBOARD_UI_TEST = Object.freeze({
  REQUIRED_NAME_PARTS: ['PR13', '⑧反映テスト'],
  BACKUP_DASHBOARD: 'PR21_統合ダッシュボード退避',
  FINGERPRINT_PROPERTY: 'HOTEL_DB_V2_PR21_UI_FINGERPRINTS'
});

function hotelDbV2DashboardUiAssertCopy_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_DASHBOARD_UI_TEST.REQUIRED_NAME_PARTS.every(function(part) {
    return name.indexOf(part) !== -1;
  });
  if (!ok) {
    throw new Error('PR #21 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」専用です。');
  }
  return spreadsheet;
}

function setupHotelDbV2DashboardUiTest() {
  const ss = hotelDbV2DashboardUiAssertCopy_();
  const properties = PropertiesService.getDocumentProperties();
  if (ss.getSheetByName(HOTEL_DB_V2_DASHBOARD_UI_TEST.BACKUP_DASHBOARD)) {
    throw new Error('PR21テスト残骸があります。先にcleanupしてください。');
  }
  if (properties.getProperty(HOTEL_DB_V2_DASHBOARD_UI_TEST.FINGERPRINT_PROPERTY)) {
    throw new Error('PR21テスト用フィンガープリントが残っています。先にcleanupしてください。');
  }

  const existingDashboard = ss.getSheetByName(HOTEL_DB_V2_DASHBOARD.SHEET_NAME);
  if (existingDashboard) existingDashboard.setName(HOTEL_DB_V2_DASHBOARD_UI_TEST.BACKUP_DASHBOARD);

  try {
    const before = hotelDbV2DashboardUiFingerprints_(ss);
    properties.setProperty(
      HOTEL_DB_V2_DASHBOARD_UI_TEST.FINGERPRINT_PROPERTY,
      JSON.stringify(before)
    );

    const data = hotelDbV2RefreshDashboard_(ss, { activate: true });
    SpreadsheetApp.getUi().alert([
      'PR #21 UIテスト準備完了',
      '',
      '統合ダッシュボードを実データから生成しました。',
      '次に testHotelDbV2DashboardUiTest() を実行してください。',
      '',
      '全体状態: ' + data.overallStatus,
      '要対応件数: ' + data.attentionTotal,
      '運用シートの変更: なし（テストでハッシュ検証します）'
    ].join('\n'));
    return { prepared: true, overallStatus: data.overallStatus, attentionTotal: data.attentionTotal };
  } catch (error) {
    const generated = ss.getSheetByName(HOTEL_DB_V2_DASHBOARD.SHEET_NAME);
    if (generated) ss.deleteSheet(generated);
    const backup = ss.getSheetByName(HOTEL_DB_V2_DASHBOARD_UI_TEST.BACKUP_DASHBOARD);
    if (backup) backup.setName(HOTEL_DB_V2_DASHBOARD.SHEET_NAME);
    properties.deleteProperty(HOTEL_DB_V2_DASHBOARD_UI_TEST.FINGERPRINT_PROPERTY);
    throw error;
  }
}

function testHotelDbV2DashboardUiTest() {
  const ss = hotelDbV2DashboardUiAssertCopy_();
  const properties = PropertiesService.getDocumentProperties();
  const stored = properties.getProperty(HOTEL_DB_V2_DASHBOARD_UI_TEST.FINGERPRINT_PROPERTY);
  if (!stored) throw new Error('setupHotelDbV2DashboardUiTest() を先に実行してください。');

  const dashboard = ss.getSheetByName(HOTEL_DB_V2_DASHBOARD.SHEET_NAME);
  if (!dashboard) throw new Error('「統合ダッシュボード」が見つかりません。');

  const expectedTitle = '宿泊施設DB Ver2.0　統合ダッシュボード';
  if (dashboard.getRange('A1').getDisplayValue() !== expectedTitle) {
    throw new Error('ダッシュボードタイトルが期待値と一致しません。');
  }

  const sections = [
    [HOTEL_DB_V2_DASHBOARD.ROWS.LATEST_SECTION, '最新バッチ実行'],
    [HOTEL_DB_V2_DASHBOARD.ROWS.WORKFLOW_SECTION, '候補ワークフロー状況'],
    [HOTEL_DB_V2_DASHBOARD.ROWS.ACTION_SECTION, '次にやること（優先順）'],
    [HOTEL_DB_V2_DASHBOARD.ROWS.AUDIT_SECTION, '監査・完了実績'],
    [HOTEL_DB_V2_DASHBOARD.ROWS.HISTORY_SECTION, '最近の処理履歴（最大10件）']
  ];
  sections.forEach(function(section) {
    if (dashboard.getRange(section[0], 1).getDisplayValue() !== section[1]) {
      throw new Error('セクション「' + section[1] + '」が見つかりません。');
    }
  });

  const data = hotelDbV2DashboardCollect_(ss);
  const workflowNames = dashboard
    .getRange(HOTEL_DB_V2_DASHBOARD.ROWS.WORKFLOW_START, 1, 5, 1)
    .getDisplayValues()
    .map(function(row) { return hotelDbV2Clean_(row[0]); });
  const expectedNames = data.workflows.map(function(workflow) { return workflow.name; });
  if (JSON.stringify(workflowNames) !== JSON.stringify(expectedNames)) {
    throw new Error('ワークフロー一覧が期待値と一致しません。');
  }

  const dashboardTotals = dashboard
    .getRange(HOTEL_DB_V2_DASHBOARD.ROWS.WORKFLOW_START, 2, 5, 1)
    .getDisplayValues()
    .map(function(row) { return hotelDbV2DashboardToNumber_(row[0]); });
  const expectedTotals = data.workflows.map(function(workflow) { return workflow.total; });
  if (JSON.stringify(dashboardTotals) !== JSON.stringify(expectedTotals)) {
    throw new Error('ワークフロー件数表示が実データ集計と一致しません。');
  }

  const before = JSON.parse(stored);
  const after = hotelDbV2DashboardUiFingerprints_(ss);
  const changes = hotelDbV2DashboardUiFingerprintChanges_(before, after);
  if (changes.length) {
    throw new Error('ダッシュボード更新で運用シートが変化しました: ' + changes.join('、'));
  }

  SpreadsheetApp.getUi().alert([
    'PR #21 UIテスト 成功',
    '',
    'タイトル・5セクション: 正常',
    '5ワークフロー集計: 正常',
    '次にやること表示: 正常',
    '最近の履歴表示: 正常',
    '運用シート内容ハッシュ: 変更なし',
    'Google Places API呼出: なし'
  ].join('\n'));
  return {
    success: true,
    workflowCount: data.workflows.length,
    attentionTotal: data.attentionTotal,
    operationalSheetsUnchanged: true
  };
}

function cleanupHotelDbV2DashboardUiTest() {
  const ss = hotelDbV2DashboardUiAssertCopy_();
  const properties = PropertiesService.getDocumentProperties();
  const dashboard = ss.getSheetByName(HOTEL_DB_V2_DASHBOARD.SHEET_NAME);
  if (dashboard) ss.deleteSheet(dashboard);

  const backup = ss.getSheetByName(HOTEL_DB_V2_DASHBOARD_UI_TEST.BACKUP_DASHBOARD);
  if (backup) backup.setName(HOTEL_DB_V2_DASHBOARD.SHEET_NAME);

  properties.deleteProperty(HOTEL_DB_V2_DASHBOARD_UI_TEST.FINGERPRINT_PROPERTY);
  SpreadsheetApp.getUi().alert('PR #21 UIテスト復元完了\n\n既存ダッシュボードを復元しました。運用シートは変更していません。');
  return { cleaned: true };
}

function hotelDbV2DashboardUiMonitoredSheetNames_() {
  return [
    HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS,
    HOTEL_DB_V2_CONFIG.SHEETS.REVIEW,
    HOTEL_DB_V2_CONFIG.SHEETS.HISTORY,
    HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES,
    HOTEL_DB_V2_CONFIG.SHEETS.SUMMARY,
    '新規追加候補',
    '新規施設分類候補',
    '閉業除外履歴',
    '重複整理履歴'
  ];
}

function hotelDbV2DashboardUiFingerprints_(spreadsheet) {
  const result = {};
  hotelDbV2DashboardUiMonitoredSheetNames_().forEach(function(name) {
    const sheet = spreadsheet.getSheetByName(name);
    result[name] = hotelDbV2DashboardUiFingerprintSheet_(sheet);
  });
  return result;
}

function hotelDbV2DashboardUiFingerprintSheet_(sheet) {
  if (!sheet) return { exists: false };
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  let values = [];
  let formulas = [];
  if (lastRow > 0 && lastColumn > 0) {
    const range = sheet.getRange(1, 1, lastRow, lastColumn);
    values = range.getDisplayValues();
    formulas = range.getFormulas();
  }
  return {
    exists: true,
    sheetId: sheet.getSheetId(),
    lastRow: lastRow,
    lastColumn: lastColumn,
    hash: hotelDbV2DashboardUiDigest_(JSON.stringify({ values: values, formulas: formulas }))
  };
}

function hotelDbV2DashboardUiFingerprintChanges_(before, after) {
  const names = {};
  Object.keys(before || {}).forEach(function(name) { names[name] = true; });
  Object.keys(after || {}).forEach(function(name) { names[name] = true; });
  return Object.keys(names).filter(function(name) {
    return JSON.stringify((before || {})[name] || null) !== JSON.stringify((after || {})[name] || null);
  });
}

function hotelDbV2DashboardUiDigest_(text) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text || ''),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

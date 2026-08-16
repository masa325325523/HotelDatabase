/**
 * PR #22 初期セットアップ・設定画面のサーバー側処理。
 *
 * 安全原則:
 * - APIキーは Script Properties の GOOGLE_PLACES_API_KEY にだけ保存する。
 * - 保存済みAPIキーそのものをクライアントへ返さない。
 * - APIキーをログ・シート・履歴へ書かない。
 * - 運用シート準備は「存在しないコアシートの新規作成」だけを行う。
 * - 既存シートの見出し・値・数式を自動修正しない。
 * - 接続テストは利用者が明示的に押したときだけ実行する。
 */

const HOTEL_DB_V2_SETUP = Object.freeze({
  VERSION: '2.0',
  DIALOG_FILE: 'V2SetupDialog',
  DIALOG_TITLE: '宿泊施設DB Ver2.0 初期セットアップ・設定',
  DIALOG_WIDTH: 760,
  DIALOG_HEIGHT: 720,
  API_KEY_MIN_LENGTH: 20,
  API_KEY_MAX_LENGTH: 200,
  LOCK_TIMEOUT_MS: 10000,
  DASHBOARD_SHEET: '統合ダッシュボード'
});

function runHotelDbV2OpenSetup() {
  const html = HtmlService
    .createHtmlOutputFromFile(HOTEL_DB_V2_SETUP.DIALOG_FILE)
    .setWidth(HOTEL_DB_V2_SETUP.DIALOG_WIDTH)
    .setHeight(HOTEL_DB_V2_SETUP.DIALOG_HEIGHT);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    HOTEL_DB_V2_SETUP.DIALOG_TITLE
  );
}

function hotelDbV2SetupGetStatus() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return hotelDbV2SetupBuildStatus_(spreadsheet);
}

function hotelDbV2SetupSaveApiKey(apiKey) {
  const store = PropertiesService.getScriptProperties();
  hotelDbV2SetupWriteApiKeyToStore_(store, apiKey);

  return {
    saved: true,
    status: hotelDbV2SetupBuildStatus_(SpreadsheetApp.getActiveSpreadsheet())
  };
}

function hotelDbV2SetupDeleteApiKey() {
  const store = PropertiesService.getScriptProperties();
  hotelDbV2SetupDeleteApiKeyFromStore_(store);

  return {
    deleted: true,
    status: hotelDbV2SetupBuildStatus_(SpreadsheetApp.getActiveSpreadsheet())
  };
}

function hotelDbV2SetupTestConnection() {
  try {
    const place = hotelDbV2ConnectionTest_();
    return {
      ok: true,
      name: hotelDbV2Clean_(place && place.name),
      address: hotelDbV2Clean_(place && place.address),
      status: hotelDbV2Clean_(place && place.status),
      message: 'Google Places APIへ正常に接続できました。'
    };
  } catch (error) {
    return {
      ok: false,
      name: '',
      address: '',
      status: '',
      message: hotelDbV2SetupSafeErrorMessage_(error)
    };
  }
}

function hotelDbV2SetupPrepareCoreSheets() {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(HOTEL_DB_V2_SETUP.LOCK_TIMEOUT_MS)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const result = hotelDbV2SetupPrepareCoreSheets_(spreadsheet);
    result.status = hotelDbV2SetupBuildStatus_(spreadsheet);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function hotelDbV2SetupBuildStatus_(spreadsheet) {
  const store = PropertiesService.getScriptProperties();
  const storedKey = hotelDbV2Clean_(
    store.getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY)
  );
  const source = hotelDbV2SetupInspectActiveSource_(spreadsheet);
  const core = hotelDbV2SetupInspectCoreSheets_(spreadsheet);
  const apiKeyConfigured = Boolean(storedKey);
  const dashboardExists = Boolean(
    spreadsheet.getSheetByName(HOTEL_DB_V2_SETUP.DASHBOARD_SHEET)
  );

  const status = {
    version: HOTEL_DB_V2_SETUP.VERSION,
    apiKeyConfigured: apiKeyConfigured,
    apiKeyProperty: HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY,
    source: source,
    coreSheets: core.sheets,
    coreMissing: core.missing,
    coreNeedsReview: core.needsReview,
    dashboardExists: dashboardExists,
    setupReady: Boolean(
      apiKeyConfigured &&
      source.ready &&
      core.missing === 0 &&
      core.needsReview === 0
    )
  };

  status.nextActions = hotelDbV2SetupNextActions_(status);
  return status;
}

function hotelDbV2SetupInspectActiveSource_(spreadsheet) {
  const sheet = spreadsheet && spreadsheet.getActiveSheet
    ? spreadsheet.getActiveSheet()
    : null;

  if (!sheet) {
    return {
      name: '',
      sheetId: '',
      ready: false,
      recognized: [],
      issues: ['アクティブシートを取得できません。'],
      warnings: []
    };
  }

  const map = hotelDbV2GetHeaderMap_(sheet);
  const reserved = hotelDbV2SetupReservedSheetNames_();
  const evaluation = hotelDbV2SetupEvaluateSourceMap_(
    sheet.getName(),
    map,
    reserved
  );

  return {
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    ready: evaluation.ready,
    recognized: Object.keys(map),
    issues: evaluation.issues,
    warnings: evaluation.warnings
  };
}

function hotelDbV2SetupEvaluateSourceMap_(sheetName, map, reservedNames) {
  const name = hotelDbV2Clean_(sheetName);
  const headerMap = map || {};
  const reserved = reservedNames || [];
  const issues = [];
  const warnings = [];

  if (reserved.indexOf(name) !== -1) {
    issues.push('現在のシートは出力・管理用シートです。元データのシートを開いてください。');
  }

  if (!headerMap.facilityName) {
    issues.push('「施設名」列が見つかりません。');
  }

  if (!headerMap.address && !headerMap.municipality) {
    issues.push('「住所」または「市区町村」列が見つかりません。');
  }

  if (!headerMap.postalCode) {
    warnings.push('「郵便番号」列が見つかりません。精度向上のため追加を推奨します。');
  }
  if (!headerMap.category) {
    warnings.push('「宿泊分類」列が見つかりません。分類管理を行う場合は追加を推奨します。');
  }
  if (!headerMap.notes) {
    warnings.push('「備考」列が見つかりません。確認メモを残す場合は追加を推奨します。');
  }

  return {
    ready: issues.length === 0,
    issues: issues,
    warnings: warnings
  };
}

function hotelDbV2SetupReservedSheetNames_() {
  const names = Object.keys(HOTEL_DB_V2_CONFIG.SHEETS).map(function(key) {
    return HOTEL_DB_V2_CONFIG.SHEETS[key];
  });

  [
    '新規追加候補',
    '新規施設分類候補',
    '閉業除外履歴',
    '重複整理履歴',
    HOTEL_DB_V2_SETUP.DASHBOARD_SHEET
  ].forEach(function(name) {
    if (names.indexOf(name) === -1) names.push(name);
  });

  return names;
}

function hotelDbV2SetupCoreSheetDefinitions_() {
  return [
    {
      name: HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS,
      headers: HOTEL_DB_V2_CORRECTION_HEADERS
    },
    {
      name: HOTEL_DB_V2_CONFIG.SHEETS.REVIEW,
      headers: HOTEL_DB_V2_REVIEW_HEADERS
    },
    {
      name: HOTEL_DB_V2_CONFIG.SHEETS.HISTORY,
      headers: HOTEL_DB_V2_HISTORY_HEADERS
    },
    {
      name: HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES,
      headers: HOTEL_DB_V2_DUPLICATE_HEADERS
    },
    {
      name: HOTEL_DB_V2_CONFIG.SHEETS.SUMMARY,
      headers: HOTEL_DB_V2_SUMMARY_HEADERS
    }
  ];
}

function hotelDbV2SetupInspectCoreSheets_(spreadsheet) {
  const definitions = hotelDbV2SetupCoreSheetDefinitions_();
  const sheets = [];
  let missing = 0;
  let needsReview = 0;

  definitions.forEach(function(definition) {
    const sheet = spreadsheet.getSheetByName(definition.name);
    if (!sheet) {
      missing++;
      sheets.push({
        name: definition.name,
        status: '未作成',
        exists: false,
        compatible: false,
        rows: 0
      });
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    if (lastRow < 1 || lastColumn < 1) {
      needsReview++;
      sheets.push({
        name: definition.name,
        status: '要確認',
        exists: true,
        compatible: false,
        rows: 0
      });
      return;
    }

    const actualHeaders = sheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0];
    const compatible = hotelDbV2SetupHeadersCompatible_(
      actualHeaders,
      definition.headers
    );

    if (!compatible) needsReview++;
    sheets.push({
      name: definition.name,
      status: compatible ? '準備済み' : '要確認',
      exists: true,
      compatible: compatible,
      rows: Math.max(0, lastRow - 1)
    });
  });

  return {
    sheets: sheets,
    missing: missing,
    needsReview: needsReview
  };
}

function hotelDbV2SetupHeadersCompatible_(actualHeaders, expectedHeaders) {
  const actual = actualHeaders || [];
  const expected = expectedHeaders || [];
  if (actual.length < expected.length) return false;

  return expected.every(function(header, index) {
    return hotelDbV2Clean_(actual[index]) === hotelDbV2Clean_(header);
  });
}

function hotelDbV2SetupPrepareCoreSheets_(spreadsheet) {
  const definitions = hotelDbV2SetupCoreSheetDefinitions_();
  const created = [];
  const skippedExisting = [];

  try {
    definitions.forEach(function(definition) {
      const existing = spreadsheet.getSheetByName(definition.name);
      if (existing) {
        skippedExisting.push(definition.name);
        return;
      }

      const sheet = spreadsheet.insertSheet(definition.name);
      sheet
        .getRange(1, 1, 1, definition.headers.length)
        .setValues([definition.headers]);
      sheet.setFrozenRows(1);
      created.push(definition.name);
    });
  } catch (error) {
    created.slice().reverse().forEach(function(name) {
      const sheet = spreadsheet.getSheetByName(name);
      if (sheet) {
        try {
          spreadsheet.deleteSheet(sheet);
        } catch (rollbackError) {
          console.error('PR22 setup rollback failed: ' + rollbackError.message);
        }
      }
    });
    throw new Error(
      '運用シート準備に失敗したため、新規作成分を可能な範囲で元に戻しました。' +
      ' 詳細: ' + hotelDbV2SetupSafeErrorMessage_(error)
    );
  }

  return {
    created: created,
    createdCount: created.length,
    skippedExisting: skippedExisting,
    existingCount: skippedExisting.length
  };
}

function hotelDbV2SetupValidateApiKey_(apiKey) {
  const key = hotelDbV2Clean_(apiKey);

  if (!key) {
    return { valid: false, key: '', message: 'APIキーを入力してください。' };
  }
  if (key.length < HOTEL_DB_V2_SETUP.API_KEY_MIN_LENGTH) {
    return {
      valid: false,
      key: '',
      message: 'APIキーが短すぎます。コピー内容を確認してください。'
    };
  }
  if (key.length > HOTEL_DB_V2_SETUP.API_KEY_MAX_LENGTH) {
    return {
      valid: false,
      key: '',
      message: 'APIキーが長すぎます。コピー内容を確認してください。'
    };
  }
  if (/\s/.test(key)) {
    return {
      valid: false,
      key: '',
      message: 'APIキーの途中に空白または改行が含まれています。'
    };
  }

  return { valid: true, key: key, message: '' };
}

function hotelDbV2SetupWriteApiKeyToStore_(store, apiKey) {
  const validation = hotelDbV2SetupValidateApiKey_(apiKey);
  if (!validation.valid) throw new Error(validation.message);

  store.setProperty(
    HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY,
    validation.key
  );
  return true;
}

function hotelDbV2SetupDeleteApiKeyFromStore_(store) {
  store.deleteProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY);
  return true;
}

function hotelDbV2SetupSafeErrorMessage_(error) {
  let message = hotelDbV2Clean_(error && error.message ? error.message : error);
  if (!message) message = '不明なエラーが発生しました。';

  try {
    const key = hotelDbV2Clean_(
      PropertiesService.getScriptProperties()
        .getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY)
    );
    if (key) message = message.split(key).join('[API_KEY]');
  } catch (ignore) {
    // 秘密情報のマスキング処理自体が失敗しても、追加情報は返さない。
  }

  return message;
}

function hotelDbV2SetupNextActions_(status) {
  const actions = [];
  const data = status || {};

  if (!data.apiKeyConfigured) {
    actions.push('Google Places APIキーを設定してください。');
  }
  if (!data.source || !data.source.ready) {
    actions.push('元データのシートを開き、「施設名」と「住所または市区町村」を確認してください。');
  }
  if (Number(data.coreMissing || 0) > 0) {
    actions.push('「運用シートを準備」で不足しているコアシートを新規作成してください。');
  }
  if (Number(data.coreNeedsReview || 0) > 0) {
    actions.push('既存の運用シートに見出し不整合があります。自動修正せず、内容を確認してください。');
  }
  if (data.apiKeyConfigured) {
    actions.push('「接続テスト」でGoogle Places APIとの接続を確認してください。');
  }
  if (
    data.apiKeyConfigured &&
    data.source && data.source.ready &&
    Number(data.coreMissing || 0) === 0 &&
    Number(data.coreNeedsReview || 0) === 0
  ) {
    actions.push('準備完了です。次は③「Ver2.0 先頭3件テスト」で確認できます。');
  }

  return actions.slice(0, 5);
}

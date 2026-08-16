/**
 * PR #22 コピー版UIテスト。
 * 対象: 宿泊施設DB_PR13_⑧反映テスト のみ。
 *
 * 実APIキーそのものは保存・表示せず、SHA-256ハッシュだけで不変確認する。
 */

const HOTEL_DB_V2_PR22_UI_TEST = Object.freeze({
  STATE_PROPERTY: 'HOTEL_DB_V2_PR22_UI_TEST_STATE',
  COPY_NAME_PARTS: Object.freeze(['PR13', '⑧反映テスト']),
  OPERATIONAL_SHEETS: Object.freeze([
    '修正候補',
    '要確認',
    '修正履歴',
    '重複候補',
    '実行サマリー',
    '新規追加候補',
    '新規施設分類候補',
    '閉業除外履歴',
    '重複整理履歴',
    '統合ダッシュボード'
  ])
});

function setupHotelDbV2SetupUiTest() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2Pr22AssertCopy_(spreadsheet);

  const properties = PropertiesService.getDocumentProperties();
  if (properties.getProperty(HOTEL_DB_V2_PR22_UI_TEST.STATE_PROPERTY)) {
    throw new Error(
      '前回のPR #22 UIテスト状態が残っています。先に cleanupHotelDbV2SetupUiTest() を実行してください。'
    );
  }

  const originalActive = spreadsheet.getActiveSheet();
  const sourceSheet = hotelDbV2Pr22FindSourceSheet_(spreadsheet);
  if (!sourceSheet) {
    throw new Error('PR #22 UIテストに使える元データシートが見つかりません。');
  }
  sourceSheet.activate();

  const definitions = hotelDbV2SetupCoreSheetDefinitions_();
  const missingCore = definitions
    .filter(function(definition) {
      return !spreadsheet.getSheetByName(definition.name);
    })
    .map(function(definition) { return definition.name; });

  const state = {
    originalActiveSheetId: originalActive ? originalActive.getSheetId() : '',
    sourceSheetId: sourceSheet.getSheetId(),
    sourceSheetName: sourceSheet.getName(),
    missingCoreBefore: missingCore,
    createdByTest: [],
    keyHashBefore: hotelDbV2Pr22ApiKeyHash_(),
    sheetHashesBefore: hotelDbV2Pr22OperationalHashes_(spreadsheet, sourceSheet)
  };

  properties.setProperty(
    HOTEL_DB_V2_PR22_UI_TEST.STATE_PROPERTY,
    JSON.stringify(state)
  );

  const status = hotelDbV2SetupGetStatus();
  SpreadsheetApp.getUi().alert([
    'PR #22 UIテスト 準備完了',
    '',
    'テスト元データ: ' + sourceSheet.getName(),
    'APIキー状態: ' + (status.apiKeyConfigured ? '設定済み' : '未設定'),
    '元データ診断: ' + (status.source.ready ? '使用可能' : '要確認'),
    '不足コアシート: ' + status.coreMissing + '件',
    '',
    '実APIキーは保存・表示していません。',
    '次に testHotelDbV2SetupUiTest() を実行してください。'
  ].join('\n'));

  return {
    sourceSheet: sourceSheet.getName(),
    apiKeyConfigured: status.apiKeyConfigured,
    sourceReady: status.source.ready,
    coreMissing: status.coreMissing
  };
}

function testHotelDbV2SetupUiTest() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2Pr22AssertCopy_(spreadsheet);
  const properties = PropertiesService.getDocumentProperties();
  const raw = properties.getProperty(HOTEL_DB_V2_PR22_UI_TEST.STATE_PROPERTY);
  if (!raw) {
    throw new Error('先に setupHotelDbV2SetupUiTest() を実行してください。');
  }
  const state = JSON.parse(raw);
  const sourceSheet = spreadsheet.getSheetById(Number(state.sourceSheetId));
  if (!sourceSheet) throw new Error('UIテスト用の元データシートが見つかりません。');
  sourceSheet.activate();

  const failures = [];
  function check(label, condition) {
    if (!condition) failures.push(label);
  }

  const status = hotelDbV2SetupGetStatus();
  check('元データ診断が使用可能ではありません。', status.source && status.source.ready === true);
  check('元データシート名が一致しません。', status.source && status.source.name === state.sourceSheetName);
  check('APIキー状態フラグがbooleanではありません。', typeof status.apiKeyConfigured === 'boolean');
  check('APIキー本体らしきフィールドがあります。', !Object.prototype.hasOwnProperty.call(status, 'apiKey'));

  const actualKey = hotelDbV2Clean_(
    PropertiesService.getScriptProperties()
      .getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY)
  );
  const statusJson = JSON.stringify(status);
  if (actualKey) {
    check('保存済みAPIキーがstatus応答へ露出しています。', statusJson.indexOf(actualKey) === -1);
  }

  const html = HtmlService
    .createHtmlOutputFromFile(HOTEL_DB_V2_SETUP.DIALOG_FILE)
    .getContent();
  check('設定画面HTMLを読み込めません。', Boolean(html));
  check('設定画面タイトルがありません。', html.indexOf('初期セットアップ・設定') !== -1);
  check('APIキー入力がpassword型ではありません。', html.indexOf('type="password"') !== -1);
  check('保存済みキーを再表示しない説明がありません。', html.indexOf('再表示しません') !== -1);
  if (actualKey) {
    check('設定画面HTMLへAPIキーが埋め込まれています。', html.indexOf(actualKey) === -1);
  }

  const prepareResult = hotelDbV2SetupPrepareCoreSheets();
  const expectedCreated = state.missingCoreBefore.slice().sort();
  const actualCreated = (prepareResult.created || []).slice().sort();
  check(
    '不足コアシート以外を新規作成しました。',
    JSON.stringify(expectedCreated) === JSON.stringify(actualCreated)
  );

  state.createdByTest = actualCreated;
  properties.setProperty(
    HOTEL_DB_V2_PR22_UI_TEST.STATE_PROPERTY,
    JSON.stringify(state)
  );

  actualCreated.forEach(function(name) {
    const definition = hotelDbV2SetupCoreSheetDefinitions_().filter(function(item) {
      return item.name === name;
    })[0];
    const sheet = spreadsheet.getSheetByName(name);
    check('作成シートが見つかりません: ' + name, Boolean(sheet));
    if (sheet && definition) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
      check(
        '新規作成シートの見出しが不正です: ' + name,
        hotelDbV2SetupHeadersCompatible_(headers, definition.headers)
      );
    }
  });

  const keyHashAfter = hotelDbV2Pr22ApiKeyHash_();
  check('UIテスト中にScript PropertiesのAPIキーが変わりました。', keyHashAfter === state.keyHashBefore);

  const hashesAfter = hotelDbV2Pr22OperationalHashes_(spreadsheet, sourceSheet);
  Object.keys(state.sheetHashesBefore).forEach(function(name) {
    const before = state.sheetHashesBefore[name];
    if (before === '__MISSING__' && actualCreated.indexOf(name) !== -1) return;
    check(
      '運用シート内容・数式が変わりました: ' + name,
      hashesAfter[name] === before
    );
  });

  const refreshed = hotelDbV2SetupGetStatus();
  check('準備後もコア不足が残っています。', refreshed.coreMissing === 0);

  if (failures.length) {
    SpreadsheetApp.getUi().alert([
      'PR #22 UIテスト 失敗',
      '',
      failures.join('\n'),
      '',
      'cleanupは実行せず、状態を確認してください。'
    ].join('\n'));
    throw new Error('PR #22 UIテスト失敗: ' + failures.join(' / '));
  }

  SpreadsheetApp.getUi().alert([
    'PR #22 UIテスト 成功',
    '',
    '設定状態取得: 正常',
    'APIキー本体の非露出: 成功',
    '設定画面HTML読込: 正常',
    '不足コアシートだけ新規作成: 成功',
    '既存運用シート内容・数式: 変更なし',
    'Script Properties APIキー: 変更なし',
    'Google Places API呼出: なし',
    '',
    '次に「⚙ 初期セットアップ・設定」を開き、画面を目視確認してください。'
  ].join('\n'));

  return {
    success: true,
    created: actualCreated.length,
    apiKeyUnchanged: true,
    operationalSheetsUnchanged: true
  };
}

function cleanupHotelDbV2SetupUiTest() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  hotelDbV2Pr22AssertCopy_(spreadsheet);
  const properties = PropertiesService.getDocumentProperties();
  const raw = properties.getProperty(HOTEL_DB_V2_PR22_UI_TEST.STATE_PROPERTY);
  if (!raw) {
    SpreadsheetApp.getUi().alert('PR #22 UIテストの復元対象はありません。');
    return { restored: true, nothingToDo: true };
  }

  const state = JSON.parse(raw);
  const created = state.createdByTest || [];

  created.slice().reverse().forEach(function(name) {
    if ((state.missingCoreBefore || []).indexOf(name) === -1) return;
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) spreadsheet.deleteSheet(sheet);
  });

  const sourceSheet = spreadsheet.getSheetById(Number(state.sourceSheetId));
  const hashesAfterCleanup = hotelDbV2Pr22OperationalHashes_(spreadsheet, sourceSheet);
  const mismatches = [];
  Object.keys(state.sheetHashesBefore || {}).forEach(function(name) {
    if (hashesAfterCleanup[name] !== state.sheetHashesBefore[name]) {
      mismatches.push(name);
    }
  });

  if (hotelDbV2Pr22ApiKeyHash_() !== state.keyHashBefore) {
    mismatches.push('Script Properties APIキー');
  }

  if (mismatches.length) {
    throw new Error('cleanup後の復元確認に失敗しました: ' + mismatches.join('、'));
  }

  const original = spreadsheet.getSheetById(Number(state.originalActiveSheetId));
  if (original) original.activate();
  properties.deleteProperty(HOTEL_DB_V2_PR22_UI_TEST.STATE_PROPERTY);

  SpreadsheetApp.getUi().alert([
    'PR #22 UIテスト復元完了',
    '',
    'テストで新規作成したシートを元の状態へ戻しました。',
    '既存運用シートは変更していません。',
    'Script PropertiesのAPIキーも変更していません。'
  ].join('\n'));

  return { restored: true, mismatches: 0 };
}

function hotelDbV2Pr22AssertCopy_(spreadsheet) {
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_PR22_UI_TEST.COPY_NAME_PARTS.every(function(part) {
    return name.indexOf(part) !== -1;
  });
  if (!ok) {
    throw new Error(
      'PR #22 UIテストはコピー版「宿泊施設DB_PR13_⑧反映テスト」でのみ実行できます。'
    );
  }
}

function hotelDbV2Pr22FindSourceSheet_(spreadsheet) {
  const reserved = hotelDbV2SetupReservedSheetNames_();
  const sheets = spreadsheet.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];
    if (reserved.indexOf(sheet.getName()) !== -1) continue;
    const evaluation = hotelDbV2SetupEvaluateSourceMap_(
      sheet.getName(), hotelDbV2GetHeaderMap_(sheet), reserved
    );
    if (evaluation.ready) return sheet;
  }
  return null;
}

function hotelDbV2Pr22OperationalHashes_(spreadsheet, sourceSheet) {
  const result = {};
  HOTEL_DB_V2_PR22_UI_TEST.OPERATIONAL_SHEETS.forEach(function(name) {
    result[name] = hotelDbV2Pr22SheetHash_(spreadsheet.getSheetByName(name));
  });
  if (sourceSheet) {
    result['__SOURCE__' + sourceSheet.getSheetId()] = hotelDbV2Pr22SheetHash_(sourceSheet);
  }
  return result;
}

function hotelDbV2Pr22SheetHash_(sheet) {
  if (!sheet) return '__MISSING__';
  const lastRow = Math.max(1, sheet.getLastRow());
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const range = sheet.getRange(1, 1, lastRow, lastColumn);
  return hotelDbV2Pr22Sha256_(JSON.stringify({
    name: sheet.getName(),
    values: range.getDisplayValues(),
    formulas: range.getFormulas()
  }));
}

function hotelDbV2Pr22ApiKeyHash_() {
  const key = hotelDbV2Clean_(
    PropertiesService.getScriptProperties()
      .getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY)
  );
  return key ? hotelDbV2Pr22Sha256_(key) : '__NOT_SET__';
}

function hotelDbV2Pr22Sha256_(value) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return digest.map(function(byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

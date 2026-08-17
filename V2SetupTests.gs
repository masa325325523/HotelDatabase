/**
 * PR #22 初期セットアップ・設定画面 自己診断。
 * 外部APIを呼ばず、実際のScript Properties・シートを書き換えない。
 */

function runHotelDbV2SetupTests() {
  const failures = [];
  let passed = 0;

  function check(label, condition) {
    if (condition) {
      passed++;
    } else {
      failures.push(label);
    }
  }

  function hasText(items, text) {
    return (items || []).some(function(item) {
      return String(item).indexOf(text) !== -1;
    });
  }

  // 1) APIキー検証: 11項目
  let result = hotelDbV2SetupValidateApiKey_('');
  check('API-01 空文字は無効', result.valid === false);
  result = hotelDbV2SetupValidateApiKey_('   ');
  check('API-02 空白だけは無効', result.valid === false);
  result = hotelDbV2SetupValidateApiKey_('1234567890123456789');
  check('API-03 19文字は無効', result.valid === false);
  result = hotelDbV2SetupValidateApiKey_('12345678901234567890');
  check('API-04 20文字は有効', result.valid === true);
  result = hotelDbV2SetupValidateApiKey_(new Array(201).join('a'));
  check('API-05 200文字は有効', result.valid === true);
  result = hotelDbV2SetupValidateApiKey_(new Array(202).join('a'));
  check('API-06 201文字は無効', result.valid === false);
  result = hotelDbV2SetupValidateApiKey_('1234567890 1234567890');
  check('API-07 途中空白は無効', result.valid === false);
  result = hotelDbV2SetupValidateApiKey_('1234567890\n1234567890');
  check('API-08 途中改行は無効', result.valid === false);
  result = hotelDbV2SetupValidateApiKey_('  12345678901234567890  ');
  check('API-09 外側空白は除去して有効', result.valid === true);
  check('API-10 保存候補値はtrim済み', result.key === '12345678901234567890');
  result = hotelDbV2SetupValidateApiKey_('example_key_1234567890123456789012345');
  check('API-11 記号を含む十分な長さのキー文字列を許容', result.valid === true);

  // 2) Script Properties書込みの純粋テスト: 6項目
  const fakeStore = {
    values: {},
    setProperty: function(key, value) { this.values[key] = value; },
    deleteProperty: function(key) { delete this.values[key]; }
  };
  check(
    'STORE-01 書込み成功',
    hotelDbV2SetupWriteApiKeyToStore_(fakeStore, '12345678901234567890') === true
  );
  check(
    'STORE-02 正しいプロパティ名',
    Object.prototype.hasOwnProperty.call(fakeStore.values, 'GOOGLE_PLACES_API_KEY')
  );
  check(
    'STORE-03 入力値を正しく保存',
    fakeStore.values.GOOGLE_PLACES_API_KEY === '12345678901234567890'
  );
  let invalidThrew = false;
  try {
    hotelDbV2SetupWriteApiKeyToStore_(fakeStore, 'short');
  } catch (error) {
    invalidThrew = true;
  }
  check('STORE-04 無効キーは例外', invalidThrew);
  check(
    'STORE-05 無効入力で既存値を壊さない',
    fakeStore.values.GOOGLE_PLACES_API_KEY === '12345678901234567890'
  );
  hotelDbV2SetupDeleteApiKeyFromStore_(fakeStore);
  check(
    'STORE-06 削除で対象プロパティだけ消える',
    !Object.prototype.hasOwnProperty.call(fakeStore.values, 'GOOGLE_PLACES_API_KEY')
  );

  // 3) 元データ判定: 15項目
  const fullMap = {
    postalCode: 1,
    municipality: 2,
    address: 3,
    facilityName: 4,
    category: 5,
    notes: 6
  };
  result = hotelDbV2SetupEvaluateSourceMap_('大阪市旭区', fullMap, []);
  check('SRC-01 完全な元データはready', result.ready === true);
  check('SRC-02 完全な元データはissuesなし', result.issues.length === 0);
  check('SRC-03 完全な元データはwarningsなし', result.warnings.length === 0);

  result = hotelDbV2SetupEvaluateSourceMap_('元DB', { address: 1 }, []);
  check('SRC-04 施設名なしはready=false', result.ready === false);
  check('SRC-05 施設名なし理由', hasText(result.issues, '施設名'));

  result = hotelDbV2SetupEvaluateSourceMap_('元DB', { facilityName: 1 }, []);
  check('SRC-06 住所系なしはready=false', result.ready === false);
  check('SRC-07 住所系なし理由', hasText(result.issues, '住所'));

  result = hotelDbV2SetupEvaluateSourceMap_('元DB', { facilityName: 1, address: 2 }, []);
  check('SRC-08 住所だけでも必須条件を満たす', result.ready === true);
  result = hotelDbV2SetupEvaluateSourceMap_('元DB', { facilityName: 1, municipality: 2 }, []);
  check('SRC-09 市区町村だけでも必須条件を満たす', result.ready === true);

  result = hotelDbV2SetupEvaluateSourceMap_('元DB', { facilityName: 1, address: 2 }, []);
  check('SRC-10 郵便番号なし警告', hasText(result.warnings, '郵便番号'));
  check('SRC-11 宿泊分類なし警告', hasText(result.warnings, '宿泊分類'));
  check('SRC-12 備考なし警告', hasText(result.warnings, '備考'));

  result = hotelDbV2SetupEvaluateSourceMap_('修正候補', fullMap, ['修正候補']);
  check('SRC-13 予約シートはready=false', result.ready === false);
  check('SRC-14 予約シート理由', hasText(result.issues, '出力・管理用'));
  result = hotelDbV2SetupEvaluateSourceMap_('通常元DB', fullMap, ['修正候補']);
  check('SRC-15 非予約シートはready', result.ready === true);

  // 4) 予約済みシート名: 11項目
  const reserved = hotelDbV2SetupReservedSheetNames_();
  [
    '修正候補', '要確認', '修正履歴', '重複候補', '実行サマリー',
    '新規追加候補', '新規施設分類候補', '閉業除外履歴', '重複整理履歴',
    '統合ダッシュボード'
  ].forEach(function(name, index) {
    check('RES-' + String(index + 1).padStart(2, '0') + ' ' + name, reserved.indexOf(name) !== -1);
  });
  check(
    'RES-11 予約名に重複なし',
    reserved.filter(function(value, index, array) {
      return array.indexOf(value) === index;
    }).length === reserved.length
  );

  // 5) 見出し互換判定: 8項目
  check(
    'HDR-01 完全一致',
    hotelDbV2SetupHeadersCompatible_(['A', 'B'], ['A', 'B']) === true
  );
  check(
    'HDR-02 末尾追加列は許容',
    hotelDbV2SetupHeadersCompatible_(['A', 'B', 'C'], ['A', 'B']) === true
  );
  check(
    'HDR-03 列不足は不適合',
    hotelDbV2SetupHeadersCompatible_(['A'], ['A', 'B']) === false
  );
  check(
    'HDR-04 先頭不一致は不適合',
    hotelDbV2SetupHeadersCompatible_(['X', 'B'], ['A', 'B']) === false
  );
  check(
    'HDR-05 中間不一致は不適合',
    hotelDbV2SetupHeadersCompatible_(['A', 'X', 'C'], ['A', 'B', 'C']) === false
  );
  check(
    'HDR-06 actual空は不適合',
    hotelDbV2SetupHeadersCompatible_([], ['A']) === false
  );
  check(
    'HDR-07 expected空は適合',
    hotelDbV2SetupHeadersCompatible_(['A'], []) === true
  );
  check(
    'HDR-08 前後空白はcleanして比較',
    hotelDbV2SetupHeadersCompatible_([' A ', 'B'], ['A', 'B']) === true
  );

  // 6) 次アクション: 10項目
  result = hotelDbV2SetupNextActions_({
    apiKeyConfigured: false,
    source: { ready: false },
    coreMissing: 2,
    coreNeedsReview: 0
  });
  check('NEXT-01 APIキー案内', hasText(result, 'APIキー'));
  check('NEXT-02 元データ案内', hasText(result, '元データ'));
  check('NEXT-03 運用シート準備案内', hasText(result, '運用シート'));
  check('NEXT-04 未設定時は接続テスト案内なし', !hasText(result, '接続テスト'));

  result = hotelDbV2SetupNextActions_({
    apiKeyConfigured: true,
    source: { ready: true },
    coreMissing: 0,
    coreNeedsReview: 0
  });
  check('NEXT-05 設定済みは接続テスト案内', hasText(result, '接続テスト'));
  check('NEXT-06 準備完了は③案内', hasText(result, '③'));
  check('NEXT-07 最大5件', result.length <= 5);

  result = hotelDbV2SetupNextActions_({
    apiKeyConfigured: true,
    source: { ready: true },
    coreMissing: 0,
    coreNeedsReview: 1
  });
  check('NEXT-08 見出し不整合案内', hasText(result, '見出し不整合'));

  result = hotelDbV2SetupNextActions_({
    apiKeyConfigured: true,
    source: { ready: true },
    coreMissing: 1,
    coreNeedsReview: 0
  });
  check('NEXT-09 コア不足時は③案内なし', !hasText(result, '③'));
  check('NEXT-10 コア不足でも接続テスト案内あり', hasText(result, '接続テスト'));

  // 7) コアシート定義: 12項目
  const definitions = hotelDbV2SetupCoreSheetDefinitions_();
  check('DEF-01 コアシートは5種類', definitions.length === 5);
  check(
    'DEF-02 シート名は一意',
    definitions.map(function(item) { return item.name; })
      .filter(function(value, index, array) { return array.indexOf(value) === index; })
      .length === 5
  );
  const expectedNames = ['修正候補', '要確認', '修正履歴', '重複候補', '実行サマリー'];
  expectedNames.forEach(function(name, index) {
    check(
      'DEF-' + String(index + 3).padStart(2, '0') + ' ' + name,
      definitions[index] && definitions[index].name === name
    );
  });
  definitions.forEach(function(item, index) {
    check(
      'DEF-' + String(index + 8).padStart(2, '0') + ' 見出し定義あり',
      Array.isArray(item.headers) && item.headers.length > 0
    );
  });

  // 8) 固定安全契約: 5項目
  check(
    'SAFE-01 Script Property名固定',
    HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY === 'GOOGLE_PLACES_API_KEY'
  );
  check('SAFE-02 APIキー最小長20', HOTEL_DB_V2_SETUP.API_KEY_MIN_LENGTH === 20);
  check('SAFE-03 APIキー最大長200', HOTEL_DB_V2_SETUP.API_KEY_MAX_LENGTH === 200);
  check('SAFE-04 HTMLファイル名固定', HOTEL_DB_V2_SETUP.DIALOG_FILE === 'V2SetupDialog');
  check('SAFE-05 ダッシュボード名固定', HOTEL_DB_V2_SETUP.DASHBOARD_SHEET === '統合ダッシュボード');

  // 11 + 6 + 15 + 11 + 8 + 10 + 12 + 5 = 78
  const expectedCount = 78;
  if (passed + failures.length !== expectedCount) {
    failures.push(
      'TEST-COUNT 自己診断定義数が期待値と不一致: actual=' +
      (passed + failures.length) + ', expected=' + expectedCount
    );
  }

  const lines = [
    failures.length ? 'PR #22 自己診断 失敗' : 'PR #22 自己診断 成功',
    '',
    '成功件数: ' + passed + '件',
    '失敗件数: ' + failures.length + '件',
    'Google Places API呼出: なし',
    '実Script Properties変更: なし',
    '元DB・候補・履歴の変更: なし'
  ];

  if (failures.length) {
    lines.push('', '失敗:', failures.join('\n'));
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));

  if (failures.length) {
    throw new Error('PR #22 自己診断で失敗があります: ' + failures.join(' / '));
  }

  return {
    success: passed,
    failed: failures.length,
    expected: expectedCount
  };
}

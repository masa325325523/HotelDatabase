/**
 * Code.gs
 * Google Places照合の実行入口とスプレッドシート用メニュー。
 *
 * 前提:
 * - Places.gs が同じApps Scriptプロジェクトに存在すること
 * - スクリプトプロパティ GOOGLE_PLACES_API_KEY が設定済みであること
 * - 1行目に「施設名」または「住所」の見出しがあること
 */

const APP_RUNNER = Object.freeze({
  TEST_ROWS: 3,
  DEFAULT_BATCH_SIZE: 100,
  LOCK_TIMEOUT_MS: 5000
});

/**
 * スプレッドシートを開いたときに専用メニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('宿泊施設DB')
    .addItem('① 設定・見出しを診断', 'runPlacesDiagnosis')
    .addItem('② Places API接続テスト', 'runPlacesApiConnectionTest')
    .addSeparator()
    .addItem('③ 先頭3件だけテスト', 'runPlacesTest3')
    .addItem('④ Places照合を実行', 'runPlacesOnly')
    .addItem('⑤ 既存Place IDを再確認', 'runRefreshExistingPlaces')
    .addSeparator()
    .addItem('⑥ 重複候補を確認', 'runDuplicateCheck')
    .addToUi();
}

/**
 * APIキー、対象シート、必要見出しを診断する。
 */
function runPlacesDiagnosis() {
  const result = diagnosePlacesConfiguration();
  const lines = [
    'Places設定診断',
    '',
    'APIキー: ' + (result.apiKeyConfigured ? '設定済み' : '未設定'),
    '対象シート: ' + (result.activeSheet || '取得不可'),
    '認識した列: ' + Object.keys(result.headerMap || {}).join(', ')
  ];

  if (result.errors && result.errors.length) {
    lines.push('', 'エラー:', result.errors.join('\n'));
  } else {
    lines.push('', '診断結果: 問題ありません。');
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
  return result;
}

/**
 * 1回だけPlaces APIへ検索を送り、接続を確認する。
 */
function runPlacesApiConnectionTest() {
  return withPlacesExecutionLock_('Places API接続テスト', function() {
    const results = searchPlaceFromGoogle(
      'ホテルルートイン名古屋今池駅前 愛知県名古屋市千種区'
    );

    if (!results || !results.length) {
      throw new Error('検索結果がありませんでした。APIキー・API有効化・請求設定を確認してください。');
    }

    const place = results[0];
    SpreadsheetApp.getUi().alert([
      '接続成功',
      '',
      '施設名: ' + safeString(place.name),
      '住所: ' + safeString(place.formatted_address),
      'Place ID: ' + safeString(place.place_id),
      '営業状態: ' + translateBusinessStatus(place.business_status)
    ].join('\n'));

    return place;
  });
}

/**
 * アクティブシートの先頭3件だけを照合する安全テスト。
 * 本番データを削除せず、出力列を追加して結果を書き込む。
 */
function runPlacesTest3() {
  return withPlacesExecutionLock_('先頭3件テスト', function() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    const response = ui.alert(
      '先頭3件だけテスト',
      '現在のシート「' + sheet.getName() + '」の2～4行目を照合します。\n' +
      'Google情報の出力列が不足している場合は右端へ追加します。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return { cancelled: true };
    }

    addPlacesOutputHeaders();
    const summary = enrichSheetWithPlaces(sheet, {
      startRow: 2,
      maxRows: APP_RUNNER.TEST_ROWS,
      skipExisting: false
    });

    showPlacesSummary_('先頭3件テスト完了', summary);
    return summary;
  });
}

/**
 * アクティブシートを最大100件ずつ照合する通常実行。
 * Place IDが保存済みの行は検索を省略する。
 */
function runPlacesOnly() {
  return withPlacesExecutionLock_('Places照合', function() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    const response = ui.alert(
      'Places照合を実行',
      '現在のシート「' + sheet.getName() + '」を最大' +
      APP_RUNNER.DEFAULT_BATCH_SIZE + '件処理します。\n' +
      'Place ID保存済みの行はスキップします。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return { cancelled: true };
    }

    addPlacesOutputHeaders();
    const summary = enrichSheetWithPlaces(sheet, {
      startRow: 2,
      maxRows: APP_RUNNER.DEFAULT_BATCH_SIZE,
      skipExisting: true
    });

    showPlacesSummary_('Places照合完了', summary);
    return summary;
  });
}

/**
 * 保存済みPlace IDを使って、営業状態や住所などを再取得する。
 */
function runRefreshExistingPlaces() {
  return withPlacesExecutionLock_('Place ID再確認', function() {
    const result = refreshExistingPlaceDetails();
    SpreadsheetApp.getUi().alert(
      'Place ID再確認完了\n\n更新件数: ' + Number(result.updated || 0) + '件'
    );
    return result;
  });
}

/**
 * 施設名＋住所の完全重複候補を実行ログとダイアログへ表示する。
 */
function runDuplicateCheck() {
  const duplicates = findDuplicateFacilities();
  const preview = duplicates.slice(0, 10).map(function(item) {
    return '行' + item.firstRow + ' と 行' + item.duplicateRow;
  });

  const lines = [
    '重複候補確認',
    '',
    '候補数: ' + duplicates.length + '件'
  ];

  if (preview.length) {
    lines.push('', '先頭10件:', preview.join('\n'));
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
  return duplicates;
}

/**
 * 同時実行を防ぎ、例外を利用者向けダイアログに変換する。
 */
function withPlacesExecutionLock_(label, callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(APP_RUNNER.LOCK_TIMEOUT_MS)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  try {
    return callback();
  } catch (error) {
    console.error(label + ': ' + error.stack);
    SpreadsheetApp.getUi().alert(
      label + 'でエラーが発生しました。\n\n' + error.message
    );
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 一括処理結果を読みやすい形で表示する。
 */
function showPlacesSummary_(title, summary) {
  const errors = summary.errors || [];
  const lines = [
    title,
    '',
    '処理件数: ' + Number(summary.processed || 0),
    '更新件数: ' + Number(summary.updated || 0),
    'スキップ: ' + Number(summary.skipped || 0),
    'エラー: ' + errors.length
  ];

  if (summary.nextStartRow) {
    lines.push('次回開始行: ' + summary.nextStartRow);
  }

  if (errors.length) {
    const preview = errors.slice(0, 5).map(function(item) {
      return '行' + item.row + ': ' + item.message;
    });
    lines.push('', '先頭5件のエラー:', preview.join('\n'));
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

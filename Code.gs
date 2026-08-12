/**
 * Code.gs
 * 全国宿泊施設データベース Ver2.0 実行入口
 */

const HOTEL_DB_V2_RUNNER = Object.freeze({
  LOCK_TIMEOUT_MS: 10000
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('宿泊施設DB')
    .addItem('① Ver2.0 設定・見出し診断', 'runHotelDbV2Diagnosis')
    .addItem('② Ver2.0 API接続テスト', 'runHotelDbV2ConnectionTest')
    .addSeparator()
    .addItem('③ Ver2.0 先頭3件テスト', 'runHotelDbV2Test3')
    .addItem('④ Ver2.0 本番バッチ実行（50件）', 'runHotelDbV2Batch')
    .addItem('⑤ 保存済みPlace IDを再確認（50件）', 'runHotelDbV2RefreshExisting')
    .addItem('⑥ 続きの開始行をリセット', 'runHotelDbV2ResetCheckpoint')
    .addSeparator()
    .addItem('⑦ 重複候補を更新', 'runHotelDbV2Duplicates')
    .addItem('⑧ 承認済み修正候補を反映', 'runHotelDbV2ApplyApprovedCorrections')
    .addItem('⑨ 修正候補を自動仕分け', 'runHotelDbV2TriageCorrections')
    .addItem('⑩ 重複候補を自動仕分け', 'runHotelDbV2DuplicateTriage')
    .addItem('⑪ 低スコア要確認を自動仕分け', 'runHotelDbV2LowScoreReviewTriageWithMunicipalityRefinement')
    .addToUi();
}

function runHotelDbV2Diagnosis() {
  const result = HOTEL_DB_V2.diagnose();
  const lines = [
    '宿泊施設DB Ver2.0 診断',
    '',
    'API: ' + result.apiVersion,
    'APIキー: ' + (result.apiKeyConfigured ? '設定済み' : '未設定'),
    '対象シート: ' + (result.activeSheet || '取得不可'),
    '認識した列: ' + Object.keys(result.headerMap || {}).join(', '),
    '本番処理の次回開始行: ' + (result.checkpoint || 2)
  ];

  if (result.errors && result.errors.length) {
    lines.push('', 'エラー:', result.errors.join('\n'));
  } else {
    lines.push('', '診断結果: 問題ありません。');
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
  return result;
}

function runHotelDbV2ConnectionTest() {
  return withHotelDbV2Lock_('Ver2.0 API接続テスト', function() {
    const place = HOTEL_DB_V2.connectionTest();
    SpreadsheetApp.getUi().alert([
      '接続成功',
      '',
      '施設名: ' + place.name,
      '住所: ' + place.address,
      'Place ID: ' + place.placeId,
      '営業状態: ' + place.status
    ].join('\n'));
    return place;
  });
}

function runHotelDbV2Test3() {
  return withHotelDbV2Lock_('Ver2.0 先頭3件テスト', function() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const response = ui.alert(
      'Ver2.0 先頭3件テスト',
      '現在のシート「' + sheet.getName() + '」の2～4行目を処理します。\n\n' +
      '元の郵便番号・住所・施設名は自動変更しません。\n' +
      '差分は「修正候補」、閉業や低一致は「要確認」へ出力します。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const summary = HOTEL_DB_V2.test3();
    showHotelDbV2Summary_('Ver2.0 先頭3件テスト完了', summary);
    return summary;
  });
}

function runHotelDbV2Batch() {
  return withHotelDbV2Lock_('Ver2.0 本番バッチ', function() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const diagnosis = HOTEL_DB_V2.diagnose();
    const response = ui.alert(
      'Ver2.0 本番バッチ実行',
      '対象シート: 「' + sheet.getName() + '」\n' +
      '開始行: ' + (diagnosis.checkpoint || 2) + '\n' +
      '最大処理件数: 50件\n\n' +
      '元データは自動修正・自動削除しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const summary = HOTEL_DB_V2.runBatch();
    showHotelDbV2Summary_('Ver2.0 本番バッチ完了', summary);
    return summary;
  });
}

function runHotelDbV2RefreshExisting() {
  return withHotelDbV2Lock_('保存済みPlace ID再確認', function() {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const response = ui.alert(
      '保存済みPlace IDを再確認',
      '現在のシート「' + sheet.getName() + '」で、Place IDが保存済みの行だけを' +
      '最大50件再確認します。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const summary = HOTEL_DB_V2.refreshExisting();
    showHotelDbV2Summary_('Place ID再確認完了', summary);
    return summary;
  });
}

function runHotelDbV2ResetCheckpoint() {
  const result = HOTEL_DB_V2.resetCheckpoint();
  SpreadsheetApp.getUi().alert(
    '開始行をリセットしました。\n\n対象シート: ' + result.sheet + '\n次回は2行目から処理します。'
  );
  return result;
}

function runHotelDbV2Duplicates() {
  return withHotelDbV2Lock_('重複候補更新', function() {
    const result = HOTEL_DB_V2.refreshDuplicates();
    SpreadsheetApp.getUi().alert(
      '重複候補を更新しました。\n\n候補数: ' + result.candidates + '件\n' +
      '自動削除はしていません。「重複候補」シートで確認してください。'
    );
    return result;
  });
}

function runHotelDbV2ApplyApprovedCorrections() {
  return withHotelDbV2Lock_('承認済み修正反映', function() {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '承認済み修正候補を反映',
      '「修正候補」シートの状態が「承認」の行だけ、元シートへ反映します。\n\n' +
      '元データが候補作成後に変更されている場合は反映せず「要再確認」にします。\n' +
      '続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = HOTEL_DB_V2.applyApprovedCorrections();
    ui.alert([
      '承認済み修正の反映完了',
      '',
      '承認対象: ' + result.approved,
      '反映済み: ' + result.applied,
      '要再確認: ' + result.conflicts,
      'エラー: ' + result.errors
    ].join('\n'));
    return result;
  });
}

function withHotelDbV2Lock_(label, callback) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(HOTEL_DB_V2_RUNNER.LOCK_TIMEOUT_MS)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }

  try {
    return callback();
  } catch (error) {
    console.error(label + ': ' + (error.stack || error.message));
    SpreadsheetApp.getUi().alert(
      label + 'でエラーが発生しました。\n\n' + error.message
    );
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function showHotelDbV2Summary_(title, summary) {
  const lines = [
    title,
    '',
    '処理件数: ' + Number(summary.processed || 0),
    '営業中: ' + Number(summary.operational || 0),
    '修正候補: ' + Number(summary.corrections || 0),
    '要確認: ' + Number(summary.needReview || 0),
    '候補なし: ' + Number(summary.notFound || 0),
    '閉業: ' + Number(summary.closed || 0),
    '一時休業: ' + Number(summary.temporaryClosed || 0),
    '開業予定: ' + Number(summary.futureOpening || 0),
    'スキップ: ' + Number(summary.skipped || 0),
    'エラー: ' + Number(summary.errors || 0),
    '集計整合: ' + (summary.reconciliation || '未確認')
  ];

  if (summary.nextStartRow) {
    lines.push('', '次回開始行: ' + summary.nextStartRow);
  } else {
    lines.push('', 'このシートの対象範囲は完了しました。');
  }

  SpreadsheetApp.getUi().alert(lines.join('\n'));
}

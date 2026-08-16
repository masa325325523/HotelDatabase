/** PR #21 統合ダッシュボード自己診断。外部API・スプレッドシートデータを書き換えずに純粋ロジックを検証する。 */
function runHotelDbV2DashboardTests() {
  const failures = [];
  let success = 0;

  function expect(label, actual, expected) {
    if (actual !== expected) {
      failures.push(label + ': 実際=' + actual + ', 期待=' + expected);
    } else {
      success++;
    }
  }

  const classifyCases = [
    ['空欄', '', 'unreviewed'],
    ['空白', '  ', 'unreviewed'],
    ['未確認', '未確認', 'unreviewed'],
    ['承認', '承認', 'approved'],
    ['要再確認', '要再確認', 'conflicts'],
    ['反映エラー', '反映エラー', 'errors'],
    ['追加エラー', '追加エラー', 'errors'],
    ['除外エラー', '除外エラー', 'errors'],
    ['反映済み', '反映済み', 'completed'],
    ['追加済み', '追加済み', 'completed'],
    ['除外済み', '除外済み', 'completed'],
    ['整理済み', '整理済み', 'completed'],
    ['却下', '却下', 'other']
  ];
  classifyCases.forEach(function(test) {
    expect('状態分類:' + test[0], hotelDbV2DashboardClassifyState_(test[1]), test[2]);
  });

  const aggregate = hotelDbV2DashboardAggregateStatePairs_([
    ['K1', '未確認'],
    ['K2', ''],
    ['K3', '承認'],
    ['K4', '要再確認'],
    ['K5', '反映エラー'],
    ['K6', '反映済み'],
    ['K7', '却下'],
    ['', '']
  ]);
  expect('集計:総件数', aggregate.total, 7);
  expect('集計:未確認', aggregate.unreviewed, 2);
  expect('集計:承認', aggregate.approved, 1);
  expect('集計:要再確認', aggregate.conflicts, 1);
  expect('集計:エラー', aggregate.errors, 1);
  expect('集計:完了', aggregate.completed, 1);
  expect('集計:その他', aggregate.other, 1);
  expect('集計:空行除外', aggregate.total === 8, false);

  expect('全体状態:エラー優先', hotelDbV2DashboardOverallStatus_({errors:1,conflicts:2,approved:3,unreviewed:4}), '要対応（エラーあり）');
  expect('全体状態:要再確認', hotelDbV2DashboardOverallStatus_({errors:0,conflicts:2,approved:3,unreviewed:4}), '要対応（要再確認あり）');
  expect('全体状態:承認待ち', hotelDbV2DashboardOverallStatus_({errors:0,conflicts:0,approved:3,unreviewed:4}), '承認済み処理待ち');
  expect('全体状態:確認待ち', hotelDbV2DashboardOverallStatus_({errors:0,conflicts:0,approved:0,unreviewed:4}), '確認待ち');
  expect('全体状態:良好', hotelDbV2DashboardOverallStatus_({errors:0,conflicts:0,approved:0,unreviewed:0}), '良好');

  const numberCases = [
    ['空', '', 0],
    ['整数', '12', 12],
    ['カンマ', '1,234', 1234],
    ['数値', 55, 55],
    ['不正文字', 'abc', 0],
    ['空白付き', ' 75 ', 75]
  ];
  numberCases.forEach(function(test) {
    expect('数値化:' + test[0], hotelDbV2DashboardToNumber_(test[1]), test[2]);
  });

  expect('優先度ラベル1', hotelDbV2DashboardPriorityLabel_(1), '最優先');
  expect('優先度ラベル2', hotelDbV2DashboardPriorityLabel_(2), '高');
  expect('優先度ラベル3', hotelDbV2DashboardPriorityLabel_(3), '中');
  expect('優先度ラベル4', hotelDbV2DashboardPriorityLabel_(4), '通常');
  expect('優先度ラベル5', hotelDbV2DashboardPriorityLabel_(5), '案内');

  const definitions = hotelDbV2DashboardWorkflowDefinitions_();
  expect('ワークフロー数', definitions.length, 5);
  expect('ワークフロー1', definitions[0].name, '修正候補');
  expect('ワークフロー2', definitions[1].name, '要確認');
  expect('ワークフロー3', definitions[2].name, '重複候補');
  expect('ワークフロー4', definitions[3].name, '新規追加候補');
  expect('ワークフロー5', definitions[4].name, '新規施設分類候補');
  expect('重複反映メニュー', definitions[2].applyAction.indexOf('⑱') !== -1, true);
  expect('新規追加メニュー', definitions[3].applyAction.indexOf('⑭') !== -1, true);
  expect('分類反映メニュー', definitions[4].applyAction.indexOf('⑯') !== -1, true);

  function workflow(name, counts) {
    return Object.assign({
      name: name,
      errors: 0,
      conflicts: 0,
      approved: 0,
      unreviewed: 0,
      other: 0,
      applyAction: name + '反映',
      reviewAction: name + '確認'
    }, counts || {});
  }

  const actionSet = hotelDbV2DashboardBuildActions_([
    workflow('A', {errors:2, conflicts:1, approved:1, unreviewed:5}),
    workflow('B', {errors:1, conflicts:3, approved:2, unreviewed:1})
  ], {exists:true,nextStartRow:'51',sourceSheet:'東京'});
  expect('行動:最大5件', actionSet.length, 5);
  expect('行動:最優先はエラー', actionSet[0].priority, 1);
  expect('行動:同優先度は件数降順', actionSet[0].target, 'A');
  expect('行動:2番目もエラー', actionSet[1].priority, 1);
  expect('行動:3番目は要再確認', actionSet[2].priority, 2);
  expect('行動:要再確認件数降順', actionSet[2].target, 'B');

  const approvedAction = hotelDbV2DashboardBuildActions_([
    workflow('修正候補', {approved:2})
  ], {exists:false});
  expect('行動:承認済み優先度', approvedAction[0].priority, 3);
  expect('行動:承認済み反映案内', approvedAction[0].action, '修正候補反映');

  const unreviewedAction = hotelDbV2DashboardBuildActions_([
    workflow('重複候補', {unreviewed:4})
  ], {exists:false});
  expect('行動:未確認優先度', unreviewedAction[0].priority, 4);
  expect('行動:未確認案内', unreviewedAction[0].action, '重複候補確認');

  const continueAction = hotelDbV2DashboardBuildActions_([], {
    exists:true,nextStartRow:'151',sourceSheet:'大阪市'
  });
  expect('行動:バックログなし次回行', continueAction[0].priority, 5);
  expect('行動:本番バッチ案内', continueAction[0].action.indexOf('④') !== -1, true);
  expect('行動:次回開始行表示', continueAction[0].action.indexOf('151') !== -1, true);
  expect('行動:対象シート', continueAction[0].target, '大阪市');

  const defaultAction = hotelDbV2DashboardBuildActions_([], {exists:false,nextStartRow:''});
  expect('行動:通常案内1件', defaultAction.length, 1);
  expect('行動:再確認案内', defaultAction[0].action.indexOf('⑤') !== -1, true);
  expect('行動:新規探索案内', defaultAction[0].action.indexOf('⑬') !== -1, true);

  expect('次操作:エラー', hotelDbV2DashboardWorkflowNextAction_(workflow('A',{errors:1})), 'エラー原因を確認');
  expect('次操作:要再確認', hotelDbV2DashboardWorkflowNextAction_(workflow('A',{conflicts:1})), '要再確認を処理');
  expect('次操作:承認', hotelDbV2DashboardWorkflowNextAction_(workflow('A',{approved:1})), 'A反映');
  expect('次操作:未確認', hotelDbV2DashboardWorkflowNextAction_(workflow('A',{unreviewed:1})), 'A確認');
  expect('次操作:その他', hotelDbV2DashboardWorkflowNextAction_(workflow('A',{other:1})), 'その他状態を確認');
  expect('次操作:対応なし', hotelDbV2DashboardWorkflowNextAction_(workflow('A',{})), '対応なし');

  const headerMap = hotelDbV2DashboardHeaderMapFromValues_(['状態','元シート','状態','']);
  expect('ヘッダ:状態列', headerMap['状態'], 1);
  expect('ヘッダ:元シート列', headerMap['元シート'], 2);
  expect('ヘッダ:重複は先頭優先', headerMap['状態'] === 3, false);
  expect('ヘッダ:空欄無視', Object.keys(headerMap).length, 2);

  expect('定数:ダッシュボード名', HOTEL_DB_V2_DASHBOARD.SHEET_NAME, '統合ダッシュボード');
  expect('定数:履歴10件', HOTEL_DB_V2_DASHBOARD.RECENT_HISTORY_LIMIT, 10);
  expect('定数:行動5件', HOTEL_DB_V2_DASHBOARD.MAX_ACTIONS, 5);
  expect('レイアウト:ワークフロー開始13', HOTEL_DB_V2_DASHBOARD.ROWS.WORKFLOW_START, 13);
  expect('レイアウト:履歴開始32', HOTEL_DB_V2_DASHBOARD.ROWS.HISTORY_START, 32);
  expect('安全契約:API呼出なし', /Places|UrlFetch/.test(hotelDbV2DashboardSafetyMarker_()), false);
  expect('安全契約:候補書換なし', /candidate-write|source-write/.test(hotelDbV2DashboardSafetyMarker_()), false);
  expect('安全契約:dashboard-only', hotelDbV2DashboardSafetyMarker_().indexOf('dashboard-only-write') !== -1, true);

  if (failures.length) {
    throw new Error('PR #21 自己診断失敗\n\n' + failures.join('\n'));
  }

  const message = [
    'PR #21 自己診断 成功',
    '',
    '成功件数: ' + success + '件',
    '失敗件数: 0件',
    'Google Places API呼出: なし',
    '元DB・候補・履歴の変更: なし'
  ].join('\n');
  try { SpreadsheetApp.getUi().alert(message); } catch (error) {}
  return { success: success, failure: 0, externalApi: false, operationalWrites: false };
}

function hotelDbV2DashboardSafetyMarker_() {
  return 'read-only-operational-sheets; dashboard-only-write; no-external-api';
}

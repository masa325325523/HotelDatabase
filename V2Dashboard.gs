/**
 * PR #21 統合ダッシュボード。
 *
 * 安全原則:
 * - 元DB・候補シート・履歴シートは読み取りのみ。
 * - 書き換えるのは「統合ダッシュボード」シートだけ。
 * - 欠けている運用シートがあっても0件・未作成として表示し、処理を止めない。
 * - 人の承認・候補状態を自動変更しない。
 * - Google Places APIは呼び出さない。
 */

const HOTEL_DB_V2_DASHBOARD = Object.freeze({
  SHEET_NAME: '統合ダッシュボード',
  RECENT_HISTORY_LIMIT: 10,
  MAX_ACTIONS: 5,
  ROWS: Object.freeze({
    TITLE: 1,
    SUBTITLE: 2,
    STATUS: 3,
    LATEST_SECTION: 5,
    LATEST_HEADERS_1: 6,
    LATEST_VALUES_1: 7,
    LATEST_HEADERS_2: 8,
    LATEST_VALUES_2: 9,
    WORKFLOW_SECTION: 11,
    WORKFLOW_HEADERS: 12,
    WORKFLOW_START: 13,
    ACTION_SECTION: 19,
    ACTION_HEADERS: 20,
    ACTION_START: 21,
    AUDIT_SECTION: 27,
    AUDIT_VALUES: 28,
    HISTORY_SECTION: 30,
    HISTORY_HEADERS: 31,
    HISTORY_START: 32,
    FOOTNOTE: 43
  }),
  COLORS: Object.freeze({
    NAVY: '#1F4E78',
    BLUE: '#D9EAF7',
    LIGHT_BLUE: '#EAF3F8',
    GREEN: '#E2F0D9',
    YELLOW: '#FFF2CC',
    RED: '#FCE4D6',
    GRAY: '#F2F2F2',
    DARK: '#1F1F1F',
    WHITE: '#FFFFFF'
  })
});

function runHotelDbV2RefreshDashboard() {
  return withHotelDbV2Lock_('統合ダッシュボード更新', function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const data = hotelDbV2RefreshDashboard_(spreadsheet, { activate: true });
    SpreadsheetApp.getUi().alert([
      '統合ダッシュボードを更新しました。',
      '',
      '全体状態: ' + data.overallStatus,
      '要対応件数: ' + data.attentionTotal,
      '未確認: ' + data.totals.unreviewed,
      '承認済み処理待ち: ' + data.totals.approved,
      '要再確認: ' + data.totals.conflicts,
      'エラー: ' + data.totals.errors,
      '',
      '元DB・候補・履歴の変更: なし'
    ].join('\n'));
    return data;
  });
}

function hotelDbV2RefreshDashboard_(spreadsheet, options) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const opts = options || {};
  const data = hotelDbV2DashboardCollect_(ss);
  const sheet = hotelDbV2DashboardGetOrCreate_(ss);
  hotelDbV2DashboardRender_(sheet, data);

  if (opts.activate !== false) {
    sheet.activate();
    sheet.getRange('A1').activate();
  }

  return data;
}

function hotelDbV2DashboardWorkflowDefinitions_() {
  return [
    {
      name: '修正候補',
      sheetName: HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS,
      reviewAction: '⑨ 自動仕分け後に内容を確認',
      applyAction: '⑧ 承認済み修正候補を反映'
    },
    {
      name: '要確認',
      sheetName: HOTEL_DB_V2_CONFIG.SHEETS.REVIEW,
      reviewAction: '⑪/⑫で仕分け後に内容を確認',
      applyAction: '承認内容を再確認（閉業候補は⑰）'
    },
    {
      name: '重複候補',
      sheetName: HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES,
      reviewAction: '⑩で仕分けし、必要なら残す行を指定',
      applyAction: '⑱ 承認済み重複候補を安全に整理'
    },
    {
      name: '新規追加候補',
      sheetName: '新規追加候補',
      reviewAction: '候補内容・営業実態・既存重複を確認',
      applyAction: '⑭ 承認済み新規追加候補を安全に追加'
    },
    {
      name: '新規施設分類候補',
      sheetName: '新規施設分類候補',
      reviewAction: '確定宿泊分類・確定備考を人が入力',
      applyAction: '⑯ 承認済み宿泊分類・備考を安全に反映'
    }
  ];
}

function hotelDbV2DashboardCollect_(spreadsheet) {
  const workflowDefinitions = hotelDbV2DashboardWorkflowDefinitions_();
  const workflows = workflowDefinitions.map(function(definition) {
    const sheet = spreadsheet.getSheetByName(definition.sheetName);
    const summary = hotelDbV2DashboardStateSummary_(sheet);
    return Object.assign({}, definition, summary);
  });

  const totals = workflows.reduce(function(acc, workflow) {
    acc.total += workflow.total;
    acc.unreviewed += workflow.unreviewed;
    acc.approved += workflow.approved;
    acc.conflicts += workflow.conflicts;
    acc.errors += workflow.errors;
    acc.completed += workflow.completed;
    acc.other += workflow.other;
    return acc;
  }, {
    total: 0,
    unreviewed: 0,
    approved: 0,
    conflicts: 0,
    errors: 0,
    completed: 0,
    other: 0
  });

  const latestSummary = hotelDbV2DashboardLatestSummary_(spreadsheet);
  const audit = hotelDbV2DashboardAuditSummary_(spreadsheet);
  const recentHistory = hotelDbV2DashboardRecentHistory_(spreadsheet, HOTEL_DB_V2_DASHBOARD.RECENT_HISTORY_LIMIT);
  const actions = hotelDbV2DashboardBuildActions_(workflows, latestSummary);
  const overallStatus = hotelDbV2DashboardOverallStatus_(totals);
  const attentionTotal = totals.unreviewed + totals.approved + totals.conflicts + totals.errors;

  return {
    updatedAt: hotelDbV2Timestamp_(),
    spreadsheetName: spreadsheet.getName(),
    overallStatus: overallStatus,
    attentionTotal: attentionTotal,
    latestSummary: latestSummary,
    workflows: workflows,
    totals: totals,
    actions: actions,
    audit: audit,
    recentHistory: recentHistory
  };
}

function hotelDbV2DashboardStateSummary_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) {
    return hotelDbV2DashboardEmptyStateSummary_(!!sheet);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = hotelDbV2DashboardHeaderMapFromValues_(headers);
  const stateColumn = map['状態'] || 0;
  const rowCount = sheet.getLastRow() - 1;
  const keys = sheet.getRange(2, 1, rowCount, 1).getDisplayValues();
  const states = stateColumn
    ? sheet.getRange(2, stateColumn, rowCount, 1).getDisplayValues()
    : Array(rowCount).fill(['']);

  const pairs = [];
  for (let i = 0; i < rowCount; i++) {
    pairs.push([
      hotelDbV2Clean_(keys[i][0]),
      stateColumn ? hotelDbV2Clean_(states[i][0]) : ''
    ]);
  }

  const summary = hotelDbV2DashboardAggregateStatePairs_(pairs);
  summary.exists = true;
  summary.hasStateHeader = !!stateColumn;
  return summary;
}

function hotelDbV2DashboardEmptyStateSummary_(exists) {
  return {
    exists: !!exists,
    hasStateHeader: false,
    total: 0,
    unreviewed: 0,
    approved: 0,
    conflicts: 0,
    errors: 0,
    completed: 0,
    other: 0
  };
}

function hotelDbV2DashboardAggregateStatePairs_(pairs) {
  const summary = hotelDbV2DashboardEmptyStateSummary_(true);
  summary.hasStateHeader = true;

  (pairs || []).forEach(function(pair) {
    const key = hotelDbV2Clean_(pair && pair[0]);
    const state = hotelDbV2Clean_(pair && pair[1]);
    if (!key && !state) return;

    summary.total++;
    const bucket = hotelDbV2DashboardClassifyState_(state);
    summary[bucket]++;
  });

  return summary;
}

function hotelDbV2DashboardClassifyState_(state) {
  const value = hotelDbV2Clean_(state);
  if (!value || value === '未確認') return 'unreviewed';
  if (value === '承認') return 'approved';
  if (value === '要再確認') return 'conflicts';
  if (value.indexOf('エラー') !== -1) return 'errors';
  if (
    value === '反映済み' ||
    value === '追加済み' ||
    value === '除外済み' ||
    value === '整理済み'
  ) {
    return 'completed';
  }
  return 'other';
}

function hotelDbV2DashboardLatestSummary_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.SUMMARY);
  const empty = {
    exists: false,
    timestamp: '',
    sourceSheet: '',
    sourceSheetId: '',
    startRow: '',
    endRow: '',
    processed: 0,
    operational: 0,
    corrections: 0,
    needReview: 0,
    notFound: 0,
    closed: 0,
    temporaryClosed: 0,
    futureOpening: 0,
    errors: 0,
    skipped: 0,
    nextStartRow: '',
    reconciliation: ''
  };

  if (!sheet || sheet.getLastRow() < 2) return empty;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = hotelDbV2DashboardHeaderMapFromValues_(headers);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  let row = null;
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i].some(function(value) { return hotelDbV2Clean_(value) !== ''; })) {
      row = values[i];
      break;
    }
  }
  if (!row) return empty;

  function value(header) {
    return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : '';
  }

  return {
    exists: true,
    timestamp: value('日時'),
    sourceSheet: value('元シート'),
    sourceSheetId: value('元シートID'),
    startRow: value('開始行'),
    endRow: value('終了行'),
    processed: hotelDbV2DashboardToNumber_(value('処理件数')),
    operational: hotelDbV2DashboardToNumber_(value('営業中')),
    corrections: hotelDbV2DashboardToNumber_(value('修正候補')),
    needReview: hotelDbV2DashboardToNumber_(value('要確認')),
    notFound: hotelDbV2DashboardToNumber_(value('未検出')),
    closed: hotelDbV2DashboardToNumber_(value('閉業')),
    temporaryClosed: hotelDbV2DashboardToNumber_(value('一時休業')),
    futureOpening: hotelDbV2DashboardToNumber_(value('開業予定')),
    errors: hotelDbV2DashboardToNumber_(value('エラー')),
    skipped: hotelDbV2DashboardToNumber_(value('スキップ')),
    nextStartRow: value('次回開始行'),
    reconciliation: value('整合確認')
  };
}

function hotelDbV2DashboardBuildActions_(workflows, latestSummary) {
  const actions = [];
  (workflows || []).forEach(function(workflow) {
    if (workflow.errors > 0) {
      actions.push({
        priority: 1,
        target: workflow.name,
        action: 'エラー詳細を確認し、原因を解消してから再実行',
        count: workflow.errors
      });
    }
    if (workflow.conflicts > 0) {
      actions.push({
        priority: 2,
        target: workflow.name,
        action: '「要再確認」の原因を確認し、必要なら候補を再作成',
        count: workflow.conflicts
      });
    }
    if (workflow.approved > 0) {
      actions.push({
        priority: 3,
        target: workflow.name,
        action: workflow.applyAction || '承認済み候補を安全反映',
        count: workflow.approved
      });
    }
    if (workflow.unreviewed > 0) {
      actions.push({
        priority: 4,
        target: workflow.name,
        action: workflow.reviewAction || '未確認候補を確認',
        count: workflow.unreviewed
      });
    }
  });

  actions.sort(function(left, right) {
    if (left.priority !== right.priority) return left.priority - right.priority;
    if (left.count !== right.count) return right.count - left.count;
    return String(left.target).localeCompare(String(right.target), 'ja');
  });

  if (!actions.length) {
    if (latestSummary && latestSummary.exists && hotelDbV2Clean_(latestSummary.nextStartRow)) {
      actions.push({
        priority: 5,
        target: latestSummary.sourceSheet || '元DB',
        action: '④ 本番バッチを続行（次回開始行 ' + latestSummary.nextStartRow + '）',
        count: 0
      });
    } else {
      actions.push({
        priority: 5,
        target: '全体',
        action: '⑤ Place ID再確認、または⑬ 新規宿泊施設探索を必要に応じて実行',
        count: 0
      });
    }
  }

  return actions.slice(0, HOTEL_DB_V2_DASHBOARD.MAX_ACTIONS);
}

function hotelDbV2DashboardOverallStatus_(totals) {
  const values = totals || {};
  if (Number(values.errors || 0) > 0) return '要対応（エラーあり）';
  if (Number(values.conflicts || 0) > 0) return '要対応（要再確認あり）';
  if (Number(values.approved || 0) > 0) return '承認済み処理待ち';
  if (Number(values.unreviewed || 0) > 0) return '確認待ち';
  return '良好';
}

function hotelDbV2DashboardAuditSummary_(spreadsheet) {
  const history = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
  const closedArchive = spreadsheet.getSheetByName('閉業除外履歴');
  const duplicateArchive = spreadsheet.getSheetByName('重複整理履歴');

  return {
    historyCount: hotelDbV2DashboardCountDataRows_(history),
    lastHistoryAt: hotelDbV2DashboardLastFirstColumnValue_(history),
    closedRemoved: hotelDbV2DashboardCountState_(closedArchive, '処理状態', '除外済み'),
    duplicatesConsolidated: hotelDbV2DashboardCountState_(duplicateArchive, '処理状態', '整理済み')
  };
}

function hotelDbV2DashboardCountDataRows_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  return values.reduce(function(total, row) {
    return total + (hotelDbV2Clean_(row[0]) ? 1 : 0);
  }, 0);
}

function hotelDbV2DashboardLastFirstColumnValue_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return '';
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const value = hotelDbV2Clean_(values[i][0]);
    if (value) return value;
  }
  return '';
}

function hotelDbV2DashboardCountState_(sheet, stateHeader, targetState) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = hotelDbV2DashboardHeaderMapFromValues_(headers);
  if (!map[stateHeader]) return 0;

  const values = sheet.getRange(2, map[stateHeader], sheet.getLastRow() - 1, 1).getDisplayValues();
  return values.reduce(function(total, row) {
    return total + (hotelDbV2Clean_(row[0]) === targetState ? 1 : 0);
  }, 0);
}

function hotelDbV2DashboardRecentHistory_(spreadsheet, limit) {
  const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.HISTORY);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = hotelDbV2DashboardHeaderMapFromValues_(headers);
  const max = Math.max(1, Number(limit || 10));
  const startRow = Math.max(2, sheet.getLastRow() - max + 1);
  const values = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, sheet.getLastColumn()).getDisplayValues();

  return values.filter(function(row) {
    return row.some(function(value) { return hotelDbV2Clean_(value) !== ''; });
  }).map(function(row) {
    function value(header) {
      return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : '';
    }
    return {
      timestamp: value('日時'),
      sourceSheet: value('元シート'),
      sourceRow: value('元行'),
      facilityName: value('施設名'),
      process: value('処理'),
      result: value('結果'),
      placeId: value('Place ID'),
      score: value('一致スコア'),
      detail: value('詳細')
    };
  }).reverse();
}

function hotelDbV2DashboardHeaderMapFromValues_(headers) {
  const map = {};
  (headers || []).forEach(function(header, index) {
    const value = hotelDbV2Clean_(header);
    if (value && !map[value]) map[value] = index + 1;
  });
  return map;
}

function hotelDbV2DashboardToNumber_(value) {
  const text = hotelDbV2Clean_(value).replace(/,/g, '');
  if (!text) return 0;
  const number = Number(text);
  return isFinite(number) ? number : 0;
}

function hotelDbV2DashboardGetOrCreate_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_DASHBOARD.SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(HOTEL_DB_V2_DASHBOARD.SHEET_NAME, 0);

  if (sheet.getMaxRows() < 45) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 45 - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < 9) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 9 - sheet.getMaxColumns());
  }
  return sheet;
}

function hotelDbV2DashboardRender_(sheet, data) {
  const rows = HOTEL_DB_V2_DASHBOARD.ROWS;
  const colors = HOTEL_DB_V2_DASHBOARD.COLORS;
  const fullRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  fullRange.breakApart();
  sheet.clear();
  sheet.setHiddenGridlines(true);
  sheet.setFrozenRows(3);

  sheet.getRange(rows.TITLE, 1, 1, 9).merge();
  sheet.getRange(rows.TITLE, 1)
    .setValue('宿泊施設DB Ver2.0　統合ダッシュボード')
    .setBackground(colors.NAVY)
    .setFontColor(colors.WHITE)
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('left');

  sheet.getRange(rows.SUBTITLE, 1, 1, 9).merge();
  sheet.getRange(rows.SUBTITLE, 1)
    .setValue('Google Places照合・候補・承認・監査の現在地を1画面で確認')
    .setBackground(colors.LIGHT_BLUE)
    .setFontColor(colors.DARK)
    .setFontSize(10);

  sheet.getRange(rows.STATUS, 1, 1, 9).setValues([[
    '更新日時', data.updatedAt,
    '全体状態', data.overallStatus,
    '要対応', data.attentionTotal,
    '完了', data.totals.completed,
    data.spreadsheetName
  ]]);
  sheet.getRange(rows.STATUS, 1, 1, 9).setFontWeight('bold').setVerticalAlignment('middle');
  sheet.getRange(rows.STATUS, 1).setBackground(colors.GRAY);
  sheet.getRange(rows.STATUS, 3).setBackground(colors.GRAY);
  sheet.getRange(rows.STATUS, 5).setBackground(colors.GRAY);
  sheet.getRange(rows.STATUS, 7).setBackground(colors.GRAY);
  sheet.getRange(rows.STATUS, 4).setBackground(hotelDbV2DashboardStatusColor_(data.overallStatus));

  hotelDbV2DashboardSection_(sheet, rows.LATEST_SECTION, '最新バッチ実行');
  sheet.getRange(rows.LATEST_HEADERS_1, 1, 1, 9).setValues([[
    '最新実行日時', '元シート', '開始行', '終了行', '処理件数', '営業中', '修正候補', '要確認', 'エラー'
  ]]);
  sheet.getRange(rows.LATEST_VALUES_1, 1, 1, 9).setValues([[
    data.latestSummary.timestamp || '—',
    data.latestSummary.sourceSheet || '—',
    data.latestSummary.startRow || '—',
    data.latestSummary.endRow || '—',
    data.latestSummary.processed,
    data.latestSummary.operational,
    data.latestSummary.corrections,
    data.latestSummary.needReview,
    data.latestSummary.errors
  ]]);
  sheet.getRange(rows.LATEST_HEADERS_2, 1, 1, 9).setValues([[
    '未検出', '閉業', '一時休業', '開業予定', 'スキップ', '次回開始行', '整合確認', '', ''
  ]]);
  sheet.getRange(rows.LATEST_VALUES_2, 1, 1, 9).setValues([[
    data.latestSummary.notFound,
    data.latestSummary.closed,
    data.latestSummary.temporaryClosed,
    data.latestSummary.futureOpening,
    data.latestSummary.skipped,
    data.latestSummary.nextStartRow || '完了/未設定',
    data.latestSummary.reconciliation || '—',
    '',
    ''
  ]]);
  hotelDbV2DashboardStyleTableHeader_(sheet.getRange(rows.LATEST_HEADERS_1, 1, 1, 9));
  hotelDbV2DashboardStyleTableHeader_(sheet.getRange(rows.LATEST_HEADERS_2, 1, 1, 9));

  hotelDbV2DashboardSection_(sheet, rows.WORKFLOW_SECTION, '候補ワークフロー状況');
  sheet.getRange(rows.WORKFLOW_HEADERS, 1, 1, 9).setValues([[
    'ワークフロー', '総件数', '未確認', '承認', '要再確認', 'エラー', '完了', 'その他', '次の操作'
  ]]);
  hotelDbV2DashboardStyleTableHeader_(sheet.getRange(rows.WORKFLOW_HEADERS, 1, 1, 9));

  const workflowRows = data.workflows.map(function(workflow) {
    return [
      workflow.name,
      workflow.total,
      workflow.unreviewed,
      workflow.approved,
      workflow.conflicts,
      workflow.errors,
      workflow.completed,
      workflow.other,
      hotelDbV2DashboardWorkflowNextAction_(workflow)
    ];
  });
  sheet.getRange(rows.WORKFLOW_START, 1, workflowRows.length, 9).setValues(workflowRows);
  data.workflows.forEach(function(workflow, index) {
    sheet.getRange(rows.WORKFLOW_START + index, 1, 1, 9)
      .setBackground(hotelDbV2DashboardWorkflowColor_(workflow));
  });

  hotelDbV2DashboardSection_(sheet, rows.ACTION_SECTION, '次にやること（優先順）');
  sheet.getRange(rows.ACTION_HEADERS, 1).setValue('優先度');
  sheet.getRange(rows.ACTION_HEADERS, 2).setValue('対象');
  sheet.getRange(rows.ACTION_HEADERS, 3, 1, 6).merge().setValue('推奨操作');
  sheet.getRange(rows.ACTION_HEADERS, 9).setValue('件数');
  hotelDbV2DashboardStyleTableHeader_(sheet.getRange(rows.ACTION_HEADERS, 1, 1, 9));

  for (let i = 0; i < HOTEL_DB_V2_DASHBOARD.MAX_ACTIONS; i++) {
    const rowNumber = rows.ACTION_START + i;
    sheet.getRange(rowNumber, 3, 1, 6).merge();
    const action = data.actions[i] || null;
    sheet.getRange(rowNumber, 1).setValue(action ? hotelDbV2DashboardPriorityLabel_(action.priority) : '');
    sheet.getRange(rowNumber, 2).setValue(action ? action.target : '');
    sheet.getRange(rowNumber, 3).setValue(action ? action.action : '');
    sheet.getRange(rowNumber, 9).setValue(action ? action.count : '');
    if (action) {
      sheet.getRange(rowNumber, 1, 1, 9).setBackground(hotelDbV2DashboardPriorityColor_(action.priority));
    }
  }

  hotelDbV2DashboardSection_(sheet, rows.AUDIT_SECTION, '監査・完了実績');
  sheet.getRange(rows.AUDIT_VALUES, 1, 1, 9).setValues([[
    '修正履歴件数', data.audit.historyCount,
    '閉業除外済み', data.audit.closedRemoved,
    '重複整理済み', data.audit.duplicatesConsolidated,
    '最終履歴日時', data.audit.lastHistoryAt || '—',
    ''
  ]]);
  [1, 3, 5, 7].forEach(function(column) {
    sheet.getRange(rows.AUDIT_VALUES, column).setBackground(colors.GRAY).setFontWeight('bold');
  });

  hotelDbV2DashboardSection_(sheet, rows.HISTORY_SECTION, '最近の処理履歴（最大10件）');
  sheet.getRange(rows.HISTORY_HEADERS, 1, 1, 9).setValues([[
    '日時', '元シート', '元行', '施設名', '処理', '結果', 'Place ID', '一致スコア', '詳細'
  ]]);
  hotelDbV2DashboardStyleTableHeader_(sheet.getRange(rows.HISTORY_HEADERS, 1, 1, 9));
  if (data.recentHistory.length) {
    const historyRows = data.recentHistory.map(function(item) {
      return [
        item.timestamp,
        item.sourceSheet,
        item.sourceRow,
        item.facilityName,
        item.process,
        item.result,
        item.placeId,
        item.score,
        item.detail
      ];
    });
    sheet.getRange(rows.HISTORY_START, 1, historyRows.length, 9).setValues(historyRows);
  } else {
    sheet.getRange(rows.HISTORY_START, 1).setValue('履歴なし');
  }

  sheet.getRange(rows.FOOTNOTE, 1, 1, 9).merge();
  sheet.getRange(rows.FOOTNOTE, 1)
    .setValue('このシートは自動生成です。更新時に内容・書式を再生成します。元DB・候補・履歴は変更しません。')
    .setFontColor('#666666')
    .setFontSize(9)
    .setBackground(colors.GRAY);

  hotelDbV2DashboardApplyLayout_(sheet);
}

function hotelDbV2DashboardWorkflowNextAction_(workflow) {
  if (workflow.errors > 0) return 'エラー原因を確認';
  if (workflow.conflicts > 0) return '要再確認を処理';
  if (workflow.approved > 0) return workflow.applyAction;
  if (workflow.unreviewed > 0) return workflow.reviewAction;
  if (workflow.other > 0) return 'その他状態を確認';
  return '対応なし';
}

function hotelDbV2DashboardSection_(sheet, rowNumber, title) {
  const colors = HOTEL_DB_V2_DASHBOARD.COLORS;
  sheet.getRange(rowNumber, 1, 1, 9).merge();
  sheet.getRange(rowNumber, 1)
    .setValue(title)
    .setBackground(colors.NAVY)
    .setFontColor(colors.WHITE)
    .setFontWeight('bold')
    .setFontSize(11);
}

function hotelDbV2DashboardStyleTableHeader_(range) {
  range
    .setBackground(HOTEL_DB_V2_DASHBOARD.COLORS.BLUE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
}

function hotelDbV2DashboardStatusColor_(status) {
  const colors = HOTEL_DB_V2_DASHBOARD.COLORS;
  const value = hotelDbV2Clean_(status);
  if (value.indexOf('エラー') !== -1) return colors.RED;
  if (value.indexOf('要再確認') !== -1) return colors.YELLOW;
  if (value.indexOf('処理待ち') !== -1 || value.indexOf('確認待ち') !== -1) return colors.BLUE;
  return colors.GREEN;
}

function hotelDbV2DashboardWorkflowColor_(workflow) {
  const colors = HOTEL_DB_V2_DASHBOARD.COLORS;
  if (workflow.errors > 0) return colors.RED;
  if (workflow.conflicts > 0) return colors.YELLOW;
  if (workflow.approved > 0 || workflow.unreviewed > 0) return colors.LIGHT_BLUE;
  return colors.GREEN;
}

function hotelDbV2DashboardPriorityLabel_(priority) {
  if (priority === 1) return '最優先';
  if (priority === 2) return '高';
  if (priority === 3) return '中';
  if (priority === 4) return '通常';
  return '案内';
}

function hotelDbV2DashboardPriorityColor_(priority) {
  const colors = HOTEL_DB_V2_DASHBOARD.COLORS;
  if (priority === 1) return colors.RED;
  if (priority === 2) return colors.YELLOW;
  if (priority === 3) return colors.BLUE;
  return colors.LIGHT_BLUE;
}

function hotelDbV2DashboardApplyLayout_(sheet) {
  const widths = [145, 125, 90, 165, 135, 115, 150, 105, 360];
  widths.forEach(function(width, index) {
    sheet.setColumnWidth(index + 1, width);
  });

  sheet.getRange(1, 1, 45, 9)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setFontFamily('Arial');

  sheet.setRowHeight(1, 34);
  sheet.setRowHeight(2, 24);
  sheet.setRowHeight(3, 28);
  [5, 11, 19, 27, 30].forEach(function(row) { sheet.setRowHeight(row, 26); });
  for (let row = 13; row <= 17; row++) sheet.setRowHeight(row, 34);
  for (let row = 21; row <= 25; row++) sheet.setRowHeight(row, 34);
  for (let row = 32; row <= 41; row++) sheet.setRowHeight(row, 34);

  sheet.getRange(6, 1, 4, 9).setHorizontalAlignment('center');
  sheet.getRange(12, 2, 6, 7).setHorizontalAlignment('center');
  sheet.getRange(20, 1, 6, 2).setHorizontalAlignment('center');
  sheet.getRange(20, 9, 6, 1).setHorizontalAlignment('center');
  sheet.getRange(31, 1, 11, 9).setVerticalAlignment('top');
}

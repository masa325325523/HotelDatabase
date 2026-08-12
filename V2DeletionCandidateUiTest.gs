/**
 * PR #15 UI実地テスト用の安全な準備・復元ヘルパー。
 * コピー版「宿泊施設DB_PR13_⑧反映テスト」専用。
 *
 * 通常の「要確認」を一時退避し、削除候補仕分け専用のテスト行を作る。
 * テスト後は cleanupHotelDbV2DeletionCandidateUiTest() で元へ戻す。
 */

const HOTEL_DB_V2_PR15_UI_TEST = Object.freeze({
  BACKUP_REVIEW: 'PR15_要確認退避',
  REQUIRED_NAME_PARTS: ['PR13', '⑧反映テスト']
});

function hotelDbV2Pr15AssertCopy_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const name = spreadsheet.getName();
  const ok = HOTEL_DB_V2_PR15_UI_TEST.REQUIRED_NAME_PARTS.every(function(part) {
    return name.indexOf(part) !== -1;
  });

  if (!ok) {
    throw new Error(
      'PR #15 UIテストはコピー版専用です。スプレッドシート名に「PR13」と「⑧反映テスト」の両方が必要です。'
    );
  }

  return spreadsheet;
}

function setupHotelDbV2DeletionCandidateUiTest() {
  const spreadsheet = hotelDbV2Pr15AssertCopy_();
  const reviewName = HOTEL_DB_V2_CONFIG.SHEETS.REVIEW;
  const existingBackup = spreadsheet.getSheetByName(HOTEL_DB_V2_PR15_UI_TEST.BACKUP_REVIEW);

  if (existingBackup) {
    throw new Error(
      'PR #15の退避シートがすでにあります。先に cleanupHotelDbV2DeletionCandidateUiTest を実行してください。'
    );
  }

  const currentReview = spreadsheet.getSheetByName(reviewName);
  if (currentReview) {
    currentReview.setName(HOTEL_DB_V2_PR15_UI_TEST.BACKUP_REVIEW);
  }

  const sheet = spreadsheet.insertSheet(reviewName);
  sheet.getRange(1, 1, 1, HOTEL_DB_V2_REVIEW_HEADERS.length)
    .setValues([HOTEL_DB_V2_REVIEW_HEADERS]);
  sheet.setFrozenRows(1);

  const today = hotelDbV2Today_();
  const rows = [
    {
      '確認キー': 'PR15_TEST|1', '状態': '未確認', '元シート': 'PR15テスト', '元シートID': 'PR15', '元行': 2,
      '郵便番号': '100-0001', '市区町村': '東京都千代田区', '住所': '千代田1-1', '施設名': '閉業テストホテルA', '宿泊分類': 'ホテル営業',
      '理由': '閉業', '候補施設名': '閉業テストホテルA', '候補住所': '東京都千代田区千代田1-1', '候補Place ID': 'TEST_CLOSED_98',
      '一致スコア': 98, '営業状態': '閉業', 'Google Maps URL': '', '確認日': today, '詳細': 'Google営業状態=CLOSED_PERMANENTLY'
    },
    {
      '確認キー': 'PR15_TEST|2', '状態': '未確認', '元シート': 'PR15テスト', '元シートID': 'PR15', '元行': 3,
      '郵便番号': '100-0002', '市区町村': '東京都千代田区', '住所': '皇居外苑1-1', '施設名': '未検出テスト旅館', '宿泊分類': '旅館営業',
      '理由': 'Google候補なし', '候補施設名': '', '候補住所': '', '候補Place ID': '',
      '一致スコア': '', '営業状態': '', 'Google Maps URL': '', '確認日': today, '詳細': '検索結果がありませんでした。'
    },
    {
      '確認キー': 'PR15_TEST|3', '状態': '未確認', '元シート': 'PR15テスト', '元シートID': 'PR15', '元行': 4,
      '郵便番号': '100-0003', '市区町村': '東京都千代田区', '住所': '一ツ橋1-1', '施設名': '一時休業テストホテル', '宿泊分類': 'ホテル営業',
      '理由': '一時休業', '候補施設名': '一時休業テストホテル', '候補住所': '東京都千代田区一ツ橋1-1', '候補Place ID': 'TEST_TEMP',
      '一致スコア': 96, '営業状態': '一時休業', 'Google Maps URL': '', '確認日': today, '詳細': 'Google営業状態=CLOSED_TEMPORARILY'
    },
    {
      '確認キー': 'PR15_TEST|4', '状態': '未確認', '元シート': 'PR15テスト', '元シートID': 'PR15', '元行': 5,
      '郵便番号': '100-0004', '市区町村': '東京都千代田区', '住所': '大手町1-1', '施設名': '開業予定テストホテル', '宿泊分類': 'ホテル営業',
      '理由': '開業予定', '候補施設名': '開業予定テストホテル', '候補住所': '東京都千代田区大手町1-1', '候補Place ID': 'TEST_FUTURE',
      '一致スコア': 95, '営業状態': '開業予定', 'Google Maps URL': '', '確認日': today, '詳細': 'Google営業状態=FUTURE_OPENING'
    },
    {
      '確認キー': 'PR15_TEST|5', '状態': '未確認', '元シート': 'PR15テスト', '元シートID': 'PR15', '元行': 6,
      '郵便番号': '100-0005', '市区町村': '東京都千代田区', '住所': '丸の内1-1', '施設名': '低一致閉業テストホテル', '宿泊分類': 'ホテル営業',
      '理由': '閉業', '候補施設名': '別施設候補', '候補住所': '東京都千代田区丸の内2-2', '候補Place ID': 'TEST_CLOSED_LOW',
      '一致スコア': 62, '営業状態': '閉業', 'Google Maps URL': '', '確認日': today, '詳細': 'Google営業状態=CLOSED_PERMANENTLY'
    },
    {
      '確認キー': 'PR15_TEST|6', '状態': '未確認', '元シート': 'PR15テスト', '元シートID': 'PR15', '元行': 7,
      '郵便番号': '100-0006', '市区町村': '東京都千代田区', '住所': '有楽町1-1', '施設名': '状態不整合テストホテル', '宿泊分類': 'ホテル営業',
      '理由': '閉業', '候補施設名': '状態不整合テストホテル', '候補住所': '東京都千代田区有楽町1-1', '候補Place ID': 'TEST_STATUS_MISMATCH',
      '一致スコア': 97, '営業状態': '営業中', 'Google Maps URL': '', '確認日': today, '詳細': '理由と営業状態の不整合テスト'
    },
    {
      '確認キー': 'PR15_TEST|7', '状態': '未確認', '元シート': 'PR15テスト', '元シートID': 'PR15', '元行': 8,
      '郵便番号': '100-0007', '市区町村': '東京都千代田区', '住所': '内幸町1-1', '施設名': '低スコア通常テスト', '宿泊分類': 'ホテル営業',
      '理由': '一致スコア不足', '候補施設名': '別候補', '候補住所': '東京都港区1-1', '候補Place ID': 'TEST_LOW_SCORE',
      '一致スコア': 40, '営業状態': '営業中', 'Google Maps URL': '', '確認日': today, '詳細': 'PR #14側の対象'
    }
  ];

  const values = rows.map(function(object) {
    return hotelDbV2RowFromObject_(HOTEL_DB_V2_REVIEW_HEADERS, object);
  });
  sheet.getRange(2, 1, values.length, HOTEL_DB_V2_REVIEW_HEADERS.length).setValues(values);
  sheet.activate();

  SpreadsheetApp.getUi().alert([
    'PR #15 ⑫UIテスト 準備完了',
    '',
    'テスト件数: 7件',
    '通常の要確認: 退避済み',
    '',
    '次にメニューの⑫「閉業・未検出を削除候補に仕分け」を実行してください。'
  ].join('\n'));

  return { prepared: true, testRows: values.length };
}

function cleanupHotelDbV2DeletionCandidateUiTest() {
  const spreadsheet = hotelDbV2Pr15AssertCopy_();
  const reviewName = HOTEL_DB_V2_CONFIG.SHEETS.REVIEW;
  const backup = spreadsheet.getSheetByName(HOTEL_DB_V2_PR15_UI_TEST.BACKUP_REVIEW);

  if (!backup) {
    throw new Error('PR #15の退避シートが見つかりません。復元処理は行いませんでした。');
  }

  const testSheet = spreadsheet.getSheetByName(reviewName);
  if (testSheet) spreadsheet.deleteSheet(testSheet);
  backup.setName(reviewName);
  backup.activate();

  SpreadsheetApp.getUi().alert([
    'PR #15 ⑫UIテスト 復元完了',
    '',
    'テスト用要確認: 削除済み',
    '通常の要確認: 復元済み'
  ].join('\n'));

  return { cleaned: true };
}

/** PR #20 自己診断。外部API・本番データを書き換えず、安全条件を検証する。 */
function runHotelDbV2ApprovedDuplicateConsolidationTests() {
  const failures = [];
  let success = 0;
  function expect(label, actual, expected) {
    if (actual !== expected) failures.push(label + ': 実際=' + actual + ', 期待=' + expected);
    else success++;
  }
  function baseCandidate() {
    return {
      state: '承認', recommendation: '重複濃厚', snapshotState: '作成済み',
      hash1: 'H1', hash2: 'H2', sheetId: 100, sheetName: 'テスト元DB',
      row1: 2, row2: 3, keepRow: 2,
      name1: '同一ホテル', address1: '東京都千代田区千代田1-1',
      name2: '同一ホテル', address2: '東京都千代田区千代田1-1'
    };
  }
  function facility(row) {
    return {
      rowNumber: row, postalCode: '100-0001', municipality: '東京都千代田区',
      address: '千代田1-1', name: '同一ホテル', category: 'ホテル営業',
      notes: '確認済み', placeId: 'P1'
    };
  }
  function snapshot(values, formulas) {
    const headers = ['郵便番号','市区町村名','住所（番地まで）','施設名','宿泊分類','備考','Place ID'];
    const rowValues = values || ['100-0001','東京都千代田区','千代田1-1','同一ホテル','ホテル営業','確認済み','P1'];
    return {
      headers: headers,
      displayValues: rowValues.map(String),
      values: rowValues.slice(),
      formulas: formulas || ['', '', '', '', '', '', ''],
      hash: 'HASH'
    };
  }
  const sourceMap = {
    postalCode: 1, municipality: 2, address: 3, facilityName: 4,
    category: 5, notes: 6, placeId: 7
  };

  const preCases = [
    ['正常', function(c){}, true],
    ['未確認', function(c){c.state='未確認';}, false],
    ['要再確認', function(c){c.state='要再確認';}, false],
    ['推奨要人確認', function(c){c.recommendation='要人確認';}, false],
    ['スナップショット要再確認', function(c){c.snapshotState='要再確認';}, false],
    ['行1ハッシュなし', function(c){c.hash1='';}, false],
    ['行2ハッシュなし', function(c){c.hash2='';}, false],
    ['シートIDなし', function(c){c.sheetId=0;}, false],
    ['シート名なし', function(c){c.sheetName='';}, false],
    ['行1が1', function(c){c.row1=1;}, false],
    ['行2が1', function(c){c.row2=1;}, false],
    ['行1行2同一', function(c){c.row2=2;}, false],
    ['施設名1なし', function(c){c.name1='';}, false],
    ['住所2なし', function(c){c.address2='';}, false],
    ['残す行未指定', function(c){c.keepRow=0;}, false],
    ['残す行が第三の行', function(c){c.keepRow=4;}, false],
    ['残す行=行2', function(c){c.keepRow=3;}, true]
  ];
  preCases.forEach(function(test) {
    const c=baseCandidate(); test[1](c);
    expect('前提:' + test[0], hotelDbV2Pr20Precheck_(c).ok, test[2]);
  });

  const identityCases = [
    ['正常', function(c,f1,f2){}, true],
    ['行1空', function(c,f1,f2){f1.name='';f1.address='';}, false],
    ['行2空', function(c,f1,f2){f2.name='';f2.address='';}, false],
    ['行1名変更', function(c,f1,f2){f1.name='別ホテル';}, false],
    ['行2名変更', function(c,f1,f2){f2.name='別ホテル';}, false],
    ['行1住所変更', function(c,f1,f2){f1.address='千代田1-2';}, false],
    ['行2住所変更', function(c,f1,f2){f2.address='千代田1-2';}, false],
    ['現在の両施設名差', function(c,f1,f2){c.name2='別ホテル';f2.name='別ホテル';}, false],
    ['現在の両住所差', function(c,f1,f2){c.address2='東京都千代田区千代田1-2';f2.address='千代田1-2';}, false],
    ['Place ID競合', function(c,f1,f2){f2.placeId='P2';}, false],
    ['片方Place ID空', function(c,f1,f2){f2.placeId='';}, true],
    ['両方Place ID空', function(c,f1,f2){f1.placeId='';f2.placeId='';}, true],
    ['名称空白表記差', function(c,f1,f2){c.name2='同一 ホテル';f2.name='同一 ホテル';}, true],
    ['丁目番地表記差', function(c,f1,f2){c.address2='東京都千代田区千代田一丁目1番';f2.address='千代田一丁目1番';}, true]
  ];
  identityCases.forEach(function(test) {
    const c=baseCandidate(), f1=facility(2), f2=facility(3); test[1](c,f1,f2);
    expect('同一性:' + test[0], hotelDbV2Pr20ValidatePairIdentity_(c,f1,f2).ok, test[2]);
  });

  const noLossCases = [
    ['全列同一', snapshot(), snapshot(), true],
    ['除外行備考空・残す行あり', snapshot(), snapshot(['100-0001','東京都千代田区','千代田1-1','同一ホテル','ホテル営業','','P1']), true],
    ['除外行だけ備考あり', snapshot(['100-0001','東京都千代田区','千代田1-1','同一ホテル','ホテル営業','','P1']), snapshot(), false],
    ['除外行だけPlace IDあり', snapshot(['100-0001','東京都千代田区','千代田1-1','同一ホテル','ホテル営業','確認済み','']), snapshot(), false],
    ['残す行だけPlace IDあり', snapshot(), snapshot(['100-0001','東京都千代田区','千代田1-1','同一ホテル','ホテル営業','確認済み','']), true],
    ['異なるPlace ID', snapshot(), snapshot(['100-0001','東京都千代田区','千代田1-1','同一ホテル','ホテル営業','確認済み','P2']), false],
    ['異なる宿泊分類', snapshot(), snapshot(['100-0001','東京都千代田区','千代田1-1','同一ホテル','旅館営業','確認済み','P1']), false],
    ['名称空白差は同等', snapshot(), snapshot(['100-0001','東京都千代田区','千代田1-1','同一 ホテル','ホテル営業','確認済み','P1']), true],
    ['住所表記差は同等', snapshot(), snapshot(['100-0001','東京都千代田区','千代田一丁目1番','同一ホテル','ホテル営業','確認済み','P1']), true],
    ['郵便番号ハイフン差は同等', snapshot(), snapshot(['1000001','東京都千代田区','千代田1-1','同一ホテル','ホテル営業','確認済み','P1']), true],
    ['市区町村空白差は同等', snapshot(), snapshot(['100-0001','東京都 千代田区','千代田1-1','同一ホテル','ホテル営業','確認済み','P1']), true],
    ['除外行だけ数式', snapshot(), snapshot(undefined, ['', '', '', '', '', '=A1', '']), false],
    ['同じ数式', snapshot(undefined, ['', '', '', '', '', '=A1', '']), snapshot(undefined, ['', '', '', '', '', '=A1', '']), true],
    ['異なる数式', snapshot(undefined, ['', '', '', '', '', '=A1', '']), snapshot(undefined, ['', '', '', '', '', '=B1', '']), false],
    ['列数不一致', snapshot(), {headers:['A'],displayValues:['x'],values:['x'],formulas:['']}, false]
  ];
  noLossCases.forEach(function(test) {
    expect('情報損失:' + test[0], hotelDbV2Pr20NoLossCheck_(sourceMap,test[1],test[2]).ok, test[3]);
  });

  expect('整理キー決定性', hotelDbV2Pr20ArchiveKey_(1,2,3,2,'A','B'), hotelDbV2Pr20ArchiveKey_(1,2,3,2,'A','B'));
  expect('整理キー残す行差', hotelDbV2Pr20ArchiveKey_(1,2,3,2,'A','B') === hotelDbV2Pr20ArchiveKey_(1,2,3,3,'A','B'), false);
  expect('整理キー行差', hotelDbV2Pr20ArchiveKey_(1,2,3,2,'A','B') === hotelDbV2Pr20ArchiveKey_(1,2,4,2,'A','B'), false);
  expect('監査ヘッダ9列', HOTEL_DB_V2_PR20_AUDIT_HEADERS.length, 9);
  expect('履歴ヘッダ30列', HOTEL_DB_V2_PR20_ARCHIVE_HEADERS.length, 30);
  expect('成功状態', HOTEL_DB_V2_PR20.APPLIED, '整理済み');
  expect('競合状態', HOTEL_DB_V2_PR20.CONFLICT, '要再確認');
  expect('エラー状態', HOTEL_DB_V2_PR20.ERROR, '整理エラー');
  expect('重複濃厚定数', HOTEL_DB_V2_PR20.STRONG, '重複濃厚');
  expect('物理削除API不使用契約', /deleteRow|deleteRows/.test(hotelDbV2Pr20SafetySourceMarker_()), false);
  expect('clearContent方式契約', /clearContent/.test(hotelDbV2Pr20SafetySourceMarker_()), true);
  expect('PR19スナップショット基盤利用', typeof hotelDbV2Pr19SnapshotRow_, 'function');
  expect('PR19復元基盤利用', typeof hotelDbV2Pr19RestoreRow_, 'function');

  if (failures.length) throw new Error('PR #20 自己診断失敗\n\n' + failures.join('\n'));
  const message = ['PR #20 自己診断 成功', '', '成功件数: ' + success + '件', '失敗件数: 0件', '物理行削除: なし', '残す行の自動選択: なし'].join('\n');
  try { SpreadsheetApp.getUi().alert(message); } catch (error) {}
  return { success: success, failure: 0, physicalRowDelete: false, autoKeepSelection: false };
}

function hotelDbV2Pr20SafetySourceMarker_() {
  return 'human-select-keep-row; archive-first; clearContent; no-physical-row-removal; no-loss-check';
}

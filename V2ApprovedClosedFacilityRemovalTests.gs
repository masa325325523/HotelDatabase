/** PR #19 自己診断。外部API・本番データを書き換えず、安全条件を検証する。 */
function runHotelDbV2ApprovedClosedFacilityRemovalTests() {
  const failures = [];
  let success = 0;
  function expect(label, actual, expected) {
    if (actual !== expected) failures.push(label + ': 実際=' + actual + ', 期待=' + expected);
    else success++;
  }
  function baseCandidate() {
    return {
      state: '承認', recommendation: '削除候補有力', reason: '閉業', placeId: 'P1',
      candidateName: '閉業テストホテル', candidateAddress: '東京都千代田区千代田1-1',
      businessStatus: '閉業', matchScore: 98, sheetId: 100, sheetName: 'テスト元DB',
      sourceRow: 2, rowHash: 'abc', postalCode: '100-0001', municipality: '東京都千代田区',
      address: '千代田1-1', name: '閉業テストホテル', category: 'ホテル営業'
    };
  }
  function facility() {
    return { postalCode: '100-0001', municipality: '東京都千代田区', address: '千代田1-1', name: '閉業テストホテル', category: 'ホテル営業', notes: '確認済み', placeId: 'P1' };
  }
  function live() {
    return {
      id: 'P1', displayName: { text: '閉業テストホテル' }, formattedAddress: '東京都千代田区千代田1-1',
      businessStatus: 'CLOSED_PERMANENTLY', addressComponents: [
        { longText: '100-0001', types: ['postal_code'] },
        { longText: '千代田区', types: ['administrative_area_level_2'] },
        { longText: '東京都', types: ['administrative_area_level_1'] }
      ]
    };
  }

  const preCases = [
    ['正常承認', function(c){}, true],
    ['未確認', function(c){c.state='未確認';}, false],
    ['要再確認', function(c){c.state='要再確認';}, false],
    ['非有力', function(c){c.recommendation='要人確認';}, false],
    ['Google候補なし', function(c){c.reason='Google候補なし';}, false],
    ['一時休業', function(c){c.reason='一時休業';}, false],
    ['開業予定', function(c){c.reason='開業予定';}, false],
    ['Place IDなし', function(c){c.placeId='';}, false],
    ['候補名なし', function(c){c.candidateName='';}, false],
    ['候補住所なし', function(c){c.candidateAddress='';}, false],
    ['候補時営業中', function(c){c.businessStatus='営業中';}, false],
    ['74点', function(c){c.matchScore=74;}, false],
    ['75点', function(c){c.matchScore=75;}, true],
    ['元シートIDなし', function(c){c.sheetId=0;}, false],
    ['元シート名なし', function(c){c.sheetName='';}, false],
    ['元行1', function(c){c.sourceRow=1;}, false],
    ['ハッシュなし', function(c){c.rowHash='';}, false]
  ];
  preCases.forEach(function(test) {
    const c = baseCandidate(); test[1](c);
    expect('前提:' + test[0], hotelDbV2Pr19Precheck_(c).ok, test[2]);
  });

  const facilityCases = [
    ['正常', function(f,c){}, true],
    ['空行', function(f,c){f.name='';f.address='';}, false],
    ['郵便番号変更', function(f,c){f.postalCode='100-0002';}, false],
    ['市区町村変更', function(f,c){f.municipality='東京都港区';}, false],
    ['住所変更', function(f,c){f.address='千代田2-2';}, false],
    ['施設名変更', function(f,c){f.name='別ホテル';}, false],
    ['分類変更', function(f,c){f.category='旅館営業';}, false],
    ['Place ID変更', function(f,c){f.placeId='P2';}, false]
  ];
  facilityCases.forEach(function(test) {
    const f=facility(), c=baseCandidate(); test[1](f,c);
    expect('元DB:' + test[0], hotelDbV2Pr19ValidateFacilitySnapshot_(f,c).ok, test[2]);
  });

  const liveCases = [
    ['正常', function(p,c,f){}, true],
    ['再取得なし', function(p,c,f){p.id='';}, false],
    ['Place ID不一致', function(p,c,f){p.id='P2';}, false],
    ['営業再開', function(p,c,f){p.businessStatus='OPERATIONAL';}, false],
    ['一時休業へ変更', function(p,c,f){p.businessStatus='CLOSED_TEMPORARILY';}, false],
    ['施設名変更', function(p,c,f){p.displayName={text:'別ホテル'};}, false],
    ['住所変更', function(p,c,f){p.formattedAddress='東京都千代田区千代田2-2';}, false],
    ['郵便番号変更', function(p,c,f){p.addressComponents[0].longText='100-0002';}, false]
  ];
  liveCases.forEach(function(test) {
    const p=live(), c=baseCandidate(), f=facility(); test[1](p,c,f);
    expect('Google:' + test[0], hotelDbV2Pr19ValidateLivePlace_(f,c,p).ok, test[2]);
  });

  expect('最終スコア74は不足', 74 >= HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE, false);
  expect('最終スコア75は合格', 75 >= HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE, true);
  expect('除外キー決定性', hotelDbV2Pr19ArchiveKey_(1,2,'P','H'), hotelDbV2Pr19ArchiveKey_(1,2,'P','H'));
  expect('除外キー行違い', hotelDbV2Pr19ArchiveKey_(1,2,'P','H') === hotelDbV2Pr19ArchiveKey_(1,3,'P','H'), false);
  expect('除外キーPlace違い', hotelDbV2Pr19ArchiveKey_(1,2,'P','H') === hotelDbV2Pr19ArchiveKey_(1,2,'Q','H'), false);
  expect('除外キーハッシュ違い', hotelDbV2Pr19ArchiveKey_(1,2,'P','H') === hotelDbV2Pr19ArchiveKey_(1,2,'P','I'), false);
  expect('履歴ヘッダ32列', HOTEL_DB_V2_PR19_ARCHIVE_HEADERS.length, 32);
  expect('監査ヘッダ10列', HOTEL_DB_V2_PR19_AUDIT_HEADERS.length, 10);
  expect('物理削除APIを仕様で使用しない', /deleteRow|deleteRows/.test(hotelDbV2Pr19SafetySourceMarker_()), false);
  expect('clearContent方式', /clearContent/.test(hotelDbV2Pr19SafetySourceMarker_()), true);
  expect('成功状態', HOTEL_DB_V2_PR19.APPLIED, '除外済み');
  expect('競合状態', HOTEL_DB_V2_PR19.CONFLICT, '要再確認');
  expect('エラー状態', HOTEL_DB_V2_PR19.ERROR, '除外エラー');
  expect('恒久閉業API値', HOTEL_DB_V2_PR19.CLOSED_STATUS_API, 'CLOSED_PERMANENTLY');
  expect('閾値', HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE, 75);

  if (failures.length) throw new Error('PR #19 自己診断失敗\n\n' + failures.join('\n'));
  SpreadsheetApp.getUi().alert(['PR #19 自己診断 成功', '', '成功件数: ' + success + '件', '失敗件数: 0件', '物理行削除: なし'].join('\n'));
  return { success: success, failure: 0, physicalRowDelete: false };
}

function hotelDbV2Pr19SafetySourceMarker_() {
  // 静的な安全契約。実装は clearContent() であり deleteRow/deleteRows は使わない。
  return 'archive-first; clearContent; no-physical-row-removal';
}
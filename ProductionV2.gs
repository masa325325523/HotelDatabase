/**
 * ProductionV2.gs
 * Ver2.0 公開インターフェース。
 */

const HOTEL_DB_V2 = Object.freeze({
  config: HOTEL_DB_V2_CONFIG,
  diagnose: hotelDbV2Diagnose_,
  connectionTest: hotelDbV2ConnectionTest_,
  test3: function() {
    return hotelDbV2ProcessActiveSheet_({
      startRow: 2,
      maxRows: HOTEL_DB_V2_CONFIG.TEST_ROWS,
      useCheckpoint: false,
      onlyExisting: false
    });
  },
  runBatch: function() {
    return hotelDbV2ProcessActiveSheet_({
      maxRows: HOTEL_DB_V2_CONFIG.BATCH_SIZE,
      useCheckpoint: true,
      checkpointMode: 'enrich',
      onlyExisting: false
    });
  },
  refreshExisting: function() {
    return hotelDbV2ProcessActiveSheet_({
      maxRows: HOTEL_DB_V2_CONFIG.BATCH_SIZE,
      useCheckpoint: true,
      checkpointMode: 'refresh',
      onlyExisting: true
    });
  },
  resetCheckpoint: hotelDbV2ResetCheckpoint_,
  refreshDuplicates: hotelDbV2RefreshDuplicates_,
  applyApprovedCorrections: hotelDbV2ApplyApprovedCorrections_
});

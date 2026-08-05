/**
 * ProductionV2.gs
 * Ver2.0 公開インターフェース。
 * 設定値は各処理の実行時に参照し、Apps Scriptのファイル読込順に依存しない。
 */

const HOTEL_DB_V2 = Object.freeze({
  diagnose: function() {
    return hotelDbV2Diagnose_();
  },
  connectionTest: function() {
    return hotelDbV2ConnectionTest_();
  },
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
  resetCheckpoint: function() {
    return hotelDbV2ResetCheckpoint_();
  },
  refreshDuplicates: function() {
    return hotelDbV2RefreshDuplicates_();
  },
  applyApprovedCorrections: function() {
    return hotelDbV2ApplyApprovedCorrections_();
  }
});

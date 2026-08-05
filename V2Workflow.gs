/**
 * 全国宿泊施設データベース Ver2.0
 * 自動生成された分割モジュール。関数名は衝突回避のため hotelDbV2 接頭辞を使用。
 */

function hotelDbV2ProcessOneRow_(sheet, rowNumber, map, output, summary, options) {
  const facility = hotelDbV2ReadFacility_(sheet, rowNumber, map);
  if (!facility.name && !facility.address) {
    summary.skipped++;
    return;
  }

  let best = null;
  let place = null;
  let score = 0;

  try {
    if (facility.placeId) {
      place = hotelDbV2GetPlaceDetails_(facility.placeId);
      score = hotelDbV2CalculateMatchScore_(facility, place);
      best = { place: place, score: score, accepted: true, reviewable: true };
    } else if (options.onlyExisting) {
      summary.skipped++;
      return;
    } else {
      best = hotelDbV2FindBestCandidate_(facility);
    }

    if (!best) {
      const review = hotelDbV2ReviewCandidate_(
        sheet, facility, 'Google候補なし', null, '', '検索結果がありませんでした。'
      );
      hotelDbV2Upsert_(output.review, review.key, review.object, true);
      summary.notFound++;
      summary.needReview++;
      if (map.checkedAt) sheet.getRange(rowNumber, map.checkedAt).setValue(hotelDbV2Today_());
      if (map.matchDecision) sheet.getRange(rowNumber, map.matchDecision).setValue('候補なし');
      hotelDbV2AppendHistory_(output, hotelDbV2HistoryObject_(
        sheet, facility, '照合', '候補なし', '', '', '不明', ''
      ));
      return;
    }

    if (!best.reviewable || !best.accepted) {
      const reason = best.reviewable ? '一致スコア要確認' : '一致スコア不足';
      const review = hotelDbV2ReviewCandidate_(
        sheet, facility, reason, best.place, best.score,
        '自動採用基準=' + HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE
      );
      hotelDbV2Upsert_(output.review, review.key, review.object, true);
      summary.needReview++;
      if (map.checkedAt) sheet.getRange(rowNumber, map.checkedAt).setValue(hotelDbV2Today_());
      if (map.matchScore) sheet.getRange(rowNumber, map.matchScore).setValue(best.score);
      if (map.matchDecision) sheet.getRange(rowNumber, map.matchDecision).setValue(reason);
      hotelDbV2AppendHistory_(output, hotelDbV2HistoryObject_(
        sheet, facility, '照合', reason, hotelDbV2Clean_(best.place.id), best.score,
        hotelDbV2TranslateBusinessStatus_(best.place.businessStatus), ''
      ));
      return;
    }

    place = facility.placeId ? place : hotelDbV2GetPlaceDetails_(best.place.id);
    score = hotelDbV2CalculateMatchScore_(facility, place);

    if (score < HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE) {
      const detailedData = hotelDbV2BuildGoogleData_(
        place, score, facility, '詳細取得後の一致スコア要確認'
      );
      hotelDbV2WriteGoogleData_(
        sheet, rowNumber, map, detailedData, Boolean(facility.placeId)
      );
      const review = hotelDbV2ReviewCandidate_(
        sheet, facility, '詳細取得後の一致スコア要確認', place, score,
        '自動採用基準=' + HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE
      );
      hotelDbV2Upsert_(output.review, review.key, review.object, true);
      summary.needReview++;
      hotelDbV2AppendHistory_(output, hotelDbV2HistoryObject_(
        sheet, facility, '照合', '詳細取得後の一致スコア要確認',
        hotelDbV2Clean_(place.id), score, hotelDbV2TranslateBusinessStatus_(place.businessStatus), ''
      ));
      return;
    }

    const status = hotelDbV2Clean_(place.businessStatus);
    const statusJa = hotelDbV2TranslateBusinessStatus_(status);
    let decision = '照合済み';

    if (status !== 'OPERATIONAL') {
      decision = statusJa + '・要確認';
    }

    const googleData = hotelDbV2BuildGoogleData_(place, score, facility, decision);
    hotelDbV2WriteGoogleData_(sheet, rowNumber, map, googleData, true);

    if (status === 'OPERATIONAL') {
      summary.operational++;
      const differences = hotelDbV2CompareFacility_(facility, googleData);
      if (differences.length) {
        const correction = hotelDbV2CorrectionCandidate_(
          sheet, facility, googleData, differences
        );
        hotelDbV2Upsert_(output.corrections, correction.key, correction.object, true);
        summary.corrections++;
        if (map.matchDecision) {
          sheet.getRange(rowNumber, map.matchDecision).setValue('修正候補あり');
        }
        hotelDbV2AppendHistory_(output, hotelDbV2HistoryObject_(
          sheet, facility, '照合', '修正候補あり', googleData.placeId,
          score, statusJa, differences.join('・')
        ));
      } else {
        hotelDbV2AppendHistory_(output, hotelDbV2HistoryObject_(
          sheet, facility, '照合', '一致', googleData.placeId,
          score, statusJa, ''
        ));
      }
    } else {
      const review = hotelDbV2ReviewCandidate_(
        sheet, facility, statusJa, place, score,
        'Google営業状態=' + hotelDbV2Clean_(status)
      );
      hotelDbV2Upsert_(output.review, review.key, review.object, true);
      summary.needReview++;

      if (status === 'CLOSED_PERMANENTLY') summary.closed++;
      else if (status === 'CLOSED_TEMPORARILY') summary.temporaryClosed++;
      else if (status === 'FUTURE_OPENING') summary.futureOpening++;

      hotelDbV2AppendHistory_(output, hotelDbV2HistoryObject_(
        sheet, facility, '営業状態確認', statusJa, googleData.placeId,
        score, statusJa, ''
      ));
    }
  } catch (error) {
    summary.errors++;
    const review = hotelDbV2ReviewCandidate_(
      sheet, facility, 'APIエラー', best && best.place ? best.place : null,
      best ? best.score : '', error.message
    );
    hotelDbV2Upsert_(output.review, review.key, review.object, true);
    if (map.checkedAt) sheet.getRange(rowNumber, map.checkedAt).setValue(hotelDbV2Today_());
    if (map.matchDecision) sheet.getRange(rowNumber, map.matchDecision).setValue('APIエラー');
    hotelDbV2AppendHistory_(output, hotelDbV2HistoryObject_(
      sheet, facility, '照合', 'エラー', facility.placeId, '', '不明', error.message
    ));
  }
}

function hotelDbV2NewSummary_(sheet, startRow) {
  return {
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    startRow: startRow,
    endRow: startRow - 1,
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
    nextStartRow: null
  };
}

function hotelDbV2AppendSummary_(output, summary) {
  const reconciliation = summary.processed ===
    summary.operational + summary.needReview + summary.errors + summary.skipped
    ? '一致'
    : '要確認';

  output.summary.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_SUMMARY_HEADERS, {
    '日時': hotelDbV2Timestamp_(),
    '元シート': summary.sheetName,
    '元シートID': summary.sheetId,
    '開始行': summary.startRow,
    '終了行': summary.endRow,
    '処理件数': summary.processed,
    '営業中': summary.operational,
    '修正候補': summary.corrections,
    '要確認': summary.needReview,
    '未検出': summary.notFound,
    '閉業': summary.closed,
    '一時休業': summary.temporaryClosed,
    '開業予定': summary.futureOpening,
    'エラー': summary.errors,
    'スキップ': summary.skipped,
    '次回開始行': summary.nextStartRow || '',
    '整合確認': reconciliation
  }));

  summary.reconciliation = reconciliation;
}

function hotelDbV2CheckpointKey_(spreadsheet, sheet, mode) {
  return [
    'HOTEL_DB_V2_CHECKPOINT',
    spreadsheet.getId(),
    sheet.getSheetId(),
    hotelDbV2Clean_(mode || 'enrich')
  ].join('::');
}

function hotelDbV2GetCheckpoint_(spreadsheet, sheet, mode) {
  const value = PropertiesService.getDocumentProperties()
    .getProperty(hotelDbV2CheckpointKey_(spreadsheet, sheet, mode));
  const row = Number(value || 2);
  return row >= 2 ? row : 2;
}

function hotelDbV2SetCheckpoint_(spreadsheet, sheet, row, mode) {
  PropertiesService.getDocumentProperties()
    .setProperty(hotelDbV2CheckpointKey_(spreadsheet, sheet, mode), String(row));
}

function hotelDbV2ClearCheckpoint_(spreadsheet, sheet, mode) {
  PropertiesService.getDocumentProperties()
    .deleteProperty(hotelDbV2CheckpointKey_(spreadsheet, sheet, mode));
}

function hotelDbV2ProcessActiveSheet_(options) {
  const opts = options || {};
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  let map = hotelDbV2GetHeaderMap_(sheet);
  hotelDbV2ValidateSourceSheet_(sheet, map);
  map = hotelDbV2EnsureOutputHeaders_(sheet);

  const output = hotelDbV2CreateOutputContext_(spreadsheet);
  const lastRow = sheet.getLastRow();
  const defaultStart = opts.startRow || 2;
  const checkpointMode = hotelDbV2Clean_(opts.checkpointMode || 'enrich');
  const startRow = opts.useCheckpoint
    ? hotelDbV2GetCheckpoint_(spreadsheet, sheet, checkpointMode)
    : Math.max(2, Number(defaultStart));
  const maxRows = Math.max(1, Number(opts.maxRows || HOTEL_DB_V2_CONFIG.BATCH_SIZE));
  const plannedEnd = Math.min(lastRow, startRow + maxRows - 1);
  const summary = hotelDbV2NewSummary_(sheet, startRow);
  const startedAt = Date.now();

  if (lastRow < 2 || startRow > lastRow) {
    if (opts.useCheckpoint) hotelDbV2ClearCheckpoint_(spreadsheet, sheet, checkpointMode);
    hotelDbV2AppendSummary_(output, summary);
    return summary;
  }

  for (let row = startRow; row <= plannedEnd; row++) {
    if (Date.now() - startedAt >= HOTEL_DB_V2_CONFIG.MAX_EXECUTION_MS) {
      summary.nextStartRow = row;
      if (opts.useCheckpoint) {
        hotelDbV2SetCheckpoint_(spreadsheet, sheet, row, checkpointMode);
      }
      break;
    }

    summary.processed++;
    summary.endRow = row;
    hotelDbV2ProcessOneRow_(sheet, row, map, output, summary, opts);
    if (opts.useCheckpoint) {
      hotelDbV2SetCheckpoint_(spreadsheet, sheet, row + 1, checkpointMode);
    }
    Utilities.sleep(HOTEL_DB_V2_CONFIG.REQUEST_INTERVAL_MS);
  }

  if (!summary.nextStartRow) {
    if (summary.endRow < lastRow) {
      summary.nextStartRow = summary.endRow + 1;
      if (opts.useCheckpoint) {
        hotelDbV2SetCheckpoint_(spreadsheet, sheet, summary.nextStartRow, checkpointMode);
      }
    } else if (opts.useCheckpoint) {
      hotelDbV2ClearCheckpoint_(spreadsheet, sheet, checkpointMode);
    }
  }

  hotelDbV2AppendSummary_(output, summary);
  return summary;
}

function hotelDbV2Diagnose_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const result = {
    apiVersion: 'Places API (New)',
    apiKeyConfigured: false,
    activeSheet: sheet ? sheet.getName() : '',
    headerMap: {},
    checkpoint: '',
    errors: []
  };

  try {
    result.apiKeyConfigured = Boolean(hotelDbV2GetApiKey_());
  } catch (error) {
    result.errors.push(error.message);
  }

  try {
    result.headerMap = hotelDbV2GetHeaderMap_(sheet);
    hotelDbV2ValidateSourceSheet_(sheet, result.headerMap);
    result.checkpoint = hotelDbV2GetCheckpoint_(spreadsheet, sheet, 'enrich');
  } catch (error) {
    result.errors.push(error.message);
  }

  return result;
}

function hotelDbV2ConnectionTest_() {
  const results = hotelDbV2SearchPlaces_(
    'ホテルルートイン名古屋今池駅前 愛知県名古屋市千種区'
  );
  if (!results.length) throw new Error('検索結果がありませんでした。');
  const details = hotelDbV2GetPlaceDetails_(results[0].id);
  return {
    placeId: hotelDbV2Clean_(details.id),
    name: hotelDbV2GetDisplayName_(details),
    address: hotelDbV2GetJapaneseFullAddress_(details),
    status: hotelDbV2TranslateBusinessStatus_(details.businessStatus)
  };
}


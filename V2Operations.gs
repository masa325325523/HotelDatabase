/**
 * 全国宿泊施設データベース Ver2.0
 * 自動生成された分割モジュール。関数名は衝突回避のため hotelDbV2 接頭辞を使用。
 */

function hotelDbV2RefreshDuplicates_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const map = hotelDbV2GetHeaderMap_(sheet);
  hotelDbV2ValidateSourceSheet_(sheet, map);

  const duplicateSheet = hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES, HOTEL_DB_V2_DUPLICATE_HEADERS
  );
  hotelDbV2RemoveRowsForSheetId_(duplicateSheet, sheet.getSheetId(), 3);

  const facilities = [];
  for (let row = 2; row <= sheet.getLastRow(); row++) {
    const facility = hotelDbV2ReadFacility_(sheet, row, map);
    if (!facility.name && !facility.address) continue;
    facility.normalizedName = hotelDbV2NormalizeText_(facility.name);
    facility.normalizedAddress = hotelDbV2NormalizeAddress_(
      facility.municipality + facility.address
    );
    facilities.push(facility);
  }

  const byPlaceId = {};
  const byNameAddress = {};
  const byAddress = {};

  facilities.forEach(function(facility) {
    if (facility.placeId) {
      if (!byPlaceId[facility.placeId]) byPlaceId[facility.placeId] = [];
      byPlaceId[facility.placeId].push(facility);
    }

    const exactKey = facility.normalizedName + '|' + facility.normalizedAddress;
    if (facility.normalizedName && facility.normalizedAddress) {
      if (!byNameAddress[exactKey]) byNameAddress[exactKey] = [];
      byNameAddress[exactKey].push(facility);
    }

    if (facility.normalizedAddress) {
      if (!byAddress[facility.normalizedAddress]) {
        byAddress[facility.normalizedAddress] = [];
      }
      byAddress[facility.normalizedAddress].push(facility);
    }
  });

  const pairMap = {};

  function addPair(left, right, decision, placeId, similarity) {
    const first = left.rowNumber < right.rowNumber ? left : right;
    const second = left.rowNumber < right.rowNumber ? right : left;
    const pairKey = first.rowNumber + '|' + second.rowNumber;
    const priority = {
      'Place ID一致': 3,
      '施設名＋住所一致': 2,
      '類似・要確認': 1
    };
    const existing = pairMap[pairKey];
    if (existing && priority[existing.decision] >= priority[decision]) return;

    pairMap[pairKey] = {
      left: first,
      right: second,
      decision: decision,
      placeId: placeId || '',
      similarity: similarity
    };
  }

  function addGroupPairs(group, decision, placeId) {
    if (!group || group.length < 2) return;
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const similarity = hotelDbV2SimilarityRatio_(
          group[i].normalizedName, group[j].normalizedName
        );
        addPair(group[i], group[j], decision, placeId, similarity);
      }
    }
  }

  Object.keys(byPlaceId).forEach(function(placeId) {
    addGroupPairs(byPlaceId[placeId], 'Place ID一致', placeId);
  });

  Object.keys(byNameAddress).forEach(function(key) {
    addGroupPairs(byNameAddress[key], '施設名＋住所一致', '');
  });

  Object.keys(byAddress).forEach(function(addressKey) {
    const group = byAddress[addressKey];
    if (group.length < 2) return;
    for (let i = 0; i < group.length; i++) {
      for (let j = i + 1; j < group.length; j++) {
        const similarity = hotelDbV2SimilarityRatio_(
          group[i].normalizedName, group[j].normalizedName
        );
        if (similarity >= 0.85) {
          addPair(group[i], group[j], '類似・要確認', '', similarity);
        }
      }
    }
  });

  const candidates = Object.keys(pairMap).map(function(pairKey) {
    const pair = pairMap[pairKey];
    const left = pair.left;
    const right = pair.right;
    const key = [
      sheet.getSheetId(), left.rowNumber, right.rowNumber, pair.decision
    ].join('|');

    return hotelDbV2RowFromObject_(HOTEL_DB_V2_DUPLICATE_HEADERS, {
      '重複キー': key,
      '元シート': sheet.getName(),
      '元シートID': sheet.getSheetId(),
      '判定': pair.decision,
      '行1': left.rowNumber,
      '施設名1': left.name,
      '住所1': left.municipality + left.address,
      '行2': right.rowNumber,
      '施設名2': right.name,
      '住所2': right.municipality + right.address,
      'Place ID': pair.placeId,
      '類似度': Math.round(pair.similarity * 100),
      '確認日': hotelDbV2Today_(),
      '状態': '未確認'
    });
  });

  if (candidates.length) {
    duplicateSheet.getRange(
      duplicateSheet.getLastRow() + 1, 1, candidates.length, HOTEL_DB_V2_DUPLICATE_HEADERS.length
    ).setValues(candidates);
  }

  return { candidates: candidates.length };
}

function hotelDbV2RemoveRowsForSheetId_(sheet, sheetId, columnNumber) {
  if (sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, columnNumber, sheet.getLastRow() - 1, 1)
    .getDisplayValues();
  const rowsToDelete = [];

  values.forEach(function(row, offset) {
    if (String(row[0]) === String(sheetId)) rowsToDelete.push(offset + 2);
  });

  rowsToDelete.reverse().forEach(function(rowNumber) {
    sheet.deleteRow(rowNumber);
  });
}

function hotelDbV2ApplyApprovedCorrections_() {
  return hotelDbV2ApplyApprovedCorrectionsWithContext_({});
}

function hotelDbV2ApplyApprovedCorrectionsWithContext_(options) {
  const opts = options || {};
  const spreadsheet = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const correctionSheet = opts.correctionSheet || spreadsheet.getSheetByName(
    HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS
  );
  if (!correctionSheet || correctionSheet.getLastRow() < 2) {
    return { approved: 0, applied: 0, conflicts: 0, errors: 0 };
  }

  const map = hotelDbV2HeaderIndex_(HOTEL_DB_V2_CORRECTION_HEADERS);
  const values = correctionSheet.getRange(
    2, 1, correctionSheet.getLastRow() - 1, HOTEL_DB_V2_CORRECTION_HEADERS.length
  ).getDisplayValues();
  const historySheet = opts.historySheet || hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY, HOTEL_DB_V2_HISTORY_HEADERS
  );

  const result = { approved: 0, applied: 0, conflicts: 0, errors: 0 };

  values.forEach(function(row, offset) {
    const state = hotelDbV2Clean_(row[map['状態'] - 1]);
    if (state !== '承認') return;
    result.approved++;
    const correctionRow = offset + 2;

    try {
      const sourceSheetId = Number(row[map['元シートID'] - 1]);
      const sourceRow = Number(row[map['元行'] - 1]);
      const sourceSheet = spreadsheet.getSheetById(sourceSheetId);
      if (!sourceSheet) throw new Error('元シートが見つかりません。');

      const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
      hotelDbV2ValidateSourceSheet_(sourceSheet, sourceMap);
      const current = hotelDbV2ReadFacility_(sourceSheet, sourceRow, sourceMap);

      const expected = {
        postalCode: hotelDbV2Clean_(row[map['元郵便番号'] - 1]),
        municipality: hotelDbV2Clean_(row[map['元市区町村'] - 1]),
        address: hotelDbV2Clean_(row[map['元住所'] - 1]),
        name: hotelDbV2Clean_(row[map['元施設名'] - 1])
      };

      const conflict =
        hotelDbV2Clean_(current.postalCode) !== expected.postalCode ||
        hotelDbV2Clean_(current.municipality) !== expected.municipality ||
        hotelDbV2Clean_(current.address) !== expected.address ||
        hotelDbV2Clean_(current.name) !== expected.name;

      if (conflict) {
        correctionSheet.getRange(correctionRow, map['状態']).setValue('要再確認');
        result.conflicts++;
        historySheet.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, {
          '日時': hotelDbV2Timestamp_(),
          '元シート': sourceSheet.getName(),
          '元シートID': sourceSheetId,
          '元行': sourceRow,
          '施設名': current.name,
          '処理': '承認修正反映',
          '結果': '競合・未反映',
          'Place ID': row[map['Place ID'] - 1],
          '一致スコア': row[map['一致スコア'] - 1],
          '営業状態': row[map['営業状態'] - 1],
          '詳細': '候補作成後に元データが変更されています。'
        }));
        return;
      }

      const proposed = {
        postalCode: hotelDbV2Clean_(row[map['修正郵便番号'] - 1]),
        municipality: hotelDbV2Clean_(row[map['修正市区町村'] - 1]),
        address: hotelDbV2Clean_(row[map['修正住所'] - 1]),
        facilityName: hotelDbV2Clean_(row[map['修正施設名'] - 1])
      };

      Object.keys(proposed).forEach(function(key) {
        if (sourceMap[key] && proposed[key]) {
          sourceSheet.getRange(sourceRow, sourceMap[key]).setValue(proposed[key]);
        }
      });

      correctionSheet.getRange(correctionRow, map['状態']).setValue('反映済み');
      correctionSheet.getRange(correctionRow, map['反映日時']).setValue(hotelDbV2Timestamp_());
      result.applied++;

      historySheet.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, {
        '日時': hotelDbV2Timestamp_(),
        '元シート': sourceSheet.getName(),
        '元シートID': sourceSheetId,
        '元行': sourceRow,
        '施設名': proposed.facilityName || current.name,
        '処理': '承認修正反映',
        '結果': '反映済み',
        'Place ID': row[map['Place ID'] - 1],
        '一致スコア': row[map['一致スコア'] - 1],
        '営業状態': row[map['営業状態'] - 1],
        '詳細': row[map['差分'] - 1]
      }));
    } catch (error) {
      result.errors++;
      correctionSheet.getRange(correctionRow, map['状態']).setValue('反映エラー');
    }
  });

  return result;
}

function hotelDbV2ResetCheckpoint_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  hotelDbV2ClearCheckpoint_(spreadsheet, sheet, 'enrich');
  hotelDbV2ClearCheckpoint_(spreadsheet, sheet, 'refresh');
  return { reset: true, sheet: sheet.getName() };
}

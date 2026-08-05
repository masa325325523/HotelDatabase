/**
 * 全国宿泊施設データベース Ver2.0
 * 自動生成された分割モジュール。関数名は衝突回避のため hotelDbV2 接頭辞を使用。
 */

function hotelDbV2GetHeaderMap_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};

  Object.keys(HOTEL_DB_V2_CONFIG.SOURCE_ALIASES).forEach(function(key) {
    const aliases = HOTEL_DB_V2_CONFIG.SOURCE_ALIASES[key].map(hotelDbV2NormalizeText_);
    for (let column = 0; column < headers.length; column++) {
      if (aliases.indexOf(hotelDbV2NormalizeText_(headers[column])) !== -1) {
        map[key] = column + 1;
        break;
      }
    }
  });

  return map;
}

function hotelDbV2ValidateSourceSheet_(sheet, map) {
  const reserved = Object.keys(HOTEL_DB_V2_CONFIG.SHEETS).map(function(key) {
    return HOTEL_DB_V2_CONFIG.SHEETS[key];
  });
  if (reserved.indexOf(sheet.getName()) !== -1) {
    throw new Error('出力用シートでは実行できません。元データのシートを開いてください。');
  }
  if (!map.facilityName) throw new Error('「施設名」列が見つかりません。');
  if (!map.address && !map.municipality) {
    throw new Error('「住所」または「市区町村」列が見つかりません。');
  }
}

function hotelDbV2EnsureOutputHeaders_(sheet) {
  const currentLastColumn = Math.max(1, sheet.getLastColumn());
  const current = sheet.getRange(1, 1, 1, currentLastColumn).getDisplayValues()[0];
  const normalized = current.map(hotelDbV2NormalizeText_);
  const missing = [];

  Object.keys(HOTEL_DB_V2_CONFIG.OUTPUT_HEADERS).forEach(function(key) {
    const header = HOTEL_DB_V2_CONFIG.OUTPUT_HEADERS[key];
    const aliases = HOTEL_DB_V2_CONFIG.SOURCE_ALIASES[key] || [header];
    const exists = aliases.some(function(alias) {
      return normalized.indexOf(hotelDbV2NormalizeText_(alias)) !== -1;
    });
    if (!exists) missing.push(header);
  });

  if (missing.length) {
    sheet.getRange(1, currentLastColumn + 1, 1, missing.length).setValues([missing]);
  }
  return hotelDbV2GetHeaderMap_(sheet);
}

function hotelDbV2ReadFacility_(sheet, rowNumber, map) {
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  function value(key) {
    return map[key] ? hotelDbV2Clean_(values[map[key] - 1]) : '';
  }

  return {
    rowNumber: rowNumber,
    postalCode: value('postalCode'),
    municipality: value('municipality'),
    address: value('address'),
    name: value('facilityName'),
    category: value('category'),
    notes: value('notes'),
    placeId: value('placeId')
  };
}

function hotelDbV2BuildGoogleData_(place, score, facility, decision) {
  const location = place && place.location ? place.location : {};
  return {
    placeId: hotelDbV2Clean_(place && place.id),
    googleName: hotelDbV2GetDisplayName_(place),
    googleAddress: hotelDbV2GetJapaneseFullAddress_(place),
    phone: hotelDbV2Clean_(
      place && (place.nationalPhoneNumber || place.internationalPhoneNumber)
    ),
    website: hotelDbV2Clean_(place && place.websiteUri),
    rating: place && place.rating !== undefined ? place.rating : '',
    reviewCount: place && place.userRatingCount !== undefined
      ? place.userRatingCount
      : '',
    businessStatus: hotelDbV2TranslateBusinessStatus_(place && place.businessStatus),
    mapsUrl: hotelDbV2Clean_(place && place.googleMapsUri),
    latitude: location.latitude === undefined ? '' : location.latitude,
    longitude: location.longitude === undefined ? '' : location.longitude,
    matchScore: score,
    checkedAt: hotelDbV2Today_(),
    matchDecision: decision || '',
    proposedPostalCode: hotelDbV2GetPostalCode_(place),
    proposedMunicipality: hotelDbV2GetMunicipalityForSource_(place, facility.municipality),
    proposedAddress: hotelDbV2GetAddressForSource_(place, facility.municipality),
    proposedName: hotelDbV2GetDisplayName_(place)
  };
}

function hotelDbV2WriteGoogleData_(sheet, rowNumber, map, googleData, includePlaceId) {
  const updates = {
    googleName: googleData.googleName,
    googleAddress: googleData.googleAddress,
    phone: googleData.phone,
    website: googleData.website,
    rating: googleData.rating,
    reviewCount: googleData.reviewCount,
    businessStatus: googleData.businessStatus,
    mapsUrl: googleData.mapsUrl,
    latitude: googleData.latitude,
    longitude: googleData.longitude,
    matchScore: googleData.matchScore,
    checkedAt: googleData.checkedAt,
    matchDecision: googleData.matchDecision
  };

  if (includePlaceId) updates.placeId = googleData.placeId;

  Object.keys(updates).forEach(function(key) {
    if (map[key] && updates[key] !== undefined) {
      sheet.getRange(rowNumber, map[key]).setValue(updates[key]);
    }
  });
}

function hotelDbV2GetOrCreateSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const existing = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn()))
      .getDisplayValues()[0];
    if (existing.length < headers.length || hotelDbV2Clean_(existing[0]) !== headers[0]) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  }
  return sheet;
}

function hotelDbV2HeaderIndex_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    map[header] = index + 1;
  });
  return map;
}

function hotelDbV2BuildKeyIndex_(sheet) {
  const index = {};
  if (sheet.getLastRow() < 2) return index;

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .getDisplayValues();
  values.forEach(function(row, offset) {
    const key = hotelDbV2Clean_(row[0]);
    if (key) index[key] = offset + 2;
  });
  return index;
}

function hotelDbV2CreateOutputContext_(spreadsheet) {
  const corrections = hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS, HOTEL_DB_V2_CORRECTION_HEADERS
  );
  const review = hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.REVIEW, HOTEL_DB_V2_REVIEW_HEADERS
  );
  const history = hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY, HOTEL_DB_V2_HISTORY_HEADERS
  );
  const summary = hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.SUMMARY, HOTEL_DB_V2_SUMMARY_HEADERS
  );

  return {
    corrections: {
      sheet: corrections,
      headers: HOTEL_DB_V2_CORRECTION_HEADERS,
      map: hotelDbV2HeaderIndex_(HOTEL_DB_V2_CORRECTION_HEADERS),
      index: hotelDbV2BuildKeyIndex_(corrections)
    },
    review: {
      sheet: review,
      headers: HOTEL_DB_V2_REVIEW_HEADERS,
      map: hotelDbV2HeaderIndex_(HOTEL_DB_V2_REVIEW_HEADERS),
      index: hotelDbV2BuildKeyIndex_(review)
    },
    history: history,
    summary: summary
  };
}

function hotelDbV2RowFromObject_(headers, object) {
  return headers.map(function(header) {
    return object[header] === undefined ? '' : object[header];
  });
}

function hotelDbV2Upsert_(context, key, object, preserveState) {
  const existingRow = context.index[key] || null;
  const values = hotelDbV2RowFromObject_(context.headers, object);

  if (existingRow) {
    if (preserveState && context.map['状態']) {
      const stateColumn = context.map['状態'];
      const oldValues = context.sheet
        .getRange(existingRow, 1, 1, values.length)
        .getDisplayValues()[0];
      const oldState = hotelDbV2Clean_(oldValues[stateColumn - 1]);
      const ignored = {
        '状態': true,
        '確認日': true,
        '反映日時': true
      };
      const unchanged = context.headers.every(function(header, index) {
        if (ignored[header]) return true;
        return hotelDbV2Clean_(oldValues[index]) === hotelDbV2Clean_(values[index]);
      });

      values[stateColumn - 1] = unchanged && oldState ? oldState : '未確認';
    }
    context.sheet.getRange(existingRow, 1, 1, values.length).setValues([values]);
    return existingRow;
  }

  const rowNumber = context.sheet.getLastRow() + 1;
  context.sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
  context.index[key] = rowNumber;
  return rowNumber;
}

function hotelDbV2AppendHistory_(context, object) {
  context.history.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, object));
}

function hotelDbV2CompareFacility_(facility, googleData) {
  const differences = [];

  const sourcePostal = hotelDbV2NormalizePostalCode_(facility.postalCode);
  const proposedPostal = hotelDbV2NormalizePostalCode_(googleData.proposedPostalCode);
  if (proposedPostal && sourcePostal !== proposedPostal) {
    differences.push('郵便番号');
  }

  const sourceMunicipality = hotelDbV2NormalizeText_(facility.municipality);
  const proposedMunicipality = hotelDbV2NormalizeText_(googleData.proposedMunicipality);
  if (proposedMunicipality && sourceMunicipality !== proposedMunicipality) {
    differences.push('市区町村');
  }

  const sourceAddress = hotelDbV2NormalizeAddress_(facility.address);
  const proposedAddress = hotelDbV2NormalizeAddress_(googleData.proposedAddress);
  if (proposedAddress && sourceAddress !== proposedAddress) {
    differences.push('住所');
  }

  const sourceName = hotelDbV2NormalizeText_(facility.name);
  const proposedName = hotelDbV2NormalizeText_(googleData.proposedName);
  if (proposedName && sourceName !== proposedName) {
    differences.push('施設名');
  }

  return differences;
}

function hotelDbV2CorrectionCandidate_(sheet, facility, googleData, differences) {
  const key = [sheet.getSheetId(), facility.rowNumber].join('|');
  return {
    key: key,
    object: {
      '候補キー': key,
      '状態': '未確認',
      '元シート': sheet.getName(),
      '元シートID': sheet.getSheetId(),
      '元行': facility.rowNumber,
      '元郵便番号': facility.postalCode,
      '修正郵便番号': googleData.proposedPostalCode,
      '元市区町村': facility.municipality,
      '修正市区町村': googleData.proposedMunicipality,
      '元住所': facility.address,
      '修正住所': googleData.proposedAddress,
      '元施設名': facility.name,
      '修正施設名': googleData.proposedName,
      'Place ID': googleData.placeId,
      '一致スコア': googleData.matchScore,
      '営業状態': googleData.businessStatus,
      'Google Maps URL': googleData.mapsUrl,
      '差分': differences.join('・'),
      '確認日': googleData.checkedAt,
      '反映日時': ''
    }
  };
}

function hotelDbV2ReviewCandidate_(sheet, facility, reason, candidate, score, detail) {
  const place = candidate || {};
  const key = [sheet.getSheetId(), facility.rowNumber].join('|');
  return {
    key: key,
    object: {
      '確認キー': key,
      '状態': '未確認',
      '元シート': sheet.getName(),
      '元シートID': sheet.getSheetId(),
      '元行': facility.rowNumber,
      '郵便番号': facility.postalCode,
      '市区町村': facility.municipality,
      '住所': facility.address,
      '施設名': facility.name,
      '宿泊分類': facility.category,
      '理由': reason,
      '候補施設名': hotelDbV2GetDisplayName_(place),
      '候補住所': hotelDbV2GetJapaneseFullAddress_(place),
      '候補Place ID': hotelDbV2Clean_(place.id),
      '一致スコア': score === undefined || score === null ? '' : score,
      '営業状態': hotelDbV2TranslateBusinessStatus_(place.businessStatus),
      'Google Maps URL': hotelDbV2Clean_(place.googleMapsUri),
      '確認日': hotelDbV2Today_(),
      '詳細': hotelDbV2Clean_(detail)
    }
  };
}

function hotelDbV2HistoryObject_(sheet, facility, action, result, placeId, score, status, detail) {
  return {
    '日時': hotelDbV2Timestamp_(),
    '元シート': sheet.getName(),
    '元シートID': sheet.getSheetId(),
    '元行': facility.rowNumber,
    '施設名': facility.name,
    '処理': action,
    '結果': result,
    'Place ID': placeId,
    '一致スコア': score,
    '営業状態': status,
    '詳細': detail
  };
}


/**
 * Places.gs
 * Places API (New) integration for HotelDatabase.
 *
 * Script property:
 *   GOOGLE_PLACES_API_KEY
 */

const PLACES = Object.freeze({
  API_BASE: 'https://places.googleapis.com/v1',
  LANGUAGE: 'ja',
  REGION: 'JP',
  TIMEZONE: 'Asia/Tokyo',
  REQUEST_INTERVAL_MS: 150,
  MAX_RETRIES: 3,
  RETRY_BASE_MS: 800,
  MIN_MATCH_SCORE: 55,
  AUTO_ACCEPT_SCORE: 72,
  DEFAULT_BATCH_SIZE: 100,
  SEARCH_FIELDS: [
    'places.id',
    'places.displayName',
    'places.formattedAddress',
    'places.businessStatus',
    'places.rating',
    'places.userRatingCount',
    'places.location',
    'places.types',
    'places.googleMapsUri'
  ].join(','),
  DETAILS_FIELDS: [
    'id',
    'displayName',
    'formattedAddress',
    'addressComponents',
    'nationalPhoneNumber',
    'internationalPhoneNumber',
    'websiteUri',
    'googleMapsUri',
    'businessStatus',
    'rating',
    'userRatingCount',
    'location',
    'types',
    'regularOpeningHours'
  ].join(','),
  HEADERS: Object.freeze({
    postalCode: ['郵便番号', '〒', 'postal_code'],
    municipality: ['市区町村名', '市区町村', '自治体名'],
    address: ['住所（番地まで）', '住所(番地まで)', '住所', '所在地'],
    facilityName: ['施設名', '宿泊施設名', '名称'],
    category: ['宿泊分類', '分類'],
    notes: ['備考', 'メモ'],
    placeId: ['Place ID', 'place_id', 'Google Place ID'],
    googleName: ['Google施設名', 'Google名称'],
    googleAddress: ['Google住所', 'Google Maps住所'],
    phone: ['電話番号', '電話'],
    website: ['公式サイト', 'ウェブサイト', 'website'],
    rating: ['評価', 'Google評価'],
    reviewCount: ['口コミ数', 'レビュー数'],
    businessStatus: ['営業状態', 'Google営業状態'],
    latitude: ['緯度', 'lat'],
    longitude: ['経度', 'lng'],
    checkedAt: ['最終確認日', '確認日'],
    matchScore: ['一致スコア', 'Google一致スコア']
  })
});

function getPlacesApiKey() {
  const scriptKey = PropertiesService.getScriptProperties()
    .getProperty('GOOGLE_PLACES_API_KEY');
  if (scriptKey) return String(scriptKey).trim();

  if (typeof getConfig === 'function') {
    const candidates = ['GOOGLE_PLACES_API_KEY', 'PLACES_API_KEY', 'Google Places API Key'];
    for (let i = 0; i < candidates.length; i++) {
      const value = getConfig(candidates[i]);
      if (value) return String(value).trim();
    }
  }

  throw new Error(
    'Google Places APIキーが未設定です。スクリプトプロパティに ' +
    'GOOGLE_PLACES_API_KEY を設定してください。'
  );
}

function callPlacesNewApi(path, options) {
  const opts = options || {};
  const request = {
    method: opts.method || 'get',
    headers: {
      'X-Goog-Api-Key': getPlacesApiKey(),
      'X-Goog-FieldMask': opts.fieldMask || ''
    },
    muteHttpExceptions: true,
    followRedirects: true
  };

  if (opts.payload !== undefined) {
    request.contentType = 'application/json';
    request.payload = JSON.stringify(opts.payload);
  }

  const url = PLACES.API_BASE + path;
  let lastError = null;

  for (let attempt = 0; attempt <= PLACES.MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, request);
      const httpCode = response.getResponseCode();
      const body = response.getContentText();
      let json = {};

      if (body) {
        try {
          json = JSON.parse(body);
        } catch (parseError) {
          throw new Error('Places API応答をJSON解析できません: HTTP ' + httpCode);
        }
      }

      if (httpCode >= 200 && httpCode < 300) return json;

      const apiMessage = json.error && json.error.message ? json.error.message : '';
      const retryable = httpCode === 429 || httpCode >= 500;
      lastError = new Error(
        'Places API (New) エラー: HTTP=' + httpCode +
        (apiMessage ? ', message=' + apiMessage : '')
      );
      if (!retryable || attempt === PLACES.MAX_RETRIES) throw lastError;
    } catch (error) {
      lastError = error;
      if (attempt === PLACES.MAX_RETRIES) throw error;
    }

    Utilities.sleep(PLACES.RETRY_BASE_MS * Math.pow(2, attempt));
  }

  throw lastError || new Error('Places API (New) 呼び出しに失敗しました。');
}

function safeString(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function normalizeText(value) {
  return safeString(value)
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s　・･,，.．'’"“”\-ー―‐_/\\()（）\[\]【】]/g, '');
}

function normalizeAddress(address) {
  return safeString(address)
    .normalize('NFKC')
    .replace(/^日本[、,\s]*/u, '')
    .replace(/〒\s*\d{3}-?\d{4}\s*/u, '')
    .replace(/[‐‑‒–—―ー−]/g, '-')
    .replace(/丁目/g, '-')
    .replace(/番地の?/g, '-')
    .replace(/番/g, '-')
    .replace(/号/g, '')
    .replace(/\s+/g, '')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '');
}

function normalizePostalCode(value) {
  const digits = safeString(value).replace(/\D/g, '');
  return digits.length === 7 ? digits.slice(0, 3) + '-' + digits.slice(3) : '';
}

function formatCheckedAt() {
  return Utilities.formatDate(new Date(), PLACES.TIMEZONE, 'yyyy-MM-dd');
}

function searchPlaceFromGoogle(query, options) {
  const text = safeString(query);
  if (!text) return [];

  const opts = options || {};
  const payload = {
    textQuery: text,
    languageCode: PLACES.LANGUAGE,
    regionCode: PLACES.REGION,
    pageSize: Math.max(1, Math.min(20, Number(opts.pageSize || 20)))
  };

  if (opts.locationBias) payload.locationBias = opts.locationBias;
  if (opts.locationRestriction) payload.locationRestriction = opts.locationRestriction;
  if (opts.includedType) payload.includedType = opts.includedType;
  if (opts.openNow === true) payload.openNow = true;

  const json = callPlacesNewApi('/places:searchText', {
    method: 'post',
    fieldMask: PLACES.SEARCH_FIELDS,
    payload: payload
  });

  return (json.places || []).map(adaptNewPlaceToLegacyShape_);
}

function adaptNewPlaceToLegacyShape_(place) {
  return {
    place_id: safeString(place && place.id),
    name: place && place.displayName ? safeString(place.displayName.text) : '',
    formatted_address: safeString(place && place.formattedAddress),
    business_status: safeString(place && place.businessStatus),
    rating: place && place.rating !== undefined ? place.rating : '',
    user_ratings_total: place && place.userRatingCount !== undefined ? place.userRatingCount : '',
    geometry: {
      location: {
        lat: place && place.location && place.location.latitude !== undefined
          ? place.location.latitude : '',
        lng: place && place.location && place.location.longitude !== undefined
          ? place.location.longitude : ''
      }
    },
    types: place && place.types ? place.types : [],
    url: safeString(place && place.googleMapsUri),
    _newApi: place || {}
  };
}

function buildFacilitySearchQuery(facility) {
  return [facility.name, facility.address, facility.municipality]
    .map(safeString)
    .filter(Boolean)
    .join(' ');
}

function findBestPlaceCandidate(facility) {
  const query = buildFacilitySearchQuery(facility);
  if (!query) return null;

  const results = searchPlaceFromGoogle(query);
  if (!results.length) return null;

  const ranked = results.map(function(place) {
    return { place: place, score: calculatePlaceMatchScore(facility, place) };
  }).sort(function(a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return Number(b.place.user_ratings_total || 0) - Number(a.place.user_ratings_total || 0);
  });

  const best = ranked[0];
  if (!best || best.score < PLACES.MIN_MATCH_SCORE) return null;

  return {
    place: best.place,
    score: best.score,
    accepted: best.score >= PLACES.AUTO_ACCEPT_SCORE,
    candidates: ranked.slice(0, 5)
  };
}

function calculatePlaceMatchScore(facility, place) {
  const sourceName = normalizeText(facility.name);
  const googleName = normalizeText(place.name);
  const sourceAddress = normalizeAddress(facility.address);
  const googleAddress = normalizeAddress(place.formatted_address);
  const municipality = normalizeText(facility.municipality);
  let score = 0;

  if (sourceName && googleName) {
    if (sourceName === googleName) score += 50;
    else if (sourceName.indexOf(googleName) !== -1 || googleName.indexOf(sourceName) !== -1) score += 38;
    else score += Math.round(similarityRatio(sourceName, googleName) * 32);
  }

  if (sourceAddress && googleAddress) {
    if (sourceAddress === googleAddress) score += 40;
    else if (sourceAddress.indexOf(googleAddress) !== -1 || googleAddress.indexOf(sourceAddress) !== -1) score += 34;
    else score += Math.round(similarityRatio(sourceAddress, googleAddress) * 28);
  }

  if (municipality && normalizeText(place.formatted_address).indexOf(municipality) !== -1) score += 10;
  if (place.business_status === 'OPERATIONAL') score += 3;
  if (place.business_status === 'CLOSED_PERMANENTLY') score -= 30;
  return Math.max(0, Math.min(100, score));
}

function similarityRatio(a, b) {
  if (!a || !b) return 0;
  if (a === b) return 1;
  const longer = a.length >= b.length ? a : b;
  const shorter = a.length >= b.length ? b : a;
  if (!longer.length) return 1;
  return (longer.length - levenshteinDistance(longer, shorter)) / longer.length;
}

function levenshteinDistance(a, b) {
  const matrix = [];
  let i;
  let j;
  for (i = 0; i <= b.length; i++) matrix[i] = [i];
  for (j = 0; j <= a.length; j++) matrix[0][j] = j;
  for (i = 1; i <= b.length; i++) {
    for (j = 1; j <= a.length; j++) {
      const cost = b.charAt(i - 1) === a.charAt(j - 1) ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j - 1] + cost,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j] + 1
      );
    }
  }
  return matrix[b.length][a.length];
}

function getPlaceDetails(placeId) {
  const id = safeString(placeId);
  if (!id) return null;
  const json = callPlacesNewApi('/places/' + encodeURIComponent(id), {
    method: 'get',
    fieldMask: PLACES.DETAILS_FIELDS
  });
  return adaptNewPlaceDetailsToLegacyShape_(json);
}

function adaptNewPlaceDetailsToLegacyShape_(place) {
  const components = (place.addressComponents || []).map(function(component) {
    return {
      long_name: safeString(component.longText),
      short_name: safeString(component.shortText),
      types: component.types || []
    };
  });

  return {
    place_id: safeString(place.id),
    name: place.displayName ? safeString(place.displayName.text) : '',
    formatted_address: safeString(place.formattedAddress),
    address_components: components,
    formatted_phone_number: safeString(place.nationalPhoneNumber),
    international_phone_number: safeString(place.internationalPhoneNumber),
    website: safeString(place.websiteUri),
    url: safeString(place.googleMapsUri),
    business_status: safeString(place.businessStatus),
    rating: place.rating === undefined ? '' : place.rating,
    user_ratings_total: place.userRatingCount === undefined ? '' : place.userRatingCount,
    geometry: {
      location: {
        lat: place.location && place.location.latitude !== undefined ? place.location.latitude : '',
        lng: place.location && place.location.longitude !== undefined ? place.location.longitude : ''
      }
    },
    types: place.types || [],
    opening_hours: place.regularOpeningHours || null,
    _newApi: place
  };
}

function normalizePlaceData(place) {
  if (!place) return {};
  const components = place.address_components || [];
  const location = place.geometry && place.geometry.location ? place.geometry.location : {};
  return {
    placeId: safeString(place.place_id),
    name: safeString(place.name),
    address: safeString(place.formatted_address),
    postalCode: extractAddressComponent(components, 'postal_code'),
    municipality: extractMunicipality(components),
    phone: safeString(place.formatted_phone_number || place.international_phone_number),
    website: safeString(place.website),
    googleMapsUrl: safeString(place.url),
    rating: place.rating === undefined ? '' : place.rating,
    reviews: place.user_ratings_total === undefined ? '' : place.user_ratings_total,
    businessStatus: translateBusinessStatus(place.business_status),
    rawBusinessStatus: safeString(place.business_status),
    latitude: location.lat === undefined ? '' : location.lat,
    longitude: location.lng === undefined ? '' : location.lng,
    types: place.types || [],
    checkedAt: formatCheckedAt()
  };
}

function extractAddressComponent(components, type) {
  for (let i = 0; i < components.length; i++) {
    if ((components[i].types || []).indexOf(type) !== -1) return safeString(components[i].long_name);
  }
  return '';
}

function extractMunicipality(components) {
  const priorities = ['locality', 'administrative_area_level_2', 'sublocality_level_1', 'postal_town'];
  for (let i = 0; i < priorities.length; i++) {
    const value = extractAddressComponent(components, priorities[i]);
    if (value) return value;
  }
  return '';
}

function translateBusinessStatus(status) {
  switch (safeString(status)) {
    case 'OPERATIONAL': return '営業中';
    case 'CLOSED_TEMPORARILY': return '一時休業';
    case 'CLOSED_PERMANENTLY': return '閉業';
    case 'FUTURE_OPENING': return '開業予定';
    default: return '不明';
  }
}

function buildFacilityNotes(existingNotes, facility, normalized, matchScore) {
  const notes = [];
  const existing = safeString(existingNotes);
  if (existing) notes.push(existing);
  if (normalized.businessStatus) notes.push('Google営業状態:' + normalized.businessStatus);
  if (matchScore !== '' && matchScore !== null && matchScore !== undefined) notes.push('一致スコア:' + matchScore);

  const sourcePostal = normalizePostalCode(facility.postalCode);
  const googlePostal = normalizePostalCode(normalized.postalCode);
  if (sourcePostal && googlePostal) {
    notes.push(sourcePostal === googlePostal ? '郵便番号一致' : '郵便番号不一致(' + googlePostal + ')');
  }
  if (normalized.website) notes.push('公式サイト確認');
  if (normalized.phone) notes.push('電話番号確認');
  notes.push('Google確認日:' + normalized.checkedAt);
  return uniqueStrings(notes).join('／');
}

function uniqueStrings(items) {
  const seen = {};
  return items.filter(function(item) {
    const key = safeString(item);
    if (!key || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function updateFacilityRow(sheet, rowNumber, placeData, headerMap, facility, matchScore) {
  if (!sheet) throw new Error('sheet が指定されていません。');
  if (rowNumber < 2) throw new Error('rowNumber は2以上を指定してください。');

  const map = headerMap || getHeaderMap(sheet);
  const source = facility || readFacilityFromRow(sheet, rowNumber, map);
  const normalized = placeData && placeData.placeId !== undefined ? placeData : normalizePlaceData(placeData);
  const updates = {
    placeId: normalized.placeId,
    googleName: normalized.name,
    googleAddress: normalized.address,
    phone: normalized.phone,
    website: normalized.website,
    rating: normalized.rating,
    reviewCount: normalized.reviews,
    businessStatus: normalized.businessStatus,
    latitude: normalized.latitude,
    longitude: normalized.longitude,
    checkedAt: normalized.checkedAt || formatCheckedAt(),
    matchScore: matchScore,
    notes: buildFacilityNotes(source.notes, source, normalized, matchScore)
  };

  Object.keys(updates).forEach(function(key) {
    const column = map[key];
    if (column && updates[key] !== undefined) sheet.getRange(rowNumber, column).setValue(updates[key]);
  });
  return updates;
}

function enrichActiveSheetWithPlaces() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  return enrichSheetWithPlaces(sheet, { maxRows: PLACES.DEFAULT_BATCH_SIZE });
}

function enrichSheetWithPlaces(sheet, options) {
  if (!sheet) throw new Error('対象シートがありません。');
  const opts = options || {};
  const headerMap = getHeaderMap(sheet);
  validateRequiredHeaders(headerMap);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { processed: 0, updated: 0, skipped: 0, errors: [] };

  const startRow = Math.max(2, Number(opts.startRow || 2));
  const maxRows = Math.max(1, Number(opts.maxRows || PLACES.DEFAULT_BATCH_SIZE));
  const endRow = Math.min(lastRow, startRow + maxRows - 1);
  let processed = 0;
  let updated = 0;
  let skipped = 0;
  const errors = [];

  for (let row = startRow; row <= endRow; row++) {
    processed++;
    try {
      const facility = readFacilityFromRow(sheet, row, headerMap);
      if (!facility.name && !facility.address) {
        skipped++;
        continue;
      }
      if (opts.skipExisting !== false && facility.placeId) {
        skipped++;
        continue;
      }

      const best = findBestPlaceCandidate(facility);
      if (!best || !best.accepted) {
        if (headerMap.notes) {
          const note = best ? 'Google候補要確認（一致スコア:' + best.score + '）' : 'Google候補なし';
          sheet.getRange(row, headerMap.notes).setValue(
            uniqueStrings([facility.notes, note, 'Google確認日:' + formatCheckedAt()]).join('／')
          );
        }
        skipped++;
        continue;
      }

      const details = getPlaceDetails(best.place.place_id);
      const normalized = normalizePlaceData(details || best.place);
      updateFacilityRow(sheet, row, normalized, headerMap, facility, best.score);
      updated++;
      Utilities.sleep(PLACES.REQUEST_INTERVAL_MS);
    } catch (error) {
      errors.push({ row: row, message: error.message });
      if (headerMap.notes) {
        const oldNotes = safeString(sheet.getRange(row, headerMap.notes).getValue());
        sheet.getRange(row, headerMap.notes).setValue(
          uniqueStrings([oldNotes, 'Google取得エラー:' + error.message]).join('／')
        );
      }
    }
  }

  const summary = {
    processed: processed,
    updated: updated,
    skipped: skipped,
    errors: errors,
    nextStartRow: endRow < lastRow ? endRow + 1 : null
  };
  Logger.log(JSON.stringify(summary));
  return summary;
}

function refreshExistingPlaceDetails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headerMap = getHeaderMap(sheet);
  validateRequiredHeaders(headerMap);
  if (!headerMap.placeId) throw new Error('Place ID列がありません。');

  const lastRow = sheet.getLastRow();
  let updated = 0;
  for (let row = 2; row <= lastRow; row++) {
    const facility = readFacilityFromRow(sheet, row, headerMap);
    if (!facility.placeId) continue;
    const details = getPlaceDetails(facility.placeId);
    if (!details) continue;
    updateFacilityRow(
      sheet,
      row,
      normalizePlaceData(details),
      headerMap,
      facility,
      headerMap.matchScore ? sheet.getRange(row, headerMap.matchScore).getValue() : ''
    );
    updated++;
    Utilities.sleep(PLACES.REQUEST_INTERVAL_MS);
  }
  return { updated: updated };
}

function getHeaderMap(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};
  Object.keys(PLACES.HEADERS).forEach(function(key) {
    const aliases = PLACES.HEADERS[key].map(normalizeText);
    for (let col = 0; col < headers.length; col++) {
      if (aliases.indexOf(normalizeText(headers[col])) !== -1) {
        map[key] = col + 1;
        break;
      }
    }
  });
  return map;
}

function validateRequiredHeaders(headerMap) {
  if (!headerMap.facilityName && !headerMap.address) {
    throw new Error('「施設名」または「住所」列が必要です。');
  }
}

function readFacilityFromRow(sheet, rowNumber, headerMap) {
  const lastColumn = sheet.getLastColumn();
  const row = sheet.getRange(rowNumber, 1, 1, lastColumn).getDisplayValues()[0];
  function value(key) {
    return headerMap[key] ? safeString(row[headerMap[key] - 1]) : '';
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

function findDuplicateFacilities() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headerMap = getHeaderMap(sheet);
  validateRequiredHeaders(headerMap);
  const lastRow = sheet.getLastRow();
  const seen = {};
  const duplicates = [];

  for (let row = 2; row <= lastRow; row++) {
    const facility = readFacilityFromRow(sheet, row, headerMap);
    if (!facility.name && !facility.address) continue;
    const key = normalizeText(facility.name) + '|' + normalizeAddress(facility.address);
    if (!key || key === '|') continue;
    if (seen[key]) duplicates.push({ firstRow: seen[key], duplicateRow: row, key: key });
    else seen[key] = row;
  }

  Logger.log(JSON.stringify(duplicates));
  return duplicates;
}

function addPlacesOutputHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const required = [
    'Place ID', 'Google施設名', 'Google住所', '電話番号', '公式サイト',
    '評価', '口コミ数', '営業状態', '緯度', '経度', '一致スコア', '最終確認日'
  ];
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const current = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const normalizedCurrent = current.map(normalizeText);
  const missing = required.filter(function(header) {
    return normalizedCurrent.indexOf(normalizeText(header)) === -1;
  });
  if (missing.length) sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
  return { added: missing };
}

function testPlacesAPI() {
  const results = searchPlaceFromGoogle('名古屋市 ホテル', { pageSize: 3 });
  Logger.log(JSON.stringify(results));
  return results;
}

function testFacilityMatching() {
  const sample = {
    name: '名古屋マリオットアソシアホテル',
    municipality: '名古屋市中村区',
    address: '愛知県名古屋市中村区名駅1丁目1番4号'
  };
  const result = findBestPlaceCandidate(sample);
  Logger.log(JSON.stringify(result));
  return result;
}

function diagnosePlacesConfiguration() {
  const result = {
    apiKeyConfigured: false,
    activeSheet: '',
    headerMap: {},
    apiVersion: 'Places API (New)',
    errors: []
  };

  try {
    result.apiKeyConfigured = Boolean(getPlacesApiKey());
  } catch (error) {
    result.errors.push(error.message);
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    result.activeSheet = sheet.getName();
    result.headerMap = getHeaderMap(sheet);
    validateRequiredHeaders(result.headerMap);
  } catch (error) {
    result.errors.push(error.message);
  }

  Logger.log(JSON.stringify(result));
  return result;
}

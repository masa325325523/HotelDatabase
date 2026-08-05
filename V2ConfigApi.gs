/**
 * 全国宿泊施設データベース Ver2.0
 * 自動生成された分割モジュール。関数名は衝突回避のため hotelDbV2 接頭辞を使用。
 */

const HOTEL_DB_V2_CONFIG = Object.freeze({
  API_BASE: 'https://places.googleapis.com/v1',
  API_KEY_PROPERTY: 'GOOGLE_PLACES_API_KEY',
  LANGUAGE_CODE: 'ja',
  REGION_CODE: 'JP',
  TIMEZONE: 'Asia/Tokyo',
  REQUEST_INTERVAL_MS: 180,
  MAX_RETRIES: 3,
  RETRY_BASE_MS: 800,
  MIN_MATCH_SCORE: 55,
  AUTO_ACCEPT_SCORE: 75,
  BATCH_SIZE: 50,
  TEST_ROWS: 3,
  MAX_EXECUTION_MS: 260000,
  SHEETS: Object.freeze({
    CORRECTIONS: '修正候補',
    REVIEW: '要確認',
    HISTORY: '修正履歴',
    DUPLICATES: '重複候補',
    SUMMARY: '実行サマリー'
  }),
  SOURCE_ALIASES: Object.freeze({
    postalCode: ['郵便番号', '〒', 'postal_code'],
    municipality: ['市区町村名', '市区町村', '自治体名'],
    address: ['住所（番地まで）', '住所(番地まで)', '住所', '所在地', '以下住所'],
    facilityName: ['施設名', '宿泊施設名', '名称'],
    category: ['宿泊分類', '分類'],
    notes: ['備考', '備考欄', 'メモ'],
    placeId: ['Place ID', 'place_id', 'Google Place ID'],
    googleName: ['Google施設名', 'Google名称'],
    googleAddress: ['Google住所', 'Google Maps住所'],
    phone: ['電話番号', '電話'],
    website: ['公式サイト', 'ウェブサイト', 'website'],
    rating: ['評価', 'Google評価'],
    reviewCount: ['口コミ数', 'レビュー数'],
    businessStatus: ['営業状態', 'Google営業状態'],
    mapsUrl: ['Google Maps URL', 'GoogleマップURL', 'Maps URL'],
    latitude: ['緯度', 'lat'],
    longitude: ['経度', 'lng'],
    matchScore: ['一致スコア', 'Google一致スコア'],
    checkedAt: ['最終確認日', '確認日'],
    matchDecision: ['照合判定', 'Google照合判定']
  }),
  OUTPUT_HEADERS: Object.freeze({
    placeId: 'Place ID',
    googleName: 'Google施設名',
    googleAddress: 'Google住所',
    phone: '電話番号',
    website: '公式サイト',
    rating: '評価',
    reviewCount: '口コミ数',
    businessStatus: '営業状態',
    mapsUrl: 'Google Maps URL',
    latitude: '緯度',
    longitude: '経度',
    matchScore: '一致スコア',
    checkedAt: '最終確認日',
    matchDecision: '照合判定'
  }),
  SEARCH_FIELDS: [
    'places.id',
    'places.displayName',
    'places.formattedAddress',
    'places.addressComponents',
    'places.businessStatus',
    'places.location',
    'places.types',
    'places.googleMapsUri'
  ].join(','),
  DETAILS_FIELDS: [
    'id',
    'displayName',
    'formattedAddress',
    'shortFormattedAddress',
    'addressComponents',
    'postalAddress',
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
  ].join(',')
});

const HOTEL_DB_V2_CORRECTION_HEADERS = [
  '候補キー', '状態', '元シート', '元シートID', '元行',
  '元郵便番号', '修正郵便番号',
  '元市区町村', '修正市区町村',
  '元住所', '修正住所',
  '元施設名', '修正施設名',
  'Place ID', '一致スコア', '営業状態', 'Google Maps URL',
  '差分', '確認日', '反映日時'
];

const HOTEL_DB_V2_REVIEW_HEADERS = [
  '確認キー', '状態', '元シート', '元シートID', '元行',
  '郵便番号', '市区町村', '住所', '施設名', '宿泊分類',
  '理由', '候補施設名', '候補住所', '候補Place ID',
  '一致スコア', '営業状態', 'Google Maps URL', '確認日', '詳細'
];

const HOTEL_DB_V2_HISTORY_HEADERS = [
  '日時', '元シート', '元シートID', '元行', '施設名',
  '処理', '結果', 'Place ID', '一致スコア', '営業状態', '詳細'
];

const HOTEL_DB_V2_DUPLICATE_HEADERS = [
  '重複キー', '元シート', '元シートID', '判定',
  '行1', '施設名1', '住所1',
  '行2', '施設名2', '住所2',
  'Place ID', '類似度', '確認日', '状態'
];

const HOTEL_DB_V2_SUMMARY_HEADERS = [
  '日時', '元シート', '元シートID', '開始行', '終了行',
  '処理件数', '営業中', '修正候補', '要確認', '未検出',
  '閉業', '一時休業', '開業予定', 'エラー', 'スキップ',
  '次回開始行', '整合確認'
];


function hotelDbV2GetApiKey_() {
  const value = PropertiesService.getScriptProperties()
    .getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY);
  if (!value || !String(value).trim()) {
    throw new Error(
      'スクリプトプロパティ「' + HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY + '」が未設定です。'
    );
  }
  return String(value).trim();
}

function hotelDbV2CallPlacesApi_(path, options) {
  const opts = options || {};
  const request = {
    method: opts.method || 'get',
    headers: {
      'X-Goog-Api-Key': hotelDbV2GetApiKey_(),
      'X-Goog-FieldMask': opts.fieldMask || '',
      'Accept-Language': 'ja'
    },
    muteHttpExceptions: true,
    followRedirects: true
  };

  if (opts.payload !== undefined) {
    request.contentType = 'application/json';
    request.payload = JSON.stringify(opts.payload);
  }

  const url = HOTEL_DB_V2_CONFIG.API_BASE + path;
  let lastError = null;

  for (let attempt = 0; attempt <= HOTEL_DB_V2_CONFIG.MAX_RETRIES; attempt++) {
    try {
      const response = UrlFetchApp.fetch(url, request);
      const code = response.getResponseCode();
      const body = response.getContentText();
      let json = {};

      if (body) {
        try {
          json = JSON.parse(body);
        } catch (error) {
          throw new Error('Places API応答をJSON解析できません: HTTP ' + code);
        }
      }

      if (code >= 200 && code < 300) return json;

      const message = json.error && json.error.message
        ? json.error.message
        : '詳細なし';
      const retryable = code === 429 || code >= 500;
      lastError = new Error(
        'Places API (New) エラー: HTTP=' + code + ', message=' + message
      );

      if (!retryable || attempt === HOTEL_DB_V2_CONFIG.MAX_RETRIES) throw lastError;
    } catch (error) {
      lastError = error;
      if (attempt === HOTEL_DB_V2_CONFIG.MAX_RETRIES) throw error;
    }

    Utilities.sleep(HOTEL_DB_V2_CONFIG.RETRY_BASE_MS * Math.pow(2, attempt));
  }

  throw lastError || new Error('Places API (New) 呼び出しに失敗しました。');
}

function hotelDbV2SearchPlaces_(query) {
  const text = hotelDbV2Clean_(query);
  if (!text) return [];

  const json = hotelDbV2CallPlacesApi_('/places:searchText', {
    method: 'post',
    fieldMask: HOTEL_DB_V2_CONFIG.SEARCH_FIELDS,
    payload: {
      textQuery: text,
      languageCode: HOTEL_DB_V2_CONFIG.LANGUAGE_CODE,
      regionCode: HOTEL_DB_V2_CONFIG.REGION_CODE,
      pageSize: 10,
      includeFutureOpeningBusinesses: true
    }
  });

  return json.places || [];
}

function hotelDbV2GetPlaceDetails_(placeId) {
  const id = hotelDbV2Clean_(placeId);
  if (!id) return null;

  return hotelDbV2CallPlacesApi_('/places/' + encodeURIComponent(id), {
    method: 'get',
    fieldMask: HOTEL_DB_V2_CONFIG.DETAILS_FIELDS
  });
}

function hotelDbV2Clean_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function hotelDbV2NormalizeText_(value) {
  return hotelDbV2Clean_(value)
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[\s　・･,，.．'’"“”\-ー―‐_/\\()（）\[\]【】]/g, '');
}

function hotelDbV2NormalizeAddress_(value) {
  return hotelDbV2Clean_(value)
    .normalize('NFKC')
    .replace(/^日本[、,\s]*/u, '')
    .replace(/^Japan[、,\s]*/i, '')
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

function hotelDbV2NormalizePostalCode_(value) {
  const digits = hotelDbV2Clean_(value).replace(/\D/g, '');
  return digits.length === 7
    ? digits.slice(0, 3) + '-' + digits.slice(3)
    : '';
}

function hotelDbV2SimilarityRatio_(a, b) {
  const left = hotelDbV2Clean_(a);
  const right = hotelDbV2Clean_(b);
  if (!left || !right) return 0;
  if (left === right) return 1;

  const longer = left.length >= right.length ? left : right;
  const shorter = left.length >= right.length ? right : left;
  return (longer.length - hotelDbV2Levenshtein_(longer, shorter)) / longer.length;
}

function hotelDbV2Levenshtein_(a, b) {
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

function hotelDbV2Unique_(values) {
  const seen = {};
  return (values || []).filter(function(value) {
    const text = hotelDbV2Clean_(value);
    if (!text || seen[text]) return false;
    seen[text] = true;
    return true;
  });
}

function hotelDbV2ContainsJapanese_(value) {
  return /[\u3040-\u30ff\u3400-\u9fff]/.test(hotelDbV2Clean_(value));
}

function hotelDbV2GetComponent_(place, type) {
  const components = place && place.addressComponents
    ? place.addressComponents
    : [];

  for (let i = 0; i < components.length; i++) {
    if ((components[i].types || []).indexOf(type) !== -1) {
      return hotelDbV2Clean_(components[i].longText || components[i].shortText);
    }
  }
  return '';
}

function hotelDbV2GetPostalCode_(place) {
  return hotelDbV2NormalizePostalCode_(hotelDbV2GetComponent_(place, 'postal_code'));
}

function hotelDbV2GetPrefecture_(place) {
  return hotelDbV2GetComponent_(place, 'administrative_area_level_1');
}

function hotelDbV2GetMunicipalityBase_(place) {
  const locality = hotelDbV2GetComponent_(place, 'locality');
  const admin2 = hotelDbV2GetComponent_(place, 'administrative_area_level_2');
  const ward = [
    hotelDbV2GetComponent_(place, 'ward'),
    hotelDbV2GetComponent_(place, 'sublocality_level_1')
  ].filter(function(value) {
    return /区$/.test(value);
  })[0] || '';

  if (locality && ward && locality.indexOf(ward) === -1) return locality + ward;
  if (locality) return locality;
  if (admin2 && ward && admin2.indexOf(ward) === -1) return admin2 + ward;
  return admin2 || ward || hotelDbV2GetComponent_(place, 'postal_town');
}

function hotelDbV2GetMunicipalityForSource_(place, sourceMunicipality) {
  const prefecture = hotelDbV2GetPrefecture_(place);
  const municipality = hotelDbV2GetMunicipalityBase_(place);
  const source = hotelDbV2Clean_(sourceMunicipality);

  if (source && prefecture && source.indexOf(prefecture) !== -1) {
    return hotelDbV2Unique_([prefecture, municipality]).join('');
  }
  return municipality || prefecture;
}

function hotelDbV2GetJapaneseFullAddress_(place) {
  if (!place) return '';

  const formatted = hotelDbV2Clean_(place.formattedAddress)
    .replace(/^日本[、,\s]*/u, '')
    .replace(/^Japan[、,\s]*/i, '')
    .replace(/〒\s*\d{3}-?\d{4}\s*/u, '')
    .trim();

  if (hotelDbV2ContainsJapanese_(formatted)) return formatted;

  const postalAddress = place.postalAddress || {};
  const addressLines = postalAddress.addressLines || [];
  const postalText = hotelDbV2Unique_(addressLines).join('');
  if (hotelDbV2ContainsJapanese_(postalText)) return postalText;

  const parts = [
    hotelDbV2GetComponent_(place, 'administrative_area_level_1'),
    hotelDbV2GetComponent_(place, 'administrative_area_level_2'),
    hotelDbV2GetComponent_(place, 'locality'),
    hotelDbV2GetComponent_(place, 'ward'),
    hotelDbV2GetComponent_(place, 'sublocality_level_1'),
    hotelDbV2GetComponent_(place, 'sublocality_level_2'),
    hotelDbV2GetComponent_(place, 'sublocality_level_3'),
    hotelDbV2GetComponent_(place, 'sublocality_level_4'),
    hotelDbV2GetComponent_(place, 'premise'),
    hotelDbV2GetComponent_(place, 'route'),
    hotelDbV2GetComponent_(place, 'street_number'),
    hotelDbV2GetComponent_(place, 'subpremise')
  ];

  const rebuilt = hotelDbV2Unique_(parts).join('');
  return rebuilt || formatted;
}

function hotelDbV2GetAddressForSource_(place, sourceMunicipality) {
  const fullAddress = hotelDbV2GetJapaneseFullAddress_(place);
  const prefecture = hotelDbV2GetPrefecture_(place);
  const municipality = hotelDbV2GetMunicipalityBase_(place);
  const source = hotelDbV2Clean_(sourceMunicipality);

  let result = hotelDbV2Clean_(fullAddress)
    .replace(/^日本[、,\s]*/u, '')
    .replace(/^Japan[、,\s]*/i, '')
    .replace(/〒\s*\d{3}-?\d{4}\s*/u, '');

  [prefecture, municipality, source].forEach(function(prefix) {
    const text = hotelDbV2Clean_(prefix);
    if (text && result.indexOf(text) === 0) result = result.slice(text.length);
  });

  return result.replace(/^[、,\s]+/, '').trim() || fullAddress;
}

function hotelDbV2GetDisplayName_(place) {
  return place && place.displayName ? hotelDbV2Clean_(place.displayName.text) : '';
}

function hotelDbV2TranslateBusinessStatus_(status) {
  switch (hotelDbV2Clean_(status)) {
    case 'OPERATIONAL': return '営業中';
    case 'CLOSED_TEMPORARILY': return '一時休業';
    case 'CLOSED_PERMANENTLY': return '閉業';
    case 'FUTURE_OPENING': return '開業予定';
    default: return '不明';
  }
}

function hotelDbV2BuildSearchQuery_(facility) {
  return hotelDbV2Unique_([
    facility.name,
    hotelDbV2NormalizePostalCode_(facility.postalCode),
    facility.municipality,
    facility.address,
    '宿泊施設'
  ]).join(' ');
}

function hotelDbV2CalculateMatchScore_(facility, place) {
  const sourceName = hotelDbV2NormalizeText_(facility.name);
  const googleName = hotelDbV2NormalizeText_(hotelDbV2GetDisplayName_(place));
  const sourceAddress = hotelDbV2NormalizeAddress_(
    hotelDbV2Clean_(facility.municipality) + hotelDbV2Clean_(facility.address)
  );
  const googleAddress = hotelDbV2NormalizeAddress_(hotelDbV2GetJapaneseFullAddress_(place));
  const sourcePostal = hotelDbV2NormalizePostalCode_(facility.postalCode);
  const googlePostal = hotelDbV2GetPostalCode_(place);
  const sourceMunicipality = hotelDbV2NormalizeText_(facility.municipality);
  const googleMunicipality = hotelDbV2NormalizeText_(hotelDbV2GetMunicipalityForSource_(place, facility.municipality));

  let score = 0;

  if (sourceName && googleName) {
    if (sourceName === googleName) score += 50;
    else if (sourceName.indexOf(googleName) !== -1 || googleName.indexOf(sourceName) !== -1) {
      score += 40;
    } else {
      score += Math.round(hotelDbV2SimilarityRatio_(sourceName, googleName) * 34);
    }
  }

  if (sourceAddress && googleAddress) {
    if (sourceAddress === googleAddress) score += 35;
    else if (sourceAddress.indexOf(googleAddress) !== -1 || googleAddress.indexOf(sourceAddress) !== -1) {
      score += 30;
    } else {
      score += Math.round(hotelDbV2SimilarityRatio_(sourceAddress, googleAddress) * 26);
    }
  }

  if (sourcePostal && googlePostal && sourcePostal === googlePostal) score += 10;
  if (sourceMunicipality && googleMunicipality &&
      (sourceMunicipality.indexOf(googleMunicipality) !== -1 ||
       googleMunicipality.indexOf(sourceMunicipality) !== -1)) {
    score += 5;
  }

  if (place.businessStatus === 'OPERATIONAL') score += 3;
  if (place.businessStatus === 'CLOSED_PERMANENTLY') score -= 30;

  return Math.max(0, Math.min(100, score));
}

function hotelDbV2FindBestCandidate_(facility) {
  const results = hotelDbV2SearchPlaces_(hotelDbV2BuildSearchQuery_(facility));
  if (!results.length) return null;

  const ranked = results.map(function(place) {
    return {
      place: place,
      score: hotelDbV2CalculateMatchScore_(facility, place)
    };
  }).sort(function(a, b) {
    return b.score - a.score;
  });

  const best = ranked[0];
  return {
    place: best.place,
    score: best.score,
    accepted: best.score >= HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE,
    reviewable: best.score >= HOTEL_DB_V2_CONFIG.MIN_MATCH_SCORE,
    candidates: ranked.slice(0, 5)
  };
}

function hotelDbV2Today_() {
  return Utilities.formatDate(new Date(), HOTEL_DB_V2_CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

function hotelDbV2Timestamp_() {
  return Utilities.formatDate(new Date(), HOTEL_DB_V2_CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}


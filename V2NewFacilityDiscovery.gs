/**
 * PR #16 営業中の新規宿泊施設を安全に探索する候補作成機能。
 *
 * 安全原則:
 * - 元データへ自動追加しない。
 * - 元データ・Place ID・既存状態列を変更しない。
 * - Googleで営業中と確認できた候補だけを対象にする。
 * - 既存Place ID一致、または既存の施設名+住所が一致するものは候補から除外する。
 * - 既存施設との類似が強い場合は「要人確認」に残す。
 */

const HOTEL_DB_V2_NEW_FACILITY_DISCOVERY = Object.freeze({
  SHEET_NAME: '新規追加候補',
  PAGE_SIZE: 20,
  MAX_PAGES_PER_QUERY: 3,
  MAX_EXECUTION_MS: 220000,
  SEARCH_TYPES: Object.freeze([
    Object.freeze({ type: 'hotel', label: 'ホテル' }),
    Object.freeze({ type: 'japanese_inn', label: '旅館' }),
    Object.freeze({ type: 'budget_japanese_inn', label: '簡易旅館' }),
    Object.freeze({ type: 'guest_house', label: 'ゲストハウス' }),
    Object.freeze({ type: 'hostel', label: 'ホステル' }),
    Object.freeze({ type: 'bed_and_breakfast', label: '民宿・B&B' }),
    Object.freeze({ type: 'private_guest_room', label: '民泊・個室' }),
    Object.freeze({ type: 'cottage', label: '一棟貸し・コテージ' }),
    Object.freeze({ type: 'inn', label: '宿・イン' })
  ]),
  SEARCH_FIELDS: [
    'nextPageToken',
    'places.id',
    'places.displayName',
    'places.formattedAddress',
    'places.addressComponents',
    'places.businessStatus',
    'places.location',
    'places.types',
    'places.primaryType',
    'places.googleMapsUri',
    'places.nationalPhoneNumber',
    'places.websiteUri',
    'places.rating',
    'places.userRatingCount'
  ].join(','),
  STRONG_NAME_SIMILARITY: 0.86,
  STRONG_ADDRESS_SIMILARITY: 0.88,
  POSSIBLE_NAME_SIMILARITY: 0.55,
  POSSIBLE_ADDRESS_SIMILARITY: 0.60
});

const HOTEL_DB_V2_NEW_FACILITY_HEADERS = [
  '候補キー', '状態', '探索元シート', '探索元シートID', '対象市区町村',
  '検索種別', '候補Place ID', '候補施設名', '候補住所', '候補郵便番号',
  '候補市区町村', '営業状態', 'Googleタイプ', '電話番号', '公式サイト',
  '評価', '口コミ数', '緯度', '経度', 'Google Maps URL',
  '既存照合', '推奨判定', '自動判定理由', '信頼度', '発見日'
];

function runHotelDbV2DiscoverNewFacilities() {
  return withHotelDbV2Lock_('営業中新規宿泊施設探索', function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = spreadsheet.getActiveSheet();
    const map = hotelDbV2GetHeaderMap_(sourceSheet);
    hotelDbV2ValidateSourceSheet_(sourceSheet, map);

    const activeRange = sourceSheet.getActiveRange();
    const selectedRow = activeRange ? activeRange.getRow() : 0;
    if (selectedRow < 2) {
      throw new Error('探索したい市区町村の施設データ行を1行選択してから実行してください。');
    }

    const selectedFacility = hotelDbV2ReadFacility_(sourceSheet, selectedRow, map);
    const targetMunicipality = hotelDbV2Clean_(selectedFacility.municipality);
    if (!targetMunicipality) {
      throw new Error('選択行の「市区町村」が空です。市区町村が入っている行を選択してください。');
    }

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '営業中の新規宿泊施設を探索',
      '対象シート: 「' + sourceSheet.getName() + '」\n' +
      '対象市区町村: ' + targetMunicipality + '\n\n' +
      'Google Placesで宿泊系9カテゴリを検索し、営業中の候補だけを「新規追加候補」へ出力します。\n' +
      '既存DBと一致する施設は除外します。\n\n' +
      '元データへの自動追加・自動修正・自動削除はしません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2DiscoverNewFacilities_(
      spreadsheet,
      sourceSheet,
      map,
      targetMunicipality
    );

    ui.alert([
      '営業中新規宿泊施設の探索完了',
      '',
      '対象市区町村: ' + result.targetMunicipality,
      'API検索回数: ' + result.apiRequests,
      '検索結果延べ件数: ' + result.rawResults,
      '重複除外後Place数: ' + result.uniquePlaces,
      '営業中確認: ' + result.operational,
      '既存DB一致で除外: ' + result.existingExcluded,
      '対象自治体外で除外: ' + result.outsideExcluded,
      '営業中以外で除外: ' + result.nonOperationalExcluded,
      '新規追加候補へ出力: ' + result.outputCandidates,
      '新規候補有力: ' + result.newLikely,
      '要人確認: ' + result.humanReview,
      '検索エラー: ' + result.errors,
      '',
      '元データへの自動追加・自動変更はしていません。'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2DiscoverNewFacilities_(spreadsheet, sourceSheet, sourceMap, targetMunicipality) {
  const startedAt = Date.now();
  const existing = hotelDbV2DiscoveryBuildExistingIndex_(
    spreadsheet,
    sourceSheet,
    sourceMap,
    targetMunicipality
  );

  const collected = {};
  let apiRequests = 0;
  let rawResults = 0;
  let errors = 0;
  const errorMessages = [];
  let timedOut = false;

  HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.SEARCH_TYPES.forEach(function(spec) {
    if (timedOut) return;
    if (Date.now() - startedAt > HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.MAX_EXECUTION_MS) {
      timedOut = true;
      return;
    }

    try {
      const response = hotelDbV2DiscoverySearchType_(targetMunicipality, spec, startedAt);
      apiRequests += response.apiRequests;
      rawResults += response.rawResults;
      if (response.timedOut) timedOut = true;

      response.places.forEach(function(place) {
        const id = hotelDbV2Clean_(place && place.id);
        if (!id) return;

        if (!collected[id]) {
          collected[id] = {
            place: place,
            labels: {}
          };
        }
        collected[id].labels[spec.label] = true;
      });
    } catch (error) {
      errors++;
      errorMessages.push(spec.label + ': ' + error.message);
    }
  });

  const candidateSheet = hotelDbV2GetOrCreateSheet_(
    spreadsheet,
    HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.SHEET_NAME,
    HOTEL_DB_V2_NEW_FACILITY_HEADERS
  );
  const candidateContext = {
    sheet: candidateSheet,
    headers: HOTEL_DB_V2_NEW_FACILITY_HEADERS,
    map: hotelDbV2HeaderIndex_(HOTEL_DB_V2_NEW_FACILITY_HEADERS),
    index: hotelDbV2BuildKeyIndex_(candidateSheet)
  };

  let operational = 0;
  let existingExcluded = 0;
  let outsideExcluded = 0;
  let nonOperationalExcluded = 0;
  let outputCandidates = 0;
  let newLikely = 0;
  let humanReview = 0;

  Object.keys(collected).forEach(function(placeId) {
    const item = collected[placeId];
    const decision = hotelDbV2DiscoveryEvaluateCandidate_(
      item.place,
      targetMunicipality,
      existing
    );

    if (decision.kind === 'nonOperational') {
      nonOperationalExcluded++;
      return;
    }
    if (decision.kind === 'outside') {
      outsideExcluded++;
      return;
    }
    if (decision.kind === 'existing') {
      existingExcluded++;
      return;
    }
    if (decision.kind === 'invalid') return;

    operational++;
    outputCandidates++;
    if (decision.recommendation === '新規候補有力') newLikely++;
    else humanReview++;

    const place = item.place;
    const location = place.location || {};
    const labels = Object.keys(item.labels).sort().join('・');
    const candidateMunicipality = hotelDbV2DiscoveryCandidateMunicipality_(
      place,
      targetMunicipality
    );

    hotelDbV2Upsert_(candidateContext, placeId, {
      '候補キー': placeId,
      '状態': '未確認',
      '探索元シート': sourceSheet.getName(),
      '探索元シートID': String(sourceSheet.getSheetId()),
      '対象市区町村': targetMunicipality,
      '検索種別': labels,
      '候補Place ID': placeId,
      '候補施設名': hotelDbV2GetDisplayName_(place),
      '候補住所': hotelDbV2GetJapaneseFullAddress_(place),
      '候補郵便番号': hotelDbV2GetPostalCode_(place),
      '候補市区町村': candidateMunicipality,
      '営業状態': hotelDbV2TranslateBusinessStatus_(place.businessStatus),
      'Googleタイプ': (place.types || []).join(', '),
      '電話番号': hotelDbV2Clean_(place.nationalPhoneNumber),
      '公式サイト': hotelDbV2Clean_(place.websiteUri),
      '評価': place.rating === undefined ? '' : place.rating,
      '口コミ数': place.userRatingCount === undefined ? '' : place.userRatingCount,
      '緯度': location.latitude === undefined ? '' : location.latitude,
      '経度': location.longitude === undefined ? '' : location.longitude,
      'Google Maps URL': hotelDbV2Clean_(place.googleMapsUri),
      '既存照合': decision.existingMatch || '既存一致なし',
      '推奨判定': decision.recommendation,
      '自動判定理由': decision.reason,
      '信頼度': decision.confidence,
      '発見日': hotelDbV2Today_()
    }, true);
  });

  candidateSheet.setFrozenRows(1);

  return {
    targetMunicipality: targetMunicipality,
    apiRequests: apiRequests,
    rawResults: rawResults,
    uniquePlaces: Object.keys(collected).length,
    operational: operational,
    existingExcluded: existingExcluded,
    outsideExcluded: outsideExcluded,
    nonOperationalExcluded: nonOperationalExcluded,
    outputCandidates: outputCandidates,
    newLikely: newLikely,
    humanReview: humanReview,
    errors: errors,
    errorMessages: errorMessages,
    timedOut: timedOut
  };
}

function hotelDbV2DiscoverySearchType_(targetMunicipality, spec, startedAt) {
  const basePayload = {
    textQuery: targetMunicipality + ' ' + spec.label,
    languageCode: HOTEL_DB_V2_CONFIG.LANGUAGE_CODE,
    regionCode: HOTEL_DB_V2_CONFIG.REGION_CODE,
    pageSize: HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.PAGE_SIZE,
    includedType: spec.type,
    strictTypeFiltering: true,
    includeFutureOpeningBusinesses: false
  };

  const places = [];
  let pageToken = '';
  let apiRequests = 0;
  let rawResults = 0;
  let timedOut = false;

  for (let page = 0; page < HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.MAX_PAGES_PER_QUERY; page++) {
    if (Date.now() - startedAt > HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.MAX_EXECUTION_MS) {
      timedOut = true;
      break;
    }

    const payload = {};
    Object.keys(basePayload).forEach(function(key) {
      payload[key] = basePayload[key];
    });
    if (pageToken) payload.pageToken = pageToken;

    const json = hotelDbV2CallPlacesApi_('/places:searchText', {
      method: 'post',
      fieldMask: HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.SEARCH_FIELDS,
      payload: payload
    });

    apiRequests++;
    const pagePlaces = json.places || [];
    rawResults += pagePlaces.length;
    pagePlaces.forEach(function(place) {
      places.push(place);
    });

    pageToken = hotelDbV2Clean_(json.nextPageToken);
    if (!pageToken) break;

    Utilities.sleep(HOTEL_DB_V2_CONFIG.REQUEST_INTERVAL_MS);
  }

  return {
    places: places,
    apiRequests: apiRequests,
    rawResults: rawResults,
    timedOut: timedOut
  };
}

function hotelDbV2DiscoveryBuildExistingIndex_(spreadsheet, sourceSheet, sourceMap, targetMunicipality) {
  const placeIds = {};
  const rows = [];
  const reservedNames = {};

  Object.keys(HOTEL_DB_V2_CONFIG.SHEETS).forEach(function(key) {
    reservedNames[HOTEL_DB_V2_CONFIG.SHEETS[key]] = true;
  });
  reservedNames[HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.SHEET_NAME] = true;

  spreadsheet.getSheets().forEach(function(sheet) {
    if (reservedNames[sheet.getName()]) return;
    if (sheet !== sourceSheet && /(?:^PR\d|テスト)/i.test(sheet.getName())) return;
    if (sheet.getLastRow() < 2) return;

    const map = hotelDbV2GetHeaderMap_(sheet);
    if (!map.placeId) return;

    const values = sheet
      .getRange(2, map.placeId, sheet.getLastRow() - 1, 1)
      .getDisplayValues();
    values.forEach(function(row) {
      const id = hotelDbV2Clean_(row[0]);
      if (id) placeIds[id] = true;
    });
  });

  if (sourceSheet.getLastRow() >= 2) {
    const values = sourceSheet
      .getRange(2, 1, sourceSheet.getLastRow() - 1, sourceSheet.getLastColumn())
      .getDisplayValues();

    values.forEach(function(row, offset) {
      function value(key) {
        return sourceMap[key]
          ? hotelDbV2Clean_(row[sourceMap[key] - 1])
          : '';
      }

      const municipality = value('municipality');
      const address = value('address');
      const name = value('facilityName');
      const placeId = value('placeId');
      if (!name) return;

      if (!hotelDbV2DiscoveryMunicipalityTextMatches_(targetMunicipality, municipality, address)) {
        return;
      }

      rows.push({
        rowNumber: offset + 2,
        municipality: municipality,
        address: address,
        name: name,
        placeId: placeId,
        normalizedName: hotelDbV2NormalizeText_(name),
        normalizedAddress: hotelDbV2DiscoveryNormalizeFullAddress_(municipality, address)
      });
    });
  }

  return {
    placeIds: placeIds,
    rows: rows
  };
}

function hotelDbV2DiscoveryEvaluateCandidate_(place, targetMunicipality, existing) {
  const id = hotelDbV2Clean_(place && place.id);
  if (!id) {
    return { kind: 'invalid', recommendation: '対象外', reason: 'Place IDがありません。', confidence: 0 };
  }

  const rawStatus = hotelDbV2Clean_(place.businessStatus);
  if (rawStatus !== 'OPERATIONAL') {
    return {
      kind: 'nonOperational',
      recommendation: '対象外',
      reason: 'Google営業状態が営業中ではありません。',
      confidence: 100
    };
  }

  if (!hotelDbV2DiscoveryMunicipalityMatchesPlace_(targetMunicipality, place)) {
    return {
      kind: 'outside',
      recommendation: '対象外',
      reason: '候補住所が対象市区町村と一致しません。',
      confidence: 98
    };
  }

  if (existing.placeIds[id]) {
    return {
      kind: 'existing',
      recommendation: '既存DB一致',
      reason: '既存DBに同じPlace IDがあります。',
      confidence: 100,
      existingMatch: '既存Place ID一致'
    };
  }

  const match = hotelDbV2DiscoveryFindExistingMatch_(place, existing.rows || []);
  if (match.exact) {
    return {
      kind: 'existing',
      recommendation: '既存DB一致',
      reason: match.reason,
      confidence: 98,
      existingMatch: match.label
    };
  }

  if (match.possible) {
    return {
      kind: 'candidate',
      recommendation: '要人確認',
      reason: match.reason,
      confidence: match.confidence,
      existingMatch: match.label
    };
  }

  return {
    kind: 'candidate',
    recommendation: '新規候補有力',
    reason: 'Googleで営業中を確認し、対象市区町村内で、既存Place ID・既存施設名/住所との明確な一致がありません。自動追加はせず、人が最終確認します。',
    confidence: 92,
    existingMatch: '既存一致なし'
  };
}

function hotelDbV2DiscoveryFindExistingMatch_(place, rows) {
  const candidateName = hotelDbV2NormalizeText_(hotelDbV2GetDisplayName_(place));
  const candidateAddress = hotelDbV2NormalizeAddressForComparison_(
    hotelDbV2GetJapaneseFullAddress_(place)
  );

  let bestPossible = null;

  (rows || []).forEach(function(row) {
    const sourceName = row.normalizedName || hotelDbV2NormalizeText_(row.name);
    const sourceAddress = row.normalizedAddress || hotelDbV2DiscoveryNormalizeFullAddress_(
      row.municipality,
      row.address
    );

    const nameExact = !!candidateName && candidateName === sourceName;
    const addressExact = hotelDbV2DiscoveryAddressesSame_(candidateAddress, sourceAddress);

    if (nameExact && addressExact) {
      bestPossible = {
        exact: true,
        possible: false,
        label: '既存施設名・住所一致',
        reason: '既存DBに施設名と住所が一致する行があります。',
        confidence: 98
      };
      return;
    }

    if (bestPossible && bestPossible.exact) return;

    const nameSimilarity = hotelDbV2SimilarityRatio_(candidateName, sourceName);
    const addressSimilarity = hotelDbV2SimilarityRatio_(candidateAddress, sourceAddress);

    let possible = false;
    let reason = '';

    if (addressExact && candidateName && sourceName) {
      possible = true;
      reason = '既存DBと住所が一致しますが施設名が異なります。同一施設の改称・リブランド、同一建物の別施設などを否定できないため人が確認します。';
    } else if (nameExact && candidateName.length >= 4) {
      possible = true;
      reason = '既存DBと施設名が一致しますが住所が異なります。移転・支店・元住所誤りなどを否定できないため人が確認します。';
    } else if (
      nameSimilarity >= HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.STRONG_NAME_SIMILARITY &&
      addressSimilarity >= HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.POSSIBLE_ADDRESS_SIMILARITY
    ) {
      possible = true;
      reason = '既存DBの施設名と高類似で、住所にも一定の一致があります。同一施設の可能性があるため人が確認します。';
    } else if (
      addressSimilarity >= HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.STRONG_ADDRESS_SIMILARITY &&
      nameSimilarity >= HOTEL_DB_V2_NEW_FACILITY_DISCOVERY.POSSIBLE_NAME_SIMILARITY
    ) {
      possible = true;
      reason = '既存DBの住所と高類似で、施設名にも一定の一致があります。同一施設または同一建物内施設の可能性があるため人が確認します。';
    }

    if (possible) {
      const confidence = Math.round(
        75 + Math.min(23, (nameSimilarity + addressSimilarity) * 10)
      );
      if (!bestPossible || confidence > bestPossible.confidence) {
        bestPossible = {
          exact: false,
          possible: true,
          label: '類似既存あり: ' + row.name + ' (行' + row.rowNumber + ')',
          reason: reason,
          confidence: confidence
        };
      }
    }
  });

  return bestPossible || {
    exact: false,
    possible: false,
    label: '既存一致なし',
    reason: '',
    confidence: 0
  };
}

function hotelDbV2DiscoveryAddressesSame_(left, right) {
  const a = hotelDbV2NormalizeAddressForComparison_(left);
  const b = hotelDbV2NormalizeAddressForComparison_(right);
  if (!a || !b) return false;
  if (a === b) return true;

  const shorter = a.length <= b.length ? a : b;
  const longer = a.length <= b.length ? b : a;
  return shorter.length >= 8 && longer.slice(-shorter.length) === shorter;
}

function hotelDbV2DiscoveryNormalizeFullAddress_(municipality, address) {
  const municipalityText = hotelDbV2Clean_(municipality);
  const addressText = hotelDbV2Clean_(address);
  const normalizedAddress = hotelDbV2NormalizeAddressForComparison_(addressText);
  const normalizedMunicipality = hotelDbV2NormalizeAddressForComparison_(municipalityText);

  if (!normalizedAddress) return normalizedMunicipality;
  if (!normalizedMunicipality) return normalizedAddress;
  if (normalizedAddress.indexOf(normalizedMunicipality) !== -1) return normalizedAddress;

  return hotelDbV2NormalizeAddressForComparison_(municipalityText + addressText);
}

function hotelDbV2DiscoveryCandidateMunicipality_(place, targetMunicipality) {
  return hotelDbV2GetMunicipalityForSource_(place, targetMunicipality) ||
    hotelDbV2GetMunicipalityBase_(place) ||
    hotelDbV2GetPrefecture_(place);
}

function hotelDbV2DiscoveryMunicipalityMatchesPlace_(targetMunicipality, place) {
  const candidateMunicipality = hotelDbV2DiscoveryCandidateMunicipality_(place, targetMunicipality);
  const address = hotelDbV2GetJapaneseFullAddress_(place);
  return hotelDbV2DiscoveryMunicipalityTextMatches_(
    targetMunicipality,
    candidateMunicipality,
    address
  );
}

function hotelDbV2DiscoveryMunicipalityTextMatches_(targetMunicipality, candidateMunicipality, candidateAddress) {
  const target = hotelDbV2Clean_(targetMunicipality).normalize('NFKC').replace(/\s+/g, '');
  const candidate = hotelDbV2Clean_(candidateMunicipality).normalize('NFKC').replace(/\s+/g, '');
  const address = hotelDbV2Clean_(candidateAddress).normalize('NFKC').replace(/\s+/g, '');
  if (!target) return false;

  const targetNormalized = hotelDbV2NormalizeText_(target);
  const candidateNormalized = hotelDbV2NormalizeText_(candidate);
  const addressNormalized = hotelDbV2NormalizeText_(address);

  if (
    candidateNormalized &&
    (targetNormalized.indexOf(candidateNormalized) !== -1 || candidateNormalized.indexOf(targetNormalized) !== -1)
  ) return true;
  if (addressNormalized && addressNormalized.indexOf(targetNormalized) !== -1) return true;

  const targetParts = hotelDbV2DiscoveryMunicipalityParts_(target);
  const candidateParts = hotelDbV2DiscoveryMunicipalityParts_(candidate || address);

  if (targetParts.prefecture && candidateParts.prefecture && targetParts.prefecture !== candidateParts.prefecture) {
    return false;
  }
  if (targetParts.city) {
    if (!candidateParts.city || targetParts.city !== candidateParts.city) return false;
    if (targetParts.ward && candidateParts.ward && targetParts.ward !== candidateParts.ward) return false;
    return true;
  }
  if (targetParts.ward) {
    return !!candidateParts.ward && targetParts.ward === candidateParts.ward;
  }
  if (targetParts.county) {
    if (!candidateParts.county || targetParts.county !== candidateParts.county) return false;
    if (targetParts.town && candidateParts.town && targetParts.town !== candidateParts.town) return false;
    if (targetParts.village && candidateParts.village && targetParts.village !== candidateParts.village) return false;
    return true;
  }
  if (targetParts.town) {
    return !!candidateParts.town && targetParts.town === candidateParts.town;
  }
  if (targetParts.village) {
    return !!candidateParts.village && targetParts.village === candidateParts.village;
  }

  return false;
}

function hotelDbV2DiscoveryMunicipalityParts_(value) {
  let text = hotelDbV2Clean_(value).normalize('NFKC').replace(/\s+/g, '');
  const prefectureMatch = text.match(/^(北海道|東京都|大阪府|京都府|.{2,3}県)/);
  const prefecture = prefectureMatch ? prefectureMatch[1] : '';
  if (prefecture) text = text.slice(prefecture.length);

  const cityMatch = text.match(/^(.{1,12}?市)/);
  const city = cityMatch ? cityMatch[1] : '';
  if (city) text = text.slice(city.length);

  const wardMatch = text.match(/^(.{1,12}?区)/);
  const ward = wardMatch ? wardMatch[1] : '';
  if (ward) text = text.slice(ward.length);

  const countyMatch = text.match(/^(.{1,12}?郡)/);
  const county = countyMatch ? countyMatch[1] : '';
  if (county) text = text.slice(county.length);

  const townMatch = text.match(/^(.{1,12}?町)/);
  const villageMatch = text.match(/^(.{1,12}?村)/);

  return {
    prefecture: prefecture,
    city: city,
    ward: ward,
    county: county,
    town: townMatch ? townMatch[1] : '',
    village: villageMatch ? villageMatch[1] : ''
  };
}

function hotelDbV2DiscoveryFakePlace_(options) {
  const opts = options || {};
  const municipality = opts.municipality || '大阪市北区';
  const prefecture = opts.prefecture || '大阪府';
  const address = opts.address || prefecture + municipality + '梅田1丁目1-1';
  const components = opts.addressComponents || [
    { longText: opts.postalCode || '530-0001', types: ['postal_code'] },
    { longText: prefecture, types: ['administrative_area_level_1'] },
    { longText: municipality.replace(/区$/, ''), types: ['locality'] },
    { longText: municipality.match(/([^市]+区)$/) ? municipality.match(/([^市]+区)$/)[1] : '', types: ['ward'] }
  ];

  return {
    id: opts.id === undefined ? 'TEST_NEW_1' : opts.id,
    displayName: { text: opts.name || '新規テストホテル' },
    formattedAddress: address,
    addressComponents: components,
    businessStatus: opts.businessStatus === undefined ? 'OPERATIONAL' : opts.businessStatus,
    location: { latitude: 34.7, longitude: 135.5 },
    types: opts.types || ['hotel', 'lodging'],
    primaryType: opts.primaryType || 'hotel',
    googleMapsUri: 'https://maps.google.com/?cid=test',
    nationalPhoneNumber: '06-0000-0000',
    websiteUri: 'https://example.com',
    rating: 4.2,
    userRatingCount: 100
  };
}

function runHotelDbV2NewFacilityDiscoveryTests() {
  const cases = [];

  function existingIndex(placeIds, rows) {
    return { placeIds: placeIds || {}, rows: rows || [] };
  }

  function sourceRow(name, municipality, address, placeId, rowNumber) {
    return {
      rowNumber: rowNumber || 2,
      municipality: municipality || '大阪市北区',
      address: address || '大阪府大阪市北区梅田1丁目1-1',
      name: name || '既存ホテル',
      placeId: placeId || '',
      normalizedName: hotelDbV2NormalizeText_(name || '既存ホテル'),
      normalizedAddress: hotelDbV2DiscoveryNormalizeFullAddress_(
        municipality || '大阪市北区',
        address || '大阪府大阪市北区梅田1丁目1-1'
      )
    };
  }

  cases.push({
    name: '新規の営業中施設',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({ id: 'NEW_A', name: '新規ホテルA' }),
      '大阪市北区',
      existingIndex()
    ).recommendation,
    expected: '新規候補有力'
  });

  cases.push({
    name: '既存Place IDは除外',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({ id: 'EXISTING_ID' }),
      '大阪市北区',
      existingIndex({ EXISTING_ID: true })
    ).kind,
    expected: 'existing'
  });

  cases.push({
    name: '施設名と住所が一致する既存行は除外',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({ id: 'NEW_B', name: '既存ホテル', address: '大阪府大阪市北区梅田1丁目1-1' }),
      '大阪市北区',
      existingIndex({}, [sourceRow('既存ホテル', '大阪市北区', '大阪府大阪市北区梅田1丁目1-1')])
    ).kind,
    expected: 'existing'
  });

  cases.push({
    name: '同住所で別名は人確認',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({ id: 'NEW_C', name: '新ブランドホテル', address: '大阪府大阪市北区梅田1丁目1-1' }),
      '大阪市北区',
      existingIndex({}, [sourceRow('旧ブランドホテル', '大阪市北区', '大阪府大阪市北区梅田1丁目1-1')])
    ).recommendation,
    expected: '要人確認'
  });

  cases.push({
    name: '同名で別住所は人確認',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({ id: 'NEW_D', name: '同名ホテル', address: '大阪府大阪市北区梅田2丁目2-2' }),
      '大阪市北区',
      existingIndex({}, [sourceRow('同名ホテル', '大阪市北区', '大阪府大阪市北区梅田1丁目1-1')])
    ).recommendation,
    expected: '要人確認'
  });

  cases.push({
    name: 'リブランド候補は人確認',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({ id: 'NEW_E', name: 'KOKO HOTEL 大阪梅田', address: '大阪府大阪市北区神山町8-4' }),
      '大阪市北区',
      existingIndex({}, [sourceRow('ホテルウィングインターナショナルセレクト大阪梅田', '大阪市北区', '大阪府大阪市北区神山町8-4')])
    ).recommendation,
    expected: '要人確認'
  });

  cases.push({
    name: '対象自治体外は除外',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({
        id: 'OUTSIDE', name: '神戸ホテル', prefecture: '兵庫県', municipality: '神戸市中央区',
        address: '兵庫県神戸市中央区北野町1丁目1-1',
        addressComponents: [
          { longText: '650-0002', types: ['postal_code'] },
          { longText: '兵庫県', types: ['administrative_area_level_1'] },
          { longText: '神戸市', types: ['locality'] },
          { longText: '中央区', types: ['ward'] }
        ]
      }),
      '大阪市北区', existingIndex()
    ).kind,
    expected: 'outside'
  });

  ['CLOSED_PERMANENTLY', 'CLOSED_TEMPORARILY', 'FUTURE_OPENING', ''].forEach(function(status) {
    cases.push({
      name: '営業中以外は除外: ' + (status || '状態なし'),
      actual: hotelDbV2DiscoveryEvaluateCandidate_(
        hotelDbV2DiscoveryFakePlace_({ id: 'STATUS_' + status, businessStatus: status }),
        '大阪市北区', existingIndex()
      ).kind,
      expected: 'nonOperational'
    });
  });

  cases.push({
    name: 'Place IDなしは対象外',
    actual: hotelDbV2DiscoveryEvaluateCandidate_(
      hotelDbV2DiscoveryFakePlace_({ id: '' }),
      '大阪市北区', existingIndex()
    ).kind,
    expected: 'invalid'
  });

  cases.push({
    name: '同じ大阪市でも別区は除外',
    actual: hotelDbV2DiscoveryMunicipalityTextMatches_(
      '大阪市北区', '大阪市中央区', '大阪府大阪市中央区心斎橋筋1丁目'
    ),
    expected: false
  });

  cases.push({
    name: '東京23区の同じ区は一致',
    actual: hotelDbV2DiscoveryMunicipalityTextMatches_(
      '東京都新宿区', '新宿区', '東京都新宿区西新宿1丁目'
    ),
    expected: true
  });

  cases.push({
    name: '郡単位の対象は同じ郡を許可',
    actual: hotelDbV2DiscoveryMunicipalityTextMatches_(
      '鳥取県岩美郡', '岩美郡岩美町', '鳥取県岩美郡岩美町岩井536'
    ),
    expected: true
  });

  cases.push({
    name: '町が違えば除外',
    actual: hotelDbV2DiscoveryMunicipalityTextMatches_(
      '岩美町', '八頭町', '鳥取県八頭郡八頭町郡家648'
    ),
    expected: false
  });

  cases.push({
    name: '部屋番号違いは既存完全一致にしない',
    actual: hotelDbV2DiscoveryFindExistingMatch_(
      hotelDbV2DiscoveryFakePlace_({
        id: 'ROOM_NEW', name: 'Japan Hinata 2',
        address: '愛知県名古屋市千種区春岡1丁目4-21 Fuchsia301',
        addressComponents: [
          { longText: '464-0848', types: ['postal_code'] },
          { longText: '愛知県', types: ['administrative_area_level_1'] },
          { longText: '名古屋市', types: ['locality'] },
          { longText: '千種区', types: ['ward'] }
        ]
      }),
      [sourceRow('Japan Hinata 1', '名古屋市千種区', '愛知県名古屋市千種区春岡1丁目4-21 Fuchsia901')]
    ).exact,
    expected: false
  });

  cases.push({
    name: '同一建物別室は人確認に寄せる',
    actual: hotelDbV2DiscoveryFindExistingMatch_(
      hotelDbV2DiscoveryFakePlace_({
        id: 'ROOM_NEW_2', name: 'Japan Hinata 2',
        address: '愛知県名古屋市千種区春岡1丁目4-21 Fuchsia301',
        addressComponents: [
          { longText: '464-0848', types: ['postal_code'] },
          { longText: '愛知県', types: ['administrative_area_level_1'] },
          { longText: '名古屋市', types: ['locality'] },
          { longText: '千種区', types: ['ward'] }
        ]
      }),
      [sourceRow('Japan Hinata 1', '名古屋市千種区', '愛知県名古屋市千種区春岡1丁目4-21 Fuchsia901')]
    ).possible,
    expected: true
  });

  const failures = cases.filter(function(test) {
    return test.actual !== test.expected;
  });

  if (failures.length) {
    throw new Error(
      '新規宿泊施設探索安全判定テスト失敗\n\n' +
      failures.map(function(test, index) {
        return '例' + (index + 1) + '「' + test.name + '」: 実際=' + test.actual + ' / 期待=' + test.expected;
      }).join('\n')
    );
  }

  SpreadsheetApp.getUi().alert([
    '新規宿泊施設探索 安全判定テスト 成功',
    '',
    '成功件数: ' + cases.length + '件',
    '失敗件数: 0件',
    '既存Place IDの重複追加: なし',
    '営業中以外の新規候補化: なし',
    '自治体外施設の新規候補化: なし',
    '元データへの自動追加・自動変更: なし'
  ].join('\n'));

  return {
    success: cases.length,
    failure: 0,
    sourceAutoChange: false,
    autoAdd: false
  };
}

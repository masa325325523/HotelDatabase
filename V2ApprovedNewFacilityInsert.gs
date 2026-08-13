/**
 * PR #17 承認済み「新規追加候補」を元DBへ安全に追加する。
 *
 * 安全原則:
 * - 「新規追加候補」の状態が「承認」の行だけが対象。
 * - 追加直前に Place Details を再取得して営業中・自治体・名称・住所・郵便番号を再確認する。
 * - 追加直前に既存DBとの Place ID / 名称 / 住所の重複・類似を再確認する。
 * - 少しでも競合があれば「要再確認」として追加しない。
 * - 宿泊分類は Google だけでは確定しないため自動入力しない。
 * - 成功時だけ新しい1行を追加し、既存行は変更しない。
 */

const HOTEL_DB_V2_APPROVED_NEW_FACILITY = Object.freeze({
  SHEET_NAME: '新規追加候補',
  REQUIRED_HEADERS: Object.freeze([
    '状態', '探索元シート', '探索元シートID', '対象市区町村',
    '候補Place ID', '候補施設名', '候補住所', '候補郵便番号',
    '候補市区町村', '営業状態', '検索種別'
  ]),
  AUDIT_HEADERS: Object.freeze([
    '追加処理日時', '追加先シート', '追加先行', '追加結果'
  ]),
  APPROVED_STATE: '承認',
  APPLIED_STATE: '追加済み',
  CONFLICT_STATE: '要再確認',
  ERROR_STATE: '追加エラー'
});

function runHotelDbV2ApplyApprovedNewFacilities() {
  return withHotelDbV2Lock_('承認済み新規追加候補の安全追加', function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const candidateSheet = spreadsheet.getSheetByName(
      HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME
    );
    const ui = SpreadsheetApp.getUi();

    if (!candidateSheet || candidateSheet.getLastRow() < 2) {
      ui.alert('「新規追加候補」に処理対象がありません。');
      return { approved: 0, applied: 0, conflicts: 0, errors: 0 };
    }

    const preview = hotelDbV2ApprovedNewFacilityPreview_(candidateSheet);
    if (!preview.approved) {
      ui.alert('「新規追加候補」に状態が「承認」の行はありません。');
      return { approved: 0, applied: 0, conflicts: 0, errors: 0 };
    }

    const response = ui.alert(
      '承認済み新規追加候補を安全に追加',
      '承認対象: ' + preview.approved + '件\n\n' +
      '追加直前にGoogle Placesを再確認し、営業中・自治体・名称・住所・郵便番号・既存DB重複を再判定します。\n' +
      '競合や変更があれば追加せず「要再確認」にします。\n' +
      '宿泊分類は自動入力しません。\n\n' +
      '既存行の自動修正・自動削除はしません。続行しますか？',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2ApplyApprovedNewFacilitiesWithContext_({
      spreadsheet: spreadsheet,
      candidateSheet: candidateSheet
    });

    ui.alert([
      '承認済み新規追加候補の処理完了',
      '',
      '承認対象: ' + result.approved,
      '追加済み: ' + result.applied,
      '要再確認・未追加: ' + result.conflicts,
      'エラー・未追加: ' + result.errors,
      '',
      '宿泊分類の自動入力: なし',
      '既存行の自動修正・自動削除: なし'
    ].join('\n'));

    return result;
  });
}

function hotelDbV2ApprovedNewFacilityPreview_(candidateSheet) {
  const map = hotelDbV2ApprovedNewFacilityHeaderMap_(candidateSheet);
  if (!map['状態'] || candidateSheet.getLastRow() < 2) return { approved: 0 };

  const states = candidateSheet
    .getRange(2, map['状態'], candidateSheet.getLastRow() - 1, 1)
    .getDisplayValues();
  let approved = 0;
  states.forEach(function(row) {
    if (hotelDbV2Clean_(row[0]) === HOTEL_DB_V2_APPROVED_NEW_FACILITY.APPROVED_STATE) {
      approved++;
    }
  });
  return { approved: approved };
}

function hotelDbV2ApplyApprovedNewFacilitiesWithContext_(options) {
  const opts = options || {};
  const spreadsheet = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const candidateSheet = opts.candidateSheet || spreadsheet.getSheetByName(
    HOTEL_DB_V2_APPROVED_NEW_FACILITY.SHEET_NAME
  );
  const detailsProvider = opts.placeDetailsProvider || function(placeId) {
    return hotelDbV2GetPlaceDetails_(placeId);
  };

  if (!candidateSheet || candidateSheet.getLastRow() < 2) {
    return { approved: 0, applied: 0, conflicts: 0, errors: 0 };
  }

  const map = hotelDbV2ApprovedNewFacilityEnsureHeaders_(candidateSheet);
  const historySheet = opts.historySheet || hotelDbV2GetOrCreateSheet_(
    spreadsheet,
    HOTEL_DB_V2_CONFIG.SHEETS.HISTORY,
    HOTEL_DB_V2_HISTORY_HEADERS
  );
  const values = candidateSheet
    .getRange(2, 1, candidateSheet.getLastRow() - 1, candidateSheet.getLastColumn())
    .getDisplayValues();
  const result = { approved: 0, applied: 0, conflicts: 0, errors: 0 };

  values.forEach(function(row, offset) {
    const candidateRow = offset + 2;
    const candidate = hotelDbV2ApprovedNewFacilityReadCandidate_(row, map);
    if (candidate.state !== HOTEL_DB_V2_APPROVED_NEW_FACILITY.APPROVED_STATE) return;

    result.approved++;
    let sourceSheet = null;
    let insertedRow = 0;

    try {
      const precheck = hotelDbV2ApprovedNewFacilityPrecheck_(candidate);
      if (!precheck.ok) {
        hotelDbV2ApprovedNewFacilityConflict_(
          candidateSheet, map, candidateRow, candidate, null, historySheet,
          precheck.reason
        );
        result.conflicts++;
        return;
      }

      sourceSheet = spreadsheet.getSheetById(Number(candidate.sourceSheetId));
      if (!sourceSheet) {
        hotelDbV2ApprovedNewFacilityConflict_(
          candidateSheet, map, candidateRow, candidate, null, historySheet,
          '探索元シートが見つかりません。'
        );
        result.conflicts++;
        return;
      }

      if (sourceSheet.getName() !== candidate.sourceSheetName) {
        hotelDbV2ApprovedNewFacilityConflict_(
          candidateSheet, map, candidateRow, candidate, sourceSheet, historySheet,
          '探索元シート名が候補作成時から変更されています。'
        );
        result.conflicts++;
        return;
      }

      let sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
      hotelDbV2ValidateSourceSheet_(sourceSheet, sourceMap);
      sourceMap = hotelDbV2EnsureOutputHeaders_(sourceSheet);

      const livePlace = detailsProvider(candidate.placeId);
      const liveCheck = hotelDbV2ApprovedNewFacilityValidateLive_(candidate, livePlace);
      if (!liveCheck.ok) {
        hotelDbV2ApprovedNewFacilityConflict_(
          candidateSheet, map, candidateRow, candidate, sourceSheet, historySheet,
          liveCheck.reason
        );
        result.conflicts++;
        return;
      }

      const existing = hotelDbV2DiscoveryBuildExistingIndex_(
        spreadsheet,
        sourceSheet,
        sourceMap,
        candidate.targetMunicipality
      );
      const decision = hotelDbV2DiscoveryEvaluateCandidate_(
        livePlace,
        candidate.targetMunicipality,
        existing
      );

      if (
        decision.kind !== 'candidate' ||
        decision.recommendation !== '新規候補有力'
      ) {
        hotelDbV2ApprovedNewFacilityConflict_(
          candidateSheet, map, candidateRow, candidate, sourceSheet, historySheet,
          '追加直前の既存DB再照合で安全に新規と断定できません。' +
          (decision.reason ? ' ' + decision.reason : '')
        );
        result.conflicts++;
        return;
      }

      const insertData = hotelDbV2ApprovedNewFacilityBuildInsertData_(
        candidate,
        livePlace
      );
      const insertCheck = hotelDbV2ApprovedNewFacilityValidateInsertData_(insertData);
      if (!insertCheck.ok) {
        hotelDbV2ApprovedNewFacilityConflict_(
          candidateSheet, map, candidateRow, candidate, sourceSheet, historySheet,
          insertCheck.reason
        );
        result.conflicts++;
        return;
      }

      insertedRow = hotelDbV2ApprovedNewFacilityInsertSourceRow_(
        sourceSheet,
        sourceMap,
        insertData
      );

      historySheet.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, {
        '日時': hotelDbV2Timestamp_(),
        '元シート': sourceSheet.getName(),
        '元シートID': sourceSheet.getSheetId(),
        '元行': insertedRow,
        '施設名': insertData.name,
        '処理': '承認新規追加',
        '結果': '追加済み',
        'Place ID': insertData.placeId,
        '一致スコア': '',
        '営業状態': '営業中',
        '詳細': '新規追加候補から承認追加。宿泊分類は未確定のため自動入力していません。'
      }));

      hotelDbV2ApprovedNewFacilitySetResult_(
        candidateSheet,
        map,
        candidateRow,
        HOTEL_DB_V2_APPROVED_NEW_FACILITY.APPLIED_STATE,
        sourceSheet.getName(),
        insertedRow,
        '追加済み。宿泊分類は要確認。'
      );
      result.applied++;
    } catch (error) {
      if (sourceSheet && insertedRow) {
        try {
          sourceSheet.deleteRow(insertedRow);
        } catch (rollbackError) {
          console.error('PR17追加行ロールバック失敗: ' + rollbackError.message);
        }
      }

      result.errors++;
      hotelDbV2ApprovedNewFacilitySetResult_(
        candidateSheet,
        map,
        candidateRow,
        HOTEL_DB_V2_APPROVED_NEW_FACILITY.ERROR_STATE,
        sourceSheet ? sourceSheet.getName() : candidate.sourceSheetName,
        '',
        '追加エラー: ' + error.message
      );
    }
  });

  return result;
}

function hotelDbV2ApprovedNewFacilityHeaderMap_(sheet) {
  const map = {};
  if (!sheet || sheet.getLastColumn() < 1) return map;
  const headers = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];
  headers.forEach(function(header, index) {
    const text = hotelDbV2Clean_(header);
    if (text && !map[text]) map[text] = index + 1;
  });
  return map;
}

function hotelDbV2ApprovedNewFacilityEnsureHeaders_(sheet) {
  let map = hotelDbV2ApprovedNewFacilityHeaderMap_(sheet);
  const missingRequired = HOTEL_DB_V2_APPROVED_NEW_FACILITY.REQUIRED_HEADERS.filter(
    function(header) { return !map[header]; }
  );
  if (missingRequired.length) {
    throw new Error(
      '「新規追加候補」に必要な列がありません: ' + missingRequired.join('、')
    );
  }

  const missingAudit = HOTEL_DB_V2_APPROVED_NEW_FACILITY.AUDIT_HEADERS.filter(
    function(header) { return !map[header]; }
  );
  if (missingAudit.length) {
    const startColumn = sheet.getLastColumn() + 1;
    sheet.getRange(1, startColumn, 1, missingAudit.length).setValues([missingAudit]);
    map = hotelDbV2ApprovedNewFacilityHeaderMap_(sheet);
  }
  return map;
}

function hotelDbV2ApprovedNewFacilityReadCandidate_(row, map) {
  function value(header) {
    return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : '';
  }

  return {
    state: value('状態'),
    sourceSheetName: value('探索元シート'),
    sourceSheetId: value('探索元シートID'),
    targetMunicipality: value('対象市区町村'),
    searchTypes: value('検索種別'),
    placeId: value('候補Place ID'),
    name: value('候補施設名'),
    address: value('候補住所'),
    postalCode: value('候補郵便番号'),
    candidateMunicipality: value('候補市区町村'),
    businessStatus: value('営業状態'),
    recommendation: value('推奨判定')
  };
}

function hotelDbV2ApprovedNewFacilityPrecheck_(candidate) {
  if (!candidate || candidate.state !== HOTEL_DB_V2_APPROVED_NEW_FACILITY.APPROVED_STATE) {
    return { ok: false, reason: '状態が「承認」ではありません。' };
  }
  if (!candidate.sourceSheetName || !Number(candidate.sourceSheetId)) {
    return { ok: false, reason: '探索元シート情報が不足しています。' };
  }
  if (!candidate.targetMunicipality) {
    return { ok: false, reason: '対象市区町村がありません。' };
  }
  if (!candidate.placeId) {
    return { ok: false, reason: '候補Place IDがありません。' };
  }
  if (!candidate.name || !candidate.address) {
    return { ok: false, reason: '候補施設名または候補住所がありません。' };
  }
  if (!hotelDbV2NormalizePostalCode_(candidate.postalCode)) {
    return { ok: false, reason: '候補郵便番号を7桁で確認できません。' };
  }
  if (candidate.businessStatus !== '営業中') {
    return { ok: false, reason: '候補作成時の営業状態が「営業中」ではありません。' };
  }
  if (
    candidate.candidateMunicipality &&
    !hotelDbV2DiscoveryMunicipalityTextMatches_(
      candidate.targetMunicipality,
      candidate.candidateMunicipality,
      candidate.address
    )
  ) {
    return { ok: false, reason: '候補作成時点で対象市区町村と候補住所が一致しません。' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2ApprovedNewFacilityValidateLive_(candidate, livePlace) {
  if (!livePlace || !hotelDbV2Clean_(livePlace.id)) {
    return { ok: false, reason: 'Google Placesで候補を再取得できません。' };
  }
  if (hotelDbV2Clean_(livePlace.id) !== candidate.placeId) {
    return { ok: false, reason: '再取得したPlace IDが候補と一致しません。' };
  }
  if (hotelDbV2Clean_(livePlace.businessStatus) !== 'OPERATIONAL') {
    return { ok: false, reason: '追加直前のGoogle営業状態が営業中ではありません。' };
  }
  if (!hotelDbV2DiscoveryMunicipalityMatchesPlace_(candidate.targetMunicipality, livePlace)) {
    return { ok: false, reason: '追加直前のGoogle住所が対象市区町村と一致しません。' };
  }

  const liveName = hotelDbV2GetDisplayName_(livePlace);
  const liveAddress = hotelDbV2GetJapaneseFullAddress_(livePlace);
  const livePostal = hotelDbV2GetPostalCode_(livePlace);
  if (!liveName || !liveAddress || !livePostal) {
    return { ok: false, reason: '追加直前の名称・住所・郵便番号のいずれかを確認できません。' };
  }

  if (hotelDbV2NormalizeText_(candidate.name) !== hotelDbV2NormalizeText_(liveName)) {
    return { ok: false, reason: '候補作成後にGoogle施設名が変更されています。' };
  }
  if (!hotelDbV2DiscoveryAddressesSame_(candidate.address, liveAddress)) {
    return { ok: false, reason: '候補作成後にGoogle住所が変更されています。' };
  }
  if (
    hotelDbV2NormalizePostalCode_(candidate.postalCode) !==
    hotelDbV2NormalizePostalCode_(livePostal)
  ) {
    return { ok: false, reason: '候補作成後にGoogle郵便番号が変更されています。' };
  }

  return { ok: true, reason: '' };
}

function hotelDbV2ApprovedNewFacilityBuildInsertData_(candidate, livePlace) {
  const municipality = hotelDbV2GetMunicipalityForSource_(
    livePlace,
    candidate.targetMunicipality
  ) || candidate.candidateMunicipality || candidate.targetMunicipality;
  const name = hotelDbV2GetDisplayName_(livePlace);
  const address = hotelDbV2GetAddressForSource_(livePlace, {
    municipality: municipality,
    address: '',
    name: name
  });
  const location = livePlace.location || {};

  return {
    postalCode: hotelDbV2GetPostalCode_(livePlace),
    municipality: municipality,
    address: address,
    name: name,
    category: '',
    notes: 'PR17承認新規追加／宿泊分類要確認' +
      (candidate.searchTypes ? '／Google探索種別:' + candidate.searchTypes : ''),
    placeId: hotelDbV2Clean_(livePlace.id),
    googleName: name,
    googleAddress: hotelDbV2GetJapaneseFullAddress_(livePlace),
    phone: hotelDbV2Clean_(livePlace.nationalPhoneNumber || livePlace.internationalPhoneNumber),
    website: hotelDbV2Clean_(livePlace.websiteUri),
    rating: livePlace.rating === undefined ? '' : livePlace.rating,
    reviewCount: livePlace.userRatingCount === undefined ? '' : livePlace.userRatingCount,
    businessStatus: hotelDbV2TranslateBusinessStatus_(livePlace.businessStatus),
    mapsUrl: hotelDbV2Clean_(livePlace.googleMapsUri),
    latitude: location.latitude === undefined ? '' : location.latitude,
    longitude: location.longitude === undefined ? '' : location.longitude,
    matchScore: '',
    checkedAt: hotelDbV2Today_(),
    matchDecision: '新規追加（承認）'
  };
}

function hotelDbV2ApprovedNewFacilityValidateInsertData_(data) {
  if (!data.placeId) return { ok: false, reason: '追加用Place IDがありません。' };
  if (!data.name) return { ok: false, reason: '追加用施設名がありません。' };
  if (!hotelDbV2NormalizePostalCode_(data.postalCode)) {
    return { ok: false, reason: '追加用郵便番号を確認できません。' };
  }
  if (!data.municipality) return { ok: false, reason: '追加用市区町村がありません。' };
  if (!data.address) return { ok: false, reason: '追加用住所がありません。' };
  if (data.businessStatus !== '営業中') {
    return { ok: false, reason: '追加用営業状態が営業中ではありません。' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2ApprovedNewFacilityInsertSourceRow_(sourceSheet, sourceMap, data) {
  const lastDataRow = hotelDbV2ApprovedNewFacilityLastDataRow_(sourceSheet, sourceMap);
  sourceSheet.insertRowAfter(lastDataRow);
  const newRow = lastDataRow + 1;
  const lastColumn = sourceSheet.getLastColumn();

  if (lastDataRow >= 2 && lastColumn > 0) {
    try {
      sourceSheet
        .getRange(lastDataRow, 1, 1, lastColumn)
        .copyTo(
          sourceSheet.getRange(newRow, 1, 1, lastColumn),
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        );
    } catch (formatError) {
      console.warn('PR17書式コピーを省略: ' + formatError.message);
    }
  }

  function setValue(key, value) {
    if (sourceMap[key]) sourceSheet.getRange(newRow, sourceMap[key]).setValue(value);
  }

  setValue('postalCode', data.postalCode);
  setValue('municipality', data.municipality);
  setValue('address', data.address);
  setValue('facilityName', data.name);
  setValue('category', '');
  setValue('notes', data.notes);

  hotelDbV2WriteGoogleData_(sourceSheet, newRow, sourceMap, data, true);
  return newRow;
}

function hotelDbV2ApprovedNewFacilityLastDataRow_(sheet, map) {
  if (sheet.getLastRow() < 2) return 1;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  let lastDataRow = 1;

  values.forEach(function(row, offset) {
    function value(key) {
      return map[key] ? hotelDbV2Clean_(row[map[key] - 1]) : '';
    }
    if (
      value('facilityName') || value('address') || value('postalCode') ||
      value('placeId')
    ) {
      lastDataRow = offset + 2;
    }
  });
  return lastDataRow;
}

function hotelDbV2ApprovedNewFacilitySetResult_(
  candidateSheet,
  map,
  rowNumber,
  state,
  sourceSheetName,
  sourceRow,
  resultText
) {
  candidateSheet.getRange(rowNumber, map['状態']).setValue(state);
  candidateSheet.getRange(rowNumber, map['追加処理日時']).setValue(hotelDbV2Timestamp_());
  candidateSheet.getRange(rowNumber, map['追加先シート']).setValue(sourceSheetName || '');
  candidateSheet.getRange(rowNumber, map['追加先行']).setValue(sourceRow || '');
  candidateSheet.getRange(rowNumber, map['追加結果']).setValue(resultText || '');
}

function hotelDbV2ApprovedNewFacilityConflict_(
  candidateSheet,
  map,
  candidateRow,
  candidate,
  sourceSheet,
  historySheet,
  reason
) {
  hotelDbV2ApprovedNewFacilitySetResult_(
    candidateSheet,
    map,
    candidateRow,
    HOTEL_DB_V2_APPROVED_NEW_FACILITY.CONFLICT_STATE,
    sourceSheet ? sourceSheet.getName() : candidate.sourceSheetName,
    '',
    '未追加: ' + reason
  );

  historySheet.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, {
    '日時': hotelDbV2Timestamp_(),
    '元シート': sourceSheet ? sourceSheet.getName() : candidate.sourceSheetName,
    '元シートID': candidate.sourceSheetId,
    '元行': '',
    '施設名': candidate.name,
    '処理': '承認新規追加',
    '結果': '要再確認・未追加',
    'Place ID': candidate.placeId,
    '一致スコア': '',
    '営業状態': candidate.businessStatus,
    '詳細': reason
  }));
}

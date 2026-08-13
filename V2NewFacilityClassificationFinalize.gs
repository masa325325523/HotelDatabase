/**
 * PR #18 追加済み新規施設の宿泊分類・備考を安全に確定する補助機能。
 *
 * ⑮: PR17で追加済み・宿泊分類未確定の施設を「新規施設分類候補」へ整理する。
 * ⑯: 人が「確定宿泊分類」を入力し、状態を「承認」にした候補だけ元DBへ反映する。
 *
 * 安全原則:
 * - Googleタイプは参考情報としてのみ使い、宿泊分類を自動確定しない。
 * - ⑮では元DBを書き換えない。
 * - ⑯では宿泊分類と備考以外を変更しない。
 * - Place ID・施設名・住所・現在分類・現在備考が候補作成後に変わっていたら反映しない。
 * - 既に宿泊分類が入っている行は上書きしない。
 * - 反映成功時だけ修正履歴へ記録する。
 */

const HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION = Object.freeze({
  SHEET_NAME: '新規施設分類候補',
  SOURCE_CANDIDATE_SHEET: '新規追加候補',
  NOTE_MARKER: '宿泊分類要確認',
  APPROVED_STATE: '承認',
  APPLIED_STATE: '反映済み',
  CONFLICT_STATE: '要再確認',
  ERROR_STATE: '反映エラー'
});

const HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION_HEADERS = [
  '候補キー', '状態', '元シート', '元シートID', '元行',
  'Place ID', '施設名', '住所',
  '現在宿泊分類', '現在備考', 'Googleタイプ',
  '参考分類', '参考理由', '参考信頼度',
  '確定宿泊分類', '確定備考',
  '確認日', '反映日時', '反映結果'
];

function runHotelDbV2BuildNewFacilityClassificationCandidates() {
  return withHotelDbV2Lock_('新規施設の分類候補作成', function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const sourceCandidateSheet = spreadsheet.getSheetByName(
      HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SOURCE_CANDIDATE_SHEET
    );

    if (!sourceCandidateSheet || sourceCandidateSheet.getLastRow() < 2) {
      ui.alert('「新規追加候補」に追加済み施設がありません。');
      return { scanned: 0, candidates: 0, review: 0, skipped: 0 };
    }

    const response = ui.alert(
      '追加済み新規施設の分類候補を作成',
      'PR17で追加済みになった施設のうち、宿泊分類が未確定のものを「新規施設分類候補」へ整理します。\n\n' +
      'Googleタイプから表示する「参考分類」は参考情報であり、自動確定しません。\n' +
      '元データの宿泊分類・備考は変更しません。続行しますか？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2BuildNewFacilityClassificationCandidates_(
      spreadsheet,
      sourceCandidateSheet
    );

    ui.alert([
      '新規施設の分類候補作成 完了',
      '',
      '追加済み確認: ' + result.scanned + '件',
      '分類候補へ出力: ' + result.candidates + '件',
      '要再確認候補: ' + result.review + '件',
      '対象外・確定済み: ' + result.skipped + '件',
      '',
      '元データの自動変更: なし',
      '宿泊分類の自動確定: なし'
    ].join('\n'));
    return result;
  });
}

function runHotelDbV2ApplyApprovedNewFacilityClassifications() {
  return withHotelDbV2Lock_('承認済み新規施設分類の安全反映', function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(
      HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME
    );
    const ui = SpreadsheetApp.getUi();

    if (!sheet || sheet.getLastRow() < 2) {
      ui.alert('「新規施設分類候補」に処理対象がありません。');
      return { approved: 0, applied: 0, conflicts: 0, errors: 0 };
    }

    const preview = hotelDbV2NewFacilityClassificationPreview_(sheet);
    if (!preview.approved) {
      ui.alert('「新規施設分類候補」に状態が「承認」の行はありません。');
      return { approved: 0, applied: 0, conflicts: 0, errors: 0 };
    }

    const response = ui.alert(
      '承認済みの宿泊分類・備考を安全に反映',
      '承認対象: ' + preview.approved + '件\n\n' +
      '「確定宿泊分類」を人が入力し、状態を「承認」にした行だけを反映します。\n' +
      'Place ID・施設名・住所・現在分類・現在備考が候補作成時から変わっていた場合は反映しません。\n' +
      '既に宿泊分類が入っている行は上書きしません。\n\n' +
      '変更するのは宿泊分類と備考だけです。続行しますか？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2ApplyApprovedNewFacilityClassifications_(
      spreadsheet,
      sheet
    );

    ui.alert([
      '承認済み宿泊分類・備考の反映完了',
      '',
      '承認対象: ' + result.approved + '件',
      '反映済み: ' + result.applied + '件',
      '要再確認・未反映: ' + result.conflicts + '件',
      'エラー・未反映: ' + result.errors + '件',
      '',
      '施設名・住所・Place IDの変更: なし',
      '既存分類の上書き: なし'
    ].join('\n'));
    return result;
  });
}

function hotelDbV2BuildNewFacilityClassificationCandidates_(spreadsheet, sourceCandidateSheet) {
  const sourceMap = hotelDbV2NewFacilityClassificationHeaderMap_(sourceCandidateSheet);
  const required = [
    '状態', '探索元シート', '探索元シートID', '候補Place ID',
    '候補施設名', '候補住所', '追加先シート', '追加先行'
  ];
  const missing = required.filter(function(header) { return !sourceMap[header]; });
  if (missing.length) {
    throw new Error('「新規追加候補」に必要な列がありません: ' + missing.join('、'));
  }

  const outputSheet = hotelDbV2GetOrCreateSheet_(
    spreadsheet,
    HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME,
    HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION_HEADERS
  );
  const outputMap = hotelDbV2HeaderIndex_(HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION_HEADERS);
  const outputIndex = hotelDbV2BuildKeyIndex_(outputSheet);
  const values = sourceCandidateSheet.getRange(
    2, 1, sourceCandidateSheet.getLastRow() - 1, sourceCandidateSheet.getLastColumn()
  ).getDisplayValues();

  const result = { scanned: 0, candidates: 0, review: 0, skipped: 0 };

  values.forEach(function(row) {
    function value(header) {
      return sourceMap[header] ? hotelDbV2Clean_(row[sourceMap[header] - 1]) : '';
    }

    if (value('状態') !== HOTEL_DB_V2_APPROVED_NEW_FACILITY.APPLIED_STATE) return;
    result.scanned++;

    const sourceSheetId = Number(value('探索元シートID'));
    const sourceRow = Number(value('追加先行'));
    const sourceSheet = sourceSheetId ? spreadsheet.getSheetById(sourceSheetId) : null;
    if (!sourceSheet || sourceRow < 2) {
      hotelDbV2NewFacilityClassificationUpsert_(outputSheet, outputMap, outputIndex, {
        key: [sourceSheetId || '0', sourceRow || '0', value('候補Place ID')].join('|'),
        state: '要再確認',
        sheetName: value('追加先シート') || value('探索元シート'),
        sheetId: sourceSheetId || '',
        sourceRow: sourceRow || '',
        placeId: value('候補Place ID'),
        name: value('候補施設名'),
        address: value('候補住所'),
        category: '', notes: '', googleTypes: value('Googleタイプ'),
        reference: '', reason: '追加先シートまたは追加先行を確認できません。', confidence: 0
      });
      result.review++;
      return;
    }

    const sourceSheetName = value('追加先シート') || value('探索元シート');
    const map = hotelDbV2GetHeaderMap_(sourceSheet);
    hotelDbV2ValidateSourceSheet_(sourceSheet, map);
    const facility = hotelDbV2ReadFacility_(sourceSheet, sourceRow, map);
    const sameIdentity = hotelDbV2NewFacilityClassificationIdentityMatches_({
      placeId: value('候補Place ID'),
      name: value('候補施設名'),
      address: value('候補住所')
    }, facility);

    if (sourceSheet.getName() !== sourceSheetName || !sameIdentity.ok) {
      hotelDbV2NewFacilityClassificationUpsert_(outputSheet, outputMap, outputIndex, {
        key: [sourceSheet.getSheetId(), sourceRow, value('候補Place ID')].join('|'),
        state: '要再確認',
        sheetName: sourceSheet.getName(), sheetId: sourceSheet.getSheetId(), sourceRow: sourceRow,
        placeId: facility.placeId || value('候補Place ID'), name: facility.name, address: facility.address,
        category: facility.category, notes: facility.notes, googleTypes: value('Googleタイプ'),
        reference: '', reason: sameIdentity.ok ? '追加先シート名が候補作成時から変更されています。' : sameIdentity.reason,
        confidence: 0
      });
      result.review++;
      return;
    }

    if (hotelDbV2Clean_(facility.category)) {
      result.skipped++;
      return;
    }
    if (facility.notes.indexOf(HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.NOTE_MARKER) === -1) {
      result.skipped++;
      return;
    }

    const recommendation = hotelDbV2NewFacilityClassificationRecommend_(value('Googleタイプ'));
    hotelDbV2NewFacilityClassificationUpsert_(outputSheet, outputMap, outputIndex, {
      key: [sourceSheet.getSheetId(), sourceRow, facility.placeId].join('|'),
      state: '未確認',
      sheetName: sourceSheet.getName(), sheetId: sourceSheet.getSheetId(), sourceRow: sourceRow,
      placeId: facility.placeId, name: facility.name, address: facility.address,
      category: facility.category, notes: facility.notes, googleTypes: value('Googleタイプ'),
      reference: recommendation.reference, reason: recommendation.reason, confidence: recommendation.confidence
    });
    result.candidates++;
  });

  return result;
}

function hotelDbV2NewFacilityClassificationHeaderMap_(sheet) {
  const map = {};
  if (!sheet || sheet.getLastColumn() < 1) return map;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  headers.forEach(function(header, index) {
    const text = hotelDbV2Clean_(header);
    if (text && !map[text]) map[text] = index + 1;
  });
  return map;
}

function hotelDbV2NewFacilityClassificationUpsert_(sheet, map, index, data) {
  const row = hotelDbV2RowFromObject_(HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION_HEADERS, {
    '候補キー': data.key,
    '状態': data.state,
    '元シート': data.sheetName,
    '元シートID': data.sheetId,
    '元行': data.sourceRow,
    'Place ID': data.placeId,
    '施設名': data.name,
    '住所': data.address,
    '現在宿泊分類': data.category,
    '現在備考': data.notes,
    'Googleタイプ': data.googleTypes,
    '参考分類': data.reference,
    '参考理由': data.reason,
    '参考信頼度': data.confidence,
    '確定宿泊分類': '',
    '確定備考': '',
    '確認日': hotelDbV2Today_(),
    '反映日時': '',
    '反映結果': ''
  });

  const existingRow = index[data.key];
  if (existingRow) {
    const old = sheet.getRange(existingRow, 1, 1, row.length).getDisplayValues()[0];
    const oldState = hotelDbV2Clean_(old[map['状態'] - 1]);
    const oldFinalCategory = hotelDbV2Clean_(old[map['確定宿泊分類'] - 1]);
    const oldFinalNotes = hotelDbV2Clean_(old[map['確定備考'] - 1]);
    if (oldState === '承認' || oldState === '反映済み') row[map['状態'] - 1] = oldState;
    row[map['確定宿泊分類'] - 1] = oldFinalCategory;
    row[map['確定備考'] - 1] = oldFinalNotes;
    if (oldState === '反映済み') {
      row[map['反映日時'] - 1] = old[map['反映日時'] - 1];
      row[map['反映結果'] - 1] = old[map['反映結果'] - 1];
    }
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    return existingRow;
  }

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);
  index[data.key] = newRow;
  return newRow;
}

function hotelDbV2NewFacilityClassificationRecommend_(googleTypes) {
  const text = hotelDbV2NormalizeText_(googleTypes).toLowerCase();
  if (!text) {
    return { reference: '要確認', reason: 'Googleタイプがありません。許可情報・公式情報で確認してください。', confidence: 40 };
  }
  if (text.indexOf('japanese_inn') !== -1) {
    return { reference: '旅館系（要確認）', reason: 'Googleタイプに japanese_inn が含まれます。法的な営業区分は別途確認が必要です。', confidence: 85 };
  }
  if (text.indexOf('hostel') !== -1 || text.indexOf('guest_house') !== -1 || text.indexOf('bed_and_breakfast') !== -1 || text.indexOf('budget_japanese_inn') !== -1) {
    return { reference: '簡易宿所系（要確認）', reason: 'ゲストハウス・ホステル・B&B等のGoogleタイプです。許可区分は別途確認が必要です。', confidence: 82 };
  }
  if (text.indexOf('private_guest_room') !== -1 || text.indexOf('cottage') !== -1) {
    return { reference: '住宅宿泊事業・簡易宿所系（要確認）', reason: '民泊・貸切系のGoogleタイプです。住宅宿泊事業か旅館業かを公的情報で確認してください。', confidence: 72 };
  }
  if (text.indexOf('hotel') !== -1) {
    return { reference: 'ホテル系（要確認）', reason: 'Googleタイプに hotel が含まれます。旅館業法上の分類は別途確認が必要です。', confidence: 80 };
  }
  if (text.indexOf('inn') !== -1 || text.indexOf('lodging') !== -1) {
    return { reference: '旅館・簡易宿所系（要確認）', reason: 'Googleタイプが広義の宿泊施設です。公的な許可情報で分類を確認してください。', confidence: 60 };
  }
  return { reference: '要確認', reason: 'Googleタイプだけでは宿泊分類を推定できません。', confidence: 40 };
}

function hotelDbV2NewFacilityClassificationIdentityMatches_(snapshot, facility) {
  if (!snapshot || !facility) return { ok: false, reason: '比較対象がありません。' };
  if (!snapshot.placeId || snapshot.placeId !== hotelDbV2Clean_(facility.placeId)) {
    return { ok: false, reason: 'Place IDが候補作成時と一致しません。' };
  }
  if (hotelDbV2NormalizeText_(snapshot.name) !== hotelDbV2NormalizeText_(facility.name)) {
    return { ok: false, reason: '施設名が候補作成時と一致しません。' };
  }
  if (!hotelDbV2DiscoveryAddressesSame_(snapshot.address, facility.address)) {
    return { ok: false, reason: '住所が候補作成時と一致しません。' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2NewFacilityClassificationPreview_(sheet) {
  const map = hotelDbV2NewFacilityClassificationHeaderMap_(sheet);
  if (!map['状態'] || sheet.getLastRow() < 2) return { approved: 0 };
  const states = sheet.getRange(2, map['状態'], sheet.getLastRow() - 1, 1).getDisplayValues();
  let approved = 0;
  states.forEach(function(row) {
    if (hotelDbV2Clean_(row[0]) === HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.APPROVED_STATE) approved++;
  });
  return { approved: approved };
}

function hotelDbV2ApplyApprovedNewFacilityClassifications_(spreadsheet, sheet) {
  const map = hotelDbV2NewFacilityClassificationHeaderMap_(sheet);
  const required = HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION_HEADERS.filter(function(header) {
    return !map[header];
  });
  if (required.length) throw new Error('「新規施設分類候補」に必要な列がありません: ' + required.join('、'));

  const historySheet = hotelDbV2GetOrCreateSheet_(
    spreadsheet,
    HOTEL_DB_V2_CONFIG.SHEETS.HISTORY,
    HOTEL_DB_V2_HISTORY_HEADERS
  );
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  const result = { approved: 0, applied: 0, conflicts: 0, errors: 0 };

  values.forEach(function(row, offset) {
    const rowNumber = offset + 2;
    function value(header) { return hotelDbV2Clean_(row[map[header] - 1]); }
    if (value('状態') !== HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.APPROVED_STATE) return;
    result.approved++;

    const candidate = {
      sheetName: value('元シート'), sheetId: Number(value('元シートID')), sourceRow: Number(value('元行')),
      placeId: value('Place ID'), name: value('施設名'), address: value('住所'),
      currentCategory: value('現在宿泊分類'), currentNotes: value('現在備考'),
      finalCategory: value('確定宿泊分類'), finalNotes: value('確定備考')
    };

    const precheck = hotelDbV2NewFacilityClassificationPrecheck_(candidate);
    if (!precheck.ok) {
      hotelDbV2NewFacilityClassificationSetResult_(sheet, map, rowNumber, HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.CONFLICT_STATE, '未反映: ' + precheck.reason);
      result.conflicts++;
      return;
    }

    let sourceSheet = null;
    let originalCategory = '';
    let originalNotes = '';
    try {
      sourceSheet = spreadsheet.getSheetById(candidate.sheetId);
      if (!sourceSheet || sourceSheet.getName() !== candidate.sheetName) {
        throw new Error('元シートが見つからないか、シート名が候補作成時から変更されています。');
      }
      const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
      hotelDbV2ValidateSourceSheet_(sourceSheet, sourceMap);
      if (!sourceMap.category || !sourceMap.notes) {
        throw new Error('元シートに「宿泊分類」または「備考」列がありません。');
      }
      if (candidate.sourceRow < 2 || candidate.sourceRow > sourceSheet.getLastRow()) {
        throw new Error('元行が見つかりません。');
      }

      const facility = hotelDbV2ReadFacility_(sourceSheet, candidate.sourceRow, sourceMap);
      const snapshotCheck = hotelDbV2NewFacilityClassificationSnapshotCheck_(candidate, facility);
      if (!snapshotCheck.ok) {
        hotelDbV2NewFacilityClassificationSetResult_(sheet, map, rowNumber, HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.CONFLICT_STATE, '未反映: ' + snapshotCheck.reason);
        result.conflicts++;
        return;
      }

      originalCategory = facility.category;
      originalNotes = facility.notes;
      const newNotes = hotelDbV2NewFacilityClassificationMergeNotes_(facility.notes, candidate.finalNotes);

      sourceSheet.getRange(candidate.sourceRow, sourceMap.category).setValue(candidate.finalCategory);
      sourceSheet.getRange(candidate.sourceRow, sourceMap.notes).setValue(newNotes);

      historySheet.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, {
        '日時': hotelDbV2Timestamp_(),
        '元シート': sourceSheet.getName(),
        '元シートID': sourceSheet.getSheetId(),
        '元行': candidate.sourceRow,
        '施設名': facility.name,
        '処理': '新規施設分類確定',
        '結果': '反映済み',
        'Place ID': facility.placeId,
        '一致スコア': '',
        '営業状態': '',
        '詳細': '宿泊分類=' + candidate.finalCategory + (candidate.finalNotes ? '／確定備考=' + candidate.finalNotes : '')
      }));

      hotelDbV2NewFacilityClassificationSetResult_(sheet, map, rowNumber, HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.APPLIED_STATE, '反映済み');
      result.applied++;
    } catch (error) {
      if (sourceSheet && candidate.sourceRow >= 2) {
        try {
          const rollbackMap = hotelDbV2GetHeaderMap_(sourceSheet);
          if (rollbackMap.category) sourceSheet.getRange(candidate.sourceRow, rollbackMap.category).setValue(originalCategory);
          if (rollbackMap.notes) sourceSheet.getRange(candidate.sourceRow, rollbackMap.notes).setValue(originalNotes);
        } catch (rollbackError) {
          console.error('PR18分類反映ロールバック失敗: ' + rollbackError.message);
        }
      }
      hotelDbV2NewFacilityClassificationSetResult_(sheet, map, rowNumber, HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.ERROR_STATE, '反映エラー: ' + error.message);
      result.errors++;
    }
  });

  return result;
}

function hotelDbV2NewFacilityClassificationPrecheck_(candidate) {
  if (!candidate) return { ok: false, reason: '候補データがありません。' };
  if (!candidate.sheetName || !candidate.sheetId || !candidate.sourceRow) return { ok: false, reason: '元シート情報が不足しています。' };
  if (!candidate.placeId || !candidate.name || !candidate.address) return { ok: false, reason: 'Place ID・施設名・住所のいずれかが不足しています。' };
  if (!candidate.finalCategory) return { ok: false, reason: '「確定宿泊分類」が未入力です。' };
  if (candidate.finalCategory.length > 100) return { ok: false, reason: '「確定宿泊分類」が長すぎます。' };
  if (candidate.finalNotes.length > 1000) return { ok: false, reason: '「確定備考」が長すぎます。' };
  return { ok: true, reason: '' };
}

function hotelDbV2NewFacilityClassificationSnapshotCheck_(candidate, facility) {
  const identity = hotelDbV2NewFacilityClassificationIdentityMatches_(candidate, facility);
  if (!identity.ok) return identity;
  if (hotelDbV2Clean_(facility.category)) {
    return { ok: false, reason: '元DBには既に宿泊分類が入っています。上書きしません。' };
  }
  if (hotelDbV2Clean_(candidate.currentCategory) !== hotelDbV2Clean_(facility.category)) {
    return { ok: false, reason: '現在宿泊分類が候補作成時から変更されています。' };
  }
  if (hotelDbV2Clean_(candidate.currentNotes) !== hotelDbV2Clean_(facility.notes)) {
    return { ok: false, reason: '備考が候補作成時から変更されています。' };
  }
  if (facility.notes.indexOf(HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.NOTE_MARKER) === -1) {
    return { ok: false, reason: '宿泊分類要確認の印が元備考からなくなっています。' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2NewFacilityClassificationMergeNotes_(currentNotes, finalNotes) {
  let base = hotelDbV2Clean_(currentNotes);
  base = base.replace(/／?宿泊分類要確認／?/g, '／');
  base = base.replace(/^／+|／+$/g, '').replace(/／{2,}/g, '／');
  const additions = [];
  if (base) additions.push(base);
  additions.push('宿泊分類確認済');
  const finalText = hotelDbV2Clean_(finalNotes);
  if (finalText && additions.indexOf(finalText) === -1) additions.push(finalText);
  return additions.join('／');
}

function hotelDbV2NewFacilityClassificationSetResult_(sheet, map, rowNumber, state, resultText) {
  sheet.getRange(rowNumber, map['状態']).setValue(state);
  sheet.getRange(rowNumber, map['反映日時']).setValue(hotelDbV2Timestamp_());
  sheet.getRange(rowNumber, map['反映結果']).setValue(resultText || '');
}

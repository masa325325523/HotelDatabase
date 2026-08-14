/**
 * PR #19 承認済み恒久閉業施設の安全なアーカイブ除外。
 *
 * 安全原則:
 * - 「要確認」で状態=承認 かつ 削除推奨判定=削除候補有力 の恒久閉業だけ対象。
 * - 元DB行全体の候補作成時ハッシュを再確認する。
 * - 実行直前に Places Details を再取得し、恒久閉業・名称・住所・郵便番号・一致スコアを再確認する。
 * - 閉業除外履歴へ完全保存し、読み戻し検証に成功してから元DBの内容だけ clearContent() する。
 * - deleteRow()/deleteRows() は使用しない。下の元行番号を一切ずらさない。
 * - 失敗時は元の値・数式を復元し、ハッシュ一致まで確認する。
 */

const HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME = '閉業除外履歴';

const HOTEL_DB_V2_PR19_AUDIT_HEADERS = Object.freeze([
  '除外候補作成日時', '除外元行ハッシュ', '除外処理日時', '除外履歴キー',
  '最終Google確認日時', '最終Google施設名', '最終Google住所',
  '最終Google営業状態', '最終一致スコア', '除外結果'
]);

const HOTEL_DB_V2_PR19_ARCHIVE_HEADERS = Object.freeze([
  '除外キー', '処理状態', 'アーカイブ日時', '除外完了日時',
  '元シート', '元シートID', '元行', '郵便番号', '市区町村', '住所',
  '施設名', '宿泊分類', '備考', 'Place ID', 'Google施設名', 'Google住所',
  '営業状態', '一致スコア', 'Google Maps URL', '候補確認日',
  '削除推奨判定', '削除判定理由', '最終Google確認日時',
  '最終Google施設名', '最終Google住所', '最終Google営業状態',
  '最終一致スコア', '元行ハッシュ', '元見出しJSON', '元行値JSON',
  '元行数式JSON', '詳細'
]);

const HOTEL_DB_V2_PR19 = Object.freeze({
  APPROVED: '承認',
  APPLIED: '除外済み',
  CONFLICT: '要再確認',
  ERROR: '除外エラー',
  LIKELY: '削除候補有力',
  CLOSED_REASON: '閉業',
  CLOSED_STATUS_JA: '閉業',
  CLOSED_STATUS_API: 'CLOSED_PERMANENTLY'
});

function runHotelDbV2ApplyApprovedClosedFacilityRemovals() {
  return withHotelDbV2Lock_('承認済み閉業候補の安全除外', function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const review = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);
    const ui = SpreadsheetApp.getUi();
    if (!review || review.getLastRow() < 2) {
      ui.alert('「要確認」に処理対象がありません。');
      return hotelDbV2Pr19EmptyResult_();
    }

    const map = hotelDbV2Pr19EnsureReviewHeaders_(review);
    const approved = hotelDbV2Pr19CountApproved_(review, map);
    if (!approved) {
      ui.alert('「要確認」に状態が「承認」の行はありません。');
      return hotelDbV2Pr19EmptyResult_();
    }

    const response = ui.alert(
      '承認済み閉業候補を安全に除外',
      '承認対象: ' + approved + '件\n\n' +
      '・削除候補有力＋人が承認した恒久閉業施設だけ対象です。\n' +
      '・実行直前に元DBとGoogle Placesを再確認します。\n' +
      '・閉業除外履歴へ完全保存できた施設だけ元DBから除外します。\n' +
      '・元DBの行そのものは削除せず、内容だけを空にします。\n' +
      '・Google候補なし、一時休業、開業予定は除外しません。\n\n続行しますか？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2ApplyApprovedClosedFacilityRemovals_({
      spreadsheet: spreadsheet,
      reviewSheet: review
    });

    ui.alert([
      '承認済み閉業候補の安全除外 完了', '',
      '承認対象: ' + result.approved + '件',
      '除外済み: ' + result.applied + '件',
      '要再確認・未除外: ' + result.conflicts + '件',
      'エラー・未除外: ' + result.errors + '件',
      '集計整合: ' + result.reconciliation, '',
      '物理的な行削除: なし',
      '閉業除外履歴への保存: 実施'
    ].join('\n'));
    return result;
  });
}

function hotelDbV2Pr19EmptyResult_() {
  return { approved: 0, applied: 0, conflicts: 0, errors: 0, reconciliation: '一致' };
}

function hotelDbV2Pr19CountApproved_(sheet, map) {
  if (!map['状態'] || sheet.getLastRow() < 2) return 0;
  return sheet.getRange(2, map['状態'], sheet.getLastRow() - 1, 1)
    .getDisplayValues()
    .reduce(function(total, row) {
      return total + (hotelDbV2Clean_(row[0]) === HOTEL_DB_V2_PR19.APPROVED ? 1 : 0);
    }, 0);
}

function hotelDbV2Pr19EnsureReviewHeaders_(sheet) {
  let map = hotelDbV2DeletionHeaderMap_(sheet);
  const required = [
    '状態', '元シート', '元シートID', '元行', '郵便番号', '市区町村', '住所',
    '施設名', '宿泊分類', '理由', '候補施設名', '候補住所', '候補Place ID',
    '一致スコア', '営業状態', 'Google Maps URL', '確認日',
    '削除推奨判定', '削除判定理由', '削除信頼度'
  ];
  const missing = required.filter(function(header) { return !map[header]; });
  if (missing.length) throw new Error('「要確認」に必要な列がありません: ' + missing.join('、'));

  const missingAudit = HOTEL_DB_V2_PR19_AUDIT_HEADERS.filter(function(header) { return !map[header]; });
  if (missingAudit.length) {
    sheet.getRange(1, sheet.getLastColumn() + 1, 1, missingAudit.length).setValues([missingAudit]);
    map = hotelDbV2DeletionHeaderMap_(sheet);
  }
  return map;
}

function hotelDbV2Pr19ReadCandidate_(row, map) {
  function value(header) { return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : ''; }
  return {
    state: value('状態'), sheetName: value('元シート'), sheetId: Number(value('元シートID')),
    sourceRow: Number(value('元行')), postalCode: value('郵便番号'), municipality: value('市区町村'),
    address: value('住所'), name: value('施設名'), category: value('宿泊分類'),
    reason: value('理由'), candidateName: value('候補施設名'), candidateAddress: value('候補住所'),
    placeId: value('候補Place ID'), matchScore: Number(value('一致スコア') || 0),
    businessStatus: value('営業状態'), mapsUrl: value('Google Maps URL'), checkedAt: value('確認日'),
    recommendation: value('削除推奨判定'), deletionReason: value('削除判定理由'),
    confidence: Number(value('削除信頼度') || 0), snapshotAt: value('除外候補作成日時'),
    rowHash: value('除外元行ハッシュ'), archiveKey: value('除外履歴キー')
  };
}

function hotelDbV2Pr19Precheck_(candidate) {
  if (!candidate || candidate.state !== HOTEL_DB_V2_PR19.APPROVED) return { ok: false, reason: '状態が「承認」ではありません。' };
  if (candidate.recommendation !== HOTEL_DB_V2_PR19.LIKELY) return { ok: false, reason: '削除推奨判定が「削除候補有力」ではありません。' };
  if (candidate.reason !== HOTEL_DB_V2_PR19.CLOSED_REASON) return { ok: false, reason: '理由が「閉業」ではありません。' };
  if (!candidate.placeId) return { ok: false, reason: '候補Place IDがありません。' };
  if (!candidate.candidateName || !candidate.candidateAddress) return { ok: false, reason: '候補施設名または候補住所がありません。' };
  if (candidate.businessStatus !== HOTEL_DB_V2_PR19.CLOSED_STATUS_JA) return { ok: false, reason: '候補作成時の営業状態が「閉業」ではありません。' };
  if (candidate.matchScore < HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE) return { ok: false, reason: '候補作成時の一致スコアが自動採用基準未満です。' };
  if (!candidate.sheetId || !candidate.sheetName || candidate.sourceRow < 2) return { ok: false, reason: '元シート情報が不足しています。' };
  if (!candidate.rowHash) return { ok: false, reason: '除外元行ハッシュがありません。⑫を再実行してください。' };
  return { ok: true, reason: '' };
}

function hotelDbV2ApplyApprovedClosedFacilityRemovals_(options) {
  const opts = options || {};
  const spreadsheet = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const reviewSheet = opts.reviewSheet || spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW);
  if (!reviewSheet || reviewSheet.getLastRow() < 2) return hotelDbV2Pr19EmptyResult_();
  const map = hotelDbV2Pr19EnsureReviewHeaders_(reviewSheet);
  const archiveSheet = opts.archiveSheet || hotelDbV2GetOrCreateSheet_(spreadsheet, HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME, HOTEL_DB_V2_PR19_ARCHIVE_HEADERS);
  const historySheet = opts.historySheet || hotelDbV2GetOrCreateSheet_(spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY, HOTEL_DB_V2_HISTORY_HEADERS);
  const detailsProvider = opts.placeDetailsProvider || function(placeId) { return hotelDbV2GetPlaceDetails_(placeId); };
  const scoreProvider = opts.scoreProvider || function(facility, place) { return hotelDbV2CalculateMatchScore_(facility, place); };
  const values = reviewSheet.getRange(2, 1, reviewSheet.getLastRow() - 1, reviewSheet.getLastColumn()).getDisplayValues();
  const result = hotelDbV2Pr19EmptyResult_();

  values.forEach(function(row, offset) {
    const reviewRow = offset + 2;
    const candidate = hotelDbV2Pr19ReadCandidate_(row, map);
    if (candidate.state !== HOTEL_DB_V2_PR19.APPROVED) return;
    result.approved++;

    const precheck = hotelDbV2Pr19Precheck_(candidate);
    if (!precheck.ok) {
      hotelDbV2Pr19Conflict_(reviewSheet, map, reviewRow, precheck.reason);
      result.conflicts++;
      return;
    }

    let sourceSheet = null;
    let snapshot = null;
    let archiveRow = 0;
    let livePlace = null;
    let liveScore = 0;
    let didClear = false;

    try {
      sourceSheet = spreadsheet.getSheetById(candidate.sheetId);
      const sourceCheck = hotelDbV2Pr19ValidateSource_(sourceSheet, candidate);
      if (!sourceCheck.ok) {
        hotelDbV2Pr19Conflict_(reviewSheet, map, reviewRow, sourceCheck.reason);
        result.conflicts++;
        return;
      }
      snapshot = hotelDbV2Pr19SnapshotRow_(sourceSheet, candidate.sourceRow);
      if (snapshot.hash !== candidate.rowHash) {
        hotelDbV2Pr19Conflict_(reviewSheet, map, reviewRow, '元DB行が候補作成時から変更されています（行ハッシュ不一致）。');
        result.conflicts++;
        return;
      }

      const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
      hotelDbV2ValidateSourceSheet_(sourceSheet, sourceMap);
      const facility = hotelDbV2ReadFacility_(sourceSheet, candidate.sourceRow, sourceMap);
      const identityCheck = hotelDbV2Pr19ValidateFacilitySnapshot_(facility, candidate);
      if (!identityCheck.ok) {
        hotelDbV2Pr19Conflict_(reviewSheet, map, reviewRow, identityCheck.reason);
        result.conflicts++;
        return;
      }

      livePlace = detailsProvider(candidate.placeId);
      const liveCheck = hotelDbV2Pr19ValidateLivePlace_(facility, candidate, livePlace);
      if (!liveCheck.ok) {
        hotelDbV2Pr19SetLiveAudit_(reviewSheet, map, reviewRow, livePlace, '要再確認');
        hotelDbV2Pr19Conflict_(reviewSheet, map, reviewRow, liveCheck.reason);
        result.conflicts++;
        return;
      }

      liveScore = Number(scoreProvider(facility, livePlace) || 0);
      hotelDbV2Pr19SetLiveAudit_(reviewSheet, map, reviewRow, livePlace, liveScore);
      if (liveScore < HOTEL_DB_V2_CONFIG.AUTO_ACCEPT_SCORE) {
        hotelDbV2Pr19Conflict_(reviewSheet, map, reviewRow, '最終一致スコアが自動採用基準未満です。');
        result.conflicts++;
        return;
      }

      const archiveKey = hotelDbV2Pr19ArchiveKey_(candidate.sheetId, candidate.sourceRow, candidate.placeId, snapshot.hash);
      const existingArchive = hotelDbV2Pr19FindArchive_(archiveSheet, archiveKey);
      if (existingArchive && existingArchive.state === HOTEL_DB_V2_PR19.APPLIED) {
        hotelDbV2Pr19SetResult_(reviewSheet, map, reviewRow, HOTEL_DB_V2_PR19.APPLIED, '既に同一除外キーで除外済みです。', archiveKey);
        result.applied++;
        return;
      }

      archiveRow = hotelDbV2Pr19AppendArchive_(archiveSheet, candidate, facility, livePlace, liveScore, snapshot, archiveKey);
      const archiveCheck = hotelDbV2Pr19VerifyArchive_(archiveSheet, archiveRow, archiveKey, snapshot.hash, candidate.placeId, facility.name);
      if (!archiveCheck.ok) throw new Error('アーカイブ検証失敗: ' + archiveCheck.reason);

      sourceSheet.getRange(candidate.sourceRow, 1, 1, sourceSheet.getLastColumn()).clearContent();
      didClear = true;
      const cleared = hotelDbV2Pr19RowIsEmpty_(sourceSheet, candidate.sourceRow);
      if (!cleared) throw new Error('元DB行の内容クリア後検証に失敗しました。');

      hotelDbV2Pr19FinalizeArchive_(archiveSheet, archiveRow, '除外済み', '正常に元DBから内容除外。物理行削除なし。');
      hotelDbV2Pr19SetResult_(reviewSheet, map, reviewRow, HOTEL_DB_V2_PR19.APPLIED, '除外済み', archiveKey);
      historySheet.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, {
        '日時': hotelDbV2Timestamp_(), '元シート': candidate.sheetName, '元シートID': candidate.sheetId,
        '元行': candidate.sourceRow, '施設名': facility.name, '処理': '承認閉業除外', '結果': '除外済み',
        'Place ID': candidate.placeId, '一致スコア': liveScore, '営業状態': '閉業',
        '詳細': '閉業除外履歴キー=' + archiveKey + '／物理行削除なし'
      }));
      hotelDbV2Pr19InvalidateLinkedCandidates_(spreadsheet, candidate.sheetId, candidate.sourceRow);
      result.applied++;
    } catch (error) {
      if (didClear && sourceSheet && snapshot) {
        try {
          hotelDbV2Pr19RestoreRow_(sourceSheet, candidate.sourceRow, snapshot);
          const restored = hotelDbV2Pr19SnapshotRow_(sourceSheet, candidate.sourceRow);
          if (restored.hash !== snapshot.hash) throw new Error('ロールバック後ハッシュ不一致');
          if (archiveRow) hotelDbV2Pr19FinalizeArchive_(archiveSheet, archiveRow, 'ロールバック済み', error.message);
        } catch (rollbackError) {
          if (archiveRow) hotelDbV2Pr19FinalizeArchive_(archiveSheet, archiveRow, 'ロールバック要確認', error.message + '／' + rollbackError.message);
        }
      }
      hotelDbV2Pr19SetResult_(reviewSheet, map, reviewRow, HOTEL_DB_V2_PR19.ERROR, '未除外: ' + error.message, candidate.archiveKey || '');
      result.errors++;
    }
  });

  result.reconciliation = result.approved === result.applied + result.conflicts + result.errors ? '一致' : '要確認';
  return result;
}

function hotelDbV2Pr19ValidateSource_(sourceSheet, candidate) {
  if (!sourceSheet) return { ok: false, reason: '元シートが見つかりません。' };
  if (sourceSheet.getName() !== candidate.sheetName) return { ok: false, reason: '元シート名が候補作成時から変更されています。' };
  if (candidate.sourceRow < 2 || candidate.sourceRow > sourceSheet.getLastRow()) return { ok: false, reason: '元行が見つかりません。' };
  if (sourceSheet.getName() === HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME) return { ok: false, reason: '閉業除外履歴は元DBとして処理できません。' };
  return { ok: true, reason: '' };
}

function hotelDbV2Pr19ValidateFacilitySnapshot_(facility, candidate) {
  if (!facility || (!facility.name && !facility.address)) return { ok: false, reason: '元行が空です。' };
  if (hotelDbV2NormalizePostalCode_(facility.postalCode) !== hotelDbV2NormalizePostalCode_(candidate.postalCode)) return { ok: false, reason: '郵便番号が候補作成時から変更されています。' };
  if (hotelDbV2NormalizeText_(facility.municipality) !== hotelDbV2NormalizeText_(candidate.municipality)) return { ok: false, reason: '市区町村が候補作成時から変更されています。' };
  if (!hotelDbV2DiscoveryAddressesSame_(facility.address, candidate.address)) return { ok: false, reason: '住所が候補作成時から変更されています。' };
  if (hotelDbV2NormalizeText_(facility.name) !== hotelDbV2NormalizeText_(candidate.name)) return { ok: false, reason: '施設名が候補作成時から変更されています。' };
  if (hotelDbV2Clean_(facility.category) !== hotelDbV2Clean_(candidate.category)) return { ok: false, reason: '宿泊分類が候補作成時から変更されています。' };
  if (hotelDbV2Clean_(facility.placeId) !== candidate.placeId) return { ok: false, reason: '元DB Place IDが候補Place IDと一致しません。' };
  return { ok: true, reason: '' };
}

function hotelDbV2Pr19ValidateLivePlace_(facility, candidate, livePlace) {
  if (!livePlace || !hotelDbV2Clean_(livePlace.id)) return { ok: false, reason: 'Google Placesで候補を再取得できません。' };
  if (hotelDbV2Clean_(livePlace.id) !== candidate.placeId) return { ok: false, reason: '再取得したPlace IDが候補と一致しません。' };
  if (hotelDbV2Clean_(livePlace.businessStatus) !== HOTEL_DB_V2_PR19.CLOSED_STATUS_API) return { ok: false, reason: '最終Google確認で恒久閉業ではありません。' };
  const liveName = hotelDbV2GetDisplayName_(livePlace);
  const liveAddress = hotelDbV2GetJapaneseFullAddress_(livePlace);
  const livePostal = hotelDbV2GetPostalCode_(livePlace);
  if (!liveName || !liveAddress || !hotelDbV2NormalizePostalCode_(livePostal)) return { ok: false, reason: '最終Google確認で名称・住所・郵便番号を確認できません。' };
  if (hotelDbV2NormalizeText_(liveName) !== hotelDbV2NormalizeText_(candidate.candidateName)) return { ok: false, reason: 'Google施設名が候補作成時から変更されています。' };
  if (!hotelDbV2DiscoveryAddressesSame_(liveAddress, candidate.candidateAddress)) return { ok: false, reason: 'Google住所が候補作成時から変更されています。' };
  if (hotelDbV2NormalizePostalCode_(livePostal) !== hotelDbV2NormalizePostalCode_(facility.postalCode)) return { ok: false, reason: 'Google郵便番号が元DB郵便番号と一致しません。' };
  return { ok: true, reason: '' };
}

function hotelDbV2Pr19SnapshotRow_(sheet, rowNumber) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const range = sheet.getRange(rowNumber, 1, 1, lastColumn);
  const displayValues = range.getDisplayValues()[0];
  const values = range.getValues()[0];
  const formulas = range.getFormulas()[0];
  const payload = JSON.stringify({ headers: headers, displayValues: displayValues, formulas: formulas });
  return { headers: headers, displayValues: displayValues, values: values, formulas: formulas, hash: hotelDbV2Pr19Sha256_(payload) };
}

function hotelDbV2Pr19Sha256_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return bytes.map(function(value) { const b = value < 0 ? value + 256 : value; return ('0' + b.toString(16)).slice(-2); }).join('');
}

function hotelDbV2Pr19ArchiveKey_(sheetId, row, placeId, hash) {
  return [sheetId, row, placeId, hash].join('|');
}

function hotelDbV2Pr19AppendArchive_(sheet, candidate, facility, livePlace, liveScore, snapshot, archiveKey) {
  const row = hotelDbV2RowFromObject_(HOTEL_DB_V2_PR19_ARCHIVE_HEADERS, {
    '除外キー': archiveKey, '処理状態': 'アーカイブ済み・除外前', 'アーカイブ日時': hotelDbV2Timestamp_(),
    '除外完了日時': '', '元シート': candidate.sheetName, '元シートID': candidate.sheetId, '元行': candidate.sourceRow,
    '郵便番号': facility.postalCode, '市区町村': facility.municipality, '住所': facility.address,
    '施設名': facility.name, '宿泊分類': facility.category, '備考': facility.notes, 'Place ID': facility.placeId,
    'Google施設名': candidate.candidateName, 'Google住所': candidate.candidateAddress, '営業状態': candidate.businessStatus,
    '一致スコア': candidate.matchScore, 'Google Maps URL': candidate.mapsUrl, '候補確認日': candidate.checkedAt,
    '削除推奨判定': candidate.recommendation, '削除判定理由': candidate.deletionReason,
    '最終Google確認日時': hotelDbV2Timestamp_(), '最終Google施設名': hotelDbV2GetDisplayName_(livePlace),
    '最終Google住所': hotelDbV2GetJapaneseFullAddress_(livePlace), '最終Google営業状態': hotelDbV2TranslateBusinessStatus_(livePlace.businessStatus),
    '最終一致スコア': liveScore, '元行ハッシュ': snapshot.hash,
    '元見出しJSON': JSON.stringify(snapshot.headers), '元行値JSON': JSON.stringify(snapshot.values),
    '元行数式JSON': JSON.stringify(snapshot.formulas), '詳細': 'PR19承認閉業除外。物理行削除なし。'
  });
  const rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  SpreadsheetApp.flush();
  return rowNumber;
}

function hotelDbV2Pr19VerifyArchive_(sheet, rowNumber, key, hash, placeId, name) {
  const map = hotelDbV2HeaderIndex_(HOTEL_DB_V2_PR19_ARCHIVE_HEADERS);
  const row = sheet.getRange(rowNumber, 1, 1, HOTEL_DB_V2_PR19_ARCHIVE_HEADERS.length).getDisplayValues()[0];
  function value(header) { return hotelDbV2Clean_(row[map[header] - 1]); }
  if (value('除外キー') !== key) return { ok: false, reason: '除外キー不一致' };
  if (value('元行ハッシュ') !== hash) return { ok: false, reason: '元行ハッシュ不一致' };
  if (value('Place ID') !== placeId) return { ok: false, reason: 'Place ID不一致' };
  if (hotelDbV2NormalizeText_(value('施設名')) !== hotelDbV2NormalizeText_(name)) return { ok: false, reason: '施設名不一致' };
  return { ok: true, reason: '' };
}

function hotelDbV2Pr19FindArchive_(sheet, key) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (hotelDbV2Clean_(values[i][0]) === key) return { row: i + 2, state: hotelDbV2Clean_(values[i][1]) };
  }
  return null;
}

function hotelDbV2Pr19FinalizeArchive_(sheet, rowNumber, state, detail) {
  const map = hotelDbV2HeaderIndex_(HOTEL_DB_V2_PR19_ARCHIVE_HEADERS);
  sheet.getRange(rowNumber, map['処理状態']).setValue(state);
  if (state === '除外済み') sheet.getRange(rowNumber, map['除外完了日時']).setValue(hotelDbV2Timestamp_());
  sheet.getRange(rowNumber, map['詳細']).setValue(detail || '');
}

function hotelDbV2Pr19RowIsEmpty_(sheet, rowNumber) {
  const range = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn());
  const values = range.getDisplayValues()[0];
  const formulas = range.getFormulas()[0];
  return values.every(function(v) { return hotelDbV2Clean_(v) === ''; }) && formulas.every(function(v) { return hotelDbV2Clean_(v) === ''; });
}

function hotelDbV2Pr19RestoreRow_(sheet, rowNumber, snapshot) {
  const range = sheet.getRange(rowNumber, 1, 1, snapshot.values.length);
  range.clearContent();
  for (let i = 0; i < snapshot.values.length; i++) {
    const cell = sheet.getRange(rowNumber, i + 1);
    if (snapshot.formulas[i]) cell.setFormula(snapshot.formulas[i]);
    else cell.setValue(snapshot.values[i]);
  }
  SpreadsheetApp.flush();
}

function hotelDbV2Pr19SetLiveAudit_(sheet, map, rowNumber, livePlace, liveScore) {
  sheet.getRange(rowNumber, map['最終Google確認日時']).setValue(hotelDbV2Timestamp_());
  sheet.getRange(rowNumber, map['最終Google施設名']).setValue(livePlace ? hotelDbV2GetDisplayName_(livePlace) : '');
  sheet.getRange(rowNumber, map['最終Google住所']).setValue(livePlace ? hotelDbV2GetJapaneseFullAddress_(livePlace) : '');
  sheet.getRange(rowNumber, map['最終Google営業状態']).setValue(livePlace ? hotelDbV2TranslateBusinessStatus_(livePlace.businessStatus) : '');
  sheet.getRange(rowNumber, map['最終一致スコア']).setValue(typeof liveScore === 'number' ? liveScore : '');
}

function hotelDbV2Pr19SetResult_(sheet, map, rowNumber, state, text, archiveKey) {
  sheet.getRange(rowNumber, map['状態']).setValue(state);
  sheet.getRange(rowNumber, map['除外処理日時']).setValue(hotelDbV2Timestamp_());
  if (archiveKey !== undefined) sheet.getRange(rowNumber, map['除外履歴キー']).setValue(archiveKey || '');
  sheet.getRange(rowNumber, map['除外結果']).setValue(text || '');
}

function hotelDbV2Pr19Conflict_(sheet, map, rowNumber, reason) {
  hotelDbV2Pr19SetResult_(sheet, map, rowNumber, HOTEL_DB_V2_PR19.CONFLICT, '未除外: ' + reason, '');
}

function hotelDbV2Pr19InvalidateLinkedCandidates_(spreadsheet, sheetId, sourceRow) {
  hotelDbV2Pr19InvalidateSimpleSheet_(spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS), sheetId, sourceRow, '元シートID', ['元行']);
  hotelDbV2Pr19InvalidateSimpleSheet_(spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES), sheetId, sourceRow, '元シートID', ['行1', '行2']);
  hotelDbV2Pr19InvalidateSimpleSheet_(spreadsheet.getSheetByName('新規施設分類候補'), sheetId, sourceRow, '元シートID', ['元行']);
}

function hotelDbV2Pr19InvalidateSimpleSheet_(sheet, sheetId, sourceRow, sheetIdHeader, rowHeaders) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(header, index) { const h = hotelDbV2Clean_(header); if (h) map[h] = index + 1; });
  if (!map['状態'] || !map[sheetIdHeader]) return;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
  values.forEach(function(row, offset) {
    if (String(row[map[sheetIdHeader] - 1]) !== String(sheetId)) return;
    const matchesRow = rowHeaders.some(function(header) { return map[header] && Number(row[map[header] - 1]) === Number(sourceRow); });
    if (!matchesRow) return;
    const state = hotelDbV2Clean_(row[map['状態'] - 1]);
    if (state === '未確認' || state === '承認') sheet.getRange(offset + 2, map['状態']).setValue('要再確認');
  });
}
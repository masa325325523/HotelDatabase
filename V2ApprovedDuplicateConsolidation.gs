/**
 * PR #20 承認済み重複候補の安全な整理。
 *
 * 安全原則:
 * - 状態=承認、推奨判定=重複濃厚、かつ人が「残す行」を明示した候補だけ対象。
 * - 行1/行2の候補作成時ハッシュを再確認し、候補作成後の変更があれば整理しない。
 * - 施設名・住所が現在も正規化後に完全一致することを再確認する。
 * - 除外行にしか存在しない値・数式、または両行の競合値が1つでもあれば整理しない。
 * - 重複整理履歴へ両行を完全保存し、読み戻し検証後に除外行だけ clearContent() する。
 * - deleteRow()/deleteRows() は使用しない。下の元行番号を一切ずらさない。
 * - 失敗時は除外行の値・数式を復元し、ハッシュ一致まで確認する。
 */

const HOTEL_DB_V2_PR20_ARCHIVE_SHEET_NAME = '重複整理履歴';

const HOTEL_DB_V2_PR20_AUDIT_HEADERS = Object.freeze([
  '整理候補作成日時', '行1ハッシュ', '行2ハッシュ', '整理スナップショット状態',
  '残す行', '除外予定行', '整理処理日時', '整理履歴キー', '整理結果'
]);

const HOTEL_DB_V2_PR20_ARCHIVE_HEADERS = Object.freeze([
  '整理キー', '処理状態', 'アーカイブ日時', '整理完了日時',
  '元シート', '元シートID', '重複キー',
  '行1', '施設名1', '住所1', '行1ハッシュ',
  '行2', '施設名2', '住所2', '行2ハッシュ',
  '残す行', '除外行', 'Place ID', '類似度',
  '推奨判定', '自動判定理由', '信頼度', '確認日',
  '元見出しJSON', '残す行値JSON', '残す行数式JSON',
  '除外行値JSON', '除外行数式JSON', '情報損失チェック', '詳細'
]);

const HOTEL_DB_V2_PR20 = Object.freeze({
  APPROVED: '承認',
  APPLIED: '整理済み',
  CONFLICT: '要再確認',
  ERROR: '整理エラー',
  STRONG: '重複濃厚',
  SNAPSHOT_READY: '作成済み',
  SNAPSHOT_REVIEW: '要再確認',
  SNAPSHOT_OUT: '対象外'
});

function runHotelDbV2ApplyApprovedDuplicateConsolidations() {
  return withHotelDbV2Lock_('承認済み重複候補の安全整理', function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES);
    const ui = SpreadsheetApp.getUi();

    if (!sheet || sheet.getLastRow() < 2) {
      ui.alert('「重複候補」に処理対象がありません。');
      return hotelDbV2Pr20EmptyResult_();
    }

    const map = hotelDbV2Pr20EnsureDuplicateHeaders_(sheet);
    const approved = hotelDbV2Pr20CountApproved_(sheet, map);
    if (!approved) {
      ui.alert('「重複候補」に状態が「承認」の行はありません。');
      return hotelDbV2Pr20EmptyResult_();
    }

    const response = ui.alert(
      '承認済み重複候補を安全に整理',
      '承認対象: ' + approved + '件\n\n' +
      '・「重複濃厚」かつ人が「残す行」を明示した候補だけ対象です。\n' +
      '・行1/行2が候補作成後に変わっていないか再確認します。\n' +
      '・除外行だけに残る情報や競合値が1つでもあれば整理しません。\n' +
      '・重複整理履歴へ両行を完全保存してから、除外行の内容だけ空にします。\n' +
      '・別部屋・別階・同住所別名など「要人確認」は整理しません。\n\n' +
      '続行しますか？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return { cancelled: true };

    const result = hotelDbV2ApplyApprovedDuplicateConsolidations_({
      spreadsheet: spreadsheet,
      duplicateSheet: sheet
    });

    ui.alert([
      '承認済み重複候補の安全整理 完了', '',
      '承認対象: ' + result.approved + '件',
      '整理済み: ' + result.applied + '件',
      '要再確認・未整理: ' + result.conflicts + '件',
      'エラー・未整理: ' + result.errors + '件',
      '集計整合: ' + result.reconciliation, '',
      '物理的な行削除: なし',
      '重複整理履歴への保存: 実施'
    ].join('\n'));
    return result;
  });
}

function hotelDbV2Pr20EmptyResult_() {
  return { approved: 0, applied: 0, conflicts: 0, errors: 0, reconciliation: '一致' };
}

function hotelDbV2Pr20HeaderMap_(sheet) {
  const map = {};
  if (!sheet || sheet.getLastColumn() < 1) return map;
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0]
    .forEach(function(value, index) {
      const header = hotelDbV2Clean_(value);
      if (header && !map[header]) map[header] = index + 1;
    });
  return map;
}

function hotelDbV2Pr20EnsureDuplicateHeaders_(sheet) {
  let map = hotelDbV2Pr20HeaderMap_(sheet);
  const required = [
    '重複キー', '元シート', '元シートID', '行1', '施設名1', '住所1',
    '行2', '施設名2', '住所2', 'Place ID', '類似度', '確認日', '状態',
    '推奨判定', '自動判定理由', '信頼度'
  ];
  const missing = required.filter(function(header) { return !map[header]; });
  if (missing.length) throw new Error('「重複候補」に必要な列がありません: ' + missing.join('、'));

  const missingAudit = HOTEL_DB_V2_PR20_AUDIT_HEADERS.filter(function(header) { return !map[header]; });
  if (missingAudit.length) {
    sheet.getRange(1, sheet.getLastColumn() + 1, 1, missingAudit.length).setValues([missingAudit]);
    map = hotelDbV2Pr20HeaderMap_(sheet);
  }
  return map;
}

function hotelDbV2Pr20CountApproved_(sheet, map) {
  if (!map['状態'] || sheet.getLastRow() < 2) return 0;
  return sheet.getRange(2, map['状態'], sheet.getLastRow() - 1, 1).getDisplayValues()
    .reduce(function(total, row) {
      return total + (hotelDbV2Clean_(row[0]) === HOTEL_DB_V2_PR20.APPROVED ? 1 : 0);
    }, 0);
}

function hotelDbV2Pr20ReadCandidate_(row, map) {
  function value(header) { return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : ''; }
  return {
    key: value('重複キー'), state: value('状態'), sheetName: value('元シート'),
    sheetId: Number(value('元シートID')), row1: Number(value('行1')), name1: value('施設名1'),
    address1: value('住所1'), row2: Number(value('行2')), name2: value('施設名2'),
    address2: value('住所2'), placeId: value('Place ID'), similarity: Number(value('類似度') || 0),
    checkedAt: value('確認日'), recommendation: value('推奨判定'), triageReason: value('自動判定理由'),
    confidence: Number(value('信頼度') || 0), hash1: value('行1ハッシュ'), hash2: value('行2ハッシュ'),
    snapshotState: value('整理スナップショット状態'), keepRow: Number(value('残す行')),
    archiveKey: value('整理履歴キー')
  };
}

function hotelDbV2Pr20Precheck_(candidate) {
  if (!candidate || candidate.state !== HOTEL_DB_V2_PR20.APPROVED) {
    return { ok: false, reason: '状態が「承認」ではありません。' };
  }
  if (candidate.recommendation !== HOTEL_DB_V2_PR20.STRONG) {
    return { ok: false, reason: '推奨判定が「重複濃厚」ではありません。' };
  }
  if (candidate.snapshotState !== HOTEL_DB_V2_PR20.SNAPSHOT_READY || !candidate.hash1 || !candidate.hash2) {
    return { ok: false, reason: '安全な行スナップショットがありません。⑩を再実行して再確認してください。' };
  }
  if (!candidate.sheetId || !candidate.sheetName || candidate.row1 < 2 || candidate.row2 < 2 || candidate.row1 === candidate.row2) {
    return { ok: false, reason: '元シートまたは行1/行2情報が不正です。' };
  }
  if (!candidate.name1 || !candidate.address1 || !candidate.name2 || !candidate.address2) {
    return { ok: false, reason: '施設名または住所が不足しています。' };
  }
  if (candidate.keepRow !== candidate.row1 && candidate.keepRow !== candidate.row2) {
    return { ok: false, reason: '「残す行」には行1または行2の行番号を明示してください。' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2ApplyApprovedDuplicateConsolidations_(options) {
  const opts = options || {};
  const spreadsheet = opts.spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  const duplicateSheet = opts.duplicateSheet || spreadsheet.getSheetByName(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES);
  if (!duplicateSheet || duplicateSheet.getLastRow() < 2) return hotelDbV2Pr20EmptyResult_();

  const map = hotelDbV2Pr20EnsureDuplicateHeaders_(duplicateSheet);
  const archiveSheet = opts.archiveSheet || hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_PR20_ARCHIVE_SHEET_NAME, HOTEL_DB_V2_PR20_ARCHIVE_HEADERS
  );
  const historySheet = opts.historySheet || hotelDbV2GetOrCreateSheet_(
    spreadsheet, HOTEL_DB_V2_CONFIG.SHEETS.HISTORY, HOTEL_DB_V2_HISTORY_HEADERS
  );
  const values = duplicateSheet.getRange(2, 1, duplicateSheet.getLastRow() - 1, duplicateSheet.getLastColumn()).getDisplayValues();
  const result = hotelDbV2Pr20EmptyResult_();

  values.forEach(function(row, offset) {
    const duplicateRow = offset + 2;
    const candidate = hotelDbV2Pr20ReadCandidate_(row, map);
    if (candidate.state !== HOTEL_DB_V2_PR20.APPROVED) return;
    result.approved++;

    const precheck = hotelDbV2Pr20Precheck_(candidate);
    if (!precheck.ok) {
      hotelDbV2Pr20Conflict_(duplicateSheet, map, duplicateRow, precheck.reason);
      result.conflicts++;
      return;
    }

    const removeRow = candidate.keepRow === candidate.row1 ? candidate.row2 : candidate.row1;
    let sourceSheet = null;
    let snapshot1 = null;
    let snapshot2 = null;
    let removeSnapshot = null;
    let archiveRow = 0;
    let archiveKey = '';
    let didClear = false;

    try {
      sourceSheet = spreadsheet.getSheetById(candidate.sheetId);
      const sourceCheck = hotelDbV2Pr20ValidateSource_(sourceSheet, candidate);
      if (!sourceCheck.ok) {
        hotelDbV2Pr20Conflict_(duplicateSheet, map, duplicateRow, sourceCheck.reason);
        result.conflicts++;
        return;
      }

      snapshot1 = hotelDbV2Pr19SnapshotRow_(sourceSheet, candidate.row1);
      snapshot2 = hotelDbV2Pr19SnapshotRow_(sourceSheet, candidate.row2);
      if (snapshot1.hash !== candidate.hash1 || snapshot2.hash !== candidate.hash2) {
        hotelDbV2Pr20Conflict_(duplicateSheet, map, duplicateRow, '行1または行2が候補作成時から変更されています（行ハッシュ不一致）。');
        result.conflicts++;
        return;
      }

      const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
      hotelDbV2ValidateSourceSheet_(sourceSheet, sourceMap);
      const facility1 = hotelDbV2ReadFacility_(sourceSheet, candidate.row1, sourceMap);
      const facility2 = hotelDbV2ReadFacility_(sourceSheet, candidate.row2, sourceMap);
      const identityCheck = hotelDbV2Pr20ValidatePairIdentity_(candidate, facility1, facility2);
      if (!identityCheck.ok) {
        hotelDbV2Pr20Conflict_(duplicateSheet, map, duplicateRow, identityCheck.reason);
        result.conflicts++;
        return;
      }

      const keepSnapshot = candidate.keepRow === candidate.row1 ? snapshot1 : snapshot2;
      removeSnapshot = removeRow === candidate.row1 ? snapshot1 : snapshot2;
      const noLoss = hotelDbV2Pr20NoLossCheck_(sourceMap, keepSnapshot, removeSnapshot);
      if (!noLoss.ok) {
        hotelDbV2Pr20Conflict_(duplicateSheet, map, duplicateRow, '情報損失防止チェック: ' + noLoss.reason);
        result.conflicts++;
        return;
      }

      archiveKey = hotelDbV2Pr20ArchiveKey_(
        candidate.sheetId, candidate.row1, candidate.row2, candidate.keepRow, snapshot1.hash, snapshot2.hash
      );
      const existingArchive = hotelDbV2Pr20FindArchive_(archiveSheet, archiveKey);
      if (existingArchive && existingArchive.state === HOTEL_DB_V2_PR20.APPLIED) {
        hotelDbV2Pr20Conflict_(duplicateSheet, map, duplicateRow, '重複整理履歴では整理済みですが、元DB行が現在も存在します。履歴と元DBの整合を人が確認してください。');
        result.conflicts++;
        return;
      }

      archiveRow = hotelDbV2Pr20AppendArchive_(
        archiveSheet, candidate, snapshot1, snapshot2, candidate.keepRow, removeRow, archiveKey
      );
      const archiveCheck = hotelDbV2Pr20VerifyArchive_(
        archiveSheet, archiveRow, archiveKey, snapshot1.hash, snapshot2.hash, candidate.keepRow, removeRow
      );
      if (!archiveCheck.ok) throw new Error('アーカイブ検証失敗: ' + archiveCheck.reason);

      sourceSheet.getRange(removeRow, 1, 1, sourceSheet.getLastColumn()).clearContent();
      didClear = true;
      if (!hotelDbV2Pr19RowIsEmpty_(sourceSheet, removeRow)) {
        throw new Error('除外行の内容クリア後検証に失敗しました。');
      }

      hotelDbV2Pr20FinalizeArchive_(archiveSheet, archiveRow, HOTEL_DB_V2_PR20.APPLIED, '重複行を安全整理。物理行削除なし。');
      hotelDbV2Pr20SetResult_(duplicateSheet, map, duplicateRow, HOTEL_DB_V2_PR20.APPLIED, removeRow, archiveKey, '整理済み');
      historySheet.appendRow(hotelDbV2RowFromObject_(HOTEL_DB_V2_HISTORY_HEADERS, {
        '日時': hotelDbV2Timestamp_(), '元シート': candidate.sheetName, '元シートID': candidate.sheetId,
        '元行': removeRow, '施設名': removeRow === candidate.row1 ? candidate.name1 : candidate.name2,
        '処理': '承認重複整理', '結果': '整理済み', 'Place ID': candidate.placeId,
        '一致スコア': candidate.similarity, '営業状態': '',
        '詳細': '残す行=' + candidate.keepRow + '／除外行=' + removeRow + '／重複整理履歴キー=' + archiveKey + '／物理行削除なし'
      }));

      // 関連候補の無効化は二次処理。失敗しても元DB整理を巻き戻さない。
      hotelDbV2Pr20InvalidateLinkedCandidates_(spreadsheet, candidate.sheetId, removeRow, duplicateSheet, duplicateRow);
      result.applied++;
    } catch (error) {
      if (didClear && sourceSheet && removeSnapshot) {
        try {
          hotelDbV2Pr19RestoreRow_(sourceSheet, removeRow, removeSnapshot);
          const restored = hotelDbV2Pr19SnapshotRow_(sourceSheet, removeRow);
          if (restored.hash !== removeSnapshot.hash) throw new Error('ロールバック後ハッシュ不一致');
          if (archiveRow) hotelDbV2Pr20FinalizeArchive_(archiveSheet, archiveRow, 'ロールバック済み', error.message);
        } catch (rollbackError) {
          if (archiveRow) {
            hotelDbV2Pr20FinalizeArchive_(
              archiveSheet, archiveRow, 'ロールバック要確認', error.message + '／' + rollbackError.message
            );
          }
        }
      } else if (archiveRow) {
        try {
          hotelDbV2Pr20FinalizeArchive_(archiveSheet, archiveRow, '整理エラー・元DB不変', error.message);
        } catch (archiveError) {
          console.error('PR20アーカイブ状態更新失敗: ' + archiveError.message);
        }
      }
      hotelDbV2Pr20SetResult_(
        duplicateSheet, map, duplicateRow, HOTEL_DB_V2_PR20.ERROR,
        removeRow, archiveKey, '未整理: ' + error.message
      );
      result.errors++;
    }
  });

  result.reconciliation = result.approved === result.applied + result.conflicts + result.errors ? '一致' : '要確認';
  return result;
}

function hotelDbV2Pr20ValidateSource_(sourceSheet, candidate) {
  if (!sourceSheet) return { ok: false, reason: '元シートが見つかりません。' };
  if (sourceSheet.getName() !== candidate.sheetName) return { ok: false, reason: '元シート名が候補作成時から変更されています。' };
  if (sourceSheet.getName() === HOTEL_DB_V2_PR20_ARCHIVE_SHEET_NAME ||
      (typeof HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME !== 'undefined' && sourceSheet.getName() === HOTEL_DB_V2_PR19_ARCHIVE_SHEET_NAME)) {
    return { ok: false, reason: '履歴シートは元DBとして整理できません。' };
  }
  if (candidate.row1 < 2 || candidate.row2 < 2 || candidate.row1 > sourceSheet.getLastRow() || candidate.row2 > sourceSheet.getLastRow()) {
    return { ok: false, reason: '行1または行2が見つかりません。' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2Pr20ValidatePairIdentity_(candidate, facility1, facility2) {
  if (!facility1 || !facility2 || (!facility1.name && !facility1.address) || (!facility2.name && !facility2.address)) {
    return { ok: false, reason: '行1または行2が空です。' };
  }
  if (hotelDbV2NormalizeText_(facility1.name) !== hotelDbV2NormalizeText_(candidate.name1) ||
      hotelDbV2NormalizeAddress_(facility1.municipality + facility1.address) !== hotelDbV2NormalizeAddress_(candidate.address1)) {
    return { ok: false, reason: '行1の施設名または住所が候補作成時と一致しません。' };
  }
  if (hotelDbV2NormalizeText_(facility2.name) !== hotelDbV2NormalizeText_(candidate.name2) ||
      hotelDbV2NormalizeAddress_(facility2.municipality + facility2.address) !== hotelDbV2NormalizeAddress_(candidate.address2)) {
    return { ok: false, reason: '行2の施設名または住所が候補作成時と一致しません。' };
  }
  if (hotelDbV2NormalizeText_(facility1.name) !== hotelDbV2NormalizeText_(facility2.name)) {
    return { ok: false, reason: '現在の施設名が完全一致ではありません。別施設の可能性があるため整理しません。' };
  }
  if (hotelDbV2NormalizeAddress_(facility1.municipality + facility1.address) !==
      hotelDbV2NormalizeAddress_(facility2.municipality + facility2.address)) {
    return { ok: false, reason: '現在の住所が完全一致ではありません。別部屋・別階・別施設の可能性があるため整理しません。' };
  }
  if (facility1.placeId && facility2.placeId && facility1.placeId !== facility2.placeId) {
    return { ok: false, reason: '両行に異なるPlace IDが保存されています。重複と断定せず整理しません。' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2Pr20NoLossCheck_(sourceMap, keepSnapshot, removeSnapshot) {
  if (!keepSnapshot || !removeSnapshot || keepSnapshot.headers.length !== removeSnapshot.headers.length) {
    return { ok: false, reason: '比較用スナップショットの列構成が一致しません。' };
  }
  for (let i = 0; i < keepSnapshot.headers.length; i++) {
    if (hotelDbV2Clean_(keepSnapshot.headers[i]) !== hotelDbV2Clean_(removeSnapshot.headers[i])) {
      return { ok: false, reason: '比較用スナップショットの見出し構成が一致しません。' };
    }
  }

  const conflicts = [];
  for (let i = 0; i < removeSnapshot.headers.length; i++) {
    const header = hotelDbV2Clean_(removeSnapshot.headers[i]) || ('列' + (i + 1));
    const removeFormula = hotelDbV2Clean_(removeSnapshot.formulas[i]);
    const keepFormula = hotelDbV2Clean_(keepSnapshot.formulas[i]);
    if (removeFormula && removeFormula !== keepFormula) {
      conflicts.push(header + '（除外行だけに異なる数式）');
      continue;
    }

    const removeValue = hotelDbV2Clean_(removeSnapshot.displayValues[i]);
    const keepValue = hotelDbV2Clean_(keepSnapshot.displayValues[i]);
    if (!removeValue) continue;
    if (!keepValue) {
      conflicts.push(header + '（除外行だけに値あり）');
      continue;
    }
    if (!hotelDbV2Pr20CellEquivalent_(i + 1, sourceMap, keepValue, removeValue)) {
      conflicts.push(header + '（両行で値が競合）');
    }
  }

  if (conflicts.length) {
    return {
      ok: false,
      reason: conflicts.slice(0, 8).join('、') + (conflicts.length > 8 ? ' ほか' + (conflicts.length - 8) + '件' : '')
    };
  }
  return { ok: true, reason: '除外行にしかない情報・競合値なし' };
}

function hotelDbV2Pr20CellEquivalent_(column, sourceMap, left, right) {
  if (column === sourceMap.postalCode) {
    return hotelDbV2NormalizePostalCode_(left) === hotelDbV2NormalizePostalCode_(right);
  }
  if (column === sourceMap.municipality || column === sourceMap.facilityName || column === sourceMap.googleName) {
    return hotelDbV2NormalizeText_(left) === hotelDbV2NormalizeText_(right);
  }
  if (column === sourceMap.address || column === sourceMap.googleAddress) {
    return hotelDbV2NormalizeAddress_(left) === hotelDbV2NormalizeAddress_(right);
  }
  return hotelDbV2Clean_(left) === hotelDbV2Clean_(right);
}

function hotelDbV2Pr20ArchiveKey_(sheetId, row1, row2, keepRow, hash1, hash2) {
  return [sheetId, row1, row2, keepRow, hash1, hash2].join('|');
}

function hotelDbV2Pr20AppendArchive_(sheet, candidate, snapshot1, snapshot2, keepRow, removeRow, archiveKey) {
  const keepSnapshot = keepRow === candidate.row1 ? snapshot1 : snapshot2;
  const removeSnapshot = removeRow === candidate.row1 ? snapshot1 : snapshot2;
  const values = hotelDbV2RowFromObject_(HOTEL_DB_V2_PR20_ARCHIVE_HEADERS, {
    '整理キー': archiveKey, '処理状態': 'アーカイブ済み・整理前', 'アーカイブ日時': hotelDbV2Timestamp_(),
    '整理完了日時': '', '元シート': candidate.sheetName, '元シートID': candidate.sheetId, '重複キー': candidate.key,
    '行1': candidate.row1, '施設名1': candidate.name1, '住所1': candidate.address1, '行1ハッシュ': snapshot1.hash,
    '行2': candidate.row2, '施設名2': candidate.name2, '住所2': candidate.address2, '行2ハッシュ': snapshot2.hash,
    '残す行': keepRow, '除外行': removeRow, 'Place ID': candidate.placeId, '類似度': candidate.similarity,
    '推奨判定': candidate.recommendation, '自動判定理由': candidate.triageReason, '信頼度': candidate.confidence,
    '確認日': candidate.checkedAt, '元見出しJSON': JSON.stringify(snapshot1.headers),
    '残す行値JSON': JSON.stringify(keepSnapshot.values), '残す行数式JSON': JSON.stringify(keepSnapshot.formulas),
    '除外行値JSON': JSON.stringify(removeSnapshot.values), '除外行数式JSON': JSON.stringify(removeSnapshot.formulas),
    '情報損失チェック': '合格', '詳細': 'PR20承認重複整理。人が残す行を指定。物理行削除なし。'
  });
  const rowNumber = sheet.getLastRow() + 1;
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
  SpreadsheetApp.flush();
  return rowNumber;
}

function hotelDbV2Pr20VerifyArchive_(sheet, rowNumber, key, hash1, hash2, keepRow, removeRow) {
  const map = hotelDbV2HeaderIndex_(HOTEL_DB_V2_PR20_ARCHIVE_HEADERS);
  const row = sheet.getRange(rowNumber, 1, 1, HOTEL_DB_V2_PR20_ARCHIVE_HEADERS.length).getDisplayValues()[0];
  function value(header) { return hotelDbV2Clean_(row[map[header] - 1]); }
  if (value('整理キー') !== key) return { ok: false, reason: '整理キー不一致' };
  if (value('行1ハッシュ') !== hash1 || value('行2ハッシュ') !== hash2) return { ok: false, reason: '行ハッシュ不一致' };
  if (Number(value('残す行')) !== Number(keepRow) || Number(value('除外行')) !== Number(removeRow)) {
    return { ok: false, reason: '残す行・除外行不一致' };
  }
  return { ok: true, reason: '' };
}

function hotelDbV2Pr20FindArchive_(sheet, key) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (hotelDbV2Clean_(values[i][0]) === key) {
      return { row: i + 2, state: hotelDbV2Clean_(values[i][1]) };
    }
  }
  return null;
}

function hotelDbV2Pr20FinalizeArchive_(sheet, rowNumber, state, detail) {
  const map = hotelDbV2HeaderIndex_(HOTEL_DB_V2_PR20_ARCHIVE_HEADERS);
  sheet.getRange(rowNumber, map['処理状態']).setValue(state);
  if (state === HOTEL_DB_V2_PR20.APPLIED) {
    sheet.getRange(rowNumber, map['整理完了日時']).setValue(hotelDbV2Timestamp_());
  }
  sheet.getRange(rowNumber, map['詳細']).setValue(detail || '');
}

function hotelDbV2Pr20SetResult_(sheet, map, rowNumber, state, removeRow, archiveKey, detail) {
  sheet.getRange(rowNumber, map['状態']).setValue(state);
  sheet.getRange(rowNumber, map['除外予定行']).setValue(removeRow || '');
  sheet.getRange(rowNumber, map['整理処理日時']).setValue(hotelDbV2Timestamp_());
  sheet.getRange(rowNumber, map['整理履歴キー']).setValue(archiveKey || '');
  sheet.getRange(rowNumber, map['整理結果']).setValue(detail || '');
}

function hotelDbV2Pr20Conflict_(sheet, map, rowNumber, reason) {
  hotelDbV2Pr20SetResult_(
    sheet, map, rowNumber, HOTEL_DB_V2_PR20.CONFLICT, '', '', '未整理: ' + reason
  );
}

function hotelDbV2Pr20PrepareTriageSnapshot_(spreadsheet, duplicateSheet, rowNumber, map, decision) {
  const result = { created: 0, review: 0, approvalReset: 0 };
  if (!map['整理スナップショット状態']) return result;

  function set(header, value) {
    if (map[header]) duplicateSheet.getRange(rowNumber, map[header]).setValue(value);
  }
  const row = duplicateSheet.getRange(rowNumber, 1, 1, duplicateSheet.getLastColumn()).getDisplayValues()[0];
  function value(header) { return map[header] ? hotelDbV2Clean_(row[map[header] - 1]) : ''; }
  const state = value('状態');

  if (!decision || decision.recommendation !== HOTEL_DB_V2_PR20.STRONG) {
    set('整理スナップショット状態', HOTEL_DB_V2_PR20.SNAPSHOT_OUT);
    return result;
  }

  try {
    const sheetId = Number(value('元シートID'));
    const row1 = Number(value('行1'));
    const row2 = Number(value('行2'));
    const sourceSheet = sheetId ? spreadsheet.getSheetById(sheetId) : null;
    if (!sourceSheet || sourceSheet.getName() !== value('元シート') ||
        row1 < 2 || row2 < 2 || row1 > sourceSheet.getLastRow() || row2 > sourceSheet.getLastRow()) {
      throw new Error('元シートまたは行1/行2を確認できません。');
    }

    const sourceMap = hotelDbV2GetHeaderMap_(sourceSheet);
    hotelDbV2ValidateSourceSheet_(sourceSheet, sourceMap);
    const facility1 = hotelDbV2ReadFacility_(sourceSheet, row1, sourceMap);
    const facility2 = hotelDbV2ReadFacility_(sourceSheet, row2, sourceMap);
    const identity = hotelDbV2Pr20ValidatePairIdentity_({
      name1: value('施設名1'), address1: value('住所1'),
      name2: value('施設名2'), address2: value('住所2')
    }, facility1, facility2);
    if (!identity.ok) throw new Error(identity.reason);

    const snapshot1 = hotelDbV2Pr19SnapshotRow_(sourceSheet, row1);
    const snapshot2 = hotelDbV2Pr19SnapshotRow_(sourceSheet, row2);
    const oldHash1 = value('行1ハッシュ');
    const oldHash2 = value('行2ハッシュ');
    const changed = (oldHash1 && oldHash1 !== snapshot1.hash) || (oldHash2 && oldHash2 !== snapshot2.hash);
    const missingOld = !oldHash1 || !oldHash2;

    if (state === HOTEL_DB_V2_PR20.APPLIED) return result;

    set('整理候補作成日時', hotelDbV2Timestamp_());
    set('行1ハッシュ', snapshot1.hash);
    set('行2ハッシュ', snapshot2.hash);
    set('整理スナップショット状態', HOTEL_DB_V2_PR20.SNAPSHOT_READY);
    set('除外予定行', '');
    set('整理処理日時', '');
    set('整理履歴キー', '');

    if (state === HOTEL_DB_V2_PR20.APPROVED && (changed || missingOld)) {
      set('状態', HOTEL_DB_V2_PR20.CONFLICT);
      set('整理結果', 'スナップショットが新規作成または更新されたため、内容を再確認して再承認してください。');
      result.approvalReset++;
    } else if (state !== HOTEL_DB_V2_PR20.CONFLICT) {
      set('整理結果', '');
    }
    result.created++;
    return result;
  } catch (error) {
    set('整理スナップショット状態', HOTEL_DB_V2_PR20.SNAPSHOT_REVIEW);
    set('整理結果', 'スナップショット未作成: ' + error.message);
    result.review++;
    return result;
  }
}

function hotelDbV2Pr20InvalidateLinkedCandidates_(spreadsheet, sheetId, sourceRow, currentDuplicateSheet, currentDuplicateRow) {
  try {
    function invalidateByColumns(sheetName, idHeader, rowHeaders, completedStates, auditResultHeader) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;
      const map = hotelDbV2Pr20HeaderMap_(sheet);
      if (!map[idHeader] || !map['状態']) return;
      const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getDisplayValues();
      values.forEach(function(row, offset) {
        const rowNumber = offset + 2;
        if (sheet === currentDuplicateSheet && rowNumber === currentDuplicateRow) return;
        if (String(row[map[idHeader] - 1]) !== String(sheetId)) return;
        const matchesRow = rowHeaders.some(function(header) {
          return map[header] && Number(row[map[header] - 1]) === Number(sourceRow);
        });
        if (!matchesRow) return;
        const state = hotelDbV2Clean_(row[map['状態'] - 1]);
        if (completedStates.indexOf(state) !== -1) return;
        sheet.getRange(rowNumber, map['状態']).setValue('要再確認');
        if (auditResultHeader && map[auditResultHeader]) {
          const currentText = hotelDbV2Clean_(row[map[auditResultHeader] - 1]);
          const marker = '元施設がPR20で重複整理除外済みです。';
          sheet.getRange(rowNumber, map[auditResultHeader]).setValue(
            currentText ? currentText + '／' + marker : marker
          );
        }
      });
    }

    // 既存の差分・詳細は上書きしない。専用の結果列があるシートだけ追記する。
    invalidateByColumns(HOTEL_DB_V2_CONFIG.SHEETS.CORRECTIONS, '元シートID', ['元行'], ['反映済み'], '');
    invalidateByColumns(HOTEL_DB_V2_CONFIG.SHEETS.REVIEW, '元シートID', ['元行'], ['除外済み'], '');
    invalidateByColumns(HOTEL_DB_V2_CONFIG.SHEETS.DUPLICATES, '元シートID', ['行1', '行2'], ['整理済み'], '整理結果');
    if (typeof HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION !== 'undefined') {
      invalidateByColumns(
        HOTEL_DB_V2_NEW_FACILITY_CLASSIFICATION.SHEET_NAME,
        '元シートID', ['元行'], ['反映済み'], '反映結果'
      );
    }
  } catch (error) {
    console.warn('PR20関連候補の要再確認化を一部省略: ' + error.message);
  }
}

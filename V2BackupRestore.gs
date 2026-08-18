/**
 * PR #24 バックアップ・復元。
 *
 * 安全原則:
 * - 現在のスプレッドシートを自動上書きしない。
 * - バックアップはDrive上の完全コピーとして作成する。
 * - 復元は別の「復元候補」スプレッドシートを作るだけにする。
 * - コピー前後の値・数式フィンガープリントを比較し、検証失敗時は採用しない。
 * - Google Places APIは呼ばない。
 * - APIキー本体をバックアップ管理メタデータ・画面・ログへ返さない。
 */

const HOTEL_DB_V2_BACKUP = Object.freeze({
  VERSION: '2.0',
  DIALOG_FILE: 'V2BackupRestoreDialog',
  DIALOG_TITLE: '宿泊施設DB Ver2.0 バックアップ・復元',
  DIALOG_WIDTH: 860,
  DIALOG_HEIGHT: 760,
  LOCK_TIMEOUT_MS: 10000,
  FOLDER_PROPERTY: 'HOTEL_DB_V2_BACKUP_FOLDER_ID',
  FOLDER_NAME: '宿泊施設DB_Backups',
  BACKUP_PREFIX: '【宿泊施設DBバックアップ】',
  RECOVERY_PREFIX: '【復元候補】',
  META_PREFIX: 'HOTEL_DB_V2_BACKUP_META:',
  META_MARKER: 'HOTEL_DB_V2_BACKUP_V1',
  RECOVERY_MARKER: 'HOTEL_DB_V2_RECOVERY_V1',
  MAX_LIST: 20,
  HASH_BLOCK_ROWS: 250,
  SAFETY_MARKER: 'no-in-place-restore; verified-copy; no-places-api; api-key-not-returned'
});

function runHotelDbV2OpenBackupRestore() {
  const html = HtmlService
    .createHtmlOutputFromFile(HOTEL_DB_V2_BACKUP.DIALOG_FILE)
    .setWidth(HOTEL_DB_V2_BACKUP.DIALOG_WIDTH)
    .setHeight(HOTEL_DB_V2_BACKUP.DIALOG_HEIGHT);
  SpreadsheetApp.getUi().showModalDialog(html, HOTEL_DB_V2_BACKUP.DIALOG_TITLE);
}

function hotelDbV2BackupGetStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return hotelDbV2BackupBuildStatus_(ss);
}

function hotelDbV2BackupCreate() {
  return hotelDbV2BackupWithLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    hotelDbV2BackupAssertSourceFile_(ss);
    const folder = hotelDbV2BackupResolveFolder_(true);
    const created = hotelDbV2BackupCreateInFolder_(ss, folder, {});
    return {
      created: true,
      backup: created,
      status: hotelDbV2BackupBuildStatus_(ss)
    };
  });
}

function hotelDbV2BackupCreateRecovery(backupFileId) {
  return hotelDbV2BackupWithLock_(function() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    hotelDbV2BackupAssertSourceFile_(ss);
    const folder = hotelDbV2BackupResolveFolder_(false);
    if (!folder) throw new Error('バックアップ保存先がまだ作成されていません。先にバックアップを作成してください。');
    const backupFile = hotelDbV2BackupFindFileInFolder_(folder, backupFileId);
    if (!backupFile) throw new Error('指定したバックアップが管理フォルダ内に見つかりません。');
    const destination = hotelDbV2BackupResolveRecoveryDestination_(ss, folder);
    const recovered = hotelDbV2BackupCreateRecoveryFromFile_(ss, backupFile, destination, {});
    return {
      created: true,
      recovery: recovered,
      status: hotelDbV2BackupBuildStatus_(ss)
    };
  });
}

function hotelDbV2BackupWithLock_(callback) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(HOTEL_DB_V2_BACKUP.LOCK_TIMEOUT_MS)) {
    throw new Error('別の処理が実行中です。完了後にもう一度実行してください。');
  }
  try {
    return callback();
  } catch (error) {
    throw new Error(hotelDbV2BackupSafeErrorMessage_(error));
  } finally {
    lock.releaseLock();
  }
}

function hotelDbV2BackupBuildStatus_(ss) {
  hotelDbV2BackupAssertSpreadsheet_(ss);
  const folder = hotelDbV2BackupResolveFolder_(false);
  const backups = folder ? hotelDbV2BackupListInFolder_(folder, ss.getId()) : [];
  return {
    version: HOTEL_DB_V2_BACKUP.VERSION,
    source: { id: ss.getId(), name: ss.getName(), url: ss.getUrl() },
    folderConfigured: Boolean(folder),
    folderId: folder ? folder.getId() : null,
    folderName: folder ? folder.getName() : HOTEL_DB_V2_BACKUP.FOLDER_NAME,
    folderUrl: folder ? folder.getUrl() : null,
    backups: backups,
    backupCount: backups.length,
    inPlaceRestore: false,
    externalApiCalled: false,
    apiKeyValue: null,
    apiKeyHash: null,
    safetyMarker: HOTEL_DB_V2_BACKUP.SAFETY_MARKER
  };
}

function hotelDbV2BackupResolveFolder_(createIfMissing) {
  const store = PropertiesService.getScriptProperties();
  const savedId = hotelDbV2Clean_(store.getProperty(HOTEL_DB_V2_BACKUP.FOLDER_PROPERTY));
  if (savedId) {
    try {
      const saved = DriveApp.getFolderById(savedId);
      if (!saved.isTrashed()) return saved;
    } catch (ignore) {}
  }
  if (!createIfMissing) return null;
  const folder = DriveApp.createFolder(HOTEL_DB_V2_BACKUP.FOLDER_NAME);
  store.setProperty(HOTEL_DB_V2_BACKUP.FOLDER_PROPERTY, folder.getId());
  return folder;
}

function hotelDbV2BackupCreateInFolder_(ss, folder, options) {
  const opts = options || {};
  hotelDbV2BackupAssertSpreadsheet_(ss);
  if (!folder) throw new Error('バックアップ保存先フォルダを取得できません。');
  SpreadsheetApp.flush();
  const sourceFile = DriveApp.getFileById(ss.getId());
  const createdAt = new Date();
  const name = hotelDbV2BackupBuildBackupName_(ss.getName(), createdAt, opts.nameSuffix || '');
  let backupFile = null;
  try {
    backupFile = sourceFile.makeCopy(name, folder);
    const backupSs = SpreadsheetApp.openById(backupFile.getId());
    const backupFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(backupSs);
    const sourceFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(ss);
    if (backupFingerprint.fingerprint !== sourceFingerprint.fingerprint) {
      throw new Error('バックアップコピーと元データの検証結果が一致しません。編集を止めてから再実行してください。');
    }
    const metadata = hotelDbV2BackupBuildMetadata_({
      sourceSpreadsheetId: ss.getId(), sourceName: ss.getName(), createdAt: createdAt,
      fingerprint: backupFingerprint.fingerprint, sheetCount: backupFingerprint.sheetCount,
      usedCells: backupFingerprint.usedCells
    });
    backupFile.setDescription(HOTEL_DB_V2_BACKUP.META_PREFIX + JSON.stringify(metadata));
    return hotelDbV2BackupFileSummary_(backupFile, metadata);
  } catch (error) {
    if (backupFile) {
      try { backupFile.setTrashed(true); } catch (cleanupError) {
        console.error('PR24 invalid backup cleanup failed: ' + cleanupError.message);
      }
    }
    throw error;
  }
}

function hotelDbV2BackupCreateRecoveryFromFile_(sourceSs, backupFile, destinationFolder, options) {
  const opts = options || {};
  hotelDbV2BackupAssertSpreadsheet_(sourceSs);
  if (!backupFile) throw new Error('バックアップファイルを取得できません。');
  if (!destinationFolder) throw new Error('復元候補の保存先を取得できません。');
  const metadata = hotelDbV2BackupReadMetadata_(backupFile);
  const validation = hotelDbV2BackupValidateMetadata_(metadata, sourceSs.getId());
  if (!validation.valid) throw new Error(validation.message);
  const backupSs = SpreadsheetApp.openById(backupFile.getId());
  const backupFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(backupSs);
  if (backupFingerprint.fingerprint !== metadata.fingerprint) {
    throw new Error('バックアップ作成後に内容が変更されています。安全のため復元候補を作成しません。');
  }
  const createdAt = new Date();
  const recoveryName = hotelDbV2BackupBuildRecoveryName_(metadata.sourceName, createdAt, opts.nameSuffix || '');
  let recoveryFile = null;
  try {
    recoveryFile = backupFile.makeCopy(recoveryName, destinationFolder);
    const recoverySs = SpreadsheetApp.openById(recoveryFile.getId());
    const recoveryFingerprint = hotelDbV2BackupFingerprintSpreadsheet_(recoverySs);
    if (recoveryFingerprint.fingerprint !== metadata.fingerprint) {
      throw new Error('復元候補コピーの検証結果がバックアップと一致しません。');
    }
    const recoveryMetadata = {
      marker: HOTEL_DB_V2_BACKUP.RECOVERY_MARKER,
      version: HOTEL_DB_V2_BACKUP.VERSION,
      sourceSpreadsheetId: sourceSs.getId(),
      sourceName: metadata.sourceName,
      backupFileId: backupFile.getId(),
      createdAt: createdAt.toISOString(),
      fingerprint: metadata.fingerprint,
      sheetCount: metadata.sheetCount,
      usedCells: metadata.usedCells
    };
    recoveryFile.setDescription(HOTEL_DB_V2_BACKUP.META_PREFIX + JSON.stringify(recoveryMetadata));
    return {
      id: recoveryFile.getId(), name: recoveryFile.getName(), url: recoveryFile.getUrl(),
      createdAt: recoveryMetadata.createdAt, backupFileId: backupFile.getId(),
      fingerprint: hotelDbV2BackupShortFingerprint_(metadata.fingerprint), verified: true,
      currentFileOverwritten: false
    };
  } catch (error) {
    if (recoveryFile) {
      try { recoveryFile.setTrashed(true); } catch (cleanupError) {
        console.error('PR24 invalid recovery cleanup failed: ' + cleanupError.message);
      }
    }
    throw error;
  }
}

function hotelDbV2BackupResolveRecoveryDestination_(ss, fallbackFolder) {
  try {
    const sourceFile = DriveApp.getFileById(ss.getId());
    const parents = sourceFile.getParents();
    if (parents.hasNext()) return parents.next();
  } catch (ignore) {}
  return fallbackFolder || DriveApp.getRootFolder();
}

function hotelDbV2BackupListInFolder_(folder, sourceSpreadsheetId) {
  const items = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (file.isTrashed()) continue;
    const metadata = hotelDbV2BackupReadMetadata_(file);
    const validation = hotelDbV2BackupValidateMetadata_(metadata, sourceSpreadsheetId);
    if (!validation.valid) continue;
    items.push(hotelDbV2BackupFileSummary_(file, metadata));
  }
  items.sort(function(a, b) { return String(b.createdAt).localeCompare(String(a.createdAt)); });
  return items.slice(0, HOTEL_DB_V2_BACKUP.MAX_LIST);
}

function hotelDbV2BackupFindFileInFolder_(folder, fileId) {
  const target = hotelDbV2Clean_(fileId);
  if (!target) return null;
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    if (!file.isTrashed() && file.getId() === target) return file;
  }
  return null;
}

function hotelDbV2BackupBuildMetadata_(input) {
  const data = input || {};
  return {
    marker: HOTEL_DB_V2_BACKUP.META_MARKER,
    version: HOTEL_DB_V2_BACKUP.VERSION,
    sourceSpreadsheetId: hotelDbV2Clean_(data.sourceSpreadsheetId),
    sourceName: hotelDbV2Clean_(data.sourceName),
    createdAt: data.createdAt instanceof Date ? data.createdAt.toISOString() : hotelDbV2Clean_(data.createdAt),
    fingerprint: hotelDbV2Clean_(data.fingerprint),
    sheetCount: Number(data.sheetCount || 0), usedCells: Number(data.usedCells || 0), verified: true
  };
}

function hotelDbV2BackupReadMetadata_(file) {
  if (!file) return null;
  let description = '';
  try { description = hotelDbV2Clean_(file.getDescription()); } catch (ignore) { return null; }
  if (description.indexOf(HOTEL_DB_V2_BACKUP.META_PREFIX) !== 0) return null;
  const json = description.slice(HOTEL_DB_V2_BACKUP.META_PREFIX.length);
  try {
    const parsed = JSON.parse(json);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (ignore) { return null; }
}

function hotelDbV2BackupValidateMetadata_(metadata, expectedSourceId) {
  const data = metadata || {};
  if (data.marker !== HOTEL_DB_V2_BACKUP.META_MARKER) return { valid:false, message:'管理対象のバックアップではありません。' };
  if (hotelDbV2Clean_(data.version) !== HOTEL_DB_V2_BACKUP.VERSION) return { valid:false, message:'バックアップ形式のバージョンが一致しません。' };
  if (!hotelDbV2Clean_(data.sourceSpreadsheetId)) return { valid:false, message:'元スプレッドシートIDが記録されていません。' };
  if (expectedSourceId && data.sourceSpreadsheetId !== expectedSourceId) return { valid:false, message:'このスプレッドシート用のバックアップではありません。' };
  if (!/^[0-9a-f]{64}$/.test(hotelDbV2Clean_(data.fingerprint))) return { valid:false, message:'バックアップ検証用フィンガープリントが不正です。' };
  if (Number(data.sheetCount || 0) < 1) return { valid:false, message:'バックアップのシート数が不正です。' };
  if (!data.verified) return { valid:false, message:'未検証のバックアップです。' };
  return { valid:true, message:'' };
}

function hotelDbV2BackupFileSummary_(file, metadata) {
  return {
    id: file.getId(), name: file.getName(), url: file.getUrl(),
    createdAt: hotelDbV2Clean_(metadata.createdAt), sourceName: hotelDbV2Clean_(metadata.sourceName),
    sheetCount: Number(metadata.sheetCount || 0), usedCells: Number(metadata.usedCells || 0),
    fingerprint: hotelDbV2BackupShortFingerprint_(metadata.fingerprint), verified: true
  };
}

function hotelDbV2BackupBuildBackupName_(sourceName, date, suffix) {
  const stamp = hotelDbV2BackupTimestamp_(date);
  const tail = hotelDbV2Clean_(suffix);
  return HOTEL_DB_V2_BACKUP.BACKUP_PREFIX + stamp + '_' + hotelDbV2BackupSafeName_(sourceName) + (tail ? '_' + tail : '');
}

function hotelDbV2BackupBuildRecoveryName_(sourceName, date, suffix) {
  const stamp = hotelDbV2BackupTimestamp_(date);
  const tail = hotelDbV2Clean_(suffix);
  return HOTEL_DB_V2_BACKUP.RECOVERY_PREFIX + stamp + '_' + hotelDbV2BackupSafeName_(sourceName) + (tail ? '_' + tail : '');
}

function hotelDbV2BackupTimestamp_(date) {
  const value = date instanceof Date ? date : new Date(date);
  return Utilities.formatDate(value, HOTEL_DB_V2_CONFIG.TIMEZONE || 'Asia/Tokyo', 'yyyy-MM-dd_HH-mm-ss');
}

function hotelDbV2BackupSafeName_(value) {
  return hotelDbV2Clean_(value).replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').slice(0, 80) || '宿泊施設DB';
}

function hotelDbV2BackupShortFingerprint_(value) {
  const text = hotelDbV2Clean_(value);
  return text ? text.slice(0, 12) : '';
}

function hotelDbV2BackupFingerprintSpreadsheet_(ss) {
  hotelDbV2BackupAssertSpreadsheet_(ss);
  const sheets = ss.getSheets();
  const sheetSummaries = [];
  let usedCells = 0;
  sheets.forEach(function(sheet, index) {
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    const blockHashes = [];
    usedCells += Math.max(0, lastRow) * Math.max(0, lastColumn);
    if (lastRow > 0 && lastColumn > 0) {
      for (let startRow = 1; startRow <= lastRow; startRow += HOTEL_DB_V2_BACKUP.HASH_BLOCK_ROWS) {
        const rowCount = Math.min(HOTEL_DB_V2_BACKUP.HASH_BLOCK_ROWS, lastRow - startRow + 1);
        const range = sheet.getRange(startRow, 1, rowCount, lastColumn);
        const values = range.getValues();
        const formulas = range.getFormulas();
        const tokens = values.map(function(row, r) {
          return row.map(function(value, c) {
            const formula = formulas[r][c];
            return formula ? 'F:' + formula : 'V:' + hotelDbV2BackupSerializeValue_(value);
          });
        });
        blockHashes.push(hotelDbV2BackupDigest_(JSON.stringify(tokens)));
      }
    }
    const sheetPayload = {
      index:index, name:sheet.getName(), rows:lastRow, columns:lastColumn,
      hidden:sheet.isSheetHidden(), frozenRows:sheet.getFrozenRows(), frozenColumns:sheet.getFrozenColumns(), blocks:blockHashes
    };
    sheetSummaries.push({
      index:index, name:sheet.getName(), rows:lastRow, columns:lastColumn,
      hash:hotelDbV2BackupDigest_(JSON.stringify(sheetPayload))
    });
  });
  return {
    fingerprint: hotelDbV2BackupDigest_(JSON.stringify(sheetSummaries)),
    sheetCount: sheets.length, usedCells: usedCells, sheets: sheetSummaries
  };
}

function hotelDbV2BackupSerializeValue_(value) {
  if (value instanceof Date) return 'date:' + value.toISOString();
  if (value === null || value === undefined) return 'null:';
  const type = typeof value;
  if (type === 'number') return 'number:' + String(value);
  if (type === 'boolean') return 'boolean:' + (value ? '1' : '0');
  return 'string:' + String(value);
}

function hotelDbV2BackupDigest_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text || ''), Utilities.Charset.UTF_8);
  return bytes.map(function(value) {
    const normalized = value < 0 ? value + 256 : value;
    return ('0' + normalized.toString(16)).slice(-2);
  }).join('');
}

function hotelDbV2BackupAssertSpreadsheet_(ss) {
  if (!ss || typeof ss.getId !== 'function' || !hotelDbV2Clean_(ss.getId())) throw new Error('対象スプレッドシートを取得できません。');
}

function hotelDbV2BackupAssertSourceFile_(ss) {
  hotelDbV2BackupAssertSpreadsheet_(ss);
  try {
    const file = DriveApp.getFileById(ss.getId());
    const metadata = hotelDbV2BackupReadMetadata_(file);
    if (metadata && (metadata.marker === HOTEL_DB_V2_BACKUP.META_MARKER || metadata.marker === HOTEL_DB_V2_BACKUP.RECOVERY_MARKER)) {
      throw new Error('バックアップまたは復元候補のコピーからは新しいバックアップを作成できません。元の運用ファイルを開いてください。');
    }
  } catch (error) {
    if (error && String(error.message || '').indexOf('バックアップまたは復元候補') !== -1) throw error;
  }
}

function hotelDbV2BackupSafeErrorMessage_(error) {
  let message = hotelDbV2Clean_(error && error.message ? error.message : error);
  if (!message) message = '不明なエラーが発生しました。';
  try {
    const key = hotelDbV2Clean_(PropertiesService.getScriptProperties().getProperty(HOTEL_DB_V2_CONFIG.API_KEY_PROPERTY));
    if (key) message = message.split(key).join('[API_KEY]');
  } catch (ignore) {}
  return message;
}

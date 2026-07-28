const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const importUrlVectors = require('./fixtures/import-url-vectors');

const codeJs = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const readme = fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8');
const RECEIPT_HEADERS = [
  'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
  'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
  'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId',
  'mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId'
];
const MAP_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧', '状態', 'タグ',
  'イベント時刻', '更新時刻', 'アイコン', '音声ID'
];

function columnNumber(label) {
  return label.split('').reduce((total, character) => total * 26 + character.charCodeAt(0) - 64, 0);
}

function parseA1(value) {
  const range = String(value).match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
  if (range) {
    const row = Number(range[2]);
    const column = columnNumber(range[1]);
    const endRow = Number(range[4] || range[2]);
    const endColumn = columnNumber(range[3] || range[1]);
    return { row, column, numRows: endRow - row + 1, numColumns: endColumn - column + 1 };
  }
  const wholeColumn = String(value).match(/^([A-Z]+):([A-Z]+)$/);
  if (wholeColumn) {
    const column = columnNumber(wholeColumn[1]);
    return { row: 1, column, numRows: 1000, numColumns: columnNumber(wholeColumn[2]) - column + 1 };
  }
  throw new Error(`Unsupported A1 range: ${value}`);
}

function lastUsedRow(rows) {
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if ((rows[index] || []).some((value) => value !== '' && value != null)) return index + 1;
  }
  return 0;
}

function makeHarness(options = {}) {
  const audit = {
    sheetLookups: [], inserts: [], columnInserts: [], reads: [], writes: [], alerts: [],
    errors: [], events: [],
    uuidCalls: 0,
    locks: { attempts: 0, flushes: 0, releases: 0, held: false, nestedAttempts: 0 },
    drive: {
      creates: 0, guideCreates: 0, renames: 0, shares: 0, folderGets: 0, fileGets: 0,
      searches: 0, trashes: 0, moves: 0, folderCreates: 0, mediaFolderCreates: 0,
      callsWhileLocked: 0,
      renameIds: [], shareIds: [], trashIds: [], moveIds: []
    }
  };
  const sheets = new Map();
  const folders = new Map();
  const filesById = new Map();
  let uuid = 0;
  let mapAppendFailed = false;
  let sharingFailed = false;
  let fileGetFailed = false;
  let receiptFileSavedWriteFailed = false;
  let createAfterFileFailed = false;
  let mapAttachUpdateFailed = false;
  let mapAttachTimestampFailed = false;

  function makeSheet(name, initialRows = []) {
    const rows = initialRows.map((row) => row.slice());
    const formulas = (options.sheetFormulas && options.sheetFormulas[name]
      ? options.sheetFormulas[name]
      : initialRows.map((row) => row.map(() => ''))).map((row) => row.slice());
    let maxRows = Math.max(1000, rows.length);
    let maxColumns = options.sheetMaxColumns && options.sheetMaxColumns[name] != null
      ? Number(options.sheetMaxColumns[name])
      : Math.max(26, rows.reduce((max, row) => Math.max(max, row.length), 0));
    const sheet = {
      name,
      rows,
      formulas,
      getLastRow() { return lastUsedRow(rows); },
      getLastColumn() { return rows.reduce((max, row) => Math.max(max, row.length), 0); },
      getMaxRows() { return maxRows; },
      getMaxColumns() { return maxColumns; },
      insertRowsAfter(_after, count) { maxRows += count; },
      insertColumnsAfter(after, count) {
        if (after < 1 || after > maxColumns || count < 1) throw new Error('Invalid column insertion');
        audit.columnInserts.push({ sheet: name, after, count, lockHeld: audit.locks.held });
        maxColumns += count;
      },
      insertRowBefore(rowNumber) {
        rows.splice(rowNumber - 1, 0, []);
        formulas.splice(rowNumber - 1, 0, []);
      },
      appendRow(row) {
        if (name === 'map_info' && options.failMapAppendOnce && !mapAppendFailed) {
          mapAppendFailed = true;
          throw new Error('map append failed');
        }
        rows.push(row.slice());
        formulas.push(row.map(() => ''));
        audit.events.push({ type: 'sheet-append', sheet: name });
        audit.writes.push({ sheet: name, method: 'appendRow', values: row.slice(), lockHeld: audit.locks.held });
        if (name === 'import_receipts' && typeof options.afterReceiptAppend === 'function') {
          options.afterReceiptAppend(rows[rows.length - 1]);
        }
      },
      getDataRange() {
        return sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), Math.max(1, sheet.getLastColumn()));
      },
      getRange(rowOrA1, column, numRows = 1, numColumns = 1) {
        const info = typeof rowOrA1 === 'string'
          ? parseA1(rowOrA1)
          : { row: rowOrA1, column, numRows, numColumns };
        if (info.column + info.numColumns - 1 > maxColumns) {
          throw new Error(`Range exceeds physical columns on ${name}`);
        }
        const range = {
          getValues() {
            audit.reads.push({ sheet: name, ...info, lockHeld: audit.locks.held });
            return Array.from({ length: info.numRows }, (_unused, rowOffset) =>
              Array.from({ length: info.numColumns }, (_unusedColumn, columnOffset) =>
                (rows[info.row - 1 + rowOffset] || [])[info.column - 1 + columnOffset] ?? ''
              )
            );
          },
          getValue() { return range.getValues()[0][0]; },
          getFormulas() {
            return Array.from({ length: info.numRows }, (_unused, rowOffset) =>
              Array.from({ length: info.numColumns }, (_unusedColumn, columnOffset) =>
                (formulas[info.row - 1 + rowOffset] || [])[info.column - 1 + columnOffset] || ''
              )
            );
          },
          setValues(values) {
            if (name === 'map_info'
                && info.column === 7
                && info.numColumns === 2
                && options.failMapAttachUpdateOnce
                && !mapAttachUpdateFailed) {
              mapAttachUpdateFailed = true;
              throw new Error('pin photo attach update failed');
            }
            if (name === 'map_info'
                && info.column === 14
                && info.numColumns === 1
                && options.failMapAttachTimestampOnce
                && !mapAttachTimestampFailed) {
              mapAttachTimestampFailed = true;
              throw new Error('pin photo attach timestamp failed');
            }
            if (name === 'import_receipts'
                && options.failReceiptFileSavedWriteOnce
                && !receiptFileSavedWriteFailed
                && values[0]
                && values[0][receiptColumn('state')] === 'file_saved') {
              receiptFileSavedWriteFailed = true;
              throw new Error('receipt file metadata write failed');
            }
            if (name === 'import_receipts'
                && options.failReceiptCompleteWriteOnce
                && !options.__receiptCompleteWriteFailed
                && values[0]
                && values[0][receiptColumn('state')] === 'completed') {
              options.__receiptCompleteWriteFailed = true;
              throw new Error('receipt completion write failed');
            }
            if (name === 'import_receipts'
                && options.failSourceMoveJournalWrites
                && values[0]
                && values[0][receiptColumn('state')] === 'completed'
                && values[0][receiptColumn('lastErrorCode')]) {
              throw new Error('source move journal write failed');
            }
            audit.writes.push({
              sheet: name, method: 'setValues', ...info,
              values: values.map((row) => row.slice()), lockHeld: audit.locks.held
            });
            values.forEach((sourceRow, rowOffset) => {
              const targetRow = info.row - 1 + rowOffset;
              while (rows.length <= targetRow) {
                rows.push([]);
                formulas.push([]);
              }
              sourceRow.forEach((value, columnOffset) => {
                const targetColumn = info.column - 1 + columnOffset;
                while (rows[targetRow].length <= targetColumn) rows[targetRow].push('');
                while (formulas[targetRow].length <= targetColumn) formulas[targetRow].push('');
                rows[targetRow][targetColumn] = value;
                formulas[targetRow][targetColumn] = '';
              });
            });
            return range;
          },
          setValue(value) { return range.setValues([[value]]); },
          setBackground() { return range; },
          setFontColor() { return range; },
          setFontWeight() { return range; },
          setNumberFormat() { return range; },
          setShowHyperlink() { return range; },
          activate() { return range; }
        };
        return range;
      },
      setFrozenRows() {},
      setColumnWidth() {},
      activate() {}
    };
    sheets.set(name, sheet);
    return sheet;
  }

  function makeFile(folder, id, name, metadata = {}) {
    const file = {
      id,
      name,
      trashed: metadata.trashed === true,
      parentFolders: [folder],
      getId() { return id; },
      getName() { return file.name; },
      getMimeType() { return metadata.mimeType || 'image/jpeg'; },
      getSize() { return metadata.sizeBytes == null ? 3 : metadata.sizeBytes; },
      getLastUpdated() { return new Date('2026-07-11T12:00:00.000Z'); },
      getBlob() { return { getBytes: () => (metadata.bytes || [0xff, 0xd8, 0xff]).slice() }; },
      getSharingAccess() {
        metadata.__sharingReadCalls = Number(metadata.__sharingReadCalls || 0) + 1;
        if (metadata.failSharingReadAt === metadata.__sharingReadCalls) {
          throw new Error('Service unavailable temporarily');
        }
        if (metadata.failSharingReadOnce && !metadata.__sharingReadFailed) {
          metadata.__sharingReadFailed = true;
          throw new Error('Service unavailable temporarily');
        }
        if (metadata.failSharingRead) throw new Error('private sharing read failure for photo_SECRET123456');
        return metadata.sharingAccess || (metadata.managed ? 'PRIVATE' : 'ANYONE_WITH_LINK');
      },
      getSharingPermission() {
        return metadata.sharingPermission || (metadata.managed ? 'NONE' : 'VIEW');
      },
      isTrashed() { return file.trashed; },
      setTrashed(value) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.trashes += 1;
        audit.drive.trashIds.push(id);
        if (metadata.failTrashAlways) throw new Error('private trash failure');
        if (metadata.failTrashOnce && !metadata.__trashFailed) {
          metadata.__trashFailed = true;
          throw new Error('private trash failure');
        }
        file.trashed = !!value;
        return file;
      },
      setName(next) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.renames += 1;
        audit.drive.renameIds.push(id);
        if (metadata.failRenameAlways) throw new Error('private rename failure');
        file.name = String(next);
        return file;
      },
      setSharing(access, permission) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.shares += 1;
        audit.drive.shareIds.push(id);
        if (metadata.failSharingAlways || options.failSharingAlways) {
          throw new Error(options.sharingErrorMessage
            || 'Workspace policy denied sharing for photo_SECRET123456 https://drive.google.com/private');
        }
        if (options.failSharingOnce && !sharingFailed) {
          sharingFailed = true;
          throw new Error('sharing failed');
        }
        metadata.sharingAccess = access;
        metadata.sharingPermission = permission;
        return file;
      },
      moveTo(targetFolder) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.moves += 1;
        audit.drive.moveIds.push(id);
        audit.events.push({ type: 'drive-move', id, targetFolderId: targetFolder.id });
        if (metadata.failMoveAlways) throw new Error('private source move failure');
        if (metadata.failMoveOnce && !metadata.__moveFailed) {
          metadata.__moveFailed = true;
          throw new Error('private transient source move failure');
        }
        if (!metadata.moveWithoutParentChange) {
          file.parentFolders.forEach((parent) => {
            parent.files = parent.files.filter((candidate) => candidate !== file);
          });
          file.parentFolders = [targetFolder];
          if (!targetFolder.files.includes(file)) targetFolder.files.push(file);
        }
        return file;
      },
      getParents() {
        let index = 0;
        const parents = file.parentFolders.slice();
        return { hasNext: () => index < parents.length, next() { return parents[index++]; } };
      }
    };
    folder.files.push(file);
    filesById.set(id, file);
    return file;
  }

  function makeFolder(id, metadata = {}) {
    const folder = {
      id,
      name: metadata.name || id,
      files: [],
      folders: [],
      parentFolders: [],
      getId: () => id,
      getName: () => folder.name,
      isTrashed: () => metadata.trashed === true,
      getParents() {
        let index = 0;
        const parents = folder.parentFolders.slice();
        return { hasNext: () => index < parents.length, next() { return parents[index++]; } };
      },
      getFolders() {
        let index = 0;
        const children = folder.folders.slice();
        return { hasNext: () => index < children.length, next() { return children[index++]; } };
      },
      getFiles() {
        let index = 0;
        const children = folder.files.slice();
        return { hasNext: () => index < children.length, next() { return children[index++]; } };
      },
      createFolder(name) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.mediaFolderCreates += 1;
        if (String(name) === 'original') audit.drive.folderCreates += 1;
        if (options.failOriginalFolderCreate && String(name) === 'original') {
          throw new Error('private folder create failure');
        }
        const child = makeFolder(`folder-media-${audit.drive.mediaFolderCreates}`, { name: String(name) });
        child.parentFolders = [folder];
        folder.folders.push(child);
        return child;
      },
      seedFolder(idValue, name, folderMetadata = {}) {
        const child = makeFolder(idValue, { ...folderMetadata, name });
        child.parentFolders = [folder];
        folder.folders.push(child);
        return child;
      },
      createFile(blob) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        const isGuide = blob && blob.name === 'ここに直接ファイルを入れてください.txt';
        if (isGuide) audit.drive.guideCreates += 1;
        else audit.drive.creates += 1;
        if (options.failCreatePermission) {
          throw new Error('Access denied: DriveApp. You do not have permission to create files.');
        }
        const file = makeFile(folder,
          `file-${String(audit.drive.creates + audit.drive.guideCreates).padStart(10, '0')}`,
          blob.name, {
          managed: true,
          sharingAccess: options.managedSharingAccess || 'PRIVATE',
          sharingPermission: options.managedSharingPermission || 'NONE',
          failSharingAlways: options.failManagedSharingAlways === true,
          failRenameAlways: options.failManagedRenameAlways === true
        });
        if (typeof options.afterCreateFile === 'function') options.afterCreateFile({ file, sheets, folder });
        if (options.failCreateAfterFileOnce && !createAfterFileFailed) {
          createAfterFileFailed = true;
          throw new Error('drive create response lost');
        }
        return file;
      },
      getFilesByName(name) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.searches += 1;
        const matches = folder.files.filter((file) => !file.trashed && file.name === name);
        let index = 0;
        return { hasNext: () => index < matches.length, next: () => matches[index++] };
      },
      seedFile(idValue, name, metadata) { return makeFile(folder, idValue, name, metadata); }
    };
    folders.set(id, folder);
    return folder;
  }

  Object.entries(options.sheets || {}).forEach(([name, rows]) => makeSheet(name, rows));
  const defaultFolder = makeFolder('folder-1', { name: 'Selected' });
  const rootFolder = makeFolder('123456789012345', { name: 'Root' });
  let mediaFolders = null;
  if (options.seedMediaStructure !== false) {
    const photos = rootFolder.seedFolder('media_photos_AAAAA', 'photos');
    const audio = rootFolder.seedFolder('media_audio_AAAAAA', 'audio');
    const original = rootFolder.seedFolder('media_original_AAA', 'original');
    const originalPhotos = original.seedFolder('media_original_photo', 'photos');
    const originalAudio = original.seedFolder('media_original_audio', 'audio');
    rootFolder.seedFile('media_guide_AAAAAA', 'ここに直接ファイルを入れてください.txt', {
      mimeType: 'text/plain', sizeBytes: 0, bytes: []
    });
    mediaFolders = { photos, audio, original, originalPhotos, originalAudio };
  }
  const spreadsheet = {
    getSheetByName(name) { audit.sheetLookups.push(name); return sheets.get(name) || null; },
    insertSheet(name) { audit.inserts.push(name); return makeSheet(name); }
  };
  const lock = {
    tryLock() {
      audit.locks.attempts += 1;
      if (options.lockAcquired === false) return false;
      if (audit.locks.held) {
        audit.locks.nestedAttempts += 1;
        return false;
      }
      audit.locks.held = true;
      return true;
    },
    releaseLock() { audit.locks.releases += 1; audit.locks.held = false; }
  };
  const context = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      flush() {
        assert.equal(audit.locks.held, true, 'SpreadsheetApp.flush must run before lock release');
        audit.locks.flushes += 1;
      },
      getUi: () => ({
        ButtonSet: { OK: 'OK' },
        createMenu: () => ({ addItem() { return this; }, addToUi() {} }),
        alert(...args) { audit.alerts.push(args); }
      })
    },
    CacheService: {
      getScriptCache: () => ({ get: () => options.validToken === false ? null : '1', put() {} })
    },
    LockService: { getScriptLock: () => lock },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
      getUuid() { audit.uuidCalls += 1; return `uuid-${++uuid}`; },
      computeDigest(_algorithm, value) { return Array.from(crypto.createHash('sha256').update(String(value)).digest()); },
      base64Decode(value) { return Array.from(Buffer.from(String(value), 'base64')); },
      newBlob(bytes, mime, name) { return { bytes, mime, name }; },
      formatDate: () => '2026/07/11 12:00:00'
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/example/exec' }) },
    DriveApp: {
      Access: { ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', PRIVATE: 'PRIVATE' },
      Permission: { VIEW: 'VIEW' },
      getFolderById(id) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.folderGets += 1;
        if (options.folderErrorCode) {
          const error = new Error('provider detail');
          error.code = options.folderErrorCode;
          throw error;
        }
        const folder = folders.get(id);
        if (!folder) throw new Error('folder missing');
        return folder;
      },
      getFileById(id) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.fileGets += 1;
        if (options.failFileGetOnce && !fileGetFailed) {
          fileGetFailed = true;
          throw new Error('Service unavailable temporarily');
        }
        const file = filesById.get(id);
        if (!file) throw new Error('file missing');
        return file;
      }
    },
    HtmlService: {}, Maps: {}, console: {
      error(value) { audit.errors.push(String(value)); },
      warn: console.warn.bind(console),
      log: console.log.bind(console)
    }
  };
  vm.runInNewContext(
    `${codeJs}\nthis.__api = {\n`
      + 'setupSheet,\n'
      + 'saveImportPhotoItem: typeof saveImportPhotoItem === "undefined" ? null : saveImportPhotoItem,\n'
      + 'saveImportPinItem: typeof saveImportPinItem === "undefined" ? null : saveImportPinItem,\n'
      + 'saveImportAudioItem: typeof saveImportAudioItem === "undefined" ? null : saveImportAudioItem,\n'
      + 'saveMapData: typeof saveMapData === "undefined" ? null : saveMapData,\n'
      + 'getMapData: typeof getMapData === "undefined" ? null : getMapData,\n'
      + 'PinData: typeof PinData === "undefined" ? null : PinData,\n'
      + 'encodeSpreadsheetLiteral_: typeof encodeSpreadsheetLiteral_ === "undefined" ? null : encodeSpreadsheetLiteral_,\n'
      + 'decodeSpreadsheetLiteral_: typeof decodeSpreadsheetLiteral_ === "undefined" ? null : decodeSpreadsheetLiteral_,\n'
      + 'normalizeImportPhotoPayload_: typeof normalizeImportPhotoPayload_ === "undefined" ? null : normalizeImportPhotoPayload_,\n'
      + 'normalizeImportPinPayload_: typeof normalizeImportPinPayload_ === "undefined" ? null : normalizeImportPinPayload_,\n'
      + 'hashImportPayload_: typeof hashImportPayload_ === "undefined" ? null : hashImportPayload_,\n'
      + 'importPhotoAttachTempFileName_: typeof importPhotoAttachTempFileName_ === "undefined" ? null : importPhotoAttachTempFileName_,\n'
      + 'hashImportPinPayload_: typeof hashImportPinPayload_ === "undefined" ? null : hashImportPinPayload_,\n'
      + 'openImportReceiptsSheet_: typeof openImportReceiptsSheet_ === "undefined" ? null : openImportReceiptsSheet_,\n'
      + 'isValidImportHttpUrl_: typeof isValidImportHttpUrl_ === "undefined" ? null : isValidImportHttpUrl_,\n'
      + 'IMPORT_RECEIPT_HEADERS: typeof IMPORT_RECEIPT_HEADERS === "undefined" ? null : IMPORT_RECEIPT_HEADERS\n'
      + '};',
    context
  );
  return {
    api: context.__api, audit, sheets, folders, filesById, defaultFolder, rootFolder,
    mediaFolders, context
  };
}

function baseSheets(includeReceipts = true) {
  const sheets = {
    map_info: [MAP_HEADERS],
    config: [
      ['設定項目', '値', '説明'],
      ['IMAGE_DRIVE_URL', 'https://drive.google.com/drive/folders/123456789012345', ''],
      ['RENAME_FILE_WITH_TITLE', 'false', ''],
      ['EDIT_KEY', 'key', ''], ['WEB_APP_URL', '', ''], ['EDIT_URL', '', '']
    ]
  };
  if (includeReceipts) sheets.import_receipts = [RECEIPT_HEADERS];
  return sheets;
}

function payload(overrides = {}) {
  return Object.assign({
    __editToken: 'valid-token',
    jobId: 'job-1', itemId: 'item-1', idempotencyKey: 'job-1:item-1',
    base64: 'data:image/jpeg;base64,/9j/', filename: 'photo.jpg',
    title: '写真', description: '説明', eventAt: '2026-07-11T10:30',
    lat: 35.5, lng: 139.5, color: '#e53935', icon: 'photo',
    links: ['https://example.com'], targetFolderId: 'folder-1', status: '', tags: ['観察']
  }, overrides);
}

function pinPayload(overrides = {}) {
  return Object.assign({
    __editToken: 'valid-token',
    jobId: 'csv-job-1', itemId: 'csv-item-1', idempotencyKey: 'csv-job-1:csv-item-1',
    title: 'CSVピン', description: '説明', eventAt: '2026-07-11T10:30',
    lat: 35.5, lng: 139.5, color: '#e53935', icon: 'default',
    links: ['https://example.com/path'], status: '', tags: ['観察']
  }, overrides);
}

function existingPhotoLessPinRow(overrides = {}) {
  const values = {
    timestamp: '2026/07/01 09:00:00', title: '既存ピン', description: '既存の説明',
    lat: 35.1, lng: 139.2, color: '#4caf50', fileId: '', imageUrl: '',
    id: 'pin-existing-0001', links: 'https://example.com/existing', status: '対応中',
    tags: '既存|保護', eventAt: '2026-07-01T08:30', updatedAt: '2026/07/11 11:59:00',
    icon: 'nature',
    ...overrides
  };
  return [
    values.timestamp, values.title, values.description, values.lat, values.lng, values.color,
    values.fileId, values.imageUrl, values.id, values.links, values.status, values.tags,
    values.eventAt, values.updatedAt, values.icon
  ];
}

function attachPayload(overrides = {}) {
  return payload({
    jobId: 'attach-job-1', itemId: 'attach-item-1',
    idempotencyKey: 'attach-job-1:attach-item-1',
    operationMode: 'attach-existing-pin',
    targetPinId: 'pin-existing-0001',
    expectedUpdatedAt: '2026/07/11 11:59:00',
    ...overrides
  });
}

function receiptColumn(name) { return RECEIPT_HEADERS.indexOf(name); }

test('setupSheet creates import_receipts with fixed headers and preserves existing rows', () => {
  const first = makeHarness({ sheets: baseSheets(false) });
  first.api.setupSheet();
  assert.equal(first.audit.inserts.includes('import_receipts'), true);
  assert.deepEqual(first.sheets.get('import_receipts').rows[0].slice(0, RECEIPT_HEADERS.length), RECEIPT_HEADERS);
  assert.match(first.audit.alerts.at(-1)[1], /import_receipts/);

  const existing = ['digest', 'job', 'item', 'hash', 'completed'];
  const second = makeHarness({ sheets: { ...baseSheets(false), import_receipts: [existing] } });
  second.api.setupSheet();
  assert.deepEqual(second.sheets.get('import_receipts').rows[0].slice(0, RECEIPT_HEADERS.length), RECEIPT_HEADERS);
  assert.deepEqual(second.sheets.get('import_receipts').rows[1].slice(0, existing.length), existing);

  const partialHeaders = RECEIPT_HEADERS.slice();
  partialHeaders[4] = 'wrong-state-header';
  const repaired = makeHarness({ sheets: {
    ...baseSheets(false),
    import_receipts: [partialHeaders, ['preserved-key', 'preserved-job']]
  } });
  repaired.api.setupSheet();
  assert.deepEqual(repaired.sheets.get('import_receipts').rows[0].slice(0, RECEIPT_HEADERS.length), RECEIPT_HEADERS);
  assert.deepEqual(repaired.sheets.get('import_receipts').rows[1], ['preserved-key', 'preserved-job']);

  const oldHeaders = RECEIPT_HEADERS.slice(0, 17);
  const oldRow = ['legacy-key', 'legacy-job', 'legacy-item', 'legacy-hash', 'completed'];
  const migrated = makeHarness({ sheets: {
    ...baseSheets(false), import_receipts: [oldHeaders, oldRow]
  } });
  migrated.api.setupSheet();
  assert.deepEqual(migrated.sheets.get('import_receipts').rows[0].slice(0, RECEIPT_HEADERS.length), RECEIPT_HEADERS);
  assert.deepEqual(migrated.sheets.get('import_receipts').rows[1], oldRow);
});

test('setupSheet grows an exact 17-column receipt grid by appending headers once', () => {
  const legacyHeaders = RECEIPT_HEADERS.slice(0, 17);
  const legacyRow = legacyHeaders.map((header) => ({
    idempotencyKey: 'legacy-key',
    state: 'completed',
    pinId: 'legacy-pin',
    imageUrl: 'https://example.com/photo.jpg'
  })[header] || '');
  legacyRow[11] = 'computed-image-url';
  const formulas = [legacyHeaders.map(() => ''), legacyRow.map(() => '')];
  formulas[0][16] = '="sourceDriveFileId"';
  formulas[1][11] = '=IMAGE("https://example.com/photo.jpg")';
  const harness = makeHarness({
    sheetMaxColumns: { import_receipts: 17 },
    sheetFormulas: { import_receipts: formulas },
    sheets: {
      ...baseSheets(false),
      import_receipts: [legacyHeaders, legacyRow]
    }
  });

  harness.api.setupSheet();
  harness.api.setupSheet();

  const receiptSheet = harness.sheets.get('import_receipts');
  assert.deepEqual(receiptSheet.rows[0], RECEIPT_HEADERS);
  assert.deepEqual(receiptSheet.rows[1], legacyRow);
  assert.equal(receiptSheet.formulas[0][16], '="sourceDriveFileId"');
  assert.equal(receiptSheet.formulas[1][11], '=IMAGE("https://example.com/photo.jpg")');
  assert.deepEqual(harness.audit.columnInserts, [{
    sheet: 'import_receipts', after: 17, count: 4, lockHeld: false
  }]);
  assert.deepEqual(
    JSON.parse(JSON.stringify(
      harness.audit.writes.filter((write) => write.sheet === 'import_receipts')
    )),
    [{
      sheet: 'import_receipts', method: 'setValues', row: 1, column: 18,
      numRows: 1, numColumns: 4,
      values: [['mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId']],
      lockHeld: false
    }]
  );
});

test('setupSheet preserves an appended receipt header formula that evaluates empty', () => {
  const headerValues = RECEIPT_HEADERS.slice(0, 17).concat(['', '', '', '']);
  const headerFormulas = [headerValues.map(() => '')];
  headerFormulas[0][17] = '=""';
  const harness = makeHarness({
    sheetMaxColumns: { import_receipts: 21 },
    sheetFormulas: { import_receipts: headerFormulas },
    sheets: {
      ...baseSheets(false),
      import_receipts: [headerValues]
    }
  });

  harness.api.setupSheet();

  const receiptSheet = harness.sheets.get('import_receipts');
  assert.equal(receiptSheet.formulas[0][17], '=""');
  assert.equal(receiptSheet.rows[0][17], '');
  assert.deepEqual(receiptSheet.rows[0].slice(18), RECEIPT_HEADERS.slice(18));
  assert.deepEqual(
    harness.audit.writes.filter((write) => write.sheet === 'import_receipts'
      && (write.method === 'setValue' || write.method === 'setValues')
      && write.row === 1
      && write.column <= 18
      && write.column + write.numColumns - 1 >= 18),
    []
  );
});

test('authentication happens before Spreadsheet and Drive access', () => {
  const harness = makeHarness({ validToken: false, sheets: baseSheets() });
  assert.throws(() => harness.api.saveImportPhotoItem(payload()), /編集権限/);
  assert.deepEqual(harness.audit.sheetLookups, []);
  assert.equal(harness.audit.drive.folderGets, 0);
});

test('saveImportPinItem authenticates before Spreadsheet, Lock, UUID, or Drive access', () => {
  const harness = makeHarness({ validToken: false, sheets: baseSheets() });
  assert.equal(typeof harness.api.saveImportPinItem, 'function');
  assert.throws(() => harness.api.saveImportPinItem(pinPayload()), /編集権限/);
  assert.deepEqual(harness.audit.sheetLookups, []);
  assert.equal(harness.audit.locks.attempts, 0);
  assert.equal(harness.audit.uuidCalls, 0);
  assert.equal(Object.values(harness.audit.drive).filter(Number.isFinite).reduce((sum, value) => sum + value, 0), 0);
});

test('saveImportPinItem stores one photo-less pin and completes an empty-storage receipt without Drive', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const result = harness.api.saveImportPinItem(pinPayload());
  assert.equal(result.ok, true);
  assert.equal(result.deduplicated, false);
  assert.deepEqual(JSON.parse(JSON.stringify(result.pin)), {
    timestamp: '2026/07/11 12:00:00', title: 'CSVピン', description: '説明',
    lat: 35.5, lng: 139.5, color: '#e53935', fileId: '', imageUrl: '',
    id: 'uuid-2', links: ['https://example.com/path'], status: '', tags: ['観察'],
    eventAt: '2026-07-11T10:30', updatedAt: '', icon: 'default', folderUrl: '',
    hasAudio: false
  });
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('state')], 'completed');
  ['targetFolderId', 'tempFileName', 'fileId', 'imageUrl', 'folderUrl'].forEach((name) => {
    assert.equal(receipt[receiptColumn(name)], '');
  });
  assert.equal(receipt[receiptColumn('sourceDriveFileId')], '');
  assert.match(receipt[receiptColumn('payloadHash')], /^[0-9a-f]{64}$/);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.deepEqual(harness.audit.sheetLookups, ['import_receipts', 'map_info']);
  assert.equal(harness.audit.locks.attempts, 2);
  assert.equal(harness.audit.drive.creates, 0);
  assert.equal(harness.audit.drive.renames, 0);
  assert.equal(harness.audit.drive.shares, 0);
  assert.equal(harness.audit.drive.folderGets, 0);
  assert.equal(harness.audit.drive.fileGets, 0);
  assert.equal(harness.audit.drive.searches, 0);
});

test('saveImportPinItem deduplicates replay and rejects changed or photo payloads for the same key', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const first = harness.api.saveImportPinItem(pinPayload());
  const replay = harness.api.saveImportPinItem(pinPayload());
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.pin.id, first.pin.id);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);

  const changed = harness.api.saveImportPinItem(pinPayload({ title: '異なるピン' }));
  assert.equal(changed.errorCode, 'IDEMPOTENCY_PAYLOAD_CONFLICT');
  assert.equal(changed.retryable, false);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);

  const photo = harness.api.saveImportPhotoItem(payload({
    jobId: 'csv-job-1', itemId: 'csv-item-1', idempotencyKey: 'csv-job-1:csv-item-1'
  }));
  assert.equal(photo.errorCode, 'IDEMPOTENCY_PAYLOAD_CONFLICT');
  assert.equal(photo.retryable, false);
  assert.equal(harness.audit.drive.creates, 0);
});

test('completed photo-less replay performs no receipt or map writes when no repair is needed', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  assert.equal(harness.api.saveImportPinItem(pinPayload()).ok, true);
  const writesBefore = harness.audit.writes.length;
  const mapRowsBefore = harness.sheets.get('map_info').rows.length;
  const replay = harness.api.saveImportPinItem(pinPayload());
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(harness.audit.writes.length, writesBefore);
  assert.equal(harness.sheets.get('map_info').rows.length, mapRowsBefore);
  assert.equal(harness.audit.drive.creates + harness.audit.drive.folderGets
    + harness.audit.drive.fileGets + harness.audit.drive.searches, 0);
});

test('saveImportPinItem does not auto-create a missing receipt sheet', () => {
  const harness = makeHarness({ sheets: baseSheets(false) });
  const result = harness.api.saveImportPinItem(pinPayload());
  assert.equal(result.errorCode, 'IMPORT_RECEIPT_SHEET_MISSING');
  assert.deepEqual(harness.audit.inserts, []);
  assert.deepEqual(harness.audit.sheetLookups, ['import_receipts']);
  assert.equal(harness.audit.locks.attempts, 0);
  assert.equal(harness.audit.uuidCalls, 0);
});

test('saveImportPinItem strict validation rejects every unsafe field before Spreadsheet, Lock, UUID, map, or Drive', () => {
  const cases = [
    [{ jobId: '' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ jobId: '=formula', idempotencyKey: '=formula:csv-item-1' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ itemId: '' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ itemId: '+formula', idempotencyKey: 'csv-job-1:+formula' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ idempotencyKey: '' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ idempotencyKey: 'wrong:key' }, 'INVALID_IDEMPOTENCY_KEY'],
    [{ title: '' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ title: 'x'.repeat(81) }, 'INVALID_IMPORT_PAYLOAD'],
    [{ description: 'x'.repeat(401) }, 'INVALID_IMPORT_PAYLOAD'],
    [{ lat: 35, lng: null }, 'INVALID_IMPORT_PAYLOAD'],
    [{ lat: '35', lng: '139' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ lat: 91, lng: 139 }, 'INVALID_IMPORT_PAYLOAD'],
    [{ lat: 35, lng: 181 }, 'INVALID_IMPORT_PAYLOAD'],
    [{ color: '#abcdef' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ icon: 'unknown' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ status: 'unknown' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ tags: 'not-an-array' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ tags: ['1', '2', '3', '4', '5', '6'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: 'https://example.com' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: [42] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['ftp://example.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://example.com:99999'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['http://999.999.999.999'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://%zz.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://exa^mple.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://exa|mple.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://%2F.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://%00.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://\u007f.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://\u0080.com'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ links: ['https://[::ffff:192.168.001.1]/'] }, 'INVALID_IMPORT_PAYLOAD'],
    [{ eventAt: '2026-02-30T10:00' }, 'INVALID_IMPORT_PAYLOAD'],
    [{ eventAt: 'invalid' }, 'INVALID_IMPORT_PAYLOAD']
  ];
  cases.forEach(([changes, expectedCode]) => {
    const harness = makeHarness({ sheets: baseSheets() });
    const result = harness.api.saveImportPinItem(pinPayload(changes));
    assert.equal(result.ok, false, JSON.stringify(changes));
    assert.equal(result.errorCode, expectedCode, JSON.stringify(changes));
    assert.deepEqual(harness.audit.sheetLookups, [], JSON.stringify(changes));
    assert.equal(harness.audit.locks.attempts, 0, JSON.stringify(changes));
    assert.equal(harness.audit.uuidCalls, 0, JSON.stringify(changes));
    assert.equal(harness.audit.writes.length, 0, JSON.stringify(changes));
    assert.equal(harness.audit.drive.creates + harness.audit.drive.folderGets
      + harness.audit.drive.fileGets + harness.audit.drive.searches, 0, JSON.stringify(changes));
  });
});

test('spreadsheet literal codec is reversible for formula prefixes, apostrophes, and its own marker', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const values = [
    '=1+1', '+SUM(A1:A2)', '-10+20', '@command', '\tTAB', '\rCR',
    "'通常の文章", "'=数式ではなく文章", "'+先頭プラスを含む文章", "''二重アポストロフィ",
    '通常の文章', '\u200Bdtp-sheet:v1:v:=marker-like-user-text', ''
  ];

  values.forEach((value) => {
    const encoded = harness.api.encodeSpreadsheetLiteral_(value);
    assert.equal(/^[=+\-@\t\r]/.test(encoded), false, value);
    assert.equal(harness.api.decodeSpreadsheetLiteral_(encoded), value, value);
  });
});

test('saveImportPinItem protects title, description, and tags while preserving exact replay results', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const request = pinPayload({
    title: '=1+1',
    description: '+SUM(A1:A2)',
    tags: ['=IMPORTXML()', '+SUM', '-10+20', '@command', '通常タグ']
  });
  const result = harness.api.saveImportPinItem(request);
  assert.equal(result.ok, true);
  assert.equal(result.pin.title, '=1+1');
  assert.equal(result.pin.description, '+SUM(A1:A2)');
  assert.deepEqual(JSON.parse(JSON.stringify(result.pin.tags)), request.tags);

  const stored = harness.sheets.get('map_info').rows[1];
  assert.equal(/^[=+\-@\t\r]/.test(stored[1]), false);
  assert.equal(/^[=+\-@\t\r]/.test(stored[2]), false);
  assert.equal(/^[=+\-@\t\r]/.test(stored[11]), false);
  assert.equal(harness.api.decodeSpreadsheetLiteral_(stored[1]), request.title);
  assert.equal(harness.api.decodeSpreadsheetLiteral_(stored[2]), request.description);
  assert.equal(harness.api.decodeSpreadsheetLiteral_(stored[11]), request.tags.join('|'));

  const writesBeforeReplay = harness.audit.writes.length;
  const replay = harness.api.saveImportPinItem(request);
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.pin.title, result.pin.title);
  assert.equal(replay.pin.description, result.pin.description);
  assert.deepEqual(JSON.parse(JSON.stringify(replay.pin.tags)), request.tags);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.deepEqual(harness.audit.writes.slice(writesBeforeReplay), []);
  assert.equal(harness.audit.drive.creates, 0);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.match(String(receipt[receiptColumn('payloadHash')]), /^[0-9a-f]{64}$/);
  assert.equal(JSON.stringify(receipt).includes('dtp-sheet:v1:'), false);
});

test('PinData and getMapData decode protected cells while legacy and apostrophe text remain unchanged', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const encode = harness.api.encodeSpreadsheetLiteral_;
  const protectedRow = [
    '', encode('=title'), encode('+description'), 35, 139, '#e53935', '', '', 'protected-id', '', '',
    encode('@tag|-10+20'), '', '', 'default'
  ];
  harness.sheets.get('map_info').rows.push(protectedRow);
  harness.sheets.get('map_info').rows.push([
    '', "'=数式ではなく文章", "''二重アポストロフィ", '', '', '#e53935', '', '', 'legacy-id', '', '',
    "'通常タグ", '', '', 'default'
  ]);

  const direct = harness.api.PinData.rowToPin(protectedRow);
  assert.equal(direct.title, '=title');
  assert.equal(direct.description, '+description');
  assert.deepEqual(JSON.parse(JSON.stringify(direct.tags)), ['@tag', '-10+20']);

  const pins = harness.api.getMapData();
  assert.equal(pins[0].title, '=title');
  assert.equal(pins[0].description, '+description');
  assert.deepEqual(JSON.parse(JSON.stringify(pins[0].tags)), ['@tag', '-10+20']);
  assert.equal(pins[1].title, "'=数式ではなく文章");
  assert.equal(pins[1].description, "''二重アポストロフィ");
  assert.deepEqual(JSON.parse(JSON.stringify(pins[1].tags)), ["'通常タグ"]);
  assert.equal(JSON.stringify(pins).includes('dtp-sheet:v1:'), false);
});

test('saveMapData round-trips user apostrophes without treating them as protection prefixes', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const values = [
    "'通常の文章", "'=数式ではなく文章", "'+先頭プラスを含む文章", "''二重アポストロフィ"
  ];
  values.forEach((value, index) => {
    const result = harness.api.saveMapData({
      __editToken: 'valid-token', title: value, description: value,
      tags: [value], color: '#e53935', icon: 'default', status: ''
    });
    assert.equal(result.ok, true, String(index));
  });

  const pins = harness.api.getMapData();
  assert.deepEqual(JSON.parse(JSON.stringify(pins.map((pin) => pin.title))), values);
  assert.deepEqual(JSON.parse(JSON.stringify(pins.map((pin) => pin.description))), values);
  assert.deepEqual(JSON.parse(JSON.stringify(pins.map((pin) => pin.tags[0]))), values);
});

test('saveMapData and photo import share formula-safe B, C, and L storage', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const withoutPhoto = harness.api.saveMapData({
    __editToken: 'valid-token', title: '=no-photo', description: '+description',
    tags: ['@tag'], color: '#e53935', icon: 'default', status: ''
  });
  assert.equal(withoutPhoto.ok, true);
  assert.equal(harness.audit.drive.creates, 0);

  const withPhoto = harness.api.saveMapData({
    __editToken: 'valid-token', title: '=normal', description: '@description',
    tags: ['-tag'], color: '#e53935', icon: 'default', status: '',
    base64: 'data:image/jpeg;base64,/9j/', filename: 'normal.jpg', targetFolderId: 'folder-1'
  });
  assert.equal(withPhoto.ok, true);

  const photoRequest = payload({
    title: '+photo', description: '=description', tags: ['@photo-tag']
  });
  const imported = harness.api.saveImportPhotoItem(photoRequest);
  assert.equal(imported.ok, true);
  assert.equal(imported.pin.title, '+photo');
  assert.deepEqual(JSON.parse(JSON.stringify(imported.pin.tags)), ['@photo-tag']);

  const writesBeforeReplay = harness.audit.writes.length;
  const replay = harness.api.saveImportPhotoItem(photoRequest);
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.pin.title, imported.pin.title);
  assert.deepEqual(JSON.parse(JSON.stringify(replay.pin.tags)), ['@photo-tag']);
  assert.deepEqual(harness.audit.writes.slice(writesBeforeReplay), []);

  for (const row of harness.sheets.get('map_info').rows.slice(1)) {
    assert.equal(/^[=+\-@\t\r]/.test(row[1]), false);
    assert.equal(/^[=+\-@\t\r]/.test(row[2]), false);
    assert.equal(/^[=+\-@\t\r]/.test(row[11]), false);
    assert.equal(row[10], '');
  }
  assert.equal(harness.sheets.get('map_info').rows.length, 4);
  assert.equal(harness.audit.drive.creates, 2);
});

test('saveImportPinItem normalizes strict common fields while preserving blank status and unplaced coordinates', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const result = harness.api.saveImportPinItem(pinPayload({
    title: '  正規化  ', description: '', lat: null, lng: null, color: '#E53935',
    links: ['', 'https://example.com', 'https://example.com'],
    tags: ['#観察', '観察'], eventAt: '2026-07-11T10:30:59', status: ''
  }));
  assert.equal(result.ok, true);
  assert.equal(result.pin.title, '正規化');
  assert.equal(result.pin.lat, null);
  assert.equal(result.pin.lng, null);
  assert.equal(result.pin.color, '#e53935');
  assert.equal(result.pin.status, '');
  assert.deepEqual(JSON.parse(JSON.stringify(result.pin.links)), ['https://example.com']);
  assert.deepEqual(JSON.parse(JSON.stringify(result.pin.tags)), ['観察']);
  assert.equal(result.pin.eventAt, '2026-07-11T10:30:59');
});

test('strict URL validation accepts a parseable hostname without requiring a dotted public domain', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const result = harness.api.saveImportPinItem(pinPayload({
    links: [
      'http://intranet/path',
      'https://[2001:db8::1]/pin',
      'https://[::ffff:192.0.2.1]/pin'
    ]
  }));
  assert.equal(result.ok, true);
  assert.deepEqual(JSON.parse(JSON.stringify(result.pin.links)), [
    'http://intranet/path',
    'https://[2001:db8::1]/pin',
    'https://[::ffff:192.0.2.1]/pin'
  ]);
});

test('server fallback validation uses the shared URL vectors', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  importUrlVectors.allowed.forEach((link) => {
    assert.equal(harness.api.isValidImportHttpUrl_(link), true, link);
  });
  importUrlVectors.rejected.forEach((link) => {
    assert.equal(harness.api.isValidImportHttpUrl_(link), false, link);
  });
});

test('strict server eventAt matches real year and leap boundaries', () => {
  const accepted = makeHarness({ sheets: baseSheets() });
  assert.equal(accepted.api.saveImportPinItem(pinPayload({ eventAt: '0001-01-01T00:00:00' })).ok, true);
  const leap = makeHarness({ sheets: baseSheets() });
  assert.equal(leap.api.saveImportPinItem(pinPayload({ eventAt: '2000-02-29T23:59:59' })).ok, true);
  const yearZero = makeHarness({ sheets: baseSheets() });
  assert.equal(
    yearZero.api.saveImportPinItem(pinPayload({ eventAt: '0000-12-31T23:59' })).errorCode,
    'INVALID_IMPORT_PAYLOAD'
  );
});

test('pin payload hash is domain-separated and reflects normalized ordered content', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  function hash(changes) {
    return harness.api.hashImportPinPayload_(
      harness.api.normalizeImportPinPayload_(pinPayload(changes))
    );
  }
  const base = hash({ color: '#E53935', tags: ['#A', 'a'], links: ['', 'https://a.example', 'https://a.example'] });
  assert.equal(base, hash({ color: '#e53935', tags: ['A'], links: ['https://a.example'] }));
  assert.notEqual(hash({ tags: ['A', 'B'] }), hash({ tags: ['B', 'A'] }));
  assert.notEqual(
    hash({ links: ['https://a.example', 'https://b.example'] }),
    hash({ links: ['https://b.example', 'https://a.example'] })
  );
  const normalized = harness.api.normalizeImportPinPayload_(pinPayload());
  const withoutDomain = crypto.createHash('sha256').update(JSON.stringify({
    jobId: normalized.jobId, itemId: normalized.itemId,
    title: normalized.title, description: normalized.description, eventAt: normalized.eventAt,
    lat: normalized.lat, lng: normalized.lng, color: normalized.color, icon: normalized.icon,
    links: normalized.links.slice(), status: normalized.status, tags: normalized.tags.slice()
  })).digest('hex');
  assert.notEqual(harness.api.hashImportPinPayload_(normalized), withoutDomain);
});

test('photo receipt key cannot be reused for a photo-less pin', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  assert.equal(harness.api.saveImportPhotoItem(payload()).ok, true);
  const pin = harness.api.saveImportPinItem(pinPayload({
    jobId: 'job-1', itemId: 'item-1', idempotencyKey: 'job-1:item-1'
  }));
  assert.equal(pin.errorCode, 'IDEMPOTENCY_PAYLOAD_CONFLICT');
  assert.equal(pin.retryable, false);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('photo strict normalization preserves the selected target folder for Drive sources', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const normalized = harness.api.normalizeImportPhotoPayload_(payload({ color: '#E53935' }));
  assert.equal(normalized.color, '#E53935');
  assert.equal(normalized.sourceDriveFileId, '');
  const expected = crypto.createHash('sha256').update(JSON.stringify({
    jobId: normalized.jobId, itemId: normalized.itemId, base64: normalized.base64,
    filename: normalized.filename, title: normalized.title, description: normalized.description,
    eventAt: normalized.eventAt, lat: normalized.lat, lng: normalized.lng,
    color: '#E53935', icon: normalized.icon, links: normalized.links.slice(),
    targetFolderId: normalized.targetFolderId, status: normalized.status, tags: normalized.tags.slice()
  })).digest('hex');
  assert.equal(harness.api.hashImportPayload_(normalized), expected);
  const driveNormalized = harness.api.normalizeImportPhotoPayload_(payload({
    color: '#E53935', sourceDriveFileId: 'photo_AAAAAAAAAAA'
  }));
  assert.equal(driveNormalized.sourceDriveFileId, 'photo_AAAAAAAAAAA');
  assert.equal(driveNormalized.targetFolderId, 'folder-1');
  assert.notEqual(harness.api.hashImportPayload_(driveNormalized), expected);
});

test('photo-less receipt recovery repairs missing map rows and unfinished receipts without duplicates', () => {
  const completedMissingMap = makeHarness({ sheets: baseSheets() });
  const first = completedMissingMap.api.saveImportPinItem(pinPayload());
  completedMissingMap.sheets.get('map_info').rows.pop();
  const repaired = completedMissingMap.api.saveImportPinItem(pinPayload());
  assert.equal(repaired.ok, true);
  assert.equal(repaired.deduplicated, true);
  assert.equal(repaired.pin.id, first.pin.id);
  assert.equal(completedMissingMap.sheets.get('map_info').rows.length, 2);

  const unfinished = makeHarness({ sheets: baseSheets() });
  const saved = unfinished.api.saveImportPinItem(pinPayload());
  const receipt = unfinished.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('state')] = 'failed';
  receipt[receiptColumn('lastErrorCode')] = 'IMPORT_ITEM_SAVE_FAILED';
  const recovered = unfinished.api.saveImportPinItem(pinPayload());
  assert.equal(recovered.ok, true);
  assert.equal(recovered.deduplicated, true);
  assert.equal(recovered.pin.id, saved.pin.id);
  assert.equal(unfinished.sheets.get('map_info').rows.length, 2);
  assert.equal(receipt[receiptColumn('state')], 'completed');
});

test('photo-less live lease blocks while expired lease reuses the same pin id', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const first = harness.api.saveImportPinItem(pinPayload());
  harness.sheets.get('map_info').rows.pop();
  const receipt = harness.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('state')] = 'reserved';
  receipt[receiptColumn('leaseOwner')] = 'foreign-owner';
  receipt[receiptColumn('leaseUntil')] = '2999-01-01T00:00:00.000Z';
  const busy = harness.api.saveImportPinItem(pinPayload());
  assert.equal(busy.errorCode, 'IMPORT_ITEM_IN_PROGRESS');
  assert.equal(busy.retryable, true);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);

  receipt[receiptColumn('leaseUntil')] = '2000-01-01T00:00:00.000Z';
  const recovered = harness.api.saveImportPinItem(pinPayload());
  assert.equal(recovered.ok, true);
  assert.equal(recovered.pin.id, first.pin.id);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('photo-less receipt rejects storage corruption and a lost owner before map or Drive effects', () => {
  const corrupted = makeHarness({ sheets: baseSheets() });
  corrupted.api.saveImportPinItem(pinPayload());
  corrupted.sheets.get('map_info').rows.pop();
  const receipt = corrupted.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('targetFolderId')] = 'unexpected-folder';
  const rejected = corrupted.api.saveImportPinItem(pinPayload());
  assert.equal(rejected.errorCode, 'IMPORT_RECEIPT_CORRUPTED');
  assert.equal(rejected.retryable, false);
  assert.equal(corrupted.sheets.get('map_info').rows.length, 1);
  assert.equal(corrupted.audit.drive.folderGets, 0);

  const ownerLost = makeHarness({
    sheets: baseSheets(),
    afterReceiptAppend(row) { row[receiptColumn('leaseOwner')] = 'foreign-owner'; }
  });
  const lost = ownerLost.api.saveImportPinItem(pinPayload());
  assert.equal(lost.errorCode, 'IMPORT_ITEM_LEASE_LOST');
  assert.equal(lost.retryable, true);
  assert.equal(ownerLost.sheets.get('map_info').rows.length, 1);
  assert.equal(ownerLost.audit.drive.creates, 0);
});

test('photo-less receipt rejects file_saved and unknown states instead of coercing them to reserved', () => {
  ['file_saved', 'unknown'].forEach((state) => {
    const harness = makeHarness({ sheets: baseSheets() });
    harness.api.saveImportPinItem(pinPayload());
    harness.sheets.get('map_info').rows.pop();
    const receipt = harness.sheets.get('import_receipts').rows[1];
    receipt[receiptColumn('state')] = state;
    receipt[receiptColumn('leaseOwner')] = '';
    receipt[receiptColumn('leaseUntil')] = '';
    const result = harness.api.saveImportPinItem(pinPayload());
    assert.equal(result.errorCode, 'IMPORT_RECEIPT_CORRUPTED', state);
    assert.equal(harness.sheets.get('map_info').rows.length, 1, state);
  });
});

test('photo-less map and receipt completion interruptions converge without Drive or duplicate rows', () => {
  const mapFailure = makeHarness({ sheets: baseSheets(), failMapAppendOnce: true });
  const failed = mapFailure.api.saveImportPinItem(pinPayload());
  assert.equal(failed.errorCode, 'IMPORT_MAP_ROW_FAILED');
  assert.equal(mapFailure.sheets.get('import_receipts').rows[1][receiptColumn('state')], 'failed');
  const mapRecovered = mapFailure.api.saveImportPinItem(pinPayload());
  assert.equal(mapRecovered.ok, true);
  assert.equal(mapFailure.sheets.get('map_info').rows.length, 2);

  const completionFailure = makeHarness({ sheets: baseSheets(), failReceiptCompleteWriteOnce: true });
  const interrupted = completionFailure.api.saveImportPinItem(pinPayload());
  assert.equal(interrupted.ok, false);
  assert.equal(completionFailure.sheets.get('map_info').rows.length, 2);
  const completionRecovered = completionFailure.api.saveImportPinItem(pinPayload());
  assert.equal(completionRecovered.ok, true);
  assert.equal(completionRecovered.deduplicated, true);
  assert.equal(completionFailure.sheets.get('map_info').rows.length, 2);
  assert.equal(mapFailure.audit.drive.creates + completionFailure.audit.drive.creates, 0);
});

test('photo-less receipts contain only hashes, ids, state, lease metadata, and empty storage columns', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.api.saveImportPinItem(pinPayload({
    __editToken: 'secret-edit-token', title: 'secret-title', description: 'secret-description',
    tags: ['secret-tag'], links: ['https://secret.example.com']
  }));
  const serialized = JSON.stringify(harness.sheets.get('import_receipts').rows[1]);
  ['secret-edit-token', 'secret-title', 'secret-description', 'secret-tag', 'secret.example.com'].forEach((secret) => {
    assert.equal(serialized.includes(secret), false, secret);
  });
});

test('missing receipt sheet is created by photo save without running setupSheet', () => {
  const harness = makeHarness({ sheets: baseSheets(false) });
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.ok, true, JSON.stringify(result));
  assert.deepEqual(harness.audit.inserts, ['import_receipts']);
  assert.deepEqual(harness.sheets.get('import_receipts').rows[0], RECEIPT_HEADERS);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('strict validation rejects malformed payloads before receipt, Drive, or map writes', () => {
  const cases = [
    { jobId: '' }, { itemId: '' }, { idempotencyKey: '' },
    { jobId: 'bad\u0000id', idempotencyKey: 'bad\u0000id:item-1' },
    { jobId: 'x'.repeat(129), idempotencyKey: `${'x'.repeat(129)}:item-1` },
    { idempotencyKey: 'other:item-1' }, { title: '' }, { title: 'x'.repeat(81) },
    { base64: 'data:image/png;base64,/9j/' }, { base64: 'data:image/jpeg;base64,YWJj' },
    { base64: 'not-base64' }, { filename: '' },
    { lat: 35, lng: null }, { lat: '', lng: '' }, { lat: '35', lng: '139' },
    { lat: 91, lng: 139 }, { lat: 35, lng: 181 },
    { color: '' }, { color: 'red' }, { color: '#abcdef' },
    { icon: 'unknown' }, { status: 'unknown' },
    { tags: 'not-an-array' }, { tags: ['1', '2', '3', '4', '5', '6'] },
    { sourceDriveFileId: 123 }, { sourceDriveFileId: 'https://drive.google.com/file/d/id' },
    { sourceDriveFileId: '../photo_AAAAAAAAAAA' }, { sourceDriveFileId: 'bad\nphoto' },
    { sourceDriveFileId: 'x'.repeat(201) }
  ];
  cases.forEach((changes) => {
    const harness = makeHarness({ sheets: baseSheets() });
    const result = harness.api.saveImportPhotoItem(payload(changes));
    assert.equal(result.ok, false, JSON.stringify(changes));
    assert.match(result.errorCode, /INVALID|IDEMPOTENCY/);
    assert.equal(harness.audit.writes.length, 0);
    assert.equal(harness.audit.drive.creates, 0);
    assert.equal(harness.audit.sheetLookups.length, 0);
  });
});

test('strict server color, icon, and status allowlists match the client catalogs', () => {
  const serverColors = codeJs.match(/const COLOR_OPTIONS = \[([\s\S]*?)\];/)[1]
    .match(/#[0-9a-f]{6}/g);
  const clientColors = indexHtml.match(/const PIN_COLORS = \[([\s\S]*?)\];/)[1]
    .match(/#[0-9a-f]{6}/g);
  const serverIcons = codeJs.match(/const ICON_OPTIONS = \[([^\]]*)\];/)[1]
    .match(/'([^']+)'/g).map((value) => value.slice(1, -1));
  const clientIcons = Array.from(
    indexHtml.match(/const PIN_ICONS = \[([\s\S]*?)\];/)[1].matchAll(/id: '([^']+)'/g),
    (match) => match[1]
  );
  const serverStatuses = codeJs.match(/const STATUS_OPTIONS = \[([^\]]*)\];/)[1]
    .match(/'([^']+)'/g).map((value) => value.slice(1, -1));
  const clientStatuses = indexHtml.match(/const PIN_STATUSES = \[([^\]]*)\];/)[1]
    .match(/'([^']+)'/g).map((value) => value.slice(1, -1));

  assert.deepEqual(serverColors, clientColors);
  assert.deepEqual(serverIcons, clientIcons);
  assert.deepEqual(serverStatuses, clientStatuses);
});

test('normal API rejects an invalid receipt schema without creating or rewriting it', () => {
  const harness = makeHarness({ sheets: {
    ...baseSheets(false),
    import_receipts: [['wrongHeader'], ['existing-row']]
  } });
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'IMPORT_RECEIPT_CORRUPTED');
  assert.equal(harness.audit.inserts.length, 0);
  assert.equal(harness.audit.writes.length, 0);
  assert.equal(harness.audit.drive.creates, 0);
});

test('receipt auto-migration rejects reordered duplicate and missing-middle headers without mutation', () => {
  const malformedHeaders = [
    (() => { const headers = RECEIPT_HEADERS.slice(); [headers[2], headers[3]] = [headers[3], headers[2]]; return headers; })(),
    (() => { const headers = RECEIPT_HEADERS.slice(); headers[5] = headers[4]; return headers; })(),
    RECEIPT_HEADERS.filter((header) => header !== 'leaseUntil')
  ];
  malformedHeaders.forEach((headers) => {
    const existingRow = ['preserve-me', 'job'];
    const harness = makeHarness({ sheets: {
      ...baseSheets(false), import_receipts: [headers, existingRow]
    } });
    const result = harness.api.saveImportPhotoItem(payload());
    assert.equal(result.errorCode, 'IMPORT_RECEIPT_CORRUPTED', JSON.stringify(headers));
    assert.deepEqual(harness.sheets.get('import_receipts').rows, [headers, existingRow]);
    assert.equal(harness.audit.inserts.length, 0);
    assert.equal(harness.audit.writes.length, 0);
    assert.equal(harness.audit.drive.creates, 0);
  });
});

test('photo save rejects an evaluated-empty formula in the receipt header suffix without mutation', () => {
  const headerValues = RECEIPT_HEADERS.slice(0, 17).concat(['', '', '', '']);
  const headerFormulas = [headerValues.map(() => '')];
  headerFormulas[0][17] = '=""';
  const harness = makeHarness({
    sheetMaxColumns: { import_receipts: 21 },
    sheetFormulas: { import_receipts: headerFormulas },
    sheets: {
      ...baseSheets(false),
      import_receipts: [headerValues]
    }
  });

  const result = harness.api.saveImportPhotoItem(payload());

  const receiptSheet = harness.sheets.get('import_receipts');
  assert.equal(receiptSheet.formulas[0][17], '=""');
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'IMPORT_RECEIPT_CORRUPTED');
  assert.equal(result.retryable, false);
  assert.deepEqual(receiptSheet.rows, [headerValues]);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
  assert.deepEqual(harness.audit.columnInserts, []);
  assert.deepEqual(harness.audit.writes, []);
  assert.equal(harness.audit.uuidCalls, 0);
  assert.equal(harness.audit.drive.creates, 0);
  assert.equal(harness.audit.drive.fileGets + harness.audit.drive.folderGets, 0);
});

test('legacy 16-column schema reproduces the prior IMPORT_RECEIPT_CORRUPTED failure before claim or Drive access', () => {
  const legacyHeaders = RECEIPT_HEADERS.slice(0, 16);
  const harness = makeHarness({ sheets: {
    ...baseSheets(false), import_receipts: [legacyHeaders]
  } });

  assert.throws(
    () => harness.api.openImportReceiptsSheet_(),
    (error) => error && error.code === 'IMPORT_RECEIPT_CORRUPTED' && error.retryable === false
  );
  assert.equal(harness.audit.uuidCalls, 0, 'receipt claim must not start');
  assert.equal(harness.audit.locks.attempts, 0, 'receipt claim lock must not start');
  assert.equal(harness.audit.drive.fileGets + harness.audit.drive.folderGets + harness.audit.drive.creates, 0);
  assert.equal(harness.sheets.get('map_info').rows.length, 1, 'map persistence must not start');
});

test('photo save auto-migrates the exact legacy 16-column receipt schema before Drive access', () => {
  const legacyHeaders = RECEIPT_HEADERS.slice(0, 16);
  const legacyRow = [
    'legacy-key', 'legacy-job', 'legacy-item', 'legacy-hash', 'completed', '', '',
    'legacy-pin', 'folder-1', '__drop_pin_import_legacy-pin.jpg', 'legacy-file',
    'https://drive.google.com/thumbnail?id=legacy-file',
    'https://drive.google.com/drive/folders/folder-1',
    '2026-07-10T00:00:00.000Z', '2026-07-10T00:00:00.000Z', ''
  ];
  const harness = makeHarness({
    sheetMaxColumns: { import_receipts: 16 },
    sheets: {
      ...baseSheets(false),
      import_receipts: [legacyHeaders, legacyRow]
    }
  });

  const result = harness.api.saveImportPhotoItem(payload());

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.deepEqual(harness.sheets.get('import_receipts').rows[0], RECEIPT_HEADERS);
  assert.deepEqual(harness.sheets.get('import_receipts').rows[1], legacyRow);
  const migrationWrite = harness.audit.writes.find((write) => (
    write.sheet === 'import_receipts'
      && write.row === 1
      && write.column === 17
      && write.numRows === 1
      && write.numColumns === 5
  ));
  assert.ok(migrationWrite, 'migration must write only the five appended header cells');
  assert.deepEqual(harness.audit.columnInserts, [{
    sheet: 'import_receipts', after: 16, count: 5, lockHeld: true
  }]);
  assert.equal(harness.audit.locks.nestedAttempts, 0);
  assert.equal(harness.audit.locks.releases, harness.audit.locks.attempts);

  const replay = harness.api.saveImportPhotoItem(payload());
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(harness.audit.writes.filter((write) => (
    write.sheet === 'import_receipts'
      && write.row === 1
      && write.column === 17
      && write.numRows === 1
      && write.numColumns === 5
  )).length, 1, 'repeated saves must not repeat the schema migration');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('legacy 17-column completed receipt replays as photo without rewriting its row', () => {
  const initial = makeHarness({ sheets: baseSheets() });
  const first = initial.api.saveImportPhotoItem(payload());
  assert.equal(first.ok, true, JSON.stringify(first));

  const legacyHeaders = RECEIPT_HEADERS.slice(0, 17);
  const legacyRow = initial.sheets.get('import_receipts').rows[1].slice(0, 17);
  const replayHarness = makeHarness({
    sheetMaxColumns: { import_receipts: 17 },
    sheets: {
      ...baseSheets(false),
      map_info: initial.sheets.get('map_info').rows.map((row) => row.slice()),
      import_receipts: [legacyHeaders, legacyRow]
    }
  });

  const replay = replayHarness.api.saveImportPhotoItem(payload());

  assert.equal(replay.ok, true, JSON.stringify(replay));
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.pin.imageUrl, 'https://drive.google.com/thumbnail?id=file-0000000001&sz=w1920');
  assert.deepEqual(replayHarness.sheets.get('import_receipts').rows[1], legacyRow);
  assert.deepEqual(replayHarness.audit.columnInserts, [{
    sheet: 'import_receipts', after: 17, count: 4, lockHeld: true
  }]);
});

test('legacy receipt migration lock contention is retryable and performs no schema or Drive write', () => {
  const legacyHeaders = RECEIPT_HEADERS.slice(0, 16);
  const harness = makeHarness({ lockAcquired: false, sheets: {
    ...baseSheets(false), import_receipts: [legacyHeaders]
  } });
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.errorCode, 'IMPORT_ITEM_IN_PROGRESS');
  assert.equal(result.retryable, true);
  assert.deepEqual(harness.sheets.get('import_receipts').rows, [legacyHeaders]);
  assert.equal(harness.audit.writes.length, 0);
  assert.equal(harness.audit.drive.creates, 0);
});

test('legacy receipt migration enables device and managed Drive photos', () => {
  const deviceHarness = makeHarness({ sheets: {
    ...baseSheets(false), import_receipts: [RECEIPT_HEADERS.slice(0, 16)]
  } });
  const deviceRequests = [
    payload({ itemId: 'item-1', idempotencyKey: 'job-1:item-1', filename: 'one.jpg', title: '一枚目' }),
    payload({ itemId: 'item-2', idempotencyKey: 'job-1:item-2', filename: 'two.jpg', title: '二枚目' })
  ];
  const deviceResults = deviceRequests.map((request) => deviceHarness.api.saveImportPhotoItem(request));
  assert.equal(deviceResults.every((result) => result.ok), true, JSON.stringify(deviceResults));
  assert.equal(deviceHarness.audit.drive.creates, 2);
  assert.equal(deviceHarness.sheets.get('map_info').rows.length, 3);
  assert.equal(deviceHarness.sheets.get('import_receipts').rows.length, 3);

  const driveHarness = makeHarness({ sheets: {
    ...baseSheets(false), import_receipts: [RECEIPT_HEADERS.slice(0, 16)]
  } });
  driveHarness.rootFolder.seedFile('photo_MIGRATEAAAAA', 'drive.jpg');
  const driveRequest = payload({
    sourceDriveFileId: 'photo_MIGRATEAAAAA', targetFolderId: '', title: 'Drive写真'
  });
  const driveResult = driveHarness.api.saveImportPhotoItem(driveRequest);
  assert.equal(driveResult.ok, true, JSON.stringify(driveResult));
  assert.notEqual(driveResult.pin.fileId, 'photo_MIGRATEAAAAA');
  assert.equal(driveHarness.audit.drive.creates, 1);
  assert.equal(driveHarness.sheets.get('map_info').rows.length, 2);
});

test('response loss after legacy migration retries without duplicate photo map row or receipt', () => {
  const harness = makeHarness({ failCreateAfterFileOnce: true, sheets: {
    ...baseSheets(false), import_receipts: [RECEIPT_HEADERS.slice(0, 16)]
  } });
  const request = payload({ title: '応答喪失後の再試行' });
  const lost = harness.api.saveImportPhotoItem(request);
  assert.equal(lost.errorCode, 'DRIVE_MANAGED_COPY_CREATE_FAILED');
  const retry = harness.api.saveImportPhotoItem(request);
  assert.equal(retry.ok, true, JSON.stringify(retry));
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.sheets.get('import_receipts').rows.length, 2);
});

test('first save creates one file and row; completed retry deduplicates; changed payload conflicts', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const first = harness.api.saveImportPhotoItem(payload());
  assert.equal(first.ok, true);
  assert.equal(first.deduplicated, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.renames, 1);
  assert.equal(harness.audit.drive.shares, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.audit.locks.attempts, 4);
  assert.equal(harness.audit.locks.flushes, 3);
  assert.equal(harness.audit.locks.releases, 4);
  assert.equal(harness.audit.locks.nestedAttempts, 0);
  assert.equal(harness.audit.writes.filter((write) => write.sheet === 'import_receipts').length, 3);
  assert.equal(harness.audit.writes.filter((write) => write.sheet === 'map_info').length, 1);

  const second = harness.api.saveImportPhotoItem(payload());
  assert.equal(second.ok, true);
  assert.equal(second.deduplicated, true);
  assert.equal(second.pin.id, first.pin.id);
  assert.equal(second.pin.fileId, first.pin.fileId);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.renames, 1);
  assert.equal(harness.audit.drive.shares, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);

  const conflict = harness.api.saveImportPhotoItem(payload({ title: '別内容' }));
  assert.equal(conflict.ok, false);
  assert.equal(conflict.errorCode, 'IDEMPOTENCY_PAYLOAD_CONFLICT');
  assert.equal(conflict.retryable, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('Drive save stores the managed JPEG in photos and archives the source after pin linkage', () => {
  const harness = makeHarness({ sheets: baseSheets(), seedMediaStructure: false });
  const source = harness.rootFolder.seedFile('photo_ARCHIVEAAAAA', 'source.png', {
    mimeType: 'image/png', sizeBytes: 3
  });

  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: source.id,
    targetFolderId: 'folder-1',
    title: '整理対象'
  }));

  assert.equal(result.ok, true, JSON.stringify(result));
  const photos = harness.rootFolder.folders.filter((folder) => folder.name === 'photos');
  const originals = harness.rootFolder.folders.filter((folder) => folder.name === 'original');
  const originalPhotos = originals[0].folders.filter((folder) => folder.name === 'photos');
  assert.equal(photos.length, 1);
  assert.equal(originals.length, 1);
  assert.equal(originalPhotos.length, 1);
  assert.equal(harness.audit.drive.mediaFolderCreates, 5);
  assert.deepEqual(source.parentFolders, [originalPhotos[0]]);
  assert.equal(harness.audit.drive.moves, 1);
  assert.equal(harness.audit.drive.moveIds[0], source.id);
  assert.equal(harness.defaultFolder.files.length, 0, 'Drive managed JPEG must ignore the selected folder');
  assert.equal(photos[0].files.filter((file) => file.id !== source.id && !file.trashed).length, 1);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('targetFolderId')], photos[0].id);
  assert.notEqual(receipt[receiptColumn('fileId')], source.id);
  const mapAppendAt = harness.audit.events.findIndex((event) =>
    event.type === 'sheet-append' && event.sheet === 'map_info');
  const sourceMoveAt = harness.audit.events.findIndex((event) =>
    event.type === 'drive-move' && event.id === source.id);
  assert.equal(mapAppendAt >= 0 && sourceMoveAt > mapAppendAt, true);
});

test('Drive save reuses original/photos, preserves its files, and ignores case variants and trash', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.rootFolder.seedFolder('original_TRASHAAAA', 'original', { trashed: true });
  harness.rootFolder.seedFolder('original_CASEAAAAA', 'Original');
  const existing = harness.mediaFolders.originalPhotos.seedFile(
    'photo_EXISTINGAAAA', 'existing.jpg', { mimeType: 'image/jpeg' }
  );
  const source = harness.rootFolder.seedFile('photo_REUSEORIGAAA', 'source.webp', {
    mimeType: 'image/webp', sizeBytes: 3
  });

  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: source.id,
    targetFolderId: '',
    title: '再利用'
  }));

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(harness.audit.drive.folderCreates, 0);
  assert.equal(harness.mediaFolders.originalPhotos.files.includes(existing), true);
  assert.equal(harness.mediaFolders.originalPhotos.files.includes(source), true);
  assert.equal(existing.trashed, false);
});

test('ambiguous media folders reject before claim, managed creation, or map commit', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.rootFolder.seedFolder('original_SECONDAAA', 'original');
  const source = harness.rootFolder.seedFile('photo_AMBIGUOUSAA', 'source.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3
  });

  const result = harness.api.saveImportPhotoItem(payload({ sourceDriveFileId: source.id }));

  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_MEDIA_STRUCTURE_AMBIGUOUS');
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
  assert.equal(harness.audit.drive.moves, 0);
  assert.equal(harness.audit.drive.creates, 0);
  assert.equal(harness.audit.drive.trashes, 0);
  assert.equal(harness.sheets.get('import_receipts').rows.length, 1);
});

test('media structure creation failure stops before claim or managed creation', () => {
  const harness = makeHarness({
    sheets: baseSheets(), seedMediaStructure: false, failOriginalFolderCreate: true
  });
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.errorCode, 'DRIVE_MEDIA_STRUCTURE_FAILED');
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
  assert.equal(harness.sheets.get('import_receipts').rows.length, 1);
  assert.equal(harness.audit.drive.creates, 0);
  assert.equal(harness.audit.drive.trashes, 0);
});

test('source move failures stay journaled after map linkage and a transient failure retries', () => {
  const cases = [
    [{ failMoveOnce: true }, 'DRIVE_SOURCE_MOVE_FAILED', true],
    [{ moveWithoutParentChange: true }, 'DRIVE_SOURCE_MOVE_VERIFY_FAILED', false]
  ];
  for (const [sourceMetadata, errorCode, canRecover] of cases) {
    const harness = makeHarness({ sheets: baseSheets() });
    const source = harness.rootFolder.seedFile(`photo_${errorCode}`, 'source.jpg', {
      mimeType: 'image/jpeg', sizeBytes: 3, ...sourceMetadata
    });

    const result = harness.api.saveImportPhotoItem(payload({ sourceDriveFileId: source.id }));

    assert.equal(result.ok, false, errorCode);
    assert.equal(result.errorCode, errorCode);
    assert.equal(harness.sheets.get('map_info').rows.length, 2, errorCode);
    assert.equal(harness.audit.drive.trashes, 0, errorCode);
    const receipt = harness.sheets.get('import_receipts').rows[1];
    assert.equal(receipt[receiptColumn('state')], 'completed', errorCode);
    assert.notEqual(receipt[receiptColumn('fileId')], '', errorCode);
    assert.equal(receipt[receiptColumn('lastErrorCode')], errorCode);
    if (canRecover) {
      const recovered = harness.api.saveImportPhotoItem(payload({ sourceDriveFileId: source.id }));
      assert.equal(recovered.ok, true, JSON.stringify(recovered));
      assert.deepEqual(source.parentFolders, [harness.mediaFolders.originalPhotos]);
      assert.equal(receipt[receiptColumn('lastErrorCode')], '');
    }
  }
});

test('completed replay keeps retrying repeated source move failures until physical archive succeeds', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const metadata = { mimeType: 'image/jpeg', sizeBytes: 3, failMoveAlways: true };
  const source = harness.rootFolder.seedFile('photo_REPEATMOVEAA', 'repeat.jpg', metadata);
  const request = payload({ sourceDriveFileId: source.id });

  assert.equal(harness.api.saveImportPhotoItem(request).errorCode, 'DRIVE_SOURCE_MOVE_FAILED');
  assert.equal(harness.api.saveImportPhotoItem(request).errorCode, 'DRIVE_SOURCE_MOVE_FAILED');
  metadata.failMoveAlways = false;
  const recovered = harness.api.saveImportPhotoItem(request);

  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(recovered.deduplicated, true);
  assert.deepEqual(source.parentFolders.map((folder) => folder.id), [harness.mediaFolders.originalPhotos.id]);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('completed replay keeps retrying source move verification failures until parentage recovers', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const metadata = { mimeType: 'image/jpeg', sizeBytes: 3, moveWithoutParentChange: true };
  const source = harness.rootFolder.seedFile('photo_REPEATVERIFY', 'verify.jpg', metadata);
  const request = payload({ sourceDriveFileId: source.id });

  assert.equal(harness.api.saveImportPhotoItem(request).errorCode, 'DRIVE_SOURCE_MOVE_VERIFY_FAILED');
  assert.equal(harness.api.saveImportPhotoItem(request).errorCode, 'DRIVE_SOURCE_MOVE_VERIFY_FAILED');
  metadata.moveWithoutParentChange = false;
  const recovered = harness.api.saveImportPhotoItem(request);

  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.deepEqual(source.parentFolders.map((folder) => folder.id), [harness.mediaFolders.originalPhotos.id]);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('completed replay archives by physical state after source-move journal writes fail', () => {
  const options = { sheets: baseSheets(), failSourceMoveJournalWrites: true };
  const harness = makeHarness(options);
  const metadata = { mimeType: 'image/jpeg', sizeBytes: 3, failMoveAlways: true };
  const source = harness.rootFolder.seedFile('photo_JOURNALFAILA', 'journal.jpg', metadata);
  const request = payload({ sourceDriveFileId: source.id });

  assert.equal(harness.api.saveImportPhotoItem(request).errorCode, 'DRIVE_SOURCE_MOVE_FAILED');
  metadata.failMoveAlways = false;
  options.failSourceMoveJournalWrites = false;
  const recovered = harness.api.saveImportPhotoItem(request);

  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.deepEqual(source.parentFolders.map((folder) => folder.id), [harness.mediaFolders.originalPhotos.id]);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('map failure leaves the source in Inbox and retry reuses the JPEG before archiving once', () => {
  const harness = makeHarness({ sheets: baseSheets(), failMapAppendOnce: true });
  const source = harness.rootFolder.seedFile('photo_MAPRETRYAAAA', 'source.heic', {
    mimeType: 'image/heic', sizeBytes: 3
  });
  const request = payload({ sourceDriveFileId: source.id, targetFolderId: 'folder-1' });

  const failed = harness.api.saveImportPhotoItem(request);
  assert.equal(failed.errorCode, 'IMPORT_MAP_ROW_FAILED');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.moves, 0);
  assert.deepEqual(source.parentFolders, [harness.rootFolder]);

  const retried = harness.api.saveImportPhotoItem(request);
  assert.equal(retried.ok, true, JSON.stringify(retried));
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.moves, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('local photo save stores its managed JPEG in photos without moving a source', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const result = harness.api.saveImportPhotoItem(payload({ sourceDriveFileId: '' }));
  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(harness.audit.drive.folderCreates, 0);
  assert.equal(harness.audit.drive.moves, 0);
  assert.equal(harness.mediaFolders.photos.files.filter((file) => !file.trashed).length, 1);
});

test('attach-existing-pin updates only G H and N on one photo-less row and replays idempotently', () => {
  const before = existingPhotoLessPinRow();
  const sheets = baseSheets();
  sheets.map_info.push(before.slice());
  const harness = makeHarness({ sheets });

  const result = harness.api.saveImportPhotoItem(attachPayload());

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.pin.id, 'pin-existing-0001');
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  const after = harness.sheets.get('map_info').rows[1];
  before.forEach((value, index) => {
    if ([6, 7, 13].includes(index)) return;
    assert.equal(after[index], value, `column ${index + 1}`);
  });
  assert.match(after[6], /^file-/);
  assert.match(after[7], /thumbnail\?id=/);
  assert.match(after[13], /^\d{4}-\d{2}-\d{2}T/);
  assert.deepEqual(
    harness.audit.writes.filter((write) => write.sheet === 'map_info').map((write) => ({
      method: write.method, column: write.column, numColumns: write.numColumns
    })),
    [
      { method: 'setValues', column: 7, numColumns: 2 },
      { method: 'setValues', column: 14, numColumns: 1 }
    ]
  );
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('pinId')], 'pin-existing-0001');
  assert.equal(receipt[receiptColumn('targetFolderId')], harness.mediaFolders.photos.id);
  assert.equal(receipt[receiptColumn('state')], 'completed');

  const replay = harness.api.saveImportPhotoItem(attachPayload());
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.pin.id, 'pin-existing-0001');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('different attach operations use distinct deterministic managed JPEG names', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const first = harness.api.normalizeImportPhotoPayload_(attachPayload());
  const second = harness.api.normalizeImportPhotoPayload_(attachPayload({
    jobId: 'attach-job-2', itemId: 'attach-item-2',
    idempotencyKey: 'attach-job-2:attach-item-2'
  }));

  const firstName = harness.api.importPhotoAttachTempFileName_(first.idempotencyKeyHash);
  const secondName = harness.api.importPhotoAttachTempFileName_(second.idempotencyKeyHash);
  assert.notEqual(firstName, secondName);
  assert.equal(firstName, harness.api.importPhotoAttachTempFileName_(first.idempotencyKeyHash));
  assert.match(firstName, /^__drop_pin_attach_[0-9a-f]{32}\.jpg$/);
});

test('attach-existing-pin decodes a protected existing title before managed JPEG rename', () => {
  const sheets = baseSheets();
  sheets.config[2][1] = 'true';
  const protectedTitle = '\u200Bdtp-sheet:v1:v:=既存ピン';
  sheets.map_info.push(existingPhotoLessPinRow({ title: protectedTitle }));
  const harness = makeHarness({ sheets });

  const result = harness.api.saveImportPhotoItem(attachPayload());

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.pin.title, '=既存ピン');
  assert.equal(harness.filesById.get(result.pin.fileId).name, '=既存ピン.jpg');
  assert.equal(harness.sheets.get('map_info').rows[1][1], protectedTitle);
});

test('attach-existing-pin rejects missing, already photographed, and stale targets before Drive writes', () => {
  const missing = makeHarness({ sheets: baseSheets() });
  const notFound = missing.api.saveImportPhotoItem(attachPayload());
  assert.equal(notFound.errorCode, 'PIN_PHOTO_ATTACH_TARGET_NOT_FOUND');
  assert.equal(notFound.error, '写真を追加するピンが見つかりません。画面を再読み込みしてください。');
  assert.equal(missing.audit.drive.creates, 0);

  for (const existingPhoto of [
    { fileId: 'existing-file' },
    { imageUrl: 'https://example.com/existing.jpg' }
  ]) {
    const sheets = baseSheets();
    sheets.map_info.push(existingPhotoLessPinRow(existingPhoto));
    const harness = makeHarness({ sheets });
    const result = harness.api.saveImportPhotoItem(attachPayload());
    assert.equal(result.errorCode, 'PIN_PHOTO_ATTACH_ALREADY_HAS_PHOTO');
    assert.equal(result.error, 'このピンには既に写真が登録されています。');
    assert.equal(harness.audit.drive.creates, 0);
  }

  const staleSheets = baseSheets();
  staleSheets.map_info.push(existingPhotoLessPinRow({ updatedAt: '2026/07/11 12:00:01' }));
  const stale = makeHarness({ sheets: staleSheets });
  const conflict = stale.api.saveImportPhotoItem(attachPayload());
  assert.equal(conflict.errorCode, 'PIN_PHOTO_ATTACH_CONFLICT');
  assert.equal(conflict.error, 'ピンが別の操作で更新されました。画面を再読み込みしてから再試行してください。');
  assert.equal(stale.audit.drive.creates, 0);
});

test('attach-existing-pin compensates its root JPEG when the final target check conflicts', () => {
  const before = existingPhotoLessPinRow();
  const sheets = baseSheets();
  sheets.map_info.push(before.slice());
  const harness = makeHarness({
    sheets,
    afterCreateFile({ sheets: currentSheets }) {
      currentSheets.get('map_info').rows[1][13] = '2026/07/11 12:00:01';
    }
  });

  const result = harness.api.saveImportPhotoItem(attachPayload());

  assert.equal(result.errorCode, 'PIN_PHOTO_ATTACH_CONFLICT');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.trashes, 1);
  assert.equal(harness.audit.drive.moves, 0);
  assert.equal(harness.sheets.get('map_info').rows[1][6], '');
  assert.equal(harness.sheets.get('map_info').rows[1][7], '');
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('state')], 'failed');
  assert.equal(receipt[receiptColumn('fileId')], '');
});

test('Drive photo attach moves one source only after the final target check and keeps one map row', () => {
  const before = existingPhotoLessPinRow();
  const sheets = baseSheets();
  sheets.map_info.push(before.slice());
  let photoAtManagedCreate = '';
  const harness = makeHarness({
    sheets,
    afterCreateFile({ sheets: currentSheets }) {
      photoAtManagedCreate = String(currentSheets.get('map_info').rows[1][6] || '');
    }
  });
  const source = harness.rootFolder.seedFile('photo_ATTACHDRIVEA', 'attach.heic', {
    mimeType: 'image/heic', sizeBytes: 3
  });

  const result = harness.api.saveImportPhotoItem(attachPayload({
    sourceDriveFileId: source.id,
    targetFolderId: 'folder-1'
  }));

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(result.pin.id, 'pin-existing-0001');
  assert.equal(photoAtManagedCreate, '');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.moves, 1);
  assert.deepEqual(source.parentFolders, [harness.mediaFolders.originalPhotos]);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(
    harness.sheets.get('import_receipts').rows[1][receiptColumn('targetFolderId')],
    harness.mediaFolders.photos.id
  );
});

test('Drive photo attach conflict compensates before linkage and move failure keeps the linked photo', () => {
  const conflictSheets = baseSheets();
  conflictSheets.map_info.push(existingPhotoLessPinRow());
  const conflictHarness = makeHarness({
    sheets: conflictSheets,
    afterCreateFile({ sheets }) {
      sheets.get('map_info').rows[1][13] = '2026/07/11 12:00:01';
    }
  });
  const conflictSource = conflictHarness.rootFolder.seedFile(
    'photo_ATTACHCONFLA', 'conflict.jpg', { mimeType: 'image/jpeg', sizeBytes: 3 }
  );
  const conflict = conflictHarness.api.saveImportPhotoItem(attachPayload({
    sourceDriveFileId: conflictSource.id
  }));
  assert.equal(conflict.errorCode, 'PIN_PHOTO_ATTACH_CONFLICT');
  assert.equal(conflictHarness.audit.drive.moves, 0);
  assert.equal(conflictHarness.audit.drive.trashes, 1);
  assert.deepEqual(conflictSource.parentFolders.map((folder) => folder.name), ['Root']);

  const moveSheets = baseSheets();
  moveSheets.map_info.push(existingPhotoLessPinRow());
  const moveHarness = makeHarness({ sheets: moveSheets });
  const moveSource = moveHarness.rootFolder.seedFile('photo_ATTACHMOVEAA', 'move.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, failMoveAlways: true
  });
  const moveFailure = moveHarness.api.saveImportPhotoItem(attachPayload({
    sourceDriveFileId: moveSource.id
  }));
  assert.equal(moveFailure.errorCode, 'PIN_PHOTO_ATTACH_SOURCE_MOVE_FAILED');
  assert.equal(moveHarness.audit.drive.trashes, 0);
  assert.notEqual(moveHarness.sheets.get('map_info').rows[1][6], '');
  assert.notEqual(moveHarness.sheets.get('map_info').rows[1][7], '');
});

test('photo attach map failure leaves the source in Inbox and retry moves it after linkage', () => {
  const sheets = baseSheets();
  sheets.map_info.push(existingPhotoLessPinRow());
  const harness = makeHarness({ sheets, failMapAttachUpdateOnce: true });
  const source = harness.rootFolder.seedFile('photo_ATTACHMAPAAA', 'map.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3
  });
  const request = attachPayload({ sourceDriveFileId: source.id });

  const failed = harness.api.saveImportPhotoItem(request);
  assert.equal(failed.errorCode, 'PIN_PHOTO_ATTACH_MAP_UPDATE_FAILED');
  assert.equal(failed.error, '写真ファイルは準備できましたが、ピンへ設定できませんでした。再試行してください。');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.moves, 0);
  assert.equal(harness.audit.drive.trashes, 0);
  assert.equal(harness.sheets.get('map_info').rows[1][6], '');

  const recovered = harness.api.saveImportPhotoItem(request);
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(recovered.pin.id, 'pin-existing-0001');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.moves, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('photo attach repairs a partial G H write without a new row or JPEG', () => {
  const sheets = baseSheets();
  sheets.map_info.push(existingPhotoLessPinRow());
  const harness = makeHarness({ sheets, failMapAttachTimestampOnce: true });

  const failed = harness.api.saveImportPhotoItem(attachPayload());
  assert.equal(failed.errorCode, 'PIN_PHOTO_ATTACH_MAP_UPDATE_FAILED');
  assert.match(harness.sheets.get('map_info').rows[1][6], /^file-/);
  assert.equal(harness.sheets.get('map_info').rows[1][13], '2026/07/11 11:59:00');

  const recovered = harness.api.saveImportPhotoItem(attachPayload());
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(recovered.pin.id, 'pin-existing-0001');
  assert.match(harness.sheets.get('map_info').rows[1][13], /^\d{4}-\d{2}-\d{2}T/);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('attach-existing-pin maps managed JPEG creation denial to its operation-specific error', () => {
  const sheets = baseSheets();
  sheets.map_info.push(existingPhotoLessPinRow());
  const harness = makeHarness({ sheets, failCreatePermission: true });

  const result = harness.api.saveImportPhotoItem(attachPayload());

  assert.equal(result.errorCode, 'PIN_PHOTO_ATTACH_FILE_CREATE_FAILED');
  assert.equal(result.error, '表示用の写真を作成できませんでした。Driveの保存権限を確認してください。');
  assert.equal(harness.sheets.get('map_info').rows[1][6], '');
});

test('readable anonymous JPEG saves a new managed JPEG without changing the source', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.sheets.get('config').rows[2][1] = 'true';
  const source = harness.rootFolder.seedFile('photo_READONLYAAAA', 'original-name.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'ANYONE_WITH_LINK', failRenameAlways: true
  });
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_READONLYAAAA', targetFolderId: 'folder-1', lat: null, lng: null,
    title: '改名しない写真'
  }));

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.notEqual(result.pin.fileId, 'photo_READONLYAAAA');
  assert.equal(result.pin.lat, null);
  assert.equal(result.pin.lng, null);
  assert.equal(source.name, 'original-name.jpg');
  assert.equal(harness.audit.drive.renameIds.includes('photo_READONLYAAAA'), false);
  assert.equal(harness.audit.drive.shareIds.includes('photo_READONLYAAAA'), false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('緯度')], '');
  assert.equal(harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('経度')], '');
});

test('private JPEG with no source sharing permission saves a placed pin through a managed JPEG', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const source = harness.rootFolder.seedFile('photo_PRIVATEAAAAA', 'private.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE', failSharingAlways: true
  });
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_PRIVATEAAAAA', targetFolderId: '', lat: 35.25, lng: 139.75,
    title: '管理コピー'
  }));

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.notEqual(result.pin.fileId, 'photo_PRIVATEAAAAA');
  assert.equal(result.pin.lat, 35.25);
  assert.equal(result.pin.lng, 139.75);
  assert.equal(source.name, 'private.jpg');
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.shareIds.includes('photo_PRIVATEAAAAA'), false);
  assert.deepEqual(harness.audit.drive.shareIds, [result.pin.fileId]);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('sourceDriveFileId')], 'photo_PRIVATEAAAAA');
  assert.equal(receipt[receiptColumn('fileId')], result.pin.fileId);
  assert.equal(harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('緯度')], 35.25);
  assert.equal(harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('経度')], 139.75);
});

test('managed JPEG inheriting anonymous view access skips a denied redundant sharing write', () => {
  const harness = makeHarness({
    sheets: baseSheets(),
    managedSharingAccess: 'ANYONE_WITH_LINK',
    managedSharingPermission: 'VIEW',
    failManagedSharingAlways: true
  });

  const result = harness.api.saveImportPhotoItem(payload());

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.equal(harness.audit.drive.shares, 0);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(
    harness.sheets.get('import_receipts').rows[1][receiptColumn('state')],
    'completed'
  );
});

test('JPEG, PNG, WebP, HEIC, and HEIF Drive sources each create one photos JPEG and archive the source', () => {
  [
    { id: 'photo_JPEGSOURCEAAA', name: 'photo.jpg', mimeType: 'image/jpeg' },
    { id: 'photo_PNGSOURCEAAAA', name: 'private.png', mimeType: 'image/png' },
    { id: 'photo_WEBPSOURCEAAA', name: 'private.webp', mimeType: 'image/webp' },
    { id: 'photo_HEICSOURCEAAA', name: 'private.heic', mimeType: 'image/heic' },
    { id: 'photo_HEIFSOURCEAAA', name: 'private.heif', mimeType: 'image/heif' }
  ].forEach((sample) => {
    const harness = makeHarness({ sheets: baseSheets() });
    const source = harness.rootFolder.seedFile(sample.id, sample.name, {
      mimeType: sample.mimeType, sizeBytes: 3, sharingAccess: 'PRIVATE'
    });
    const result = harness.api.saveImportPhotoItem(payload({
      jobId: `job-${sample.name}`,
      itemId: `item-${sample.name}`,
      idempotencyKey: `job-${sample.name}:item-${sample.name}`,
      sourceDriveFileId: sample.id,
      targetFolderId: 'folder-1'
    }));

    assert.equal(result.ok, true, sample.name);
    assert.notEqual(result.pin.fileId, sample.id, sample.name);
    assert.equal(harness.audit.drive.creates, 1, sample.name);
    assert.equal(source.name, sample.name, sample.name);
    assert.equal(source.trashed, false, sample.name);
    assert.equal(harness.audit.drive.renameIds.includes(sample.id), false, sample.name);
    assert.equal(harness.audit.drive.shareIds.includes(sample.id), false, sample.name);
    assert.deepEqual(source.parentFolders, [harness.mediaFolders.originalPhotos], sample.name);
    const receipt = harness.sheets.get('import_receipts').rows[1];
    assert.equal(receipt[receiptColumn('sourceDriveFileId')], sample.id, sample.name);
    assert.equal(receipt[receiptColumn('fileId')], result.pin.fileId, sample.name);
    assert.equal(receipt[receiptColumn('targetFolderId')], harness.mediaFolders.photos.id, sample.name);
  });
});

test('Workspace link-sharing denial is specific, non-retryable, journaled, and safely logged', () => {
  const harness = makeHarness({
    sheets: baseSheets(),
    failSharingAlways: true,
    sharingErrorMessage: 'Workspace policy photo_SECRET123456 https://drive.google.com/private '
      + 'editToken=valid-token teacher@example.com a.jpg R1 '
      + 'data:image/jpeg;charset=utf-8;base64,aGVs bG8='
  });
  harness.rootFolder.seedFile('photo_POLICYAAAAAA', 'policy.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_POLICYAAAAAA', targetFolderId: ''
  }));

  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_LINK_SHARING_DENIED');
  assert.equal(result.retryable, false);
  assert.match(result.error, /Google Drive共有ポリシー/);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('lastErrorCode')], 'DRIVE_LINK_SHARING_DENIED');
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
  assert.equal(harness.audit.errors.length > 0, true);
  const logged = harness.audit.errors.join('\n');
  assert.match(logged, /"operation":"drive-photo-import"/);
  assert.match(logged, /"stage":"sharing"/);
  assert.doesNotMatch(
    logged,
    /photo_SECRET123456|drive\.google\.com|data:image|valid-token|teacher@example\.com|a\.jpg|R1|aGVs|bG8/
  );
});

test('administrator-disabled link sharing is also a permanent policy denial', () => {
  const harness = makeHarness({
    sheets: baseSheets(),
    failSharingAlways: true,
    sharingErrorMessage: 'Sharing has been disabled by your administrator'
  });
  harness.rootFolder.seedFile('photo_ADMINPOLICYA', 'admin-policy.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });

  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_ADMINPOLICYA', targetFolderId: ''
  }));

  assert.equal(result.errorCode, 'DRIVE_LINK_SHARING_DENIED');
  assert.equal(result.retryable, false);
});

test('HEIC source trash failure cannot turn a completed pin into a failed response', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const source = harness.rootFolder.seedFile('photo_HEICKEEPAAAA', 'keep.heic', {
    mimeType: 'image/heic', sizeBytes: 3, sharingAccess: 'PRIVATE', failTrashAlways: true
  });
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_HEICKEEPAAAA', targetFolderId: '', title: '保持するHEIC'
  }));

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.notEqual(result.pin.fileId, 'photo_HEICKEEPAAAA');
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.trashIds.includes('photo_HEICKEEPAAAA'), false);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.sheets.get('import_receipts').rows[1][receiptColumn('state')], 'completed');
});

test('managed retry does not depend on or mutate the source sharing state', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const source = harness.rootFolder.seedFile('photo_CHANGEDAAAAA', 'changed.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'ANYONE_WITH_LINK'
  });
  const request = payload({ sourceDriveFileId: 'photo_CHANGEDAAAAA', targetFolderId: '' });
  const first = harness.api.saveImportPhotoItem(request);
  assert.equal(first.ok, true);
  harness.sheets.get('map_info').rows.splice(1);
  source.setSharing('PRIVATE');
  const mutationsBeforeRetry = {
    renames: harness.audit.drive.renames,
    shares: harness.audit.drive.shares,
    trashes: harness.audit.drive.trashes
  };

  const recovered = harness.api.saveImportPhotoItem(request);
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(recovered.pin.fileId, first.pin.fileId);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(source.name, 'changed.jpg');
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.renameIds.includes('photo_CHANGEDAAAAA'), false);
  assert.equal(harness.audit.drive.trashIds.includes('photo_CHANGEDAAAAA'), false);
  assert.equal(harness.audit.drive.renames >= mutationsBeforeRetry.renames, true);
  assert.equal(harness.audit.drive.shares >= mutationsBeforeRetry.shares, true);
  assert.equal(harness.audit.drive.trashes, mutationsBeforeRetry.trashes);
});

test('managed-copy rename failure is stage-specific and never compensates the persisted copy or source', () => {
  const harness = makeHarness({ sheets: baseSheets(), failManagedRenameAlways: true });
  const source = harness.rootFolder.seedFile('photo_RENAMECOPYAA', 'source.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_RENAMECOPYAA', targetFolderId: '', title: '名前確定'
  }));

  assert.equal(result.errorCode, 'DRIVE_MANAGED_COPY_FINALIZE_FAILED');
  assert.equal(result.retryable, true);
  assert.equal(source.name, 'source.jpg');
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.trashes, 0);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('state')], 'failed');
  assert.equal(receipt[receiptColumn('fileId')], 'file-0000000001');
  assert.equal(receipt[receiptColumn('lastErrorCode')], 'DRIVE_MANAGED_COPY_FINALIZE_FAILED');
  assert.match(harness.audit.errors.join('\n'), /"stage":"managed-copy-finalize"/);
});

test('anonymous PNG Drive source uses a photos managed JPEG and archives the unchanged source', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const source = harness.rootFolder.seedFile('photo_AAAAAAAAAAA', 'before.PNG', {
    mimeType: 'image/png', sizeBytes: 3
  });
  const request = payload({
    sourceDriveFileId: 'photo_AAAAAAAAAAA', targetFolderId: '', title: '安全/タイトル'
  });
  const first = harness.api.saveImportPhotoItem(request);
  assert.equal(first.ok, true);
  assert.notEqual(first.pin.fileId, 'photo_AAAAAAAAAAA');
  assert.equal(source.name, 'before.PNG');
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.renameIds.includes('photo_AAAAAAAAAAA'), false);
  assert.equal(harness.audit.drive.shareIds.includes('photo_AAAAAAAAAAA'), false);
  assert.deepEqual(source.parentFolders, [harness.mediaFolders.originalPhotos]);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('sourceDriveFileId')], 'photo_AAAAAAAAAAA');
  assert.equal(receipt[receiptColumn('fileId')], first.pin.fileId);
  assert.equal(receipt[receiptColumn('targetFolderId')], harness.mediaFolders.photos.id);
  assert.equal(receipt[receiptColumn('state')], 'completed');

  const replay = harness.api.saveImportPhotoItem(request);
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);

  const conflict = harness.api.saveImportPhotoItem(payload({ sourceDriveFileId: 'photo_BBBBBBBBBBB' }));
  assert.equal(conflict.errorCode, 'IDEMPOTENCY_PAYLOAD_CONFLICT');
  assert.equal(conflict.retryable, false);
});

test('a different job cannot link the same surviving Drive source while the same key still deduplicates', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.rootFolder.seedFile('photo_AAAAAAAAAAA', 'source.jpg');
  const firstRequest = payload({ sourceDriveFileId: 'photo_AAAAAAAAAAA', targetFolderId: '' });
  const first = harness.api.saveImportPhotoItem(firstRequest);
  assert.equal(first.ok, true);
  assert.equal(harness.api.saveImportPhotoItem(firstRequest).deduplicated, true);

  const second = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-2', itemId: 'item-2', idempotencyKey: 'job-2:item-2',
    sourceDriveFileId: 'photo_AAAAAAAAAAA', targetFolderId: ''
  }));
  assert.equal(second.errorCode, 'DRIVE_SOURCE_ALREADY_LINKED');
  assert.equal(second.retryable, false);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.audit.drive.creates, 1);
});

test('a completed Drive source remains excluded after its pin is deleted', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const source = harness.rootFolder.seedFile('photo_REIMPORTAAAA', 'reimport.jpg');
  const first = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_REIMPORTAAAA', targetFolderId: 'folder-1'
  }));
  assert.equal(first.ok, true);

  harness.sheets.get('map_info').rows.splice(1);
  const second = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-reimport', itemId: 'item-reimport',
    idempotencyKey: 'job-reimport:item-reimport',
    sourceDriveFileId: 'photo_REIMPORTAAAA', targetFolderId: 'folder-1'
  }));

  assert.equal(second.ok, false);
  assert.equal(second.errorCode, 'DRIVE_SOURCE_ALREADY_LINKED');
  assert.equal(second.retryable, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(source.name, 'reimport.jpg');
  assert.equal(source.trashed, false);
  assert.deepEqual(source.parentFolders, [harness.mediaFolders.originalPhotos]);
});

test('a live managed map file without a completed receipt is not accepted as a source', () => {
  const sheets = baseSheets();
  sheets.map_info.push([
    '', '既存', '', 35, 139, '#e53935', 'photo_ORPHANAAAAAA',
    'https://drive.google.com/thumbnail?id=photo_ORPHANAAAAAA', 'existing-pin',
    '', '', '', '', '', 'photo'
  ]);
  const harness = makeHarness({ sheets });
  const source = harness.rootFolder.seedFile('photo_ORPHANAAAAAA', 'orphan.jpg');
  const writesBefore = harness.audit.writes.length;
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_ORPHANAAAAAA', targetFolderId: ''
  }));
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_SOURCE_ALREADY_LINKED');
  assert.equal(result.retryable, false);
  assert.equal(harness.audit.writes.length, writesBefore);
  assert.equal(harness.audit.drive.creates, 0);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.deepEqual(source.parentFolders, [harness.rootFolder]);
});

test('completed receipts from the legacy managed-copy path remain replayable without touching either file', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const request = payload({ sourceDriveFileId: 'photo_LEGACYAAAAAA', targetFolderId: 'folder-1' });
  const normalized = harness.api.normalizeImportPhotoPayload_(request);
  normalized.targetFolderId = 'folder-1';
  const receipt = Array(RECEIPT_HEADERS.length).fill('');
  receipt[receiptColumn('idempotencyKey')] = normalized.idempotencyKeyHash;
  receipt[receiptColumn('jobId')] = normalized.jobId;
  receipt[receiptColumn('itemId')] = normalized.itemId;
  receipt[receiptColumn('payloadHash')] = harness.api.hashImportPayload_(normalized);
  receipt[receiptColumn('state')] = 'completed';
  receipt[receiptColumn('pinId')] = 'legacy-pin';
  receipt[receiptColumn('targetFolderId')] = 'folder-1';
  receipt[receiptColumn('tempFileName')] = '__drop_pin_import_legacy-pin.jpg';
  receipt[receiptColumn('fileId')] = 'legacy-copy-file';
  receipt[receiptColumn('imageUrl')] = 'https://drive.google.com/thumbnail?id=legacy-copy-file&sz=w1920';
  receipt[receiptColumn('folderUrl')] = 'https://drive.google.com/drive/folders/folder-1';
  receipt[receiptColumn('sourceDriveFileId')] = 'photo_LEGACYAAAAAA';
  harness.sheets.get('import_receipts').rows.push(receipt);
  harness.sheets.get('map_info').rows.push([
    '2026/07/11 12:00:00', '写真', '説明', 35.5, 139.5, '#e53935',
    'legacy-copy-file', receipt[receiptColumn('imageUrl')], 'legacy-pin',
    'https://example.com', '', '観察', '2026-07-11T10:30', '', 'photo'
  ]);

  const replay = harness.api.saveImportPhotoItem(request);
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.pin.fileId, 'legacy-copy-file');
  assert.equal(harness.audit.drive.creates, 0);
  assert.equal(harness.audit.drive.renames, 0);
  assert.equal(harness.audit.drive.shares, 0);
  assert.equal(harness.audit.drive.trashes, 0);
});

test('HEIC Drive source creates one managed JPEG, links it, and preserves the original', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const source = harness.rootFolder.seedFile('photo_HEICAAAAAAA', 'original.HEIC', {
    mimeType: 'image/heic', sizeBytes: 3
  });
  const request = payload({
    sourceDriveFileId: 'photo_HEICAAAAAAA', targetFolderId: '', title: '変換/写真'
  });
  const first = harness.api.saveImportPhotoItem(request);
  assert.equal(first.ok, true);
  assert.notEqual(first.pin.fileId, 'photo_HEICAAAAAAA');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.mediaFolders.photos.files.some((file) => file.id === first.pin.fileId), true);
  assert.equal(harness.filesById.get(first.pin.fileId).name, 'photo.jpg');
  assert.equal(source.trashed, false);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('state')], 'completed');
  assert.equal(receipt[receiptColumn('fileId')], first.pin.fileId);
  assert.equal(receipt[receiptColumn('sourceDriveFileId')], 'photo_HEICAAAAAAA');
  assert.equal(receipt[receiptColumn('targetFolderId')], harness.mediaFolders.photos.id);

  const replay = harness.api.saveImportPhotoItem(request);
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.renames, 1);
  assert.equal(harness.audit.drive.trashes, 0);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('HEIC map failure preserves the Inbox source and retry reuses one managed JPEG before archiving', () => {
  const harness = makeHarness({ sheets: baseSheets(), failMapAppendOnce: true });
  const source = harness.rootFolder.seedFile('photo_HEICBBBBBBB', 'retry.heif', {
    mimeType: 'image/heif', sizeBytes: 3
  });
  const request = payload({
    sourceDriveFileId: 'photo_HEICBBBBBBB', targetFolderId: '', title: '再試行'
  });
  const failed = harness.api.saveImportPhotoItem(request);
  assert.equal(failed.errorCode, 'IMPORT_MAP_ROW_FAILED');
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.moves, 0);

  const retry = harness.api.saveImportPhotoItem(request);
  assert.equal(retry.ok, true);
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.moves, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('HEIC create response loss preserves the source and retry finds the one deterministic JPEG', () => {
  const harness = makeHarness({ sheets: baseSheets(), failCreateAfterFileOnce: true });
  const source = harness.rootFolder.seedFile('photo_HEICCCCCCCC', 'response-loss.heic', {
    mimeType: 'image/heic', sizeBytes: 3
  });
  const request = payload({
    sourceDriveFileId: 'photo_HEICCCCCCCC', targetFolderId: '', title: '応答喪失'
  });
  const lost = harness.api.saveImportPhotoItem(request);
  assert.equal(lost.errorCode, 'DRIVE_MANAGED_COPY_CREATE_FAILED');
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.creates, 1);

  const retry = harness.api.saveImportPhotoItem(request);
  assert.equal(retry.ok, true);
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('HEIC source cleanup is disabled so trash restrictions never affect the completed link', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const source = harness.rootFolder.seedFile('photo_HEICDDDDDDD', 'delete-retry.heic', {
    mimeType: 'image/heic', sizeBytes: 3, failTrashOnce: true
  });
  const request = payload({
    sourceDriveFileId: 'photo_HEICDDDDDDD', targetFolderId: '', title: '削除再試行'
  });
  const first = harness.api.saveImportPhotoItem(request);
  assert.equal(first.ok, true);
  assert.equal(source.trashed, false);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.sheets.get('import_receipts').rows[1][receiptColumn('state')], 'completed');

  const retry = harness.api.saveImportPhotoItem(request);
  assert.equal(retry.ok, true);
  assert.equal(retry.deduplicated, true);
  assert.equal(source.trashed, false);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.renames, 1);
  assert.equal(harness.audit.drive.trashes, 0);
});

test('Drive source save revalidates root containment, trash, type, and size before mutation', () => {
  const cases = [
    ['outside', (harness, id) => harness.defaultFolder.seedFile(id, 'outside.jpg')],
    ['trashed', (harness, id) => harness.rootFolder.seedFile(id, 'trashed.jpg', { trashed: true })],
    ['unsupported', (harness, id) => harness.rootFolder.seedFile(id, 'unsafe.gif', { mimeType: 'image/gif' })],
    ['oversize', (harness, id) => harness.rootFolder.seedFile(id, 'large.jpg', { sizeBytes: 15 * 1024 * 1024 + 1 })]
  ];
  cases.forEach(([label, seed], index) => {
    const harness = makeHarness({ sheets: baseSheets() });
    const sourceId = `source_${label}_${String(index).padStart(10, '0')}`;
    seed(harness, sourceId);
    const result = harness.api.saveImportPhotoItem(payload({
      sourceDriveFileId: sourceId,
      targetFolderId: ''
    }));
    assert.equal(result.errorCode, 'IMPORT_DRIVE_SOURCE_INVALID', label);
    assert.equal(harness.audit.drive.creates, 0, label);
    assert.equal(harness.audit.drive.renames, 0, label);
    assert.equal(harness.audit.drive.shares, 0, label);
    assert.equal(harness.sheets.get('map_info').rows.length, 1, label);
  });
});

test('Drive photo save rejects nested and reserved-folder sources before new Drive or Sheet writes', () => {
  const cases = [
    ['nested', (harness, id) => {
      const folder = harness.rootFolder.seedFolder('nested_source_AAA', 'Nested');
      return folder.seedFile(id, 'nested.jpg', { mimeType: 'image/jpeg', sizeBytes: 3 });
    }],
    ['photos', (harness, id) => harness.mediaFolders.photos.seedFile(
      id, 'managed.jpg', { mimeType: 'image/jpeg', sizeBytes: 3 }
    )],
    ['original/photos', (harness, id) => harness.mediaFolders.originalPhotos.seedFile(
      id, 'archived.jpg', { mimeType: 'image/jpeg', sizeBytes: 3 }
    )]
  ];

  cases.forEach(([label, seed], index) => {
    const harness = makeHarness({ sheets: baseSheets() });
    const sourceId = `nested_guard_${String(index).padStart(10, '0')}`;
    seed(harness, sourceId);
    const writesBefore = harness.audit.writes.length;

    const result = harness.api.saveImportPhotoItem(payload({
      sourceDriveFileId: sourceId,
      targetFolderId: ''
    }));

    assert.equal(result.errorCode, 'IMPORT_DRIVE_SOURCE_INVALID', label);
    assert.equal(harness.audit.writes.length, writesBefore, label);
    assert.equal(harness.audit.drive.creates, 0, label);
    assert.equal(harness.audit.drive.moves, 0, label);
    assert.equal(harness.sheets.get('map_info').rows.length, 1, label);
    assert.equal(harness.sheets.get('import_receipts').rows.length, 1, label);
  });
});

test('matched photo coordinates survive a lost success response in the real save harness', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const request = payload({
    jobId: 'matched-job', itemId: 'matched-item',
    idempotencyKey: 'matched-job:matched-item', lat: 15, lng: 25
  });

  const lostResponse = harness.api.saveImportPhotoItem(request);
  assert.equal(lostResponse.ok, true);
  const retried = harness.api.saveImportPhotoItem(request);
  assert.equal(retried.ok, true);
  assert.equal(retried.deduplicated, true);
  assert.equal(retried.pin.id, lostResponse.pin.id);
  assert.equal(retried.pin.lat, 15);
  assert.equal(retried.pin.lng, 25);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('緯度')], 15);
  assert.equal(harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('経度')], 25);
});

test('receipt lookup scans only the key column before reading the matching row', () => {
  const existing = Array.from({ length: 40 }, (_unused, index) => [
    `unrelated-${index}`, `job-${index}`, `item-${index}`, `hash-${index}`, 'completed'
  ]);
  const harness = makeHarness({ sheets: {
    ...baseSheets(false),
    import_receipts: [RECEIPT_HEADERS, ...existing]
  } });
  harness.api.saveImportPhotoItem(payload());
  const receiptScans = harness.audit.reads.filter((read) =>
    read.sheet === 'import_receipts' && read.row === 2 && read.numRows >= 40
  );
  assert.equal(receiptScans.length > 0, true);
  assert.equal(receiptScans.every((read) => read.column === 1 && read.numColumns === 1), true);
});

test('receipt stores hashes and metadata but no sensitive payload bodies', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.api.saveImportPhotoItem(payload());
  const row = harness.sheets.get('import_receipts').rows[1];
  assert.equal(row.length, RECEIPT_HEADERS.length);
  assert.match(String(row[receiptColumn('idempotencyKey')]), /^[0-9a-f]{64}$/);
  assert.match(String(row[receiptColumn('payloadHash')]), /^[0-9a-f]{64}$/);
  const serialized = JSON.stringify(row);
  assert.equal(serialized.includes('data:image'), false);
  assert.equal(serialized.includes('説明'), false);
  assert.equal(serialized.includes('観察'), false);
  assert.equal(serialized.includes('valid-token'), false);
  assert.equal(row[receiptColumn('lastErrorCode')], '');
});

test('live foreign lease returns in-progress while an expired lease can reclaim and finish', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const saved = harness.api.saveImportPhotoItem(payload());
  const receipt = harness.sheets.get('import_receipts').rows[1];
  harness.sheets.get('map_info').rows.splice(1);
  receipt[receiptColumn('state')] = 'reserved';
  receipt[receiptColumn('leaseOwner')] = 'other-owner';
  receipt[receiptColumn('leaseUntil')] = '2999-01-01T00:00:00.000Z';
  const busy = harness.api.saveImportPhotoItem(payload());
  assert.equal(busy.errorCode, 'IMPORT_ITEM_IN_PROGRESS');
  assert.equal(busy.retryable, true);
  assert.equal(harness.audit.drive.creates, 1);

  receipt[receiptColumn('leaseUntil')] = '2000-01-01T00:00:00.000Z';
  const recovered = harness.api.saveImportPhotoItem(payload());
  assert.equal(recovered.ok, true);
  assert.equal(recovered.pin.id, saved.pin.id);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(receipt[receiptColumn('state')], 'completed');
});

test('Drive-create interruption loses the lease safely and retry reuses the temporary file', () => {
  let changed = false;
  const harness = makeHarness({
    sheets: baseSheets(),
    afterCreateFile({ sheets }) {
      if (changed) return;
      changed = true;
      const receipt = sheets.get('import_receipts').rows[1];
      receipt[receiptColumn('leaseOwner')] = 'new-owner';
      receipt[receiptColumn('leaseUntil')] = '2999-01-01T00:00:00.000Z';
    }
  });
  const lost = harness.api.saveImportPhotoItem(payload());
  assert.equal(lost.errorCode, 'IMPORT_ITEM_LEASE_LOST');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.trashes, 0);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  const tempName = receipt[receiptColumn('tempFileName')];
  assert.equal(harness.mediaFolders.photos.files[0].name, tempName);

  receipt[receiptColumn('leaseUntil')] = '2000-01-01T00:00:00.000Z';
  const retried = harness.api.saveImportPhotoItem(payload());
  assert.equal(retried.ok, true);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.searches >= 1, true);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('Drive create response loss is recovered by deterministic temporary file reuse', () => {
  const harness = makeHarness({ sheets: baseSheets(), failCreateAfterFileOnce: true });
  const failed = harness.api.saveImportPhotoItem(payload());
  assert.equal(failed.errorCode, 'DRIVE_MANAGED_COPY_CREATE_FAILED');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.trashes, 0);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);

  const recovered = harness.api.saveImportPhotoItem(payload());
  assert.equal(recovered.ok, true);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.mediaFolders.photos.files.filter((file) => !file.trashed).length, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('receipt file metadata write failure compensates the newly-created owned file', () => {
  const harness = makeHarness({ sheets: baseSheets(), failReceiptFileSavedWriteOnce: true });
  const failed = harness.api.saveImportPhotoItem(payload());
  assert.equal(failed.errorCode, 'IMPORT_ITEM_SAVE_FAILED');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.audit.drive.trashes, 1);
  assert.equal(harness.mediaFolders.photos.files.filter((file) => !file.trashed).length, 0);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
  const failedReceipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(failedReceipt[receiptColumn('fileId')], '');
  assert.equal(failedReceipt[receiptColumn('state')], 'failed');

  const recovered = harness.api.saveImportPhotoItem(payload());
  assert.equal(recovered.ok, true);
  assert.equal(harness.audit.drive.creates, 2);
  assert.equal(harness.mediaFolders.photos.files.filter((file) => !file.trashed).length, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('pinId')], recovered.pin.id);
  assert.equal(receipt[receiptColumn('fileId')], recovered.pin.fileId);
});

test('file_saved and failed receipts reuse file metadata without creating another file', () => {
  for (const state of ['file_saved', 'failed']) {
    const harness = makeHarness({ sheets: baseSheets() });
    harness.api.saveImportPhotoItem(payload());
    const receipt = harness.sheets.get('import_receipts').rows[1];
    harness.sheets.get('map_info').rows.splice(1);
    receipt[receiptColumn('state')] = state;
    receipt[receiptColumn('leaseOwner')] = '';
    receipt[receiptColumn('leaseUntil')] = '';
    if (state === 'failed') receipt[receiptColumn('lastErrorCode')] = 'IMPORT_MAP_ROW_FAILED';
    const result = harness.api.saveImportPhotoItem(payload());
    assert.equal(result.ok, true, state);
    assert.equal(harness.audit.drive.creates, 1, state);
    assert.equal(harness.sheets.get('map_info').rows.length, 2, state);
    assert.equal(receipt[receiptColumn('state')], 'completed', state);
  }
});

test('existing map row repairs an unfinished receipt without Drive or duplicate map writes', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const first = harness.api.saveImportPhotoItem(payload());
  const receipt = harness.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('state')] = 'file_saved';
  receipt[receiptColumn('leaseOwner')] = '';
  receipt[receiptColumn('leaseUntil')] = '';
  const creates = harness.audit.drive.creates;
  const fileGets = harness.audit.drive.fileGets;
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.ok, true);
  assert.equal(result.deduplicated, true);
  assert.equal(result.pin.id, first.pin.id);
  assert.equal(harness.audit.drive.creates, creates);
  assert.equal(harness.audit.drive.fileGets, fileGets);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(receipt[receiptColumn('state')], 'completed');
});

test('multiple deterministic temporary files are rejected without creating or selecting another', () => {
  let changed = false;
  const harness = makeHarness({
    sheets: baseSheets(),
    afterCreateFile({ sheets, folder }) {
      if (changed) return;
      changed = true;
      const receipt = sheets.get('import_receipts').rows[1];
      folder.seedFile('duplicate-temp', receipt[receiptColumn('tempFileName')]);
      receipt[receiptColumn('leaseOwner')] = 'other';
      receipt[receiptColumn('leaseUntil')] = '2999-01-01T00:00:00.000Z';
    }
  });
  harness.api.saveImportPhotoItem(payload());
  const receipt = harness.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('leaseUntil')] = '2000-01-01T00:00:00.000Z';
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'IMPORT_RECEIPT_CORRUPTED');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);
});

test('an invalid recorded fileId fails retryably without creating a replacement', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.api.saveImportPhotoItem(payload());
  const receipt = harness.sheets.get('import_receipts').rows[1];
  harness.sheets.get('map_info').rows.splice(1);
  receipt[receiptColumn('state')] = 'file_saved';
  receipt[receiptColumn('fileId')] = 'missing-file';
  receipt[receiptColumn('leaseOwner')] = '';
  receipt[receiptColumn('leaseUntil')] = '';
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_MANAGED_COPY_FINALIZE_FAILED');
  assert.equal(result.retryable, true);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(receipt[receiptColumn('state')], 'failed');
  assert.equal(receipt[receiptColumn('fileId')], 'missing-file');
  assert.equal(receipt[receiptColumn('lastErrorCode')], 'DRIVE_MANAGED_COPY_FINALIZE_FAILED');
});

test('receipt target and temporary filename corruption is rejected before managed file access', () => {
  for (const corrupt of [
    (receipt) => { receipt[receiptColumn('targetFolderId')] = 'folder-elsewhere'; },
    (receipt) => { receipt[receiptColumn('tempFileName')] = 'user-controlled.jpg'; }
  ]) {
    const harness = makeHarness({ sheets: baseSheets() });
    harness.api.saveImportPhotoItem(payload());
    harness.sheets.get('map_info').rows.splice(1);
    const receipt = harness.sheets.get('import_receipts').rows[1];
    corrupt(receipt);
    receipt[receiptColumn('leaseOwner')] = '';
    receipt[receiptColumn('leaseUntil')] = '';
    const fileGets = harness.audit.drive.fileGets;
    const creates = harness.audit.drive.creates;
    const moves = harness.audit.drive.moves;
    const result = harness.api.saveImportPhotoItem(payload());
    assert.equal(result.errorCode, 'IMPORT_RECEIPT_CORRUPTED');
    assert.equal(harness.audit.drive.fileGets, fileGets);
    assert.equal(harness.audit.drive.creates, creates);
    assert.equal(harness.audit.drive.moves, moves);
  }
});

test('recorded fileId must resolve to a file in the receipt target folder', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.api.saveImportPhotoItem(payload());
  harness.sheets.get('map_info').rows.splice(1);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('state')] = 'file_saved';
  receipt[receiptColumn('fileId')] = 'foreign-file';
  receipt[receiptColumn('leaseOwner')] = '';
  receipt[receiptColumn('leaseUntil')] = '';
  const foreignFolder = { getId: () => 'foreign-folder' };
  let renamed = false;
  let shared = false;
  harness.filesById.set('foreign-file', {
    getId: () => 'foreign-file',
    getParents() {
      let used = false;
      return { hasNext: () => !used, next() { used = true; return foreignFolder; } };
    },
    setName() { renamed = true; },
    setSharing() { shared = true; }
  });
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.errorCode, 'DRIVE_MANAGED_COPY_FINALIZE_FAILED');
  assert.equal(result.retryable, true);
  assert.equal(renamed, false);
  assert.equal(shared, false);
});

test('provider error codes during media structure verification are not exposed', () => {
  const harness = makeHarness({ sheets: baseSheets(), folderErrorCode: 'SECRET_PROVIDER_CODE' });
  const result = harness.api.saveImportPhotoItem(payload());
  assert.equal(result.errorCode, 'DRIVE_MEDIA_STRUCTURE_FAILED');
  assert.equal(result.error.includes('provider detail'), false);
  assert.equal(harness.sheets.get('import_receipts').rows.length, 1);
});

test('managed-copy creation permission denial is specific and non-retryable', () => {
  const harness = makeHarness({ sheets: baseSheets(), failCreatePermission: true });
  harness.rootFolder.seedFile('photo_CREATEPERMAAA', 'readable.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_CREATEPERMAAA', targetFolderId: ''
  }));

  assert.equal(result.ok, false);
  assert.equal(result.errorCode, 'DRIVE_MANAGED_COPY_CREATE_FAILED');
  assert.equal(result.retryable, false);
  assert.equal(
    harness.sheets.get('import_receipts').rows[1][receiptColumn('lastErrorCode')],
    'DRIVE_MANAGED_COPY_CREATE_FAILED'
  );
  assert.match(harness.audit.errors.join('\n'), /"stage":"managed-copy-create"/);
});

test('transient sharing failure is retryable, compensates only the new copy, and replays safely', () => {
  const harness = makeHarness({ sheets: baseSheets(), failSharingOnce: true });
  const failed = harness.api.saveImportPhotoItem(payload());
  assert.equal(failed.errorCode, 'DRIVE_LINK_SHARING_FAILED');
  assert.equal(failed.retryable, true);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('state')], 'failed');
  assert.equal(receipt[receiptColumn('fileId')], '');
  assert.equal(receipt[receiptColumn('lastErrorCode')], 'DRIVE_LINK_SHARING_FAILED');
  assert.equal(harness.audit.drive.trashes, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);

  const recovered = harness.api.saveImportPhotoItem(payload());
  assert.equal(recovered.ok, true);
  assert.equal(harness.audit.drive.creates, 2);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('map append failure journals only its code and retry adds one row without another file', () => {
  const harness = makeHarness({ sheets: baseSheets(), failMapAppendOnce: true });
  const failed = harness.api.saveImportPhotoItem(payload());
  assert.equal(failed.errorCode, 'IMPORT_MAP_ROW_FAILED');
  const receipt = harness.sheets.get('import_receipts').rows[1];
  assert.equal(receipt[receiptColumn('state')], 'failed');
  assert.equal(receipt[receiptColumn('lastErrorCode')], 'IMPORT_MAP_ROW_FAILED');
  assert.match(harness.audit.errors.join('\n'), /"stage":"map-row"/);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);

  const recovered = harness.api.saveImportPhotoItem(payload());
  assert.equal(recovered.ok, true);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
});

test('a stale failed file-backed Drive receipt is safely taken over by a new page job', () => {
  const harness = makeHarness({ sheets: baseSheets(), failMapAppendOnce: true });
  harness.rootFolder.seedFile('photo_FILEOWNERAAA', 'owner.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const request = payload({ sourceDriveFileId: 'photo_FILEOWNERAAA', targetFolderId: '' });
  const failed = harness.api.saveImportPhotoItem(request);
  assert.equal(failed.errorCode, 'IMPORT_MAP_ROW_FAILED');
  assert.equal(harness.audit.drive.creates, 1);
  const originalReceipt = harness.sheets.get('import_receipts').rows[1];
  const originalPinId = originalReceipt[receiptColumn('pinId')];
  const originalManagedFileId = originalReceipt[receiptColumn('fileId')];

  const recovered = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-competitor', itemId: 'item-competitor',
    idempotencyKey: 'job-competitor:item-competitor',
    sourceDriveFileId: 'photo_FILEOWNERAAA', targetFolderId: ''
  }));
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(recovered.pin.id, originalPinId);
  assert.equal(recovered.pin.fileId, originalManagedFileId);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('import_receipts').rows.length, 2);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);

  const superseded = harness.api.saveImportPhotoItem(request);
  assert.equal(superseded.errorCode, 'DRIVE_SOURCE_ALREADY_LINKED');
  assert.equal(superseded.retryable, false);
});

test('an empty current-key failure converges on another stale file-backed owner', () => {
  const options = {
    sheets: baseSheets(),
    failCreatePermission: true,
    failMapAppendOnce: true
  };
  const harness = makeHarness(options);
  harness.rootFolder.seedFile('photo_TWOSTALEOWNR', 'two-stale.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const requestB = payload({
    jobId: 'job-empty', itemId: 'item-empty', idempotencyKey: 'job-empty:item-empty',
    sourceDriveFileId: 'photo_TWOSTALEOWNR', targetFolderId: ''
  });
  assert.equal(
    harness.api.saveImportPhotoItem(requestB).errorCode,
    'DRIVE_MANAGED_COPY_CREATE_FAILED'
  );
  options.failCreatePermission = false;

  const requestA = payload({
    jobId: 'job-file', itemId: 'item-file', idempotencyKey: 'job-file:item-file',
    sourceDriveFileId: 'photo_TWOSTALEOWNR', targetFolderId: ''
  });
  assert.equal(
    harness.api.saveImportPhotoItem(requestA).errorCode,
    'IMPORT_MAP_ROW_FAILED'
  );
  const fileOwnerReceipt = harness.sheets.get('import_receipts').rows.find((row) => (
    row[receiptColumn('jobId')] === 'job-file'
  ));
  const ownedPinId = fileOwnerReceipt[receiptColumn('pinId')];
  const ownedFileId = fileOwnerReceipt[receiptColumn('fileId')];

  const recovered = harness.api.saveImportPhotoItem(requestB);
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(recovered.pin.id, ownedPinId);
  assert.equal(recovered.pin.fileId, ownedFileId);
  const activeManagedFiles = harness.mediaFolders.photos.files.filter((file) => !file.trashed);
  assert.equal(activeManagedFiles.length, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(harness.sheets.get('import_receipts').rows.slice(1).filter((row) => (
    row[receiptColumn('idempotencyKey')]
  )).length, 1);
});

test('an active file-backed Drive receipt still blocks a different page job', () => {
  const harness = makeHarness({ sheets: baseSheets(), failMapAppendOnce: true });
  harness.rootFolder.seedFile('photo_ACTIVEOWNERAA', 'active.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const request = payload({ sourceDriveFileId: 'photo_ACTIVEOWNERAA', targetFolderId: '' });
  assert.equal(
    harness.api.saveImportPhotoItem(request).errorCode,
    'IMPORT_MAP_ROW_FAILED'
  );
  const receipt = harness.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('state')] = 'file_saved';
  receipt[receiptColumn('leaseOwner')] = 'other-owner';
  receipt[receiptColumn('leaseUntil')] = '2999-01-01T00:00:00.000Z';

  const competing = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-active-other', itemId: 'item-active-other',
    idempotencyKey: 'job-active-other:item-active-other',
    sourceDriveFileId: 'photo_ACTIVEOWNERAA', targetFolderId: ''
  }));

  assert.equal(competing.errorCode, 'IMPORT_ITEM_IN_PROGRESS');
  assert.equal(competing.retryable, true);
  assert.equal(harness.audit.drive.creates, 1);

  receipt[receiptColumn('leaseUntil')] = '2000-01-01T00:00:00.000Z';
  const recovered = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-active-other', itemId: 'item-active-other',
    idempotencyKey: 'job-active-other:item-active-other',
    sourceDriveFileId: 'photo_ACTIVEOWNERAA', targetFolderId: ''
  }));
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(harness.audit.drive.creates, 1);
});

test('a new page key takes over the deterministic temp after Drive create response loss', () => {
  const harness = makeHarness({
    sheets: baseSheets(),
    failCreateAfterFileOnce: true
  });
  harness.rootFolder.seedFile('photo_NEWPAGETEMP', 'new-page.heic', {
    mimeType: 'image/heic', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const firstRequest = payload({
    jobId: 'job-temp-old', itemId: 'item-temp-old',
    idempotencyKey: 'job-temp-old:item-temp-old',
    sourceDriveFileId: 'photo_NEWPAGETEMP', targetFolderId: ''
  });
  const failed = harness.api.saveImportPhotoItem(firstRequest);
  assert.equal(failed.errorCode, 'DRIVE_MANAGED_COPY_CREATE_FAILED');
  const staleReceipt = harness.sheets.get('import_receipts').rows[1];
  const stalePinId = staleReceipt[receiptColumn('pinId')];
  assert.equal(staleReceipt[receiptColumn('fileId')], '');

  const recovered = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-temp-new', itemId: 'item-temp-new',
    idempotencyKey: 'job-temp-new:item-temp-new',
    sourceDriveFileId: 'photo_NEWPAGETEMP', targetFolderId: ''
  }));

  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(recovered.pin.id, stalePinId);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.rootFolder.files.filter((file) => (
    file.id !== 'photo_NEWPAGETEMP' && !file.trashed
  )).length, 1);
});

test('Drive source sharing visibility is not consulted before creating the managed JPEG', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const metadata = {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'ANYONE_WITH_LINK',
    failSharingReadAt: 1
  };
  harness.rootFolder.seedFile('photo_DIRECTPRIVATE', 'direct-private.jpg', metadata);
  const result = harness.api.saveImportPhotoItem(payload({
    sourceDriveFileId: 'photo_DIRECTPRIVATE', targetFolderId: ''
  }));

  assert.equal(result.ok, true, JSON.stringify(result));
  assert.notEqual(result.pin.fileId, 'photo_DIRECTPRIVATE');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(
    harness.sheets.get('import_receipts').rows[1][receiptColumn('targetFolderId')],
    harness.mediaFolders.photos.id
  );
});

test('stale no-file HEIC takeover uses the managed photos fallback target', () => {
  const harness = makeHarness({ sheets: baseSheets(), failCreateAfterFileOnce: true });
  harness.rootFolder.seedFile('photo_LEGACYTEMPHC', 'legacy-temp.heic', {
    mimeType: 'image/heic', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const firstRequest = payload({
    jobId: 'job-legacy-old', itemId: 'item-legacy-old',
    idempotencyKey: 'job-legacy-old:item-legacy-old',
    sourceDriveFileId: 'photo_LEGACYTEMPHC', targetFolderId: ''
  });
  assert.equal(
    harness.api.saveImportPhotoItem(firstRequest).errorCode,
    'DRIVE_MANAGED_COPY_CREATE_FAILED'
  );
  harness.sheets.get('import_receipts').rows[1][receiptColumn('targetFolderId')] = '';

  const recovered = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-legacy-new', itemId: 'item-legacy-new',
    idempotencyKey: 'job-legacy-new:item-legacy-new',
    sourceDriveFileId: 'photo_LEGACYTEMPHC', targetFolderId: ''
  }));

  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(
    harness.sheets.get('import_receipts').rows[1][receiptColumn('targetFolderId')],
    harness.mediaFolders.photos.id
  );
});

test('transient Drive source lookup failure is retryable and the same key later succeeds', () => {
  const harness = makeHarness({ sheets: baseSheets(), failFileGetOnce: true });
  harness.rootFolder.seedFile('photo_LOOKUPRETRYA', 'lookup.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'ANYONE_WITH_LINK'
  });
  const request = payload({ sourceDriveFileId: 'photo_LOOKUPRETRYA', targetFolderId: '' });

  const failed = harness.api.saveImportPhotoItem(request);
  assert.equal(failed.errorCode, 'DRIVE_SOURCE_CHECK_FAILED');
  assert.equal(failed.retryable, true);
  assert.equal(harness.sheets.get('map_info').rows.length, 1);

  const recovered = harness.api.saveImportPhotoItem(request);
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.notEqual(recovered.pin.fileId, 'photo_LOOKUPRETRYA');
  assert.equal(harness.audit.drive.creates, 1);
});

test('file_saved Drive retry reuses the managed JPEG without consulting source sharing', () => {
  const harness = makeHarness({ sheets: baseSheets(), failMapAppendOnce: true });
  const metadata = {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'ANYONE_WITH_LINK'
  };
  harness.rootFolder.seedFile('photo_SHARECHECKAA', 'share-check.jpg', metadata);
  const request = payload({ sourceDriveFileId: 'photo_SHARECHECKAA', targetFolderId: '' });
  assert.equal(
    harness.api.saveImportPhotoItem(request).errorCode,
    'IMPORT_MAP_ROW_FAILED'
  );
  metadata.failSharingReadOnce = true;

  const recovered = harness.api.saveImportPhotoItem(request);
  assert.equal(recovered.ok, true, JSON.stringify(recovered));
  assert.notEqual(recovered.pin.fileId, 'photo_SHARECHECKAA');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(
    harness.sheets.get('import_receipts').rows[1][receiptColumn('lastErrorCode')],
    ''
  );
});

test('map append followed by receipt completion interruption repairs without duplicate effects', () => {
  const harness = makeHarness({
    sheets: baseSheets(),
    failReceiptCompleteWriteOnce: true
  });
  const interrupted = harness.api.saveImportPhotoItem(payload());
  assert.equal(interrupted.ok, false);
  assert.equal(interrupted.errorCode, 'IMPORT_ITEM_SAVE_FAILED');
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);

  const recovered = harness.api.saveImportPhotoItem(payload());
  assert.equal(recovered.ok, true);
  assert.equal(recovered.deduplicated, true);
  assert.equal(harness.audit.drive.creates, 1);
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  assert.equal(
    harness.sheets.get('import_receipts').rows[1][receiptColumn('state')],
    'completed'
  );
});

test('a non-completed Drive receipt with a live map row converges for another job', () => {
  const harness = makeHarness({
    sheets: baseSheets(),
    failReceiptCompleteWriteOnce: true
  });
  harness.rootFolder.seedFile('photo_LIVEMAPOWNER', 'live.jpg', {
    mimeType: 'image/jpeg', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const request = payload({ sourceDriveFileId: 'photo_LIVEMAPOWNER', targetFolderId: '' });
  const interrupted = harness.api.saveImportPhotoItem(request);
  assert.equal(interrupted.errorCode, 'IMPORT_ITEM_SAVE_FAILED');
  assert.equal(harness.sheets.get('map_info').rows.length, 2);
  const existingPinId = harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('ID')];
  const existingFileId = harness.sheets.get('map_info').rows[1][MAP_HEADERS.indexOf('ファイルID')];

  const competing = harness.api.saveImportPhotoItem(payload({
    jobId: 'job-other', itemId: 'item-other', idempotencyKey: 'job-other:item-other',
    sourceDriveFileId: 'photo_LIVEMAPOWNER', targetFolderId: ''
  }));
  assert.equal(competing.ok, true, JSON.stringify(competing));
  assert.equal(competing.deduplicated, true);
  assert.equal(competing.pin.id, existingPinId);
  assert.equal(competing.pin.fileId, existingFileId);
  assert.equal(harness.audit.drive.creates, 1);
});

test('missing recorded managed Drive source copy is a finalize failure, not a create failure', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  harness.rootFolder.seedFile('photo_RECOVERYAAAA', 'recovery.heic', {
    mimeType: 'image/heic', sizeBytes: 3, sharingAccess: 'PRIVATE'
  });
  const request = payload({ sourceDriveFileId: 'photo_RECOVERYAAAA', targetFolderId: '' });
  const first = harness.api.saveImportPhotoItem(request);
  assert.equal(first.ok, true);

  harness.sheets.get('map_info').rows.splice(1);
  const receipt = harness.sheets.get('import_receipts').rows[1];
  receipt[receiptColumn('state')] = 'file_saved';
  receipt[receiptColumn('leaseOwner')] = '';
  receipt[receiptColumn('leaseUntil')] = '';
  receipt[receiptColumn('fileId')] = 'missing-managed-file';
  const result = harness.api.saveImportPhotoItem(request);

  assert.equal(result.errorCode, 'DRIVE_MANAGED_COPY_FINALIZE_FAILED');
  assert.match(harness.audit.errors.join('\n'), /"stage":"managed-copy-finalize"/);
  assert.equal(harness.audit.drive.creates, 1);
});

test('different job items claim independently and create distinct pins and files', () => {
  const harness = makeHarness({ sheets: baseSheets() });
  const first = harness.api.saveImportPhotoItem(payload());
  const second = harness.api.saveImportPhotoItem(payload({
    itemId: 'item-2', idempotencyKey: 'job-1:item-2', title: '写真2'
  }));
  assert.equal(first.ok, true);
  assert.equal(second.ok, true);
  assert.notEqual(first.pin.id, second.pin.id);
  assert.notEqual(first.pin.fileId, second.pin.fileId);
  assert.equal(harness.audit.drive.creates, 2);
  assert.equal(harness.sheets.get('map_info').rows.length, 3);
  assert.equal(harness.sheets.get('import_receipts').rows.length, 3);
});

test('README summarizes receipt-backed duplicate prevention and retry behavior', () => {
  assert.match(readme, /import_receipts/);
  assert.match(readme, /重複防止と再試行管理/);
  assert.match(readme, /応答喪失時の重複防止/);
  assert.match(readme, /既存行を削除せず[^\n]*追加・補修/);
});

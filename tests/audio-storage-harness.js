const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');

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

function iterator(values) {
  let index = 0;
  return {
    hasNext() { return index < values.length; },
    next() {
      if (index >= values.length) throw new Error('iterator exhausted');
      return values[index++];
    }
  };
}

function lastUsedRow(rows) {
  for (let index = rows.length - 1; index >= 0; index -= 1) {
    if ((rows[index] || []).some((value) => value !== '' && value != null)) return index + 1;
  }
  return 0;
}

function makeHarness(options = {}) {
  const audit = {
    events: [], errors: [], sheetWrites: 0, uuidCalls: 0,
    locks: { attempts: 0, releases: 0, held: false, nestedAttempts: 0 },
    drive: { creates: 0, trashes: 0, moves: 0, shares: 0, callsWhileLocked: 0 }
  };
  const sheets = new Map();
  const folders = new Map();
  const files = new Map();
  let uuid = 0;
  let timestampTick = 0;
  let failReceiptState = options.failReceiptState || '';
  let failMapAudioWrite = options.failMapAudioWriteOnce === true;
  let failMapAudioTimestamp = options.failMapAudioTimestampOnce === true;

  function makeSheet(name, sourceRows) {
    const rows = (sourceRows || []).map((row) => row.slice());
    let maxRows = Math.max(1000, rows.length);
    const sheet = {
      rows,
      getLastRow() { return lastUsedRow(rows); },
      getLastColumn() { return rows.reduce((max, row) => Math.max(max, row.length), 0); },
      getMaxColumns() { return Math.max(26, sheet.getLastColumn()); },
      getMaxRows() { return maxRows; },
      insertRowsAfter(_after, count) { maxRows += count; },
      getDataRange() {
        return sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), Math.max(1, sheet.getLastColumn()));
      },
      getRange(row, column, rowCount = 1, columnCount = 1) {
        const range = {
          getValues() {
            return Array.from({ length: rowCount }, (_unused, rowOffset) =>
              Array.from({ length: columnCount }, (_unusedColumn, columnOffset) =>
                (rows[row - 1 + rowOffset] || [])[column - 1 + columnOffset] ?? ''
              )
            );
          },
          getValue() { return range.getValues()[0][0]; },
          getFormulas() {
            return Array.from({ length: rowCount }, () => Array(columnCount).fill(''));
          },
          setValues(values) {
            if (name === 'import_receipts' && failReceiptState
                && values[0] && values[0][RECEIPT_HEADERS.indexOf('state')] === failReceiptState) {
              failReceiptState = '';
              throw new Error('private receipt write failure');
            }
            if (name === 'map_info' && column === 16 && failMapAudioWrite) {
              failMapAudioWrite = false;
              throw new Error('private map audio write failure');
            }
            if (name === 'map_info' && column === 14 && failMapAudioTimestamp) {
              failMapAudioTimestamp = false;
              throw new Error('private map audio timestamp failure');
            }
            if (name === 'route_pins' && options.failRelationWriteOnce
                && !options.__relationWriteFailed) {
              options.__relationWriteFailed = true;
              throw new Error('private relation write failure');
            }
            audit.sheetWrites += 1;
            audit.events.push({ type: 'sheet-write', sheet: name, column, values: values.map((entry) => entry.slice()) });
            values.forEach((sourceRow, rowOffset) => {
              const targetRow = row - 1 + rowOffset;
              while (rows.length <= targetRow) rows.push([]);
              sourceRow.forEach((value, columnOffset) => {
                const targetColumn = column - 1 + columnOffset;
                while (rows[targetRow].length <= targetColumn) rows[targetRow].push('');
                rows[targetRow][targetColumn] = value;
              });
            });
            return range;
          },
          setValue(value) { return range.setValues([[value]]); },
          setBackground() { return range; },
          setFontColor() { return range; },
          setFontWeight() { return range; }
        };
        return range;
      },
      appendRow(row) {
        if (name === 'map_info' && options.failMapAppendOnce && !options.__mapAppendFailed) {
          options.__mapAppendFailed = true;
          throw new Error('private map append failure');
        }
        rows.push(row.slice());
        audit.sheetWrites += 1;
        audit.events.push({ type: 'sheet-append', sheet: name, values: row.slice() });
      },
      deleteRow(rowNumber) {
        if (name === 'map_info' && options.failMapDeleteOnce
            && !options.__mapDeleteFailed) {
          options.__mapDeleteFailed = true;
          throw new Error('private map delete failure');
        }
        rows.splice(rowNumber - 1, 1);
        audit.sheetWrites += 1;
        audit.events.push({ type: 'sheet-delete', sheet: name, rowNumber });
      },
      deleteRows(rowNumber, count) {
        if (name === 'map_info' && options.failMapDeleteRowsOnce
            && !options.__mapDeleteRowsFailed) {
          options.__mapDeleteRowsFailed = true;
          throw new Error('private map bulk delete failure');
        }
        rows.splice(rowNumber - 1, count);
        audit.sheetWrites += 1;
        audit.events.push({ type: 'sheet-delete', sheet: name, rowNumber, count });
      },
      insertColumnsAfter() {}, setFrozenRows() {}, setColumnWidth() {}, activate() {}
    };
    sheets.set(name, sheet);
    return sheet;
  }

  function makeFolder(id, name, parentIds = []) {
    const folder = {
      id, name, parentIds: parentIds.slice(), trashed: false,
      getId: () => id,
      getName: () => folder.name,
      isTrashed: () => folder.trashed,
      getParents: () => iterator(folder.parentIds.map((parentId) => folders.get(parentId)).filter(Boolean)),
      getFolders: () => iterator(Array.from(folders.values()).filter((candidate) =>
        !candidate.trashed && candidate.parentIds.includes(id))),
      getFiles: () => iterator(Array.from(files.values()).filter((candidate) =>
        !candidate.trashed && candidate.parentIds.includes(id))),
      getFilesByName: (value) => iterator(Array.from(files.values()).filter((candidate) =>
        !candidate.trashed && candidate.parentIds.includes(id) && candidate.name === String(value))),
      createFolder(value) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        const child = makeFolder(`created_folder_${folders.size + 1}`, String(value), [id]);
        audit.events.push({ type: 'drive-folder-create', id: child.id, parentId: id });
        return child;
      },
      createFile(blob) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.creates += 1;
        const file = makeFile(
          `managed_audio_${String(audit.drive.creates).padStart(8, '0')}`,
          String(blob && (blob.name || (blob.getName && blob.getName())) || ''),
          String(blob && (blob.mime || (blob.getContentType && blob.getContentType())) || ''),
          Array.from(blob && (blob.bytes || (blob.getBytes && blob.getBytes())) || []),
          [id],
          { managed: true }
        );
        audit.events.push({ type: 'drive-create', id: file.id, parentId: id, lockHeld: audit.locks.held });
        return file;
      }
    };
    folders.set(id, folder);
    return folder;
  }

  function makeFile(id, name, mimeType, bytes, parentIds, metadata = {}) {
    const file = {
      id: String(id), name: String(name), mimeType: String(mimeType),
      bytes: Array.from(bytes || []), parentIds: parentIds.slice(), trashed: metadata.trashed === true,
      getId: () => file.id,
      getName: () => file.name,
      getMimeType: () => file.mimeType,
      getSize: () => file.bytes.length,
      getLastUpdated: () => new Date('2026-07-22T01:02:03.000Z'),
      isTrashed: () => file.trashed,
      getParents: () => iterator(file.parentIds.map((parentId) => folders.get(parentId)).filter(Boolean)),
      getBlob: () => ({
        getBytes: () => file.bytes.slice(),
        getContentType: () => file.mimeType,
        getName: () => file.name
      }),
      setName(value) { file.name = String(value); return file; },
      setSharing() { audit.drive.shares += 1; throw new Error('managed audio must remain private'); },
      getSharingAccess: () => 'PRIVATE',
      setTrashed(value) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.trashes += 1;
        audit.events.push({
          type: 'drive-trash', id: file.id, value: !!value, lockHeld: audit.locks.held
        });
        if ((options.failTrashIds || []).includes(file.id)) throw new Error('private trash failure ' + file.id);
        if (options.failTrashOnceId === file.id && !options.__trashFailed) {
          options.__trashFailed = true;
          throw new Error('private transient trash failure ' + file.id);
        }
        file.trashed = !!value;
        if (value && typeof options.afterTrash === 'function') options.afterTrash(file);
        return file;
      },
      moveTo(targetFolder) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        audit.drive.moves += 1;
        audit.events.push({ type: 'drive-move', id: file.id, parentId: targetFolder.id });
        if (options.failMoveOnceId === file.id && !options.__moveFailed) {
          options.__moveFailed = true;
          throw new Error('private transient move failure ' + file.id);
        }
        file.parentIds = [targetFolder.id];
        return file;
      }
    };
    files.set(file.id, file);
    return file;
  }

  const rootId = 'root_media_1234567890';
  const rootFolder = makeFolder(rootId, 'Root');
  const photosFolder = makeFolder('managed_photos_12345', 'photos', [rootId]);
  const audioFolder = makeFolder('managed_audio_folder1', 'audio', [rootId]);
  const originalFolder = makeFolder('original_folder_1234', 'original', [rootId]);
  const originalPhotosFolder = makeFolder('original_photos_1234', 'photos', [originalFolder.id]);
  const originalAudioFolder = makeFolder('original_audio_12345', 'audio', [originalFolder.id]);
  makeFile('media_guide_123456', 'ここに直接ファイルを入れてください.txt', 'text/plain', [], [rootId]);

  const defaultSheets = {
    map_info: [MAP_HEADERS],
    import_receipts: [RECEIPT_HEADERS],
    config: [
      ['設定項目', '値', '説明'],
      ['IMAGE_DRIVE_URL', `https://drive.google.com/drive/folders/${rootId}`, ''],
      ['EDIT_KEY', 'key', '']
    ],
    routes: [['routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle']],
    route_pins: [['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt']],
    route_cache: [['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt']]
  };
  const sourceSheets = Object.assign({}, defaultSheets, options.sheets || {});
  Object.entries(sourceSheets).forEach(([name, rows]) => makeSheet(name, rows));
  const spreadsheet = {
    getSheetByName: (name) => sheets.get(name) || null,
    insertSheet: (name) => makeSheet(name, [])
  };
  const lock = {
    tryLock() {
      audit.locks.attempts += 1;
      if (audit.locks.held) {
        audit.locks.nestedAttempts += 1;
        return false;
      }
      audit.locks.held = true;
      return true;
    },
    releaseLock() {
      assert.equal(audit.locks.held, true);
      audit.locks.held = false;
      audit.locks.releases += 1;
    }
  };
  const context = {
    Buffer, Date, JSON, Math, Number, Object, String, Array, Error, RegExp, Set, Map,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      flush() { assert.equal(audit.locks.held, true); }
    },
    CacheService: { getScriptCache: () => ({ get: () => options.validToken === false ? null : '1' }) },
    LockService: { getScriptLock: () => lock },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'SHA_256' }, Charset: { UTF_8: 'UTF_8' },
      getUuid() { audit.uuidCalls += 1; return `uuid-${++uuid}`; },
      computeDigest(_algorithm, value) {
        return Array.from(crypto.createHash('sha256').update(String(value)).digest());
      },
      base64Decode(value) { return Array.from(Buffer.from(String(value), 'base64')); },
      base64Encode(bytes) { return Buffer.from(bytes).toString('base64'); },
      newBlob(bytes, mime, name) {
        const values = Array.from(bytes || []);
        return {
          bytes: values, mime: String(mime || ''), name: String(name || ''),
          getBytes: () => values.slice(), getContentType: () => String(mime || ''),
          getName: () => String(name || '')
        };
      },
      formatDate() {
        timestampTick += 1;
        return `2026/07/22 12:00:${String(timestampTick).padStart(2, '0')}`;
      }
    },
    DriveApp: {
      Access: { ANYONE: 'ANYONE', ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', PRIVATE: 'PRIVATE' },
      Permission: { VIEW: 'VIEW' },
      getFolderById(id) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        if (String(id) === rootId && options.failStructureOnce
            && !options.__structureFailed) {
          options.__structureFailed = true;
          throw new Error('private structure failure');
        }
        const folder = folders.get(String(id));
        if (!folder) throw new Error('private folder missing');
        return folder;
      },
      getFileById(id) {
        if (audit.locks.held) audit.drive.callsWhileLocked += 1;
        const file = files.get(String(id));
        if (!file) throw new Error('private file missing');
        return file;
      }
    },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://example.invalid/exec' }) },
    HtmlService: {}, Maps: {},
    console: {
      error(value) { audit.errors.push(String(value)); },
      warn(value) { audit.errors.push(String(value)); },
      log() {}
    }
  };
  vm.createContext(context);
  vm.runInContext(codeJs + '\n' + [
    'globalThis.__audioApi = {',
    ' saveImportAudioItem: typeof saveImportAudioItem === "function" ? saveImportAudioItem : null,',
    ' removePinAudio: typeof removePinAudio === "function" ? removePinAudio : null,',
    ' readPinAudioBlobByPinId_: typeof readPinAudioBlobByPinId_ === "function" ? readPinAudioBlobByPinId_ : null,',
    ' getPinAudioData: typeof getPinAudioData === "function" ? getPinAudioData : null,',
    ' resolveSharedProjection_: typeof resolveSharedProjection_ === "function" ? resolveSharedProjection_ : null,',
    ' getSharedViewData: typeof getSharedViewData === "function" ? getSharedViewData : null,',
    ' getSharedPinAudioData: typeof getSharedPinAudioData === "function" ? getSharedPinAudioData : null,',
    ' deletePin: typeof deletePin === "function" ? deletePin : null,',
    ' bulkDeletePins: typeof bulkDeletePins === "function" ? bulkDeletePins : null,',
    ' normalizeImportAudioPayload_: typeof normalizeImportAudioPayload_ === "function" ? normalizeImportAudioPayload_ : null',
    '};'
  ].join('\n'), context);

  return {
    api: context.__audioApi, audit, sheets, folders, files,
    rootId, rootFolder, photosFolder, audioFolder, originalFolder,
    originalPhotosFolder, originalAudioFolder,
    addFile: makeFile,
    mapRow(pinId) {
      return sheets.get('map_info').rows.find((row, index) => index > 0 && String(row[8] || '') === String(pinId));
    },
    receipt() { return sheets.get('import_receipts').rows[1] || null; },
    liveManagedAudioFiles() {
      return Array.from(files.values()).filter((file) => !file.trashed && file.parentIds.length === 1
        && file.parentIds[0] === audioFolder.id);
    },
    fileExists(id) { const file = files.get(String(id)); return !!file && !file.trashed; },
    setFailReceiptState(value) { failReceiptState = String(value || ''); }
  };
}

function validMp3Bytes(size = 1024 * 1024) {
  const bytes = Buffer.alloc(size, 0);
  bytes[0] = 0x49; bytes[1] = 0x44; bytes[2] = 0x33;
  bytes[3] = 0x04; bytes[4] = 0x00;
  return Array.from(bytes);
}

function audioPayload(overrides = {}) {
  return Object.assign({
    __editToken: 'valid-token',
    idempotencyKey: 'audio-job-1:item-1', jobId: 'audio-job-1', itemId: 'item-1',
    operationMode: 'create-pin', targetPinId: '', expectedUpdatedAt: '',
    sourceKind: 'local', sourceDriveFileId: '', sourceFileName: 'voice.wav',
    audioMimeType: 'audio/mpeg', audioBase64: Buffer.from(validMp3Bytes()).toString('base64'),
    pin: {
      title: 'voice', description: '', lat: null, lng: null, color: '#e53935',
      links: [], status: '未対応', tags: [], eventAt: '', icon: 'default'
    }
  }, overrides);
}

function pinRow(overrides = {}) {
  const values = {
    timestamp: '2026/07/22 09:00:00', title: '既存ピン', description: '維持する説明',
    lat: 35.1, lng: 139.2, color: '#4caf50', fileId: '', imageUrl: '',
    id: 'pin-existing-0001', links: 'https://example.com/existing', status: '対応中',
    tags: '既存|保護', eventAt: '2026-07-22T08:30', updatedAt: '2026/07/22 11:59:00',
    icon: 'nature', audioId: '', ...overrides
  };
  return [
    values.timestamp, values.title, values.description, values.lat, values.lng, values.color,
    values.fileId, values.imageUrl, values.id, values.links, values.status, values.tags,
    values.eventAt, values.updatedAt, values.icon, values.audioId
  ];
}

module.exports = {
  makeHarness, audioPayload, validMp3Bytes, pinRow, RECEIPT_HEADERS, MAP_HEADERS
};

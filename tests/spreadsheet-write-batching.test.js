const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const TEST_EDIT_TOKEN = 'test-edit-token';

const MAP_INFO_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧', '状態', 'タグ',
  'イベント時刻', '更新時刻', 'アイコン'
];
const ROUTES_HEADERS = [
  'routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId',
  'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'
];
const ROUTE_PINS_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];
const IMPORT_RECEIPT_HEADERS = [
  'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
  'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
  'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId'
];

function importReceiptRow({ pinId, fileId, sourceDriveFileId, targetFolderId = 'managed-folder' }) {
  const values = {
    idempotencyKey: `key-${pinId}`,
    jobId: 'job',
    itemId: pinId,
    payloadHash: `hash-${pinId}`,
    state: 'completed',
    pinId,
    targetFolderId,
    fileId,
    sourceDriveFileId
  };
  return IMPORT_RECEIPT_HEADERS.map((header) => values[header] || '');
}

function cloneMatrix(matrix) {
  return matrix.map((row) => row.slice());
}

function toColumnNumber(letters) {
  return letters.split('').reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function parseA1(a1) {
  const match = /^([A-Z]+)(\d+)$/.exec(a1);
  if (!match) throw new Error(`Unsupported A1 notation: ${a1}`);
  return { row: Number(match[2]), column: toColumnNumber(match[1]) };
}

function createAudit() {
  return {
    reads: [],
    writes: [],
    dataRangeCalls: [],
    rangeListCalls: [],
    rangeListSets: [],
    driveCalls: [],
    logs: [],
    errors: [],
    lock: {
      held: false,
      available: true,
      availabilitySequence: [],
      tryCalls: [],
      releaseCalls: 0,
      nestedAttempts: 0,
      maxDepth: 0
    }
  };
}

function createSheet(name, rows, formulas, audit) {
  const sheet = {
    name,
    rows: cloneMatrix(rows),
    formulas: cloneMatrix(formulas || rows.map((row) => row.map(() => ''))),
    maxRows: rows.length,
    failNextSetValues: false,
    failDeleteRowsStarts: new Set(),
    getLastRow() {
      for (let index = Math.max(this.rows.length, this.formulas.length) - 1; index >= 0; index -= 1) {
        const values = this.rows[index] || [];
        const formulasRow = this.formulas[index] || [];
        if (values.some((value) => value !== '' && value != null)
          || formulasRow.some((value) => value !== '' && value != null)) {
          return index + 1;
        }
      }
      return 0;
    },
    getLastColumn() {
      return this.rows.reduce((max, row) => Math.max(max, row.length), 0);
    },
    getDataRange() {
      audit.dataRangeCalls.push({ sheet: name, lockHeld: audit.lock.held });
      return createRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()), audit);
    },
    getRange(row, column, numRows = 1, numColumns = 1) {
      if (typeof row === 'string') {
        const parsed = parseA1(row);
        return createRange(this, parsed.row, parsed.column, 1, 1, audit);
      }
      return createRange(this, row, column, numRows, numColumns, audit);
    },
    getRangeList(addresses) {
      if (!Array.isArray(addresses) || addresses.length === 0) {
        throw new Error('getRangeList requires at least one range');
      }
      audit.rangeListCalls.push({ sheet: name, addresses: addresses.slice(), lockHeld: audit.lock.held });
      return {
        setValue(value) {
          audit.rangeListSets.push({ sheet: name, addresses: addresses.slice(), value, lockHeld: audit.lock.held });
          addresses.forEach((address) => {
            const parsed = parseA1(address);
            setCell(sheet, parsed.row - 1, parsed.column - 1, value, '');
          });
          return this;
        }
      };
    },
    appendRow(row) {
      const rowIndex = this.getLastRow();
      this.rows[rowIndex] = row.slice();
      this.formulas[rowIndex] = row.map(() => '');
      this.maxRows = Math.max(this.maxRows, rowIndex + 1);
      audit.writes.push({ sheet: name, method: 'appendRow', values: row.slice(), lockHeld: audit.lock.held });
    },
    getMaxRows() {
      return this.maxRows;
    },
    insertRowsAfter(afterPosition, howMany) {
      assert.equal(afterPosition, this.maxRows);
      this.maxRows += howMany;
      audit.writes.push({ sheet: name, method: 'insertRowsAfter', afterPosition, howMany, lockHeld: audit.lock.held });
    },
    deleteRow(rowPosition) {
      this.rows.splice(rowPosition - 1, 1);
      this.formulas.splice(rowPosition - 1, 1);
      this.maxRows = Math.max(0, this.maxRows - 1);
      audit.writes.push({ sheet: name, method: 'deleteRow', row: rowPosition, count: 1, lockHeld: audit.lock.held });
    },
    deleteRows(startRow, howMany) {
      audit.writes.push({ sheet: name, method: 'deleteRows', row: startRow, count: howMany, lockHeld: audit.lock.held });
      if (this.failDeleteRowsStarts.has(startRow)) throw new Error('simulated deleteRows failure');
      this.rows.splice(startRow - 1, howMany);
      this.formulas.splice(startRow - 1, howMany);
      this.maxRows = Math.max(0, this.maxRows - howMany);
    }
  };
  return sheet;
}

function ensureCell(sheet, rowIndex, columnIndex) {
  while (sheet.rows.length <= rowIndex) sheet.rows.push([]);
  while (sheet.formulas.length <= rowIndex) sheet.formulas.push([]);
  while (sheet.rows[rowIndex].length <= columnIndex) sheet.rows[rowIndex].push('');
  while (sheet.formulas[rowIndex].length <= columnIndex) sheet.formulas[rowIndex].push('');
}

function setCell(sheet, rowIndex, columnIndex, value, formula) {
  ensureCell(sheet, rowIndex, columnIndex);
  sheet.rows[rowIndex][columnIndex] = value;
  sheet.formulas[rowIndex][columnIndex] = formula || '';
}

function createRange(sheet, row, column, numRows, numColumns, audit) {
  function matrixFrom(source) {
    return Array.from({ length: numRows }, (_, rowOffset) =>
      Array.from({ length: numColumns }, (_, columnOffset) => {
        const value = (source[row - 1 + rowOffset] || [])[column - 1 + columnOffset];
        return value == null ? '' : value;
      })
    );
  }

  return {
    getValue() {
      audit.reads.push({ sheet: sheet.name, method: 'getValue', row, column, numRows, numColumns, lockHeld: audit.lock.held });
      return (sheet.rows[row - 1] || [])[column - 1] || '';
    },
    getValues() {
      audit.reads.push({ sheet: sheet.name, method: 'getValues', row, column, numRows, numColumns, lockHeld: audit.lock.held });
      return matrixFrom(sheet.rows);
    },
    getFormulas() {
      audit.reads.push({ sheet: sheet.name, method: 'getFormulas', row, column, numRows, numColumns, lockHeld: audit.lock.held });
      return matrixFrom(sheet.formulas);
    },
    setValue(value) {
      audit.writes.push({ sheet: sheet.name, method: 'setValue', row, column, numRows, numColumns, value, lockHeld: audit.lock.held });
      setCell(sheet, row - 1, column - 1, value, typeof value === 'string' && value.startsWith('=') ? value : '');
      return this;
    },
    setValues(values) {
      audit.writes.push({ sheet: sheet.name, method: 'setValues', row, column, numRows, numColumns, values: cloneMatrix(values), lockHeld: audit.lock.held });
      if (sheet.failNextSetValues) {
        sheet.failNextSetValues = false;
        throw new Error('simulated Spreadsheet write failure');
      }
      values.forEach((valuesRow, rowOffset) => {
        valuesRow.forEach((value, columnOffset) => {
          const formula = typeof value === 'string' && value.startsWith('=') ? value : '';
          setCell(sheet, row - 1 + rowOffset, column - 1 + columnOffset, value, formula);
        });
      });
      return this;
    },
    clearContent() {
      audit.writes.push({ sheet: sheet.name, method: 'clearContent', row, column, numRows, numColumns, lockHeld: audit.lock.held });
      for (let rowOffset = 0; rowOffset < numRows; rowOffset += 1) {
        for (let columnOffset = 0; columnOffset < numColumns; columnOffset += 1) {
          setCell(sheet, row - 1 + rowOffset, column - 1 + columnOffset, '', '');
        }
      }
      return this;
    }
  };
}

function mapRow(overrides = {}) {
  return [
    overrides.timestamp || '2026/07/10 10:00:00',
    Object.prototype.hasOwnProperty.call(overrides, 'title') ? overrides.title : 'Original title',
    Object.prototype.hasOwnProperty.call(overrides, 'description') ? overrides.description : 'Original description',
    Object.prototype.hasOwnProperty.call(overrides, 'lat') ? overrides.lat : 35,
    Object.prototype.hasOwnProperty.call(overrides, 'lng') ? overrides.lng : 139,
    Object.prototype.hasOwnProperty.call(overrides, 'color') ? overrides.color : '#43a047',
    overrides.fileId || '',
    overrides.imageUrl || 'computed-image-url',
    overrides.id || 'pin-1',
    Object.prototype.hasOwnProperty.call(overrides, 'links') ? overrides.links : 'https://example.com/original',
    Object.prototype.hasOwnProperty.call(overrides, 'status') ? overrides.status : '対応中',
    Object.prototype.hasOwnProperty.call(overrides, 'tags') ? overrides.tags : 'alpha|beta',
    Object.prototype.hasOwnProperty.call(overrides, 'eventAt') ? overrides.eventAt : '2026-07-10T12:00',
    overrides.updatedAt || '2026-07-10T01:00:00.000Z',
    Object.prototype.hasOwnProperty.call(overrides, 'icon') ? overrides.icon : 'food'
  ];
}

function routeRow(id, orderIndex, name = id) {
  return [id, name, '#1e88e5', 'straight', false, '', '', '', '', orderIndex, true, true, true, 'solid'];
}

function routePinRow(routeId, pinId, pinOrder = 0, createdAt = 'created', updatedAt = 'updated') {
  return [routeId, pinId, pinOrder, createdAt, updatedAt];
}

function routeCacheRow(cacheKey, routeId, coordsJson = '[[35,139],[36,140]]', provider = 'osrm', createdAt = '2026-07-10T00:00:00.000Z', expiresAt = '') {
  return [cacheKey, routeId, coordsJson, provider, createdAt, expiresAt];
}

function nonEmptyDataRows(sheet) {
  return sheet.rows.slice(1, sheet.getLastRow()).map((row) => row.slice());
}

function writesFor(audit, sheetName, method) {
  return audit.writes.filter((write) => write.sheet === sheetName && (!method || write.method === method));
}

function createHarness(options = {}) {
  const audit = createAudit();
  audit.lock.available = options.lockAvailable !== false;
  audit.lock.availabilitySequence = Array.isArray(options.lockAvailabilitySequence)
    ? options.lockAvailabilitySequence.slice()
    : [];
  const mapRows = options.mapRows || [MAP_INFO_HEADERS, mapRow()];
  const routeRows = options.routeRows || [ROUTES_HEADERS];
  const defaultImportReceiptRows = [IMPORT_RECEIPT_HEADERS].concat(
    mapRows.slice(1).filter((row) => row[6]).map((row) => importReceiptRow({
      pinId: String(row[8] || 'pin'),
      fileId: String(row[6]),
      sourceDriveFileId: ''
    }))
  );
  const sheets = {
    map_info: createSheet('map_info', mapRows, options.mapFormulas, audit),
    config: createSheet('config', [
      ['設定項目', '値', '説明'],
      ['RENAME_FILE_WITH_TITLE', options.renameFileWithTitle ? 'true' : 'false', '']
    ], null, audit),
    routes: createSheet('routes', routeRows, options.routeFormulas, audit),
    route_pins: createSheet('route_pins', options.routePinRows || [ROUTE_PINS_HEADERS], options.routePinFormulas, audit),
    route_cache: createSheet('route_cache', options.routeCacheRows || [ROUTE_CACHE_HEADERS], options.routeCacheFormulas, audit),
    share_links: createSheet('share_links', options.shareRows || [['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds']], null, audit)
  };
  if (options.includeImportReceipts !== false) {
    sheets.import_receipts = createSheet(
      'import_receipts',
      Object.prototype.hasOwnProperty.call(options, 'importReceiptRows')
        ? options.importReceiptRows : defaultImportReceiptRows,
      null,
      audit
    );
  }
  Object.entries(options.failDeleteRowsStarts || {}).forEach(([sheetName, starts]) => {
    (starts || []).forEach((start) => sheets[sheetName].failDeleteRowsStarts.add(start));
  });
  const spreadsheet = {
    getSheetByName(name) {
      return sheets[name] || null;
    }
  };
  const lock = {
    tryLock(timeoutMs) {
      if (typeof options.beforeLockAttempt === 'function') {
        options.beforeLockAttempt({ attempt: audit.lock.tryCalls.length + 1, sheets, audit });
      }
      audit.lock.tryCalls.push(timeoutMs);
      if (audit.lock.held) {
        audit.lock.nestedAttempts += 1;
        return false;
      }
      const available = audit.lock.availabilitySequence.length > 0
        ? audit.lock.availabilitySequence.shift()
        : audit.lock.available;
      if (!available) return false;
      audit.lock.held = true;
      audit.lock.maxDepth = Math.max(audit.lock.maxDepth, 1);
      return true;
    },
    releaseLock() {
      assert.equal(audit.lock.held, true, 'releaseLock should only run while held');
      audit.lock.held = false;
      audit.lock.releaseCalls += 1;
    }
  };
  const driveFiles = {};
  (options.mapRows || mapRows).slice(1).forEach((row) => {
    if (row[6]) driveFiles[row[6]] = { name: options.driveFileName || 'original.jpg', trashed: false };
  });
  Object.entries(options.driveFiles || {}).forEach(([fileId, value]) => {
    driveFiles[fileId] = { name: value.name || 'original.jpg', trashed: value.trashed === true };
  });
  const context = {
    console: {
      error(message) { audit.errors.push(String(message)); },
      log: console.log.bind(console),
      warn: console.warn.bind(console)
    },
    Date,
    JSON,
    Set,
    Logger: { log(message) { audit.logs.push(String(message)); } },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      getUuid: () => 'generated-id',
      formatDate: () => '2026/07/10 10:00:00'
    },
    CacheService: {
      getScriptCache: () => ({ get: (key) => key === `EDIT_TOKEN_${TEST_EDIT_TOKEN}` ? '1' : null })
    },
    LockService: { getScriptLock: () => lock },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    DriveApp: {
      getFileById(fileId) {
        audit.driveCalls.push({ method: 'getFileById', fileId, lockHeld: audit.lock.held });
        const file = driveFiles[fileId] || (driveFiles[fileId] = { name: 'unknown.jpg', trashed: false });
        return {
          getName() {
            audit.driveCalls.push({ method: 'getName', fileId, lockHeld: audit.lock.held });
            return file.name;
          },
          setName(name) {
            audit.driveCalls.push({ method: 'setName', fileId, name, lockHeld: audit.lock.held });
            if (options.driveRenameError) throw new Error('simulated Drive rename failure');
            file.name = name;
          },
          setTrashed(trashed) {
            audit.driveCalls.push({ method: 'setTrashed', fileId, trashed, lockHeld: audit.lock.held });
            if ((options.driveTrashErrors || []).includes(fileId)) {
              throw new Error('simulated Drive trash failure');
            }
            file.trashed = trashed;
          }
        };
      }
    }
  };

  vm.runInNewContext(`${codeJs}
globalThis.__writeBatchApi = {
  updatePinDetails,
  movePin,
  unplacePin,
  bulkUpdatePinStatus,
  updateRoutesOrder,
  updatePin,
  saveRouteGroup,
  setRoutePins,
  deleteRoutePinsForRoute_,
  deleteRoutePinsForPinIds_,
  deleteRouteCacheRowsForRouteIds_,
  putRouteCache,
  invalidateRouteCacheForPin,
  invalidateRouteCacheForRoute,
  deleteRouteGroup,
  deletePin,
  bulkDeletePins,
  getRouteCache,
  readLatestRouteCacheEntryByCacheKey_,
  getSharedRoadRouteCache
};`, context);

  return { api: context.__writeBatchApi, audit, sheets, driveFiles };
}

function withEditToken(payload) {
  return { ...payload, __editToken: TEST_EDIT_TOKEN };
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function mapWrites(audit) {
  return audit.writes.filter((write) => write.sheet === 'map_info');
}

test('updatePinDetails batches a full edit into one formula-preserving A:O write', () => {
  const formulas = [MAP_INFO_HEADERS.map(() => ''), mapRow().map(() => '')];
  formulas[1][7] = '=IMAGE("https://example.com/photo")';
  const harness = createHarness({ mapFormulas: formulas });

  const result = plain(harness.api.updatePinDetails(withEditToken({
    id: 'pin-1',
    title: ' Updated ',
    description: '',
    color: '#e53935',
    icon: '',
    links: [],
    status: '',
    tags: [],
    eventAt: ''
  })));

  assert.equal(result.ok, true);
  assert.deepEqual(result.links, []);
  assert.equal(harness.sheets.map_info.rows[1][1], 'Updated');
  assert.equal(harness.sheets.map_info.rows[1][2], '');
  assert.equal(harness.sheets.map_info.rows[1][5], '#e53935');
  assert.equal(harness.sheets.map_info.rows[1][9], '');
  assert.equal(harness.sheets.map_info.rows[1][10], '');
  assert.equal(harness.sheets.map_info.rows[1][11], '');
  assert.equal(harness.sheets.map_info.rows[1][12], '');
  assert.equal(harness.sheets.map_info.rows[1][14], 'default');
  assert.equal(harness.sheets.map_info.formulas[1][7], '=IMAGE("https://example.com/photo")');
  assert.match(String(harness.sheets.map_info.rows[1][13]), /^\d{4}-\d{2}-\d{2}T/);
  const writes = mapWrites(harness.audit);
  assert.equal(writes.length, 1);
  assert.deepEqual({ method: writes[0].method, row: writes[0].row, column: writes[0].column, numRows: writes[0].numRows, numColumns: writes[0].numColumns }, {
    method: 'setValues', row: 2, column: 1, numRows: 1, numColumns: 15
  });
  assert.equal(writes[0].lockHeld, true);
  assert.equal(harness.audit.reads.every((read) => read.lockHeld), true);
});

test('updatePinDetails preserves omitted values while normalizing an omitted icon', () => {
  const harness = createHarness({ mapRows: [MAP_INFO_HEADERS, mapRow({ icon: 'legacy-invalid' })] });

  const result = plain(harness.api.updatePinDetails(withEditToken({ id: 'pin-1', title: 'Partial' })));

  const row = harness.sheets.map_info.rows[1];
  assert.equal(result.ok, true);
  assert.deepEqual(result.links, ['https://example.com/original']);
  assert.equal(row[2], 'Original description');
  assert.equal(row[5], '#43a047');
  assert.equal(row[9], 'https://example.com/original');
  assert.equal(row[10], '対応中');
  assert.equal(row[11], 'alpha|beta');
  assert.equal(row[12], '2026-07-10T12:00');
  assert.equal(row[14], 'default');
});

test('updatePinDetails keeps existing null compatibility and referenceUrls support', () => {
  let harness = createHarness({ mapRows: [MAP_INFO_HEADERS, mapRow({ icon: 'legacy-invalid' })] });

  harness.api.updatePinDetails(withEditToken({
    id: 'pin-1', title: 'Nulls', description: null, links: null,
    color: null, icon: null, status: null, tags: null, eventAt: null
  }));
  let row = harness.sheets.map_info.rows[1];
  assert.equal(row[2], '');
  assert.equal(row[5], '#43a047');
  assert.equal(row[9], '');
  assert.equal(row[10], '対応中');
  assert.equal(row[11], 'alpha|beta');
  assert.equal(row[12], '2026-07-10T12:00');
  assert.equal(row[14], 'default');

  harness = createHarness();
  const result = plain(harness.api.updatePinDetails(withEditToken({
    id: 'pin-1', title: 'References', referenceUrls: ['https://example.com/new', 'invalid']
  })));
  row = harness.sheets.map_info.rows[1];
  assert.equal(row[9], 'https://example.com/new');
  assert.deepEqual(result.links, ['https://example.com/new']);
});

test('updatePinDetails renames photos only when enabled and never while locked', () => {
  let harness = createHarness({
    renameFileWithTitle: true,
    mapRows: [MAP_INFO_HEADERS, mapRow({ fileId: 'photo-file' })]
  });
  harness.api.updatePinDetails(withEditToken({ id: 'pin-1', title: 'Renamed' }));
  assert.equal(harness.audit.driveCalls.some((call) => call.method === 'setName'), true);
  assert.equal(harness.audit.driveCalls.every((call) => call.lockHeld === false), true);

  harness = createHarness({
    renameFileWithTitle: false,
    mapRows: [MAP_INFO_HEADERS, mapRow({ fileId: 'photo-file' })]
  });
  harness.api.updatePinDetails(withEditToken({ id: 'pin-1', title: 'No rename' }));
  assert.equal(harness.audit.driveCalls.length, 0);

  harness = createHarness({ renameFileWithTitle: true });
  harness.api.updatePinDetails(withEditToken({ id: 'pin-1', title: 'No photo' }));
  assert.equal(harness.audit.driveCalls.length, 0);
});

test('updatePinDetails never renames a directly linked Drive import source', () => {
  const harness = createHarness({
    renameFileWithTitle: true,
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-source', fileId: 'source-file' })],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-source', fileId: 'source-file', sourceDriveFileId: 'source-file'
      })
    ]
  });

  const result = plain(harness.api.updatePinDetails(withEditToken({
    id: 'pin-source', title: 'Source stays unchanged'
  })));

  assert.equal(result.ok, true);
  assert.equal(harness.sheets.map_info.rows[1][1], 'Source stays unchanged');
  assert.equal(harness.audit.driveCalls.length, 0);
  assert.equal(harness.driveFiles['source-file'].name, 'original.jpg');
});

test('updatePinDetails fails closed on Drive rename when ownership receipts are missing', () => {
  const harness = createHarness({
    includeImportReceipts: false,
    renameFileWithTitle: true,
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-unknown', fileId: 'unknown-file' })]
  });

  const result = plain(harness.api.updatePinDetails(withEditToken({
    id: 'pin-unknown', title: 'Metadata only'
  })));

  assert.equal(result.ok, true);
  assert.equal(harness.sheets.map_info.rows[1][1], 'Metadata only');
  assert.equal(harness.audit.driveCalls.length, 0);
  assert.equal(harness.driveFiles['unknown-file'].name, 'original.jpg');
});

test('updatePinDetails also fails closed when the ownership receipt row is missing', () => {
  const harness = createHarness({
    importReceiptRows: [IMPORT_RECEIPT_HEADERS],
    renameFileWithTitle: true,
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-unknown-row', fileId: 'unknown-row-file' })]
  });

  const result = plain(harness.api.updatePinDetails(withEditToken({
    id: 'pin-unknown-row', title: 'Metadata only'
  })));

  assert.equal(result.ok, true);
  assert.equal(harness.sheets.map_info.rows[1][1], 'Metadata only');
  assert.equal(harness.audit.driveCalls.length, 0);
  assert.equal(harness.driveFiles['unknown-row-file'].name, 'original.jpg');
});

test('Drive rename failure propagates after the row is committed and the lock is released', () => {
  const harness = createHarness({
    renameFileWithTitle: true,
    driveRenameError: true,
    mapRows: [MAP_INFO_HEADERS, mapRow({ fileId: 'photo-file' })]
  });

  assert.throws(
    () => harness.api.updatePinDetails(withEditToken({ id: 'pin-1', title: 'Committed title' })),
    /simulated Drive rename failure/
  );
  assert.equal(harness.sheets.map_info.rows[1][1], 'Committed title');
  assert.equal(harness.audit.lock.held, false);
  assert.equal(harness.audit.lock.releaseCalls, 1);
  assert.equal(harness.audit.driveCalls.every((call) => call.lockHeld === false), true);
});

test('lock failure returns a retry message without Spreadsheet access', () => {
  const harness = createHarness({ lockAvailable: false });

  const result = plain(harness.api.updatePinDetails(withEditToken({ id: 'pin-1', title: 'Blocked' })));

  assert.deepEqual(result, { ok: false, error: '別の更新処理が実行中です。少し待ってから再試行してください。' });
  assert.equal(harness.audit.reads.length, 0);
  assert.equal(harness.audit.writes.length, 0);
  assert.equal(harness.audit.lock.releaseCalls, 0);
  assert.equal(harness.audit.lock.tryCalls.length, 1);
  assert.equal(Number.isFinite(harness.audit.lock.tryCalls[0]), true);
});

test('Spreadsheet exceptions release the shared lock', () => {
  const harness = createHarness();
  harness.sheets.map_info.failNextSetValues = true;

  assert.throws(
    () => harness.api.updatePinDetails(withEditToken({ id: 'pin-1', title: 'Fails' })),
    /simulated Spreadsheet write failure/
  );
  assert.equal(harness.audit.lock.held, false);
  assert.equal(harness.audit.lock.releaseCalls, 1);
});

test('movePin and unplacePin retain two narrow writes under the lock', () => {
  let harness = createHarness();
  let result = plain(harness.api.movePin(withEditToken({ id: 'pin-1', lat: 36, lng: 140 })));
  assert.equal(result.ok, true);
  let writes = mapWrites(harness.audit);
  assert.equal(writes.length, 2);
  assert.deepEqual(writes.map((write) => [write.column, write.numColumns]), [[4, 2], [14, 1]]);
  assert.equal(writes.every((write) => write.lockHeld), true);

  harness = createHarness();
  result = plain(harness.api.unplacePin(withEditToken({ id: 'pin-1' })));
  assert.equal(result.ok, true);
  writes = mapWrites(harness.audit);
  assert.equal(writes.length, 2);
  assert.deepEqual(writes.map((write) => [write.column, write.numColumns]), [[4, 2], [14, 1]]);
  assert.equal(writes.every((write) => write.lockHeld), true);
});

test('unplacePin clears coordinates and invalidates only routes that contain the pin under one lock', () => {
  const routePinRows = [
    ROUTE_PINS_HEADERS,
    routePinRow('route-target', 'pin-1'),
    routePinRow('route-other', 'pin-2')
  ];
  const routeCacheRows = [
    ROUTE_CACHE_HEADERS,
    routeCacheRow('target-cache', 'route-target'),
    routeCacheRow('other-cache', 'route-other')
  ];
  const harness = createHarness({ routePinRows, routeCacheRows });
  const originalRoutePins = cloneMatrix(harness.sheets.route_pins.rows);

  const result = plain(harness.api.unplacePin(withEditToken({ id: 'pin-1' })));

  assert.deepEqual(result, { ok: true });
  assert.deepEqual(harness.sheets.map_info.rows[1].slice(3, 5), ['', '']);
  assert.match(String(harness.sheets.map_info.rows[1][13]), /^\d{4}-\d{2}-\d{2}T/);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-other']);
  assert.deepEqual(harness.sheets.route_pins.rows, originalRoutePins);
  assert.equal(harness.audit.reads.every((read) => read.lockHeld), true);
  assert.equal(harness.audit.writes.every((write) => write.lockHeld), true);
  assert.equal(harness.audit.lock.releaseCalls, 1);
  assert.equal(harness.audit.lock.held, false);
});

test('legacy updatePin takes two sequential non-nested locks', () => {
  const harness = createHarness();

  const result = plain(harness.api.updatePin(withEditToken({
    id: 'pin-1', title: 'Legacy', lat: 36, lng: 140
  })));

  assert.equal(result.ok, true);
  assert.equal(harness.audit.lock.tryCalls.length, 2);
  assert.equal(harness.audit.lock.releaseCalls, 2);
  assert.equal(harness.audit.lock.nestedAttempts, 0);
  assert.equal(harness.audit.lock.held, false);
});

test('bulkUpdatePinStatus scans once, deduplicates IDs, and skips missing IDs', () => {
  const harness = createHarness({ mapRows: [
    MAP_INFO_HEADERS,
    mapRow({ id: 'pin-a', status: '未対応' }),
    mapRow({ id: 'pin-b', status: '未対応' }),
    mapRow({ id: 'pin-c', status: '未対応' })
  ] });

  const result = plain(harness.api.bulkUpdatePinStatus(withEditToken({
    ids: ['pin-a', 'pin-a', 'missing', 'pin-c'], status: '完了'
  })));

  assert.deepEqual(result, { ok: true, updatedCount: 2 });
  assert.equal(harness.sheets.map_info.rows[1][10], '完了');
  assert.equal(harness.sheets.map_info.rows[2][10], '未対応');
  assert.equal(harness.sheets.map_info.rows[3][10], '完了');
  assert.equal(harness.sheets.map_info.rows[1][13], harness.sheets.map_info.rows[3][13]);
  assert.deepEqual(plain(harness.audit.rangeListSets.map((entry) => entry.addresses)), [['K2', 'K4'], ['N2', 'N4']]);
  assert.equal(harness.audit.rangeListSets.every((entry) => entry.lockHeld), true);
  assert.equal(harness.audit.reads.every((read) => read.lockHeld), true);
  assert.equal(harness.audit.writes.length, 0);
});

test('bulkUpdatePinStatus validates empty IDs and invalid status before locking', () => {
  let harness = createHarness();
  let result = plain(harness.api.bulkUpdatePinStatus(withEditToken({ ids: [], status: '完了' })));
  assert.equal(result.ok, false);
  assert.equal(harness.audit.lock.tryCalls.length, 0);

  harness = createHarness();
  result = plain(harness.api.bulkUpdatePinStatus(withEditToken({ ids: ['pin-1'], status: 'invalid' })));
  assert.equal(result.ok, false);
  assert.equal(harness.audit.lock.tryCalls.length, 0);
});

test('bulkUpdatePinStatus avoids empty RangeLists when no IDs match', () => {
  const harness = createHarness();

  const result = plain(harness.api.bulkUpdatePinStatus(withEditToken({ ids: ['missing'], status: '完了' })));

  assert.deepEqual(result, { ok: true, updatedCount: 0 });
  assert.equal(harness.audit.rangeListCalls.length, 0);
  assert.equal(harness.audit.rangeListSets.length, 0);
});

test('bulkUpdatePinStatus chunks RangeLists at 500 rows and uses setValue', () => {
  function run(count) {
    const rows = [MAP_INFO_HEADERS];
    const ids = [];
    for (let index = 0; index < count; index += 1) {
      const id = `pin-${index}`;
      ids.push(id);
      rows.push(mapRow({ id, status: '未対応' }));
    }
    const harness = createHarness({ mapRows: rows });
    const result = harness.api.bulkUpdatePinStatus(withEditToken({ ids, status: '対応中' }));
    return { harness, result };
  }

  let execution = run(500);
  assert.equal(execution.result.updatedCount, 500);
  assert.deepEqual(execution.harness.audit.rangeListSets.map((entry) => entry.addresses.length), [500, 500]);

  execution = run(501);
  assert.equal(execution.result.updatedCount, 501);
  assert.deepEqual(execution.harness.audit.rangeListSets.map((entry) => entry.addresses.length), [500, 500, 1, 1]);
  assert.equal(execution.harness.audit.writes.length, 0);
});

test('updateRoutesOrder writes one J range and keeps unspecified physical order and formulas', () => {
  const routeRows = [
    ROUTES_HEADERS,
    routeRow('route-a', 8),
    ['', '', '', '', '', '', '', '', '', 77, '', '', '', ''],
    routeRow('route-b', 3),
    routeRow('route-c', 4)
  ];
  const formulas = routeRows.map((row) => row.map(() => ''));
  formulas[2][9] = '=ROW()';
  const harness = createHarness({ routeRows, routeFormulas: formulas });

  const result = plain(harness.api.updateRoutesOrder(withEditToken({ orderedIds: ['route-c'] })));

  assert.equal(result.ok, true);
  assert.deepEqual(result.routeGroups.map((group) => group.routeId), ['route-c', 'route-a', 'route-b']);
  assert.equal(harness.sheets.routes.rows[1][9], 1);
  assert.equal(harness.sheets.routes.rows[3][9], 2);
  assert.equal(harness.sheets.routes.rows[4][9], 0);
  assert.equal(harness.sheets.routes.formulas[2][9], '=ROW()');
  const routeWrites = harness.audit.writes.filter((write) => write.sheet === 'routes');
  assert.equal(routeWrites.length, 1);
  assert.deepEqual([routeWrites[0].method, routeWrites[0].row, routeWrites[0].column, routeWrites[0].numRows, routeWrites[0].numColumns], ['setValues', 2, 10, 4, 1]);
  assert.equal(routeWrites[0].lockHeld, true);
  assert.equal(harness.audit.reads.every((read) => read.lockHeld), true);
});

test('updateRoutesOrder supports full, one-route, and zero-route orders', () => {
  let harness = createHarness({ routeRows: [ROUTES_HEADERS, routeRow('a', 0), routeRow('b', 1)] });
  let result = plain(harness.api.updateRoutesOrder(withEditToken({ orderedIds: ['b', 'a'] })));
  assert.deepEqual(result.routeGroups.map((group) => group.routeId), ['b', 'a']);
  assert.equal(harness.audit.writes.filter((write) => write.sheet === 'routes').length, 1);

  harness = createHarness({ routeRows: [ROUTES_HEADERS, routeRow('only', 9)] });
  result = plain(harness.api.updateRoutesOrder(withEditToken({ orderedIds: [] })));
  assert.deepEqual(result.routeGroups.map((group) => group.routeId), ['only']);
  assert.equal(harness.sheets.routes.rows[1][9], 0);

  harness = createHarness({ routeRows: [ROUTES_HEADERS] });
  result = plain(harness.api.updateRoutesOrder(withEditToken({ orderedIds: [] })));
  assert.deepEqual(result, { ok: true, routeGroups: [] });
  assert.equal(harness.audit.writes.filter((write) => write.sheet === 'routes').length, 0);
});

test('updateRoutesOrder rejects duplicate and missing route IDs before writing', () => {
  let harness = createHarness({ routeRows: [ROUTES_HEADERS, routeRow('a', 0)] });
  let result = plain(harness.api.updateRoutesOrder(withEditToken({ orderedIds: ['a', 'a'] })));
  assert.deepEqual(result, { ok: false, error: 'duplicate_route_id', routeId: 'a' });
  assert.equal(harness.audit.writes.length, 0);

  harness = createHarness({ routeRows: [ROUTES_HEADERS, routeRow('a', 0)] });
  result = plain(harness.api.updateRoutesOrder(withEditToken({ orderedIds: ['missing'] })));
  assert.deepEqual(result, { ok: false, error: 'route_not_found', routeId: 'missing' });
  assert.equal(harness.audit.writes.length, 0);
});

test('setRoutePins replaces memberships with one padded fixed-width write and preserves the header and other routes', () => {
  const originalHeader = ROUTE_PINS_HEADERS.slice();
  const harness = createHarness({
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0), routeRow('route-b', 1)],
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'new-pin' }), mapRow({ id: 'other-1' }), mapRow({ id: 'other-2' })],
    routePinRows: [
      ROUTE_PINS_HEADERS,
      routePinRow('route-a', 'old-1', 0, 'old-created-1', 'old-updated-1'),
      routePinRow('route-b', 'other-1', 0, 'keep-created-1', 'keep-updated-1'),
      routePinRow('route-a', 'old-2', 1, 'old-created-2', 'old-updated-2'),
      routePinRow('route-b', 'other-2', 1, 'keep-created-2', 'keep-updated-2'),
      routePinRow('route-a', 'old-3', 2, 'old-created-3', 'old-updated-3')
    ]
  });

  const result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: ['new-pin'] })));

  assert.deepEqual(result, { ok: true, routeId: 'route-a', pinIds: ['new-pin'] });
  assert.deepEqual(harness.sheets.route_pins.rows[0], originalHeader);
  const rows = nonEmptyDataRows(harness.sheets.route_pins);
  assert.deepEqual(rows.slice(0, 2), [
    routePinRow('route-b', 'other-1', 0, 'keep-created-1', 'keep-updated-1'),
    routePinRow('route-b', 'other-2', 1, 'keep-created-2', 'keep-updated-2')
  ]);
  assert.equal(rows[2][0], 'route-a');
  assert.equal(rows[2][1], 'new-pin');
  assert.equal(rows[2][2], 0);
  assert.equal(rows[2][3], rows[2][4]);
  assert.match(String(rows[2][3]), /^\d{4}-\d{2}-\d{2}T/);
  assert.equal(writesFor(harness.audit, 'route_pins', 'setValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_pins', 'deleteRow').length, 0);
  assert.equal(writesFor(harness.audit, 'route_pins', 'deleteRows').length, 0);
  assert.equal(writesFor(harness.audit, 'route_pins', 'appendRow').length, 0);
  const rewrite = writesFor(harness.audit, 'route_pins', 'setValues')[0];
  assert.deepEqual([rewrite.row, rewrite.column, rewrite.numRows, rewrite.numColumns], [2, 1, 5, 5]);
  assert.deepEqual(plain(rewrite.values.slice(3)), [new Array(5).fill(''), new Array(5).fill('')]);
  assert.equal(rewrite.lockHeld, true);
  assert.equal(harness.audit.lock.tryCalls.length, 1);
  assert.equal(harness.audit.lock.releaseCalls, 1);
  assert.equal(harness.audit.lock.held, false);
  assert.equal(harness.audit.dataRangeCalls.filter((call) => call.sheet === 'route_pins').length, 0);
  assert.equal(harness.audit.reads.filter((read) => read.sheet === 'route_pins' && read.method === 'getValues').length, 1);
  assert.equal(harness.audit.driveCalls.length, 0);
});

test('setRoutePins supports zero, one, multiple, and MAX_ROUTE_PINS memberships', () => {
  function run(pinCount, existingRows = []) {
    const pinIds = Array.from({ length: pinCount }, (_, index) => `pin-${index}`);
    const harness = createHarness({
      routeRows: [ROUTES_HEADERS, routeRow('route-a', 0)],
      mapRows: [MAP_INFO_HEADERS].concat(pinIds.map((id) => mapRow({ id }))),
      routePinRows: [ROUTE_PINS_HEADERS].concat(existingRows)
    });
    const result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds })));
    return { harness, result, pinIds };
  }

  let execution = run(0, [routePinRow('route-a', 'legacy', 0)]);
  assert.deepEqual(execution.result, { ok: true, routeId: 'route-a', pinIds: [] });
  assert.equal(execution.harness.sheets.route_pins.getLastRow(), 1);
  assert.equal(writesFor(execution.harness.audit, 'route_pins', 'setValues').length, 1);

  for (const count of [1, 3, 100]) {
    execution = run(count);
    assert.equal(execution.result.ok, true);
    assert.deepEqual(execution.result.pinIds, execution.pinIds);
    assert.deepEqual(nonEmptyDataRows(execution.harness.sheets.route_pins).map((row) => row[2]),
      Array.from({ length: count }, (_, index) => index));
    assert.equal(writesFor(execution.harness.audit, 'route_pins', 'setValues').length, 1);
  }
});

test('setRoutePins rejects unplaced and invalid coordinates atomically inside the mutation lock', () => {
  const invalidCoordinates = [
    { label: 'both null', lat: null, lng: null },
    { label: 'latitude null', lat: null, lng: 139 },
    { label: 'longitude null', lat: 35, lng: null },
    { label: 'latitude NaN equivalent', lat: 'not-a-number', lng: 139 },
    { label: 'longitude NaN equivalent', lat: 35, lng: Number.NaN },
    { label: 'latitude below range', lat: -90.0001, lng: 139 },
    { label: 'latitude above range', lat: 90.0001, lng: 139 },
    { label: 'longitude below range', lat: 35, lng: -180.0001 },
    { label: 'longitude above range', lat: 35, lng: 180.0001 }
  ];

  invalidCoordinates.forEach(({ label, lat, lng }) => {
    const harness = createHarness({
      routeRows: [ROUTES_HEADERS, routeRow('route-a', 0), routeRow('route-b', 1)],
      mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'candidate', lat, lng })],
      routePinRows: [
        ROUTE_PINS_HEADERS,
        routePinRow('route-a', 'existing', 0, 'existing-created', 'existing-updated'),
        routePinRow('route-b', 'other', 0, 'other-created', 'other-updated')
      ]
    });
    const before = cloneMatrix(harness.sheets.route_pins.rows);

    const result = plain(harness.api.setRoutePins(withEditToken({
      routeId: 'route-a',
      pinIds: ['candidate']
    })));

    assert.deepEqual(result, { ok: false, error: 'pin_unplaced', pinId: 'candidate' }, label);
    assert.deepEqual(harness.sheets.route_pins.rows, before, `${label}: memberships must remain unchanged`);
    assert.equal(writesFor(harness.audit, 'route_pins').length, 0, `${label}: no route_pins writes`);
    const mapReads = harness.audit.dataRangeCalls.filter((call) => call.sheet === 'map_info');
    assert.equal(mapReads.length, 1, `${label}: map_info must be read once`);
    assert.equal(mapReads[0].lockHeld, true, `${label}: validation must run inside the lock`);
  });
});

test('setRoutePins rejects a placed and unplaced mixed batch without partially replacing memberships', () => {
  const harness = createHarness({
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0), routeRow('route-b', 1)],
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'placed', lat: 35, lng: 139 }),
      mapRow({ id: 'unplaced', lat: null, lng: null })
    ],
    routePinRows: [
      ROUTE_PINS_HEADERS,
      routePinRow('route-a', 'existing-a', 0, 'a-created', 'a-updated'),
      routePinRow('route-b', 'existing-b', 0, 'b-created', 'b-updated')
    ]
  });
  const before = cloneMatrix(harness.sheets.route_pins.rows);

  const result = plain(harness.api.setRoutePins(withEditToken({
    routeId: 'route-a',
    pinIds: ['placed', 'unplaced']
  })));

  assert.deepEqual(result, { ok: false, error: 'pin_unplaced', pinId: 'unplaced' });
  assert.deepEqual(harness.sheets.route_pins.rows, before);
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);
});

test('setRoutePins route existence lookup uses the calculated routeId value without rewriting its formula', () => {
  const routeRows = [ROUTES_HEADERS, routeRow('route-a', 0)];
  const routeFormulas = routeRows.map((row) => row.map(() => ''));
  routeFormulas[1][0] = '=CONCAT("route-","a")';
  const harness = createHarness({
    routeRows,
    routeFormulas,
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a' })]
  });

  const result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: ['pin-a'] })));

  assert.deepEqual(result, { ok: true, routeId: 'route-a', pinIds: ['pin-a'] });
  assert.equal(harness.sheets.routes.formulas[1][0], '=CONCAT("route-","a")');
  assert.equal(writesFor(harness.audit, 'routes').length, 0);
});

test('setRoutePins keeps validation errors and performs no write for invalid inputs', () => {
  const baseOptions = {
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0)],
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-1' })]
  };

  let harness = createHarness(baseOptions);
  let result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: ['pin-1', 'pin-1'] })));
  assert.deepEqual(result, { ok: false, error: 'pin_ids_duplicated', pinId: 'pin-1' });
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);

  harness = createHarness(baseOptions);
  result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: ['missing'] })));
  assert.deepEqual(result, { ok: false, error: 'pin_not_found', pinId: 'missing' });
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);

  harness = createHarness(baseOptions);
  result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: null })));
  assert.deepEqual(result, { ok: false, error: 'pin_ids_invalid' });
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);

  harness = createHarness(baseOptions);
  result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'missing', pinIds: [] })));
  assert.deepEqual(result, { ok: false, error: 'route_not_found' });
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);

  const tooManyIds = Array.from({ length: 101 }, (_, index) => `pin-${index}`);
  harness = createHarness({
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0)],
    mapRows: [MAP_INFO_HEADERS].concat(tooManyIds.map((id) => mapRow({ id })))
  });
  result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: tooManyIds })));
  assert.deepEqual(result, { ok: false, error: 'too_many_pins' });
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);
});

test('setRoutePins returns the shared busy error without Spreadsheet access and releases the lock after rewrite exceptions', () => {
  let harness = createHarness({
    lockAvailable: false,
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0)]
  });
  let result = plain(harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: [] })));
  assert.deepEqual(result, { ok: false, error: '別の更新処理が実行中です。少し待ってから再試行してください。' });
  assert.equal(harness.audit.reads.length, 0);
  assert.equal(harness.audit.writes.length, 0);
  assert.equal(harness.audit.lock.releaseCalls, 0);

  harness = createHarness({
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0)],
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-1' })]
  });
  harness.sheets.route_pins.failNextSetValues = true;
  assert.throws(
    () => harness.api.setRoutePins(withEditToken({ routeId: 'route-a', pinIds: ['pin-1'] })),
    /simulated Spreadsheet write failure/
  );
  assert.equal(harness.audit.lock.held, false);
  assert.equal(harness.audit.lock.releaseCalls, 1);
});

test('route_pins route and pin deletion use stable single rewrites and preserve headers', () => {
  let harness = createHarness({ routePinRows: [
    ROUTE_PINS_HEADERS,
    routePinRow('route-a', 'pin-a', 0),
    routePinRow('route-b', 'pin-a', 0, 'b-created', 'b-updated'),
    routePinRow('route-a', 'pin-b', 1),
    routePinRow('route-c', 'pin-c', 0, 'c-created', 'c-updated')
  ] });

  let removedRoutes = plain(harness.api.deleteRoutePinsForRoute_('route-a'));
  assert.deepEqual(removedRoutes, ['route-a']);
  assert.deepEqual(harness.sheets.route_pins.rows[0], ROUTE_PINS_HEADERS);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins), [
    routePinRow('route-b', 'pin-a', 0, 'b-created', 'b-updated'),
    routePinRow('route-c', 'pin-c', 0, 'c-created', 'c-updated')
  ]);
  assert.equal(writesFor(harness.audit, 'route_pins', 'setValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_pins', 'deleteRow').length, 0);

  harness = createHarness({ routePinRows: [
    ROUTE_PINS_HEADERS,
    routePinRow('route-a', 'pin-a', 0),
    routePinRow('route-b', 'pin-a', 0),
    routePinRow('route-b', 'pin-b', 1),
    routePinRow('route-c', 'pin-c', 0)
  ] });
  removedRoutes = plain(harness.api.deleteRoutePinsForPinIds_(['pin-a', 'missing'])).sort();
  assert.deepEqual(removedRoutes, ['route-a', 'route-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => [row[0], row[1]]), [
    ['route-b', 'pin-b'], ['route-c', 'pin-c']
  ]);
  assert.equal(writesFor(harness.audit, 'route_pins', 'setValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_pins', 'deleteRow').length, 0);

  harness = createHarness({ routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a', 0)] });
  removedRoutes = plain(harness.api.deleteRoutePinsForPinIds_(['missing']));
  assert.deepEqual(removedRoutes, []);
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);
  assert.deepEqual(harness.sheets.route_pins.rows[0], ROUTE_PINS_HEADERS);
});

test('route_cache deletion preserves unrelated rows and all cache fields while returning the exact count', () => {
  const keepOne = routeCacheRow('keep-1', 'route-b', '[[1,2],[3,4]]', 'provider-b', 'created-b', 'expires-b');
  const keepTwo = routeCacheRow('keep-2', 'route-c', '[[5,6],[7,8]]', 'provider-c', 'created-c', 'expires-c');
  let harness = createHarness({ routeCacheRows: [
    ROUTE_CACHE_HEADERS,
    routeCacheRow('drop-1', 'route-a'),
    keepOne,
    routeCacheRow('drop-2', 'route-d'),
    keepTwo,
    routeCacheRow('drop-3', 'route-a')
  ] });

  let deleted = harness.api.deleteRouteCacheRowsForRouteIds_(['route-a', 'route-d']);
  assert.equal(deleted, 3);
  assert.deepEqual(harness.sheets.route_cache.rows[0], ROUTE_CACHE_HEADERS);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache), [keepOne, keepTwo]);
  assert.equal(writesFor(harness.audit, 'route_cache', 'setValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_cache', 'deleteRow').length, 0);
  assert.deepEqual(plain(harness.api.readLatestRouteCacheEntryByCacheKey_('keep-2')), {
    cacheKey: 'keep-2', routeId: 'route-c', coords: [[5, 6], [7, 8]], provider: 'provider-c', createdAt: 'created-c'
  });

  harness = createHarness({ routeCacheRows: [ROUTE_CACHE_HEADERS, keepOne] });
  deleted = harness.api.deleteRouteCacheRowsForRouteIds_(['missing']);
  assert.equal(deleted, 0);
  assert.equal(writesFor(harness.audit, 'route_cache').length, 0);
});

test('route_pins and route_cache rewrites keep formulas and extension columns attached to retained rows', () => {
  const routePinRows = [
    ROUTE_PINS_HEADERS.concat(['extension']),
    routePinRow('route-drop', 'pin-drop').concat(['drop-extra']),
    routePinRow('route-keep', 'pin-keep', 0, 'keep-created', 'computed-updated').concat(['computed-extra'])
  ];
  const routePinFormulas = routePinRows.map((row) => row.map(() => ''));
  routePinFormulas[1][0] = '=CONCAT("route-","drop")';
  routePinFormulas[2][4] = '=ROW()';
  routePinFormulas[2][5] = '=A3';
  let harness = createHarness({ routePinRows, routePinFormulas });

  assert.deepEqual(plain(harness.api.deleteRoutePinsForRoute_('route-drop')), ['route-drop']);
  assert.deepEqual(harness.sheets.route_pins.rows[0], ROUTE_PINS_HEADERS.concat(['extension']));
  assert.equal(harness.sheets.route_pins.formulas[1][4], '=ROW()');
  assert.equal(harness.sheets.route_pins.formulas[1][5], '=A3');
  assert.equal(harness.sheets.route_pins.getLastRow(), 2);
  assert.equal(writesFor(harness.audit, 'route_pins', 'setValues')[0].numColumns, 6);

  const routeCacheRows = [
    ROUTE_CACHE_HEADERS.concat(['extension']),
    routeCacheRow('drop', 'route-drop').concat(['drop-extra']),
    routeCacheRow('keep', 'route-keep', '[[1,2],[3,4]]', 'osrm', 'computed-created', 'computed-expires').concat(['computed-extra'])
  ];
  const routeCacheFormulas = routeCacheRows.map((row) => row.map(() => ''));
  routeCacheFormulas[1][1] = '=CONCAT("route-","drop")';
  routeCacheFormulas[2][4] = '=NOW()';
  routeCacheFormulas[2][5] = '=E3+1';
  routeCacheFormulas[2][6] = '=B3';
  harness = createHarness({ routeCacheRows, routeCacheFormulas });

  assert.equal(harness.api.deleteRouteCacheRowsForRouteIds_(['route-drop']), 1);
  assert.deepEqual(harness.sheets.route_cache.rows[0], ROUTE_CACHE_HEADERS.concat(['extension']));
  assert.equal(harness.sheets.route_cache.formulas[1][4], '=NOW()');
  assert.equal(harness.sheets.route_cache.formulas[1][5], '=E3+1');
  assert.equal(harness.sheets.route_cache.formulas[1][6], '=B3');
  assert.equal(harness.sheets.route_cache.getLastRow(), 2);
  assert.equal(writesFor(harness.audit, 'route_cache', 'setValues')[0].numColumns, 7);
});

test('shared road-route cache lookup still returns an unrelated retained cache after rewrite deletion', () => {
  const cacheKey = 'osrm|road|false|p1:35.00000,139.00000>p2:36.00000,140.00000';
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'p1', lat: 35, lng: 139 }),
      mapRow({ id: 'p2', lat: 36, lng: 140 })
    ],
    routeRows: [
      ROUTES_HEADERS,
      ['route-keep', 'Keep', '#1e88e5', 'road', false, '', '', '', '', 0, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PINS_HEADERS,
      routePinRow('route-keep', 'p1', 0),
      routePinRow('route-keep', 'p2', 1)
    ],
    routeCacheRows: [
      ROUTE_CACHE_HEADERS,
      routeCacheRow('drop-cache', 'route-drop'),
      routeCacheRow(cacheKey, 'route-keep', '[[35,139],[36,140]]')
    ],
    shareRows: [
      ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'],
      ['2026-07-10T00:00:00.000Z', 'Shared keep route', 'share-token', '', 'or', true, '', '', 'route-keep']
    ]
  });

  assert.equal(harness.api.deleteRouteCacheRowsForRouteIds_(['route-drop']), 1);
  assert.deepEqual(plain(harness.api.getSharedRoadRouteCache({ token: 'share-token', routeId: 'route-keep' })), {
    ok: true,
    routeId: 'route-keep',
    coords: [[35, 139], [36, 140]]
  });
});

test('public route/cache writers share one non-nested lock and report lock failure', () => {
  let harness = createHarness({
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0)],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a')]
  });
  let result = plain(harness.api.invalidateRouteCacheForRoute(withEditToken({ routeId: 'route-a' })));
  assert.deepEqual(result, { ok: true, deleted: 1 });
  assert.equal(harness.audit.lock.tryCalls.length, 1);
  assert.equal(harness.audit.lock.releaseCalls, 1);
  assert.equal(harness.audit.lock.nestedAttempts, 0);
  assert.equal(writesFor(harness.audit, 'route_cache').every((write) => write.lockHeld), true);

  harness = createHarness({ lockAvailable: false });
  result = plain(harness.api.putRouteCache(withEditToken({
    cacheKey: 'new-cache', routeId: 'route-a', coords: [[35, 139], [36, 140]], provider: 'osrm'
  })));
  assert.deepEqual(result, { ok: false, error: '別の更新処理が実行中です。少し待ってから再試行してください。' });
  assert.equal(harness.audit.reads.length, 0);
  assert.equal(harness.audit.writes.length, 0);
});

test('saveRouteGroup validates input before locking and locks Spreadsheet work', () => {
  let harness = createHarness();
  let result = plain(harness.api.saveRouteGroup(withEditToken({ name: '' })));
  assert.deepEqual(result, { ok: false, error: 'route_name_required' });
  assert.equal(harness.audit.lock.tryCalls.length, 0);
  assert.equal(harness.audit.reads.length, 0);
  assert.equal(harness.audit.writes.length, 0);

  harness = createHarness({ lockAvailable: false });
  result = plain(harness.api.saveRouteGroup(withEditToken({ name: 'New route' })));
  assert.deepEqual(result, { ok: false, error: '別の更新処理が実行中です。少し待ってから再試行してください。' });
  assert.equal(harness.audit.reads.length, 0);
  assert.equal(harness.audit.writes.length, 0);
});

test('deleteRouteGroup removes route, memberships, and cache under one lock without duplicate route_pins scans', () => {
  const harness = createHarness({
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0), routeRow('route-b', 1)],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a'), routePinRow('route-b', 'pin-b')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a'), routeCacheRow('cache-b', 'route-b')]
  });

  const result = plain(harness.api.deleteRouteGroup(withEditToken({ routeId: 'route-a' })));

  assert.deepEqual(result, { ok: true });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.routes).map((row) => row[0]), ['route-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[0]), ['route-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-b']);
  assert.equal(harness.audit.lock.tryCalls.length, 1);
  assert.equal(harness.audit.lock.releaseCalls, 1);
  assert.equal(harness.audit.dataRangeCalls.filter((call) => call.sheet === 'route_pins').length, 0);
  assert.equal(harness.audit.reads.filter((read) => read.sheet === 'route_pins' && read.method === 'getValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_pins', 'setValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_cache', 'setValues').length, 1);
});

test('bulkDeletePins uses one current-row commit scan, Drive outside the lock, and descending contiguous deleteRows runs', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-a', fileId: 'file-aaaaaaaaaa' }),
      mapRow({ id: 'pin-c' }),
      mapRow({ id: 'pin-b', fileId: 'file-bbbbbbbbbb' }),
      mapRow({ id: 'keep' }),
      mapRow({ id: 'pin-d', fileId: 'file-dddddddddd' })
    ],
    routePinRows: [
      ROUTE_PINS_HEADERS,
      routePinRow('route-a', 'pin-a'),
      routePinRow('route-fail', 'pin-b'),
      routePinRow('route-c', 'pin-c'),
      routePinRow('route-d', 'pin-d'),
      routePinRow('route-keep', 'keep')
    ],
    routeCacheRows: [
      ROUTE_CACHE_HEADERS,
      routeCacheRow('cache-a', 'route-a'),
      routeCacheRow('cache-fail', 'route-fail'),
      routeCacheRow('cache-c', 'route-c'),
      routeCacheRow('cache-d', 'route-d'),
      routeCacheRow('cache-keep', 'route-keep')
    ],
    driveTrashErrors: ['file-bbbbbbbbbb']
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({
    ids: ['pin-a', 'pin-c', 'pin-b', 'pin-d', 'missing', 'pin-a']
  })));

  assert.deepEqual(result, { ok: true, deletedCount: 3, failedIds: ['pin-b'] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-b', 'keep']);
  assert.deepEqual(writesFor(harness.audit, 'map_info', 'deleteRows').map((write) => [write.row, write.count]), [[6, 1], [2, 2]]);
  assert.equal(writesFor(harness.audit, 'map_info', 'deleteRow').length, 0);
  assert.equal(harness.audit.dataRangeCalls.filter((call) => call.sheet === 'map_info').length, 2);
  assert.equal(harness.audit.reads.filter((read) => read.sheet === 'map_info' && read.method === 'getValues').length, 2);

  const getFileCalls = harness.audit.driveCalls.filter((call) => call.method === 'getFileById');
  const trashCalls = harness.audit.driveCalls.filter((call) => call.method === 'setTrashed');
  assert.equal(getFileCalls.length, 3);
  assert.equal(trashCalls.length, 3);
  assert.equal(harness.audit.driveCalls.every((call) => call.lockHeld === false), true);

  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => [row[0], row[1]]), [
    ['route-fail', 'pin-b'], ['route-keep', 'keep']
  ]);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-fail', 'route-keep']);
  assert.equal(writesFor(harness.audit, 'route_pins', 'setValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_cache', 'setValues').length, 1);
  assert.equal(writesFor(harness.audit, 'route_pins', 'deleteRow').length, 0);
  assert.equal(writesFor(harness.audit, 'route_cache', 'deleteRow').length, 0);
  assert.equal(harness.audit.lock.tryCalls.length, 1);
  assert.equal(harness.audit.lock.releaseCalls, 1);
  assert.equal(harness.audit.lock.nestedAttempts, 0);
  assert.equal(harness.audit.writes.every((write) => write.lockHeld), true);
});

test('bulkDeletePins returns busy after Drive success, logs only pinId/stage, and converges on retry', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'secret-file-a' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a')],
    lockAvailabilitySequence: [false, true]
  });

  let result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a'] })));
  assert.deepEqual(result, { ok: false, error: '別の更新処理が実行中です。少し待ってから再試行してください。' });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-a']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['pin-a']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-a']);
  assert.equal(harness.driveFiles['secret-file-a'].trashed, true);
  assert.equal(harness.audit.logs.some((line) => line.includes('pinId=pin-a') && line.includes('spreadsheet_lock_failed_after_drive')), true);
  assert.equal(harness.audit.logs.some((line) => line.includes('secret-file-a')), false);
  assert.equal(harness.audit.writes.length, 0);

  result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a'] })));
  assert.deepEqual(result, { ok: true, deletedCount: 1, failedIds: [] });
  assert.equal(harness.audit.driveCalls.filter((call) => call.method === 'setTrashed' && call.fileId === 'secret-file-a').length, 2);
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
  assert.equal(harness.sheets.route_pins.getLastRow(), 1);
  assert.equal(harness.sheets.route_cache.getLastRow(), 1);
  assert.equal(harness.audit.lock.held, false);
  assert.equal(harness.audit.lock.releaseCalls, 1);
});

test('bulkDeletePins treats a row removed before commit as completed without deleting its old row occupant', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'file-aaaaaaaaaa' }), mapRow({ id: 'pin-b' })],
    beforeLockAttempt({ sheets }) {
      sheets.map_info.rows.splice(1, 1);
      sheets.map_info.formulas.splice(1, 1);
      sheets.map_info.maxRows -= 1;
    }
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a'] })));

  assert.deepEqual(result, { ok: true, deletedCount: 0, failedIds: [] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-b']);
  assert.equal(writesFor(harness.audit, 'map_info', 'deleteRows').length, 0);
  assert.equal(writesFor(harness.audit, 'map_info', 'deleteRow').length, 0);
});

test('bulkDeletePins re-resolves moved rows by pinId and never deletes a different pin at the stale row number', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'file-aaaaaaaaaa' }), mapRow({ id: 'pin-b' })],
    beforeLockAttempt({ sheets }) {
      sheets.map_info.rows.splice(1, 0, mapRow({ id: 'inserted' }));
      sheets.map_info.formulas.splice(1, 0, MAP_INFO_HEADERS.map(() => ''));
      sheets.map_info.maxRows += 1;
    }
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a'] })));

  assert.deepEqual(result, { ok: true, deletedCount: 1, failedIds: [] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['inserted', 'pin-b']);
  assert.deepEqual(writesFor(harness.audit, 'map_info', 'deleteRows').map((write) => [write.row, write.count]), [[3, 1]]);
});

test('bulkDeletePins isolates a fileId conflict and commits other verified photo deletions', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'old-file-aaaa' }), mapRow({ id: 'pin-b', fileId: 'file-bbbbbbbbbb' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a'), routePinRow('route-b', 'pin-b')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a'), routeCacheRow('cache-b', 'route-b')],
    beforeLockAttempt({ sheets }) {
      sheets.map_info.rows[1][6] = 'replacement-file';
    }
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a', 'pin-b'] })));

  assert.deepEqual(result, { ok: true, deletedCount: 1, failedIds: ['pin-a'] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-a']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['pin-a']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-a']);
  assert.equal(harness.audit.driveCalls.filter((call) => call.method === 'setTrashed').length, 2);
  assert.equal(harness.audit.logs.some((line) => line.includes('pinId=pin-a') && line.includes('file_id_conflict')), true);
  assert.equal(harness.audit.logs.some((line) => line.includes('old-file-aaaa') || line.includes('replacement-file') || line.includes('file-bbbbbbbbbb')), false);
  assert.equal(harness.audit.lock.releaseCalls, 1);
  assert.equal(harness.audit.lock.held, false);
});

test('bulkDeletePins reports every ID in a failed contiguous deleteRows run and leaves related route data intact', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a' }), mapRow({ id: 'pin-b' }), mapRow({ id: 'keep' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a'), routePinRow('route-b', 'pin-b')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a'), routeCacheRow('cache-b', 'route-b')],
    failDeleteRowsStarts: { map_info: [2] }
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a', 'pin-b'] })));

  assert.deepEqual(result, { ok: true, deletedCount: 0, failedIds: ['pin-b', 'pin-a'] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-a', 'pin-b', 'keep']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['pin-a', 'pin-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-a', 'route-b']);
  assert.equal(writesFor(harness.audit, 'route_pins').length, 0);
  assert.equal(writesFor(harness.audit, 'route_cache').length, 0);
  assert.equal(harness.audit.lock.held, false);
  assert.equal(harness.audit.lock.releaseCalls, 1);
});

test('bulkDeletePins commits successful runs and cleans only those IDs when another run fails', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a' }), mapRow({ id: 'keep' }), mapRow({ id: 'pin-b' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a'), routePinRow('route-b', 'pin-b')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a'), routeCacheRow('cache-b', 'route-b')],
    failDeleteRowsStarts: { map_info: [4] }
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a', 'pin-b'] })));

  assert.deepEqual(result, { ok: true, deletedCount: 1, failedIds: ['pin-b'] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['keep', 'pin-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['pin-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-b']);
  assert.equal(harness.audit.lock.held, false);
  assert.equal(harness.audit.lock.releaseCalls, 1);
});

test('bulkDeletePins retry repairs route memberships after post-map cache cleanup failure', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a')]
  });
  harness.sheets.route_cache.failNextSetValues = true;

  assert.throws(
    () => harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a'] })),
    /simulated Spreadsheet write failure/
  );
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['pin-a']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-a']);
  assert.equal(harness.audit.lock.held, false);

  const result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a'] })));
  assert.deepEqual(result, { ok: true, deletedCount: 0, failedIds: [] });
  assert.equal(harness.sheets.route_pins.getLastRow(), 1);
  assert.equal(harness.sheets.route_cache.getLastRow(), 1);
  assert.equal(harness.audit.lock.held, false);
});

test('bulkDeletePins handles all-missing and all-Drive-failure batches without unrelated cleanup', () => {
  let harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'keep' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-keep', 'keep')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-keep', 'route-keep')]
  });
  let result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['missing'] })));
  assert.deepEqual(result, { ok: true, deletedCount: 0, failedIds: [] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['keep']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-keep']);

  harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'file-aaaaaaaaaa' }), mapRow({ id: 'pin-b', fileId: 'file-bbbbbbbbbb' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a'), routePinRow('route-b', 'pin-b')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a'), routeCacheRow('cache-b', 'route-b')],
    driveTrashErrors: ['file-aaaaaaaaaa', 'file-bbbbbbbbbb'],
    lockAvailable: false
  });
  result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a', 'pin-b'] })));
  assert.deepEqual(result, { ok: true, deletedCount: 0, failedIds: ['pin-b', 'pin-a'] });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-a', 'pin-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['pin-a', 'pin-b']);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_cache).map((row) => row[1]), ['route-a', 'route-b']);
  assert.equal(harness.audit.lock.tryCalls.length, 0);
});

test('single pin and route deletion retries repair cleanup after the parent row was committed', () => {
  let harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a' })],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a')]
  });
  harness.sheets.route_cache.failNextSetValues = true;
  assert.throws(() => harness.api.deletePin(withEditToken({ id: 'pin-a' })), /simulated Spreadsheet write failure/);
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[1]), ['pin-a']);
  let result = plain(harness.api.deletePin(withEditToken({ id: 'pin-a' })));
  assert.deepEqual(result, { ok: false, error: 'id not found' });
  assert.equal(harness.sheets.route_pins.getLastRow(), 1);
  assert.equal(harness.sheets.route_cache.getLastRow(), 1);

  harness = createHarness({
    routeRows: [ROUTES_HEADERS, routeRow('route-a', 0)],
    routePinRows: [ROUTE_PINS_HEADERS, routePinRow('route-a', 'pin-a')],
    routeCacheRows: [ROUTE_CACHE_HEADERS, routeCacheRow('cache-a', 'route-a')]
  });
  harness.sheets.route_cache.failNextSetValues = true;
  assert.throws(() => harness.api.deleteRouteGroup(withEditToken({ routeId: 'route-a' })), /simulated Spreadsheet write failure/);
  assert.equal(harness.sheets.routes.getLastRow(), 1);
  assert.deepEqual(nonEmptyDataRows(harness.sheets.route_pins).map((row) => row[0]), ['route-a']);
  result = plain(harness.api.deleteRouteGroup(withEditToken({ routeId: 'route-a' })));
  assert.deepEqual(result, { ok: false, error: 'route_not_found' });
  assert.equal(harness.sheets.route_pins.getLastRow(), 1);
  assert.equal(harness.sheets.route_cache.getLastRow(), 1);
});

test('deletePin applies the same post-Drive lock, current-row, missing-row, and fileId conflict rules', () => {
  let harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'file-aaaaaaaaaa' })],
    lockAvailable: false
  });
  let result = plain(harness.api.deletePin(withEditToken({ id: 'pin-a' })));
  assert.deepEqual(result, { ok: false, error: '別の更新処理が実行中です。少し待ってから再試行してください。' });
  assert.equal(harness.sheets.map_info.getLastRow(), 2);
  assert.equal(harness.audit.logs.some((line) => line.includes('pinId=pin-a') && line.includes('spreadsheet_lock_failed_after_drive')), true);

  harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'file-aaaaaaaaaa' }), mapRow({ id: 'pin-b' })],
    beforeLockAttempt({ sheets }) {
      sheets.map_info.rows.splice(1, 1);
      sheets.map_info.formulas.splice(1, 1);
      sheets.map_info.maxRows -= 1;
    }
  });
  result = plain(harness.api.deletePin(withEditToken({ id: 'pin-a' })));
  assert.deepEqual(result, { ok: true });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-b']);
  assert.equal(writesFor(harness.audit, 'map_info').length, 0);

  harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-a', fileId: 'old-file-aaaa' })],
    beforeLockAttempt({ sheets }) { sheets.map_info.rows[1][6] = 'new-file-aaaa'; }
  });
  result = plain(harness.api.deletePin(withEditToken({ id: 'pin-a' })));
  assert.equal(result.ok, false);
  assert.match(result.error, /更新|競合|再試行/);
  assert.equal(harness.sheets.map_info.getLastRow(), 2);
  assert.equal(harness.audit.logs.some((line) => line.includes('pinId=pin-a') && line.includes('file_id_conflict')), true);
  assert.equal(harness.audit.lock.held, false);
});

test('pin deletion does not trash a Drive file still referenced by another surviving pin', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-a', fileId: 'shared-file' }),
      mapRow({ id: 'pin-b', fileId: 'shared-file' })
    ]
  });
  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-a' })));
  assert.deepEqual(result, { ok: true });
  assert.deepEqual(nonEmptyDataRows(harness.sheets.map_info).map((row) => row[8]), ['pin-b']);
  assert.equal(harness.audit.driveCalls.filter((call) => call.method === 'getFileById').length, 0);
  assert.equal(harness.driveFiles['shared-file'].trashed, false);
});

test('pin deletion removes the pin but never trashes a directly linked Drive import source', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-source', fileId: 'source-file' })],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-source', fileId: 'source-file', sourceDriveFileId: 'source-file'
      })
    ]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-source' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
  assert.equal(harness.audit.driveCalls.length, 0);
  assert.equal(harness.driveFiles['source-file'].trashed, false);
});

test('pin deletion trashes only the managed JPEG and preserves its Drive source', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-managed', fileId: 'managed-jpeg' })],
    driveFiles: { 'source-photo': { name: 'source.jpg' } },
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-managed',
        fileId: 'managed-jpeg',
        sourceDriveFileId: 'source-photo',
        targetFolderId: 'managed-folder'
      })
    ]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-managed' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
  assert.equal(harness.driveFiles['managed-jpeg'].trashed, true);
  assert.equal(harness.driveFiles['source-photo'].trashed, false);
  assert.deepEqual(
    harness.audit.driveCalls.filter((call) => call.method === 'setTrashed').map((call) => call.fileId),
    ['managed-jpeg']
  );
});

test('pin deletion fails closed on Drive cleanup when the ownership receipt sheet is missing', () => {
  const harness = createHarness({
    includeImportReceipts: false,
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-unknown', fileId: 'unknown-file' })]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-unknown' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
  assert.equal(harness.audit.driveCalls.length, 0);
  assert.equal(harness.driveFiles['unknown-file'].trashed, false);
});

test('pin deletion also fails closed when the ownership receipt row is missing', () => {
  const harness = createHarness({
    importReceiptRows: [IMPORT_RECEIPT_HEADERS],
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-unknown-row', fileId: 'unknown-row-file' })]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-unknown-row' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
  assert.equal(harness.audit.driveCalls.length, 0);
  assert.equal(harness.driveFiles['unknown-row-file'].trashed, false);
});

test('pin deletion fails closed when a receipt cannot prove local or Drive managed ownership', () => {
  const harness = createHarness({
    mapRows: [MAP_INFO_HEADERS, mapRow({ id: 'pin-ambiguous', fileId: 'ambiguous-file' })],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-ambiguous',
        fileId: 'ambiguous-file',
        sourceDriveFileId: '',
        targetFolderId: ''
      })
    ]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-ambiguous' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.driveFiles['ambiguous-file'].trashed, false);
  assert.equal(harness.audit.driveCalls.length, 0);
});

test('pin deletion treats malformed source ownership ids as unknown and fails closed', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-malformed-source', fileId: 'photo_DIRECTVALID' })
    ],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-malformed-source',
        fileId: 'photo_DIRECTVALID',
        sourceDriveFileId: 'bad',
        targetFolderId: 'managed-folder'
      })
    ]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-malformed-source' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.driveFiles['photo_DIRECTVALID'].trashed, false);
  assert.equal(harness.audit.driveCalls.length, 0);
});

test('pin deletion treats a malformed nonempty managed target id as unknown', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-malformed-target', fileId: 'photo_MANAGEDVALID' })
    ],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-malformed-target',
        fileId: 'photo_MANAGEDVALID',
        sourceDriveFileId: 'photo_SOURCEVALIDA',
        targetFolderId: 'bad'
      })
    ]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-malformed-target' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.driveFiles['photo_MANAGEDVALID'].trashed, false);
  assert.equal(harness.audit.driveCalls.length, 0);
});

test('source ownership protects the Drive file globally across legacy duplicate pin references', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-source', fileId: 'source-file' }),
      mapRow({ id: 'pin-legacy', fileId: 'source-file' })
    ],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-source', fileId: 'source-file', sourceDriveFileId: 'source-file'
      })
    ]
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({
    ids: ['pin-source', 'pin-legacy']
  })));

  assert.deepEqual(result, { ok: true, deletedCount: 2, failedIds: [] });
  assert.equal(harness.driveFiles['source-file'].trashed, false);
  assert.equal(harness.audit.driveCalls.length, 0);
});

test('managed-copy ownership also protects its source when a legacy pin references that source', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-managed', fileId: 'managed-copy' }),
      mapRow({ id: 'pin-legacy-source', fileId: 'private-source' })
    ],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-managed', fileId: 'managed-copy', sourceDriveFileId: 'private-source'
      })
    ]
  });

  const result = plain(harness.api.deletePin(withEditToken({ id: 'pin-legacy-source' })));

  assert.deepEqual(result, { ok: true });
  assert.equal(harness.driveFiles['private-source'].trashed, false);
  assert.equal(harness.audit.driveCalls.length, 0);
});

test('bulk deletion trashes a shared Drive file once only when every referencing pin is deleted', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-a', fileId: 'shared-file' }),
      mapRow({ id: 'pin-b', fileId: 'shared-file' })
    ]
  });
  const result = plain(harness.api.bulkDeletePins(withEditToken({ ids: ['pin-a', 'pin-b'] })));
  assert.deepEqual(result, { ok: true, deletedCount: 2, failedIds: [] });
  assert.equal(harness.audit.driveCalls.filter((call) => call.method === 'getFileById').length, 1);
  assert.equal(harness.audit.driveCalls.filter((call) => call.method === 'setTrashed').length, 1);
  assert.equal(harness.sheets.map_info.getLastRow(), 1);
});

test('bulk deletion preserves direct sources and trashes only app-managed display copies', () => {
  const harness = createHarness({
    mapRows: [
      MAP_INFO_HEADERS,
      mapRow({ id: 'pin-source', fileId: 'source-file' }),
      mapRow({ id: 'pin-copy', fileId: 'managed-copy' })
    ],
    importReceiptRows: [
      IMPORT_RECEIPT_HEADERS,
      importReceiptRow({
        pinId: 'pin-source', fileId: 'source-file', sourceDriveFileId: 'source-file'
      }),
      importReceiptRow({
        pinId: 'pin-copy', fileId: 'managed-copy', sourceDriveFileId: 'private-source'
      })
    ]
  });

  const result = plain(harness.api.bulkDeletePins(withEditToken({
    ids: ['pin-source', 'pin-copy']
  })));

  assert.deepEqual(result, { ok: true, deletedCount: 2, failedIds: [] });
  assert.equal(harness.driveFiles['source-file'].trashed, false);
  assert.equal(harness.driveFiles['managed-copy'].trashed, true);
  assert.deepEqual(
    harness.audit.driveCalls.filter((call) => call.method === 'setTrashed').map((call) => call.fileId),
    ['managed-copy']
  );
});

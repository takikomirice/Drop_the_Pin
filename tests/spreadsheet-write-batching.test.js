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
    rangeListCalls: [],
    rangeListSets: [],
    driveCalls: [],
    lock: {
      held: false,
      available: true,
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
    failNextSetValues: false,
    getLastRow() {
      return this.rows.length;
    },
    getLastColumn() {
      return this.rows.reduce((max, row) => Math.max(max, row.length), 0);
    },
    getDataRange() {
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
      this.rows.push(row.slice());
      this.formulas.push(row.map(() => ''));
      audit.writes.push({ sheet: name, method: 'appendRow', lockHeld: audit.lock.held });
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
      Array.from({ length: numColumns }, (_, columnOffset) =>
        (source[row - 1 + rowOffset] || [])[column - 1 + columnOffset] || ''
      )
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

function createHarness(options = {}) {
  const audit = createAudit();
  audit.lock.available = options.lockAvailable !== false;
  const mapRows = options.mapRows || [MAP_INFO_HEADERS, mapRow()];
  const routeRows = options.routeRows || [ROUTES_HEADERS];
  const sheets = {
    map_info: createSheet('map_info', mapRows, options.mapFormulas, audit),
    config: createSheet('config', [
      ['設定項目', '値', '説明'],
      ['RENAME_FILE_WITH_TITLE', options.renameFileWithTitle ? 'true' : 'false', '']
    ], null, audit),
    routes: createSheet('routes', routeRows, options.routeFormulas, audit),
    route_pins: createSheet('route_pins', options.routePinRows || [ROUTE_PINS_HEADERS], null, audit),
    route_cache: createSheet('route_cache', [ROUTE_CACHE_HEADERS], null, audit),
    share_links: createSheet('share_links', [['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds']], null, audit)
  };
  const spreadsheet = {
    getSheetByName(name) {
      return sheets[name] || null;
    }
  };
  const lock = {
    tryLock(timeoutMs) {
      audit.lock.tryCalls.push(timeoutMs);
      if (audit.lock.held) {
        audit.lock.nestedAttempts += 1;
        return false;
      }
      if (!audit.lock.available) return false;
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
    if (row[6]) driveFiles[row[6]] = { name: options.driveFileName || 'original.jpg' };
  });
  const context = {
    console,
    Date,
    JSON,
    Set,
    Logger: { log() {} },
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
        const file = driveFiles[fileId] || { name: 'unknown.jpg' };
        return {
          getName() {
            audit.driveCalls.push({ method: 'getName', fileId, lockHeld: audit.lock.held });
            return file.name;
          },
          setName(name) {
            audit.driveCalls.push({ method: 'setName', fileId, name, lockHeld: audit.lock.held });
            if (options.driveRenameError) throw new Error('simulated Drive rename failure');
            file.name = name;
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
  updatePin
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

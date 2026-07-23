const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const TEST_EDIT_TOKEN = 'test-edit-token';

const MAP_INFO_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明',
  '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧',
  '状態', 'タグ', 'イベント時刻', '更新時刻', 'アイコン', '音声ID'
];

function cloneRows(rows) {
  return rows.map((row) => row.slice());
}

function parseA1(a1) {
  const single = /^([A-Z]+)(\d+)$/.exec(a1);
  if (single) {
    const column = single[1].split('').reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
    return { row: Number(single[2]), column, numRows: 1, numColumns: 1 };
  }
  const range = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(a1);
  if (range) {
    const toColumn = (letters) => letters.split('').reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
    const startColumn = toColumn(range[1]);
    const endColumn = toColumn(range[3]);
    return {
      row: Number(range[2]),
      column: startColumn,
      numRows: Number(range[4]) - Number(range[2]) + 1,
      numColumns: endColumn - startColumn + 1
    };
  }
  const wholeColumn = /^([A-Z]+):([A-Z]+)$/.exec(a1);
  if (wholeColumn) {
    const toColumn = (letters) => letters.split('').reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
    return { row: 1, column: toColumn(wholeColumn[1]), numRows: 1, numColumns: 1 };
  }
  throw new Error(`Unsupported A1 range: ${a1}`);
}

function createSheet(name, rows, audit, options = {}) {
  let maxColumns = options.maxColumns == null
    ? Math.max(26, ...rows.map((row) => row.length))
    : Number(options.maxColumns);
  const sheet = {
    name,
    rows: cloneRows(rows),
    formulas: cloneRows(options.formulas || rows.map((row) => row.map(() => ''))),
    getLastRow() {
      return this.rows.length;
    },
    getLastColumn() {
      return this.rows.reduce((max, row) => Math.max(max, row.length), 0);
    },
    getMaxColumns() {
      return maxColumns;
    },
    getDataRange() {
      audit.dataRangeReads.push(name);
      return createRange(this, 1, 1, Math.max(this.rows.length, 1), Math.max(this.getLastColumn(), 1), audit);
    },
    getRange(rowOrA1, column, numRows, numColumns) {
      const coordinates = typeof rowOrA1 === 'string'
        ? parseA1(rowOrA1)
        : { row: rowOrA1, column, numRows: numRows || 1, numColumns: numColumns || 1 };
      if (coordinates.column + coordinates.numColumns - 1 > maxColumns) {
        throw new Error(`Range exceeds physical columns on ${name}`);
      }
      audit.rangeReads.push({ sheet: name, ...coordinates });
      return createRange(this, coordinates.row, coordinates.column, coordinates.numRows, coordinates.numColumns, audit);
    },
    insertColumnsAfter(after, count) {
      if (after < 1 || after > maxColumns || count < 1) throw new Error('Invalid column insertion');
      audit.columnInserts.push({ sheet: name, after, count });
      maxColumns += count;
    },
    insertRowBefore(rowNumber) {
      audit.writes.push({ sheet: name, method: 'insertRowBefore' });
      this.rows.splice(rowNumber - 1, 0, []);
      this.formulas.splice(rowNumber - 1, 0, []);
    },
    appendRow(row) {
      audit.writes.push({ sheet: name, method: 'appendRow' });
      this.rows.push(row.slice());
      this.formulas.push(row.map(() => ''));
    },
    setFrozenRows() {
      audit.writes.push({ sheet: name, method: 'setFrozenRows' });
    },
    setColumnWidth() {
      audit.writes.push({ sheet: name, method: 'setColumnWidth' });
    }
  };
  return sheet;
}

function createRange(sheet, row, column, numRows, numColumns, audit) {
  function ensureCell(rowIndex, columnIndex) {
    while (sheet.rows.length <= rowIndex) sheet.rows.push([]);
    while (sheet.formulas.length <= rowIndex) sheet.formulas.push([]);
    while (sheet.rows[rowIndex].length <= columnIndex) sheet.rows[rowIndex].push('');
    while (sheet.formulas[rowIndex].length <= columnIndex) sheet.formulas[rowIndex].push('');
  }

  const range = {
    getValue() {
      return (sheet.rows[row - 1] || [])[column - 1] || '';
    },
    getValues() {
      const values = Array.from({ length: numRows }, (_, rowOffset) =>
        Array.from({ length: numColumns }, (_, columnOffset) =>
          (sheet.rows[row - 1 + rowOffset] || [])[column - 1 + columnOffset] || ''
        )
      );
      if (typeof audit.afterGetValues === 'function') {
        audit.afterGetValues({ sheet: sheet.name, row, column, numRows, numColumns }, sheet);
      }
      return values;
    },
    getFormulas() {
      return Array.from({ length: numRows }, (_, rowOffset) =>
        Array.from({ length: numColumns }, (_, columnOffset) =>
          (sheet.formulas[row - 1 + rowOffset] || [])[column - 1 + columnOffset] || ''
        )
      );
    },
    setValue(value) {
      audit.writes.push({ sheet: sheet.name, method: 'setValue', row, column, numRows, numColumns });
      ensureCell(row - 1, column - 1);
      sheet.rows[row - 1][column - 1] = value;
      sheet.formulas[row - 1][column - 1] = '';
      return range;
    },
    setValues(values) {
      audit.writes.push({ sheet: sheet.name, method: 'setValues', row, column, numRows, numColumns });
      values.forEach((valuesRow, rowOffset) => {
        valuesRow.forEach((value, columnOffset) => {
          ensureCell(row - 1 + rowOffset, column - 1 + columnOffset);
          sheet.rows[row - 1 + rowOffset][column - 1 + columnOffset] = value;
          sheet.formulas[row - 1 + rowOffset][column - 1 + columnOffset] = '';
        });
      });
      return range;
    },
    setBackground() {
      audit.writes.push({ sheet: sheet.name, method: 'setBackground' });
      return range;
    },
    setFontColor() {
      audit.writes.push({ sheet: sheet.name, method: 'setFontColor' });
      return range;
    },
    setFontWeight() {
      audit.writes.push({ sheet: sheet.name, method: 'setFontWeight' });
      return range;
    },
    setNumberFormat() {
      audit.writes.push({ sheet: sheet.name, method: 'setNumberFormat' });
      return range;
    }
  };
  return range;
}

function pinRow({ id, fileId = '', imageUrl = '', audioId = '' }) {
  return [
    '2026/07/10 10:00:00', id, '', 35, 139, '#e53935', fileId, imageUrl, id,
    '', '未対応', '', '', '', 'default', audioId
  ];
}

function createHarness(options = {}) {
  const audit = {
    driveFileIds: [],
    dataRangeReads: [],
    rangeReads: [],
    columnInserts: [],
    writes: [],
    logs: [],
    afterGetValues: options.afterGetValues
  };
  const sheets = {};
  const initialSheets = options.sheets || {
    map_info: [
      MAP_INFO_HEADERS,
      pinRow({ id: 'pin-first' }),
      pinRow({ id: 'pin-photo', fileId: 'registered-file', imageUrl: 'https://example.com/photo.jpg' }),
      pinRow({ id: 'pin-last', audioId: 'private-audio-id' })
    ]
  };
  Object.entries(initialSheets).forEach(([name, rows]) => {
    sheets[name] = createSheet(name, rows, audit, {
      maxColumns: options.sheetMaxColumns && options.sheetMaxColumns[name],
      formulas: options.sheetFormulas && options.sheetFormulas[name]
    });
  });

  const spreadsheet = {
    getSheetByName(name) {
      return sheets[name] || null;
    },
    insertSheet(name) {
      audit.writes.push({ sheet: name, method: 'insertSheet' });
      sheets[name] = createSheet(name, [], audit);
      return sheets[name];
    }
  };
  const driveFiles = options.driveFiles || {
    'registered-file': { parentId: 'registered-parent' }
  };
  const context = {
    console,
    Date,
    JSON,
    Logger: { log: (message) => audit.logs.push(String(message)) },
    Utilities: {
      getUuid: () => 'generated-uuid',
      formatDate: () => '2026/07/10 10:00:00'
    },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    CacheService: {
      getScriptCache: () => ({
        get: (key) => (key === `EDIT_TOKEN_${TEST_EDIT_TOKEN}` ? '1' : null),
        put() {}
      })
    },
    LockService: {
      getScriptLock: () => ({
        tryLock: () => true,
        releaseLock() {}
      })
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      getUi: () => ({ alert() {}, ButtonSet: { OK: 'OK' } })
    },
    DriveApp: {
      getFileById(fileId) {
        audit.driveFileIds.push(fileId);
        const file = driveFiles[fileId];
        if (!file || file.error) throw new Error(file && file.error ? file.error : 'not found');
        return {
          getParents() {
            let consumed = false;
            return {
              hasNext: () => !consumed && !!file.parentId,
              next() {
                consumed = true;
                return { getId: () => file.parentId };
              }
            };
          }
        };
      }
    }
  };

  vm.runInNewContext(`${codeJs}
globalThis.__mapReadApi = {
  PinData,
  getMapData,
  findPinRowIndex_,
  setupSheet,
  getRouteGroups,
  getRouteCache,
  listShareLinks,
  openRoutesSheet_,
  unplacePin,
  updatePinDetails,
  getPinDriveMeta: typeof getPinDriveMeta === 'function' ? getPinDriveMeta : null
};`, context);

  return { api: context.__mapReadApi, audit, sheets, spreadsheet };
}

function sameRealm(value) {
  return JSON.parse(JSON.stringify(value));
}

function sourceFunctionBody(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name} to exist`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(open + 1, index);
  }
  assert.fail(`Could not parse function ${name}`);
}

function withEditToken(payload = {}) {
  return { ...payload, __editToken: TEST_EDIT_TOKEN };
}

test('getMapData reads photo-less and photo pins without Drive metadata calls', () => {
  const { api, audit } = createHarness();

  const pins = sameRealm(api.getMapData());

  assert.equal(pins.length, 3);
  assert.equal(pins[0].fileId, '');
  assert.equal(pins[0].imageUrl, '');
  assert.equal(pins[0].folderUrl, '');
  assert.equal(pins[0].hasAudio, false);
  assert.equal(Object.hasOwn(pins[0], 'audioId'), false);
  assert.equal(pins[1].fileId, 'registered-file');
  assert.equal(pins[1].imageUrl, 'https://example.com/photo.jpg');
  assert.equal(pins[1].folderUrl, '');
  assert.equal(pins[2].hasAudio, true);
  assert.equal(Object.hasOwn(pins[2], 'audioId'), false);
  assert.deepEqual(audit.driveFileIds, []);
});

test('setupSheet appends only P1 and preserves legacy map_info cells', () => {
  const legacyHeaders = MAP_INFO_HEADERS.slice(0, 15);
  const legacyRow = pinRow({ id: 'legacy-pin' }).slice(0, 15);
  legacyRow[2] = 'computed-description';
  const legacyFormulas = [legacyHeaders.map(() => ''), legacyRow.map(() => '')];
  legacyFormulas[0][14] = '="アイコン"';
  legacyFormulas[1][2] = '=FORMULA_PRESERVED';
  const { api, audit, sheets } = createHarness({
    sheets: { map_info: [legacyHeaders, legacyRow] },
    sheetMaxColumns: { map_info: 15 },
    sheetFormulas: { map_info: legacyFormulas }
  });

  api.setupSheet();
  api.setupSheet();

  assert.deepEqual(sheets.map_info.rows[0], MAP_INFO_HEADERS);
  assert.deepEqual(sheets.map_info.rows[1], legacyRow);
  assert.equal(sheets.map_info.formulas[0][14], '="アイコン"');
  assert.equal(sheets.map_info.formulas[1][2], '=FORMULA_PRESERVED');
  assert.deepEqual(audit.columnInserts, [{ sheet: 'map_info', after: 15, count: 1 }]);
  assert.deepEqual(
    audit.writes.filter((write) => write.sheet === 'map_info' && write.method === 'setValue'),
    [{ sheet: 'map_info', method: 'setValue', row: 1, column: 16, numRows: 1, numColumns: 1 }]
  );

  const existingAudioHeaders = MAP_INFO_HEADERS.slice();
  existingAudioHeaders[15] = 'existing-audio-header';
  const rowWithExistingAudio = legacyRow.concat(['existing-audio-id']);
  const existing = createHarness({ sheets: { map_info: [existingAudioHeaders, rowWithExistingAudio] } });
  existing.api.setupSheet();
  assert.equal(existing.sheets.map_info.rows[0][15], 'existing-audio-header');
  assert.deepEqual(existing.sheets.map_info.rows[1], rowWithExistingAudio);
});

test('setupSheet preserves a P1 formula that evaluates to an empty string', () => {
  const headerValues = MAP_INFO_HEADERS.slice(0, 15).concat(['']);
  const headerFormulas = [headerValues.map(() => '')];
  headerFormulas[0][15] = '=""';
  const { api, audit, sheets } = createHarness({
    sheets: { map_info: [headerValues] },
    sheetMaxColumns: { map_info: 16 },
    sheetFormulas: { map_info: headerFormulas }
  });

  api.setupSheet();

  assert.equal(sheets.map_info.formulas[0][15], '=""');
  assert.equal(sheets.map_info.rows[0][15], '');
  assert.deepEqual(
    audit.writes.filter((write) => write.sheet === 'map_info'
      && (write.method === 'setValue' || write.method === 'setValues')
      && write.row === 1
      && write.column <= 16
      && write.column + write.numColumns - 1 >= 16),
    []
  );
});

test('findPinRowIndex_ reads only column I and preserves header-inclusive indexes', () => {
  const { api, audit, sheets } = createHarness();
  const sheet = sheets.map_info;

  assert.equal(api.findPinRowIndex_(sheet, 'pin-first'), 1);
  assert.equal(api.findPinRowIndex_(sheet, 'pin-photo'), 2);
  assert.equal(api.findPinRowIndex_(sheet, 'pin-last'), 3);
  assert.equal(api.findPinRowIndex_(sheet, 'missing'), -1);
  assert.equal(api.findPinRowIndex_(sheet, 'ID'), 0);
  assert.equal(audit.dataRangeReads.length, 0);
  assert.deepEqual(
    audit.rangeReads.map(({ sheet: sheetName, row, column, numRows, numColumns }) => ({ sheet: sheetName, row, column, numRows, numColumns })),
    Array.from({ length: 5 }, () => ({ sheet: 'map_info', row: 1, column: 9, numRows: 4, numColumns: 1 }))
  );
});

test('pin mutations reject the header index while preserving the row-index contract', () => {
  const { api, sheets } = createHarness();
  const originalHeader = sheets.map_info.rows[0].slice();

  const result = sameRealm(api.unplacePin(withEditToken({ id: 'ID' })));

  assert.deepEqual(result, { ok: false, error: 'id not found' });
  assert.deepEqual(sheets.map_info.rows[0], originalHeader);
  ['updatePinDetails', 'movePin', 'unplacePin'].forEach((name) => {
    assert.match(sourceFunctionBody(codeJs, name), /rowIndex < 1/);
  });
  const deletePinBody = sourceFunctionBody(codeJs, 'deletePin');
  assert.match(deletePinBody, /buildPinDeleteSnapshots_\(rows, \[data\.id\]\)/);
  assert.match(deletePinBody, /indexCurrentPinDeleteRows_\(currentRows\)/);
  assert.match(deletePinBody, /current\.fileId !== snapshot\.fileId/);
  assert.match(deletePinBody, /sheet\.deleteRow\(current\.rowNumber\)/);
  const bulkStatusBody = sourceFunctionBody(codeJs, 'bulkUpdatePinStatus');
  assert.match(bulkStatusBody, /new Set\(data\.ids\)/);
  assert.match(bulkStatusBody, /for \(var rowIndex = 1;/);
  const bulkDeleteBody = sourceFunctionBody(codeJs, 'bulkDeletePins');
  assert.match(bulkDeleteBody, /buildPinDeleteSnapshots_\(rows, data\.ids\)/);
  assert.match(bulkDeleteBody, /indexCurrentPinDeleteRows_\(currentRows\)/);
  assert.match(bulkDeleteBody, /current\.fileId !== snapshot\.fileId/);
  assert.match(bulkDeleteBody, /failedIdSet\[snapshot\.pinId\] = true/);
  assert.match(bulkDeleteBody, /groupContiguousPinDeleteRows_\(currentEntries\)/);
});

test('getPinDriveMeta requires edit access and a registered pinId', () => {
  const { api, audit } = createHarness();
  assert.equal(typeof api.getPinDriveMeta, 'function');

  assert.throws(() => api.getPinDriveMeta({ pinId: 'pin-photo' }), /編集権限/);
  assert.throws(() => api.getPinDriveMeta({ pinId: 'pin-photo', __editToken: 'invalid' }), /編集権限/);
  assert.deepEqual(audit.driveFileIds, []);
  assert.deepEqual(sameRealm(api.getPinDriveMeta(withEditToken({}))), {
    ok: false,
    folderUrl: '',
    error: 'missing_pin_id'
  });
  assert.deepEqual(sameRealm(api.getPinDriveMeta(withEditToken({ pinId: 'missing' }))), {
    ok: false,
    folderUrl: '',
    error: 'pin_not_found'
  });
  assert.deepEqual(audit.driveFileIds, []);
});

test('getPinDriveMeta ignores client fileId and resolves only the registered file', () => {
  const { api, audit } = createHarness({
    driveFiles: {
      'registered-file': { parentId: 'registered-parent' },
      'arbitrary-file': { parentId: 'secret-parent' }
    }
  });
  assert.equal(typeof api.getPinDriveMeta, 'function');

  const noPhoto = api.getPinDriveMeta(withEditToken({ pinId: 'pin-first', fileId: 'arbitrary-file' }));
  assert.deepEqual(sameRealm(noPhoto), { ok: true, folderUrl: '' });
  assert.deepEqual(audit.driveFileIds, []);

  const photo = api.getPinDriveMeta(withEditToken({ pinId: 'pin-photo', fileId: 'arbitrary-file' }));
  assert.deepEqual(sameRealm(photo), {
    ok: true,
    folderUrl: 'https://drive.google.com/drive/folders/registered-parent'
  });
  assert.deepEqual(audit.driveFileIds, ['registered-file']);
});

test('getPinDriveMeta resolves pinId and fileId from one header-excluded G:I snapshot', () => {
  let rowsChanged = false;
  const { api, audit } = createHarness({
    sheets: {
      map_info: [
        MAP_INFO_HEADERS,
        pinRow({ id: 'pin-first' }),
        pinRow({ id: 'pin-photo', fileId: 'registered-file' }),
        pinRow({ id: 'pin-last', fileId: 'different-file' })
      ]
    },
    driveFiles: {
      'registered-file': { parentId: 'registered-parent' },
      'different-file': { parentId: 'different-parent' }
    },
    afterGetValues(range, sheet) {
      if (rowsChanged || range.sheet !== 'map_info') return;
      rowsChanged = true;
      sheet.rows.splice(1, 1);
    }
  });

  const result = sameRealm(api.getPinDriveMeta(withEditToken({
    pinId: 'pin-photo',
    fileId: 'different-file'
  })));

  assert.deepEqual(result, {
    ok: true,
    folderUrl: 'https://drive.google.com/drive/folders/registered-parent'
  });
  assert.deepEqual(audit.driveFileIds, ['registered-file']);
  assert.deepEqual(
    audit.rangeReads,
    [{ sheet: 'map_info', row: 2, column: 7, numRows: 3, numColumns: 3 }]
  );
});

test('getPinDriveMeta returns a generic retryable result when Drive lookup fails', () => {
  const { api, audit } = createHarness({
    driveFiles: { 'registered-file': { error: 'permission denied for owner@example.com' } }
  });
  assert.equal(typeof api.getPinDriveMeta, 'function');

  const result = sameRealm(api.getPinDriveMeta(withEditToken({ pinId: 'pin-photo' })));

  assert.deepEqual(result, { ok: false, folderUrl: '', error: 'drive_meta_unavailable' });
  assert.equal(JSON.stringify(result).includes('owner@example.com'), false);
  assert.equal(audit.logs.some((message) => message.includes('getPinDriveMeta')), true);
});

test('updatePinDetails does not resolve Drive metadata for its response', () => {
  const { api, audit, sheets } = createHarness();
  sheets.route_pins = createSheet('route_pins', [['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt']], audit);

  const result = sameRealm(api.updatePinDetails(withEditToken({
    id: 'pin-photo',
    title: 'Updated photo pin',
    description: '',
    color: '#e53935',
    icon: 'default',
    links: [],
    status: '未対応',
    tags: [],
    eventAt: ''
  })));

  assert.equal(result.ok, true);
  assert.equal(result.folderUrl, '');
  assert.deepEqual(audit.driveFileIds, []);
});

test('normal reads after setupSheet do not inspect or rewrite sheet headers', () => {
  const { api, audit, sheets } = createHarness({ sheets: {} });

  api.setupSheet();
  sheets.routes.appendRow(['route-1', 'Route', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']);
  sheets.route_pins.appendRow(['route-1', 'pin-1', 0, '', '']);
  audit.writes.length = 0;
  audit.rangeReads.length = 0;
  audit.dataRangeReads.length = 0;

  api.getMapData();
  api.getRouteGroups();
  api.getRouteCache({ cacheKey: 'missing' });
  api.listShareLinks(withEditToken());

  assert.deepEqual(audit.writes, []);
  assert.equal(audit.rangeReads.some((entry) => entry.row === 1 && entry.column === 1), false);
});

test('normal sheet openers fail clearly when setupSheet has not created a required sheet', () => {
  const { api } = createHarness();

  assert.throws(() => api.openRoutesSheet_(), /routes.*setupSheet\(\)/);
});

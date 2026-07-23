const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const codeJs = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
const HEADERS = [
  'presetId', 'name', 'enabled', 'orderIndex', 'tagsMode', 'tags', 'colorMode',
  'color', 'iconMode', 'icon', 'statusMode', 'status', 'createdAt', 'updatedAt'
];

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

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
  throw new Error('Unsupported A1 range: ' + value);
}

function createHarness(options = {}) {
  const audit = {
    sheetLookups: [], inserts: [], reads: [], writes: [], alerts: [],
    lock: { attempts: 0, releases: 0, held: false }
  };
  const sheets = new Map();
  let uuidSequence = 0;

  function lastUsedRow(rows) {
    for (let index = rows.length - 1; index >= 0; index -= 1) {
      if ((rows[index] || []).some((value) => value !== '' && value != null)) return index + 1;
    }
    return 0;
  }

  function makeSheet(name, initialRows = [], initialFormulas = {}) {
    const rows = clone(initialRows);
    const formulas = Object.assign({}, initialFormulas);
    let maxRows = Math.max(1000, rows.length);
    let maxColumns = Math.max(26, rows.reduce(
      (maximum, row) => Math.max(maximum, (row || []).length),
      0
    ));
    const sheet = {
      name,
      rows,
      formulas,
      getLastRow() { return lastUsedRow(rows); },
      getLastColumn() {
        return rows.reduce((maximum, row) => Math.max(maximum, (row || []).length), 0);
      },
      getMaxRows() { return maxRows; },
      insertRowsAfter(_after, count) { maxRows += count; },
      getMaxColumns() { return maxColumns; },
      insertColumnsAfter(_after, count) { maxColumns += count; },
      insertRowBefore(rowNumber) {
        rows.splice(rowNumber - 1, 0, []);
      },
      appendRow(row) {
        rows.push(clone(row));
        audit.writes.push({ sheet: name, method: 'appendRow', values: clone(row), lockHeld: audit.lock.held });
      },
      getDataRange() {
        const rowCount = Math.max(1, sheet.getLastRow());
        const columnCount = Math.max(1, sheet.getLastColumn());
        return sheet.getRange(1, 1, rowCount, columnCount);
      },
      getRange(rowOrA1, column, numRows = 1, numColumns = 1) {
        const info = typeof rowOrA1 === 'string'
          ? parseA1(rowOrA1)
          : { row: rowOrA1, column, numRows, numColumns };
        const range = {
          getValues() {
            audit.reads.push({ sheet: name, method: 'getValues', ...info });
            return Array.from({ length: info.numRows }, (_unused, rowOffset) =>
              Array.from({ length: info.numColumns }, (_unusedColumn, columnOffset) => {
                const sourceRow = rows[info.row - 1 + rowOffset] || [];
                return sourceRow[info.column - 1 + columnOffset] ?? '';
              })
            );
          },
          getFormulas() {
            audit.reads.push({ sheet: name, method: 'getFormulas', ...info });
            return Array.from({ length: info.numRows }, (_unused, rowOffset) =>
              Array.from({ length: info.numColumns }, (_unusedColumn, columnOffset) =>
                formulas[(info.row + rowOffset) + ':' + (info.column + columnOffset)] || ''
              )
            );
          },
          getValue() { return range.getValues()[0][0]; },
          setValues(values) {
            assert.equal(values.length, info.numRows);
            values.forEach((valuesRow) => assert.equal(valuesRow.length, info.numColumns));
            audit.writes.push({ sheet: name, method: 'setValues', ...info, values: clone(values), lockHeld: audit.lock.held });
            values.forEach((valuesRow, rowOffset) => {
              const targetIndex = info.row - 1 + rowOffset;
              while (rows.length <= targetIndex) rows.push([]);
              valuesRow.forEach((value, columnOffset) => {
                const targetColumn = info.column - 1 + columnOffset;
                while (rows[targetIndex].length <= targetColumn) rows[targetIndex].push('');
                rows[targetIndex][targetColumn] = value;
                const key = (info.row + rowOffset) + ':' + (info.column + columnOffset);
                formulas[key] = typeof value === 'string' && value.startsWith('=') ? value : '';
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

  Object.entries(options.sheets || {}).forEach(([name, sheetOptions]) => {
    if (Array.isArray(sheetOptions)) makeSheet(name, sheetOptions);
    else makeSheet(name, sheetOptions.rows || [], sheetOptions.formulas || {});
  });

  const spreadsheet = {
    getSheetByName(name) {
      audit.sheetLookups.push(name);
      return sheets.get(name) || null;
    },
    insertSheet(name) {
      audit.inserts.push(name);
      return makeSheet(name);
    }
  };
  const lock = {
    tryLock() {
      audit.lock.attempts += 1;
      const acquired = options.lockAcquired !== false;
      audit.lock.held = acquired;
      return acquired;
    },
    releaseLock() {
      audit.lock.releases += 1;
      audit.lock.held = false;
    }
  };
  const context = {
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      getUi: () => ({
        ButtonSet: { OK: 'OK' },
        createMenu: () => ({ addItem() { return this; }, addToUi() {} }),
        alert(...args) { audit.alerts.push(args); }
      })
    },
    CacheService: { getScriptCache: () => ({ get: () => options.validToken === false ? null : '1', put() {} }) },
    LockService: { getScriptLock: () => lock },
    Utilities: { getUuid: () => `preset-uuid-${++uuidSequence}` },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/example/exec' }) },
    HtmlService: {},
    DriveApp: {},
    Maps: {},
    console
  };
  vm.runInNewContext(`${codeJs}\nthis.__api = {\n    setupSheet, listInputPresets, saveInputPreset, deleteInputPreset, updateInputPresetOrder\n  };`, context);
  return { api: context.__api, audit, sheets, makeSheet };
}

function token(payload = {}) {
  return Object.assign({ __editToken: 'valid-edit-token' }, payload);
}

function presetRow(overrides = {}) {
  const preset = Object.assign({
    presetId: 'preset-1', name: '基本', enabled: true, orderIndex: 0,
    tagsMode: 'set', tags: '植物|観察', colorMode: 'keep', color: '',
    iconMode: 'keep', icon: '', statusMode: 'keep', status: '',
    createdAt: '2026-01-01T00:00:00.000Z', updatedAt: '2026-01-01T00:00:00.000Z'
  }, overrides);
  return HEADERS.map((header) => preset[header]);
}

function validPreset(overrides = {}) {
  return Object.assign({
    name: ' 1班・植物 ', enabled: true,
    tagsMode: 'set', tags: ['#植物', '観察', '植物'],
    colorMode: 'set', color: '#E53935',
    iconMode: 'set', icon: 'nature',
    statusMode: 'clear', status: '完了'
  }, overrides);
}

test('setupSheet creates input_presets with the fixed header', () => {
  const harness = createHarness({ sheets: {
    map_info: [['タイムスタンプ', '', '', '', '', '', '', '', 'ID']],
    config: [['設定項目', '値', '説明'], ['EDIT_KEY', 'key', ''], ['WEB_APP_URL', '', ''], ['EDIT_URL', '', '']]
  } });

  harness.api.setupSheet();

  assert.equal(harness.audit.inserts.includes('input_presets'), true);
  assert.deepEqual(harness.sheets.get('input_presets').rows[0].slice(0, HEADERS.length), HEADERS);
  assert.match(harness.audit.alerts.at(-1)[1], /input_presets/);
});

test('setupSheet preserves existing input_presets rows', () => {
  const harness = createHarness({ sheets: {
    map_info: [['タイムスタンプ', '', '', '', '', '', '', '', 'ID']],
    config: [['設定項目', '値', '説明'], ['EDIT_KEY', 'key', ''], ['WEB_APP_URL', '', ''], ['EDIT_URL', '', '']],
    input_presets: [HEADERS, presetRow()]
  } });

  harness.api.setupSheet();

  assert.equal(harness.audit.inserts.includes('input_presets'), false);
  assert.deepEqual(harness.sheets.get('input_presets').rows[0].slice(0, HEADERS.length), HEADERS);
  assert.deepEqual(harness.sheets.get('input_presets').rows[1].slice(0, HEADERS.length), presetRow());
  assert.match(harness.audit.alerts.at(-1)[1], /input_presets/);
});

test('normal list reads never create a missing input_presets sheet', () => {
  const harness = createHarness();
  assert.throws(() => harness.api.listInputPresets(token()), /input_presets.*setupSheet\(\)/);
  assert.deepEqual(harness.audit.inserts, []);
});

test('all input preset APIs require an edit token before Spreadsheet access', () => {
  const harness = createHarness({ validToken: false });
  const calls = [
    () => harness.api.listInputPresets({}),
    () => harness.api.saveInputPreset(validPreset()),
    () => harness.api.deleteInputPreset({ presetId: 'preset-1' }),
    () => harness.api.updateInputPresetOrder({ presetIds: [] })
  ];
  calls.forEach((call) => assert.throws(call, /編集権限/));
  assert.deepEqual(harness.audit.sheetLookups, []);
});

test('saveInputPreset creates a normalized preset at the tail in one fixed-width write', () => {
  const harness = createHarness({ sheets: { input_presets: [HEADERS, presetRow()] } });
  const result = harness.api.saveInputPreset(token(validPreset()));

  assert.equal(result.ok, true);
  assert.equal(result.preset.presetId, 'preset-uuid-1');
  assert.equal(result.preset.name, '1班・植物');
  assert.equal(result.preset.orderIndex, 1);
  assert.deepEqual(Array.from(result.preset.tags), ['植物', '観察']);
  assert.equal(result.preset.color, '#e53935');
  assert.equal(result.preset.status, null);
  assert.ok(result.preset.createdAt);
  assert.equal(result.preset.createdAt, result.preset.updatedAt);
  const writes = harness.audit.writes.filter((write) => write.sheet === 'input_presets');
  assert.equal(writes.length, 1);
  assert.deepEqual([writes[0].method, writes[0].row, writes[0].column, writes[0].numRows, writes[0].numColumns],
    ['setValues', 3, 1, 1, HEADERS.length]);
  assert.equal(writes[0].lockHeld, true);
});

test('saveInputPreset appends after the greatest saved order when orderIndex is omitted', () => {
  const harness = createHarness({ sheets: { input_presets: [
    HEADERS,
    presetRow({ presetId: 'preset-10', orderIndex: 10 }),
    presetRow({ presetId: 'preset-20', orderIndex: 20 })
  ] } });
  const result = harness.api.saveInputPreset(token(validPreset()));
  assert.equal(result.ok, true);
  assert.equal(result.preset.orderIndex, 21);
});

test('saveInputPreset updates an existing row while preserving createdAt and orderIndex', () => {
  const harness = createHarness({ sheets: { input_presets: [HEADERS, presetRow()] } });
  const result = harness.api.saveInputPreset(token(validPreset({ presetId: 'preset-1', orderIndex: 99 })));

  assert.equal(result.ok, true);
  assert.equal(result.preset.presetId, 'preset-1');
  assert.equal(result.preset.createdAt, '2026-01-01T00:00:00.000Z');
  assert.equal(result.preset.orderIndex, 0);
  assert.notEqual(result.preset.updatedAt, '2026-01-01T00:00:00.000Z');
  const write = harness.audit.writes.find((entry) => entry.sheet === 'input_presets');
  assert.deepEqual([write.row, write.column, write.numRows, write.numColumns], [2, 1, 1, HEADERS.length]);
});

test('saveInputPreset rejects an unknown update ID without writes', () => {
  const harness = createHarness({ sheets: { input_presets: [HEADERS] } });
  const result = harness.api.saveInputPreset(token(validPreset({ presetId: 'missing' })));
  assert.deepEqual(clone(result), { ok: false, error: 'プリセットが見つかりません。' });
  assert.deepEqual(harness.audit.writes, []);
});

test('saveInputPreset rejects invalid names, modes, values, and no-op presets without writes', () => {
  const invalidPayloads = [
    validPreset({ name: '  ' }),
    validPreset({ name: 'あ'.repeat(61) }),
    validPreset({ tagsMode: 'append' }),
    validPreset({ colorMode: 'clear' }),
    validPreset({ iconMode: 'clear' }),
    validPreset({ statusMode: 'append' }),
    validPreset({ tags: ['1', '2', '3', '4', '5', '6'] }),
    validPreset({ color: 'red' }),
    validPreset({ icon: 'unknown' }),
    validPreset({ statusMode: 'set', status: '不明' }),
    validPreset({ tagsMode: 'keep', colorMode: 'keep', iconMode: 'keep', statusMode: 'keep' })
  ];
  invalidPayloads.forEach((payload) => {
    const harness = createHarness({ sheets: { input_presets: [HEADERS] } });
    const result = harness.api.saveInputPreset(token(payload));
    assert.equal(result.ok, false, JSON.stringify(payload));
    assert.equal(harness.audit.writes.length, 0, JSON.stringify(payload));
  });
});

test('keep and clear modes discard unused values in the saved model', () => {
  const harness = createHarness({ sheets: { input_presets: [HEADERS] } });
  const result = harness.api.saveInputPreset(token(validPreset({
    tagsMode: 'clear', tags: ['discard'], colorMode: 'keep', color: '#ffffff',
    iconMode: 'keep', icon: 'photo', statusMode: 'clear', status: '完了'
  })));
  assert.equal(result.ok, true);
  assert.deepEqual(Array.from(result.preset.tags), []);
  assert.equal(result.preset.color, null);
  assert.equal(result.preset.icon, null);
  assert.equal(result.preset.status, null);
});

test('saveInputPreset enforces the 100 preset limit only for creates', () => {
  const rows = [HEADERS];
  for (let index = 0; index < 100; index += 1) {
    rows.push(presetRow({ presetId: `preset-${index}`, name: `Preset ${index}`, orderIndex: index }));
  }
  const harness = createHarness({ sheets: { input_presets: rows } });
  const result = harness.api.saveInputPreset(token(validPreset()));
  assert.equal(result.ok, false);
  assert.match(result.error, /100/);
  assert.deepEqual(harness.audit.writes, []);

  const updateHarness = createHarness({ sheets: { input_presets: rows } });
  const updated = updateHarness.api.saveInputPreset(token(validPreset({ presetId: 'preset-0' })));
  assert.equal(updated.ok, true);
  assert.equal(updated.preset.presetId, 'preset-0');
});

test('listInputPresets returns all normalized rows ordered by index, name, and ID', () => {
  const harness = createHarness({ sheets: { input_presets: [
    HEADERS,
    presetRow({ presetId: 'b', name: '同名', enabled: false, orderIndex: 1, tagsMode: 'keep', tags: 'discard', statusMode: 'clear' }),
    presetRow({ presetId: 'c', name: '同名', orderIndex: 1 }),
    presetRow({ presetId: 'a', name: '先頭', orderIndex: 0, colorMode: 'keep', color: '#ffffff' })
  ] } });
  const result = harness.api.listInputPresets(token());
  assert.equal(result.ok, true);
  assert.deepEqual(Array.from(result.presets, (preset) => preset.presetId), ['a', 'b', 'c']);
  assert.equal(result.presets[1].enabled, false);
  assert.deepEqual(Array.from(result.presets[1].tags), []);
  assert.equal(result.presets[2].color, null);
  const presetReads = harness.audit.reads.filter((read) => read.sheet === 'input_presets' && read.method === 'getValues');
  assert.equal(presetReads.length, 1);
  assert.equal(presetReads[0].numColumns, HEADERS.length);
  assert.deepEqual(harness.audit.writes, []);
});

test('deleteInputPreset rejects missing IDs and compacts order in one formula-preserving rewrite', () => {
  const rows = [
    HEADERS.concat(['custom']),
    presetRow({ presetId: 'a', orderIndex: 2 }).concat(['=A2']),
    presetRow({ presetId: 'b', orderIndex: 0 }).concat(['=A3']),
    presetRow({ presetId: 'c', orderIndex: 1 }).concat(['=A4'])
  ];
  const missingHarness = createHarness({ sheets: { input_presets: rows } });
  const missing = missingHarness.api.deleteInputPreset(token({ presetId: 'missing' }));
  assert.equal(missing.ok, false);
  assert.match(missing.error, /見つかりません/);
  assert.deepEqual(missingHarness.audit.writes, []);

  const harness = createHarness({ sheets: { input_presets: rows } });
  const result = harness.api.deleteInputPreset(token({ presetId: 'c' }));
  assert.equal(result.ok, true);
  assert.equal(result.preset.presetId, 'c');
  assert.deepEqual(Array.from(result.presets, (preset) => [preset.presetId, preset.orderIndex]), [['b', 0], ['a', 1]]);
  const writes = harness.audit.writes.filter((write) => write.sheet === 'input_presets');
  assert.equal(writes.length, 1);
  assert.equal(writes[0].numColumns, HEADERS.length + 1);
  assert.deepEqual(writes[0].values.slice(0, 2).map((row) => row.at(-1)), ['=A3', '=A2']);
});

test('updateInputPresetOrder validates a complete unique ID set before its single column write', () => {
  const rows = [HEADERS, presetRow({ presetId: 'a', orderIndex: 0 }), presetRow({ presetId: 'b', orderIndex: 1 })];
  for (const presetIds of [null, ['a'], ['a', 'a'], ['a', 'b', 'extra'], ['a', 'missing']]) {
    const harness = createHarness({ sheets: { input_presets: rows } });
    const result = harness.api.updateInputPresetOrder(token({ presetIds }));
    assert.equal(result.ok, false, JSON.stringify(presetIds));
    assert.deepEqual(harness.audit.writes, [], JSON.stringify(presetIds));
  }

  const harness = createHarness({ sheets: { input_presets: rows } });
  const result = harness.api.updateInputPresetOrder(token({ presetIds: ['b', 'a'] }));
  assert.equal(result.ok, true);
  assert.deepEqual(Array.from(result.presets, (preset) => [preset.presetId, preset.orderIndex]), [['b', 0], ['a', 1]]);
  const writes = harness.audit.writes.filter((write) => write.sheet === 'input_presets');
  assert.equal(writes.length, 1);
  assert.deepEqual([writes[0].column, writes[0].numRows, writes[0].numColumns], [4, 2, 1]);
  assert.deepEqual(writes[0].values, [[1], [0]]);
});

test('mutations return the shared busy error without Spreadsheet access when the lock fails', () => {
  const harness = createHarness({ lockAcquired: false, sheets: { input_presets: [HEADERS] } });
  const calls = [
    () => harness.api.saveInputPreset(token(validPreset())),
    () => harness.api.deleteInputPreset(token({ presetId: 'preset-1' })),
    () => harness.api.updateInputPresetOrder(token({ presetIds: [] }))
  ];
  calls.forEach((call) => {
    const result = call();
    assert.equal(result.ok, false);
    assert.match(result.error, /別の更新処理/);
  });
  assert.deepEqual(harness.audit.sheetLookups, []);
  assert.equal(harness.audit.lock.releases, 0);
});

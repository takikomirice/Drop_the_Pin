const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const codeJs = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
const TEST_EDIT_TOKEN = 'test-edit-token';

function loadApi() {
  const rows = [];
  const sheet = {
    appendRow(row) { rows.push(row.slice()); }
  };
  const context = {
    Logger: { log() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      getUuid: () => `pin-${rows.length + 1}`,
      formatDate: () => '2026/07/11 12:34:56'
    },
    CacheService: {
      getScriptCache: () => ({
        get: (key) => key === `EDIT_TOKEN_${TEST_EDIT_TOKEN}` ? '1' : null
      })
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: (name) => name === 'map_info' ? sheet : null
      })
    }
  };
  vm.runInNewContext(`${codeJs}\nglobalThis.__api = { saveMapData };`, context);
  return { saveMapData: context.__api.saveMapData, rows };
}

function payload(overrides = {}) {
  return {
    __editToken: TEST_EDIT_TOKEN,
    title: 'Status test',
    color: '#e53935',
    icon: 'default',
    ...overrides
  };
}

test('saveMapData preserves explicit blank status and defaults only an absent property', () => {
  const harness = loadApi();
  assert.equal(harness.saveMapData(payload({ status: '' })).ok, true);
  assert.equal(harness.saveMapData(payload()).ok, true);

  assert.equal(harness.rows[0][10], '');
  assert.equal(harness.rows[1][10], '未対応');
});

test('saveMapData continues to validate nonblank registration statuses', () => {
  const harness = loadApi();
  assert.throws(() => harness.saveMapData(payload({ status: '未知' })), /invalid status/);
  assert.equal(harness.rows.length, 0);
});

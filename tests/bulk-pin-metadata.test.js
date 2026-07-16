const assert = require('node:assert/strict');
const test = require('node:test');

const api = require('../Code.js');

const EDIT_TOKEN = 'bulk-metadata-token';
const HEADERS = [
  'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧', '状態', 'タグ',
  'イベント時刻', '更新時刻', 'アイコン'
];

function pinRow(id, overrides = {}) {
  return [
    'created', overrides.title || id, '', 35, 139, '#e53935', '', '', id, '',
    Object.prototype.hasOwnProperty.call(overrides, 'status') ? overrides.status : '未対応',
    Object.prototype.hasOwnProperty.call(overrides, 'tags') ? overrides.tags : 'alpha|beta',
    '', overrides.updatedAt || `updated-${id}`,
    Object.prototype.hasOwnProperty.call(overrides, 'icon') ? overrides.icon : 'default'
  ];
}

function createHarness(dataRows) {
  const rows = [HEADERS, ...dataRows.map((row) => row.slice())];
  const audit = { reads: [], writes: [], lockCalls: 0, releases: 0 };
  const sheet = {
    getLastRow() { return rows.length; },
    getRange(row, column, numRows = 1, numColumns = 1) {
      return {
        getValues() {
          audit.reads.push({ row, column, numRows, numColumns });
          return Array.from({ length: numRows }, (_, rowOffset) =>
            Array.from({ length: numColumns }, (_, columnOffset) =>
              rows[row - 1 + rowOffset][column - 1 + columnOffset]
            )
          );
        },
        setValues(values) {
          audit.writes.push({ method: 'setValues', row, column, numRows, numColumns, values });
          values.forEach((valuesRow, rowOffset) => {
            valuesRow.forEach((value, columnOffset) => {
              rows[row - 1 + rowOffset][column - 1 + columnOffset] = value;
            });
          });
          return this;
        },
        setValue(value) {
          audit.writes.push({ method: 'setValue', row, column, numRows, numColumns, value });
          rows[row - 1][column - 1] = value;
          return this;
        }
      };
    }
  };
  let held = false;
  global.CacheService = {
    getScriptCache: () => ({ get: (key) => key === `EDIT_TOKEN_${EDIT_TOKEN}` ? '1' : null })
  };
  global.LockService = {
    getScriptLock: () => ({
      tryLock() { audit.lockCalls += 1; held = true; return true; },
      releaseLock() { assert.equal(held, true); held = false; audit.releases += 1; }
    })
  };
  global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({ getSheetByName: (name) => name === 'map_info' ? sheet : null })
  };
  return { rows, audit };
}

function payload(data) {
  return { ...data, __editToken: EDIT_TOKEN };
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

test('bulkUpdatePinMetadata requires an edit token before Spreadsheet access', () => {
  const harness = createHarness([pinRow('pin-a')]);
  assert.throws(
    () => api.bulkUpdatePinMetadata({ ids: ['pin-a'], tagMode: 'replace', tags: [] }),
    /編集権限/
  );
  assert.equal(harness.audit.lockCalls, 0);
  assert.equal(harness.audit.reads.length, 0);
});

test('bulkUpdatePinMetadata deduplicates string IDs and adds tags in existing order', () => {
  const harness = createHarness([
    pinRow('1', { tags: 'alpha|Beta' }),
    pinRow('2', { tags: 'other' })
  ]);
  const result = plain(api.bulkUpdatePinMetadata(payload({
    ids: [1, '1', 2], tagMode: 'add', tags: ['beta', '#Gamma', 'gamma']
  })));

  assert.equal(result.ok, true);
  assert.equal(result.updatedCount, 2);
  assert.equal(result.unchangedCount, 0);
  assert.deepEqual(result.updates.map((update) => update.id), ['1', '2']);
  assert.deepEqual(result.updates.map((update) => update.tags), [
    ['alpha', 'Beta', 'Gamma'], ['other', 'beta', 'Gamma']
  ]);
  assert.equal(harness.rows[1][11], 'alpha|Beta|Gamma');
  assert.equal(harness.rows[2][11], 'other|beta|Gamma');
  assert.deepEqual(harness.audit.reads, [{ row: 1, column: 1, numRows: 3, numColumns: 15 }]);
});

test('bulkUpdatePinMetadata removes only requested tags and permits empty replace', () => {
  let harness = createHarness([
    pinRow('pin-a', { tags: 'alpha|Beta|gamma' }),
    pinRow('pin-b', { tags: 'beta|delta' })
  ]);
  let result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a', 'pin-b'], tagMode: 'remove', tags: ['#BETA', 'missing']
  })));
  assert.deepEqual(result.updates.map((update) => update.tags), [['alpha', 'gamma'], ['delta']]);
  assert.equal(harness.rows[1][11], 'alpha|gamma');
  assert.equal(harness.rows[2][11], 'delta');

  harness = createHarness([pinRow('pin-a', { tags: 'alpha|beta' })]);
  result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a'], tagMode: 'replace', tags: []
  })));
  assert.equal(result.ok, true);
  assert.deepEqual(result.updates[0].tags, []);
  assert.equal(harness.rows[1][11], '');
});

test('bulkUpdatePinMetadata rejects empty add/remove and tag-limit overflow without writes', () => {
  for (const tagMode of ['add', 'remove']) {
    const harness = createHarness([pinRow('pin-a')]);
    const result = plain(api.bulkUpdatePinMetadata(payload({ ids: ['pin-a'], tagMode, tags: [] })));
    assert.equal(result.ok, false);
    assert.equal(harness.audit.writes.length, 0);
  }

  let harness = createHarness([pinRow('pin-a')]);
  let result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a'], tagMode: 'replace', tags: ['a', 'b', 'c', 'd', 'e', 'f']
  })));
  assert.equal(result.ok, false);
  assert.equal(harness.audit.writes.length, 0);

  harness = createHarness([
    pinRow('pin-a', { tags: 'a|b|c|d|e' }),
    pinRow('pin-b', { tags: 'a' })
  ]);
  result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a', 'pin-b'], tagMode: 'add', tags: ['f']
  })));
  assert.equal(result.ok, false);
  assert.equal(harness.rows[2][11], 'a');
  assert.equal(harness.audit.writes.length, 0);

  harness = createHarness([pinRow('pin-a', { tags: 'a|b|c|d|e|f', icon: 'default' })]);
  result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a'], tagMode: 'none', icon: 'photo'
  })));
  assert.equal(result.ok, false);
  assert.equal(harness.rows[1][14], 'default');
  assert.equal(harness.audit.writes.length, 0);
});

test('bulkUpdatePinMetadata rejects invalid icons and missing IDs atomically', () => {
  let harness = createHarness([pinRow('pin-a', { icon: 'food' })]);
  let result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a'], tagMode: 'none', tags: [], icon: 'free-text'
  })));
  assert.equal(result.ok, false);
  assert.equal(harness.rows[1][14], 'food');
  assert.equal(harness.audit.writes.length, 0);

  harness = createHarness([pinRow('pin-a', { tags: 'alpha' })]);
  result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a', 'missing'], tagMode: 'replace', tags: ['next'], icon: 'photo'
  })));
  assert.equal(result.ok, false);
  assert.deepEqual(result.missingIds, ['missing']);
  assert.equal(harness.rows[1][11], 'alpha');
  assert.equal(harness.rows[1][14], 'default');
  assert.equal(harness.audit.writes.length, 0);
});

test('bulkUpdatePinMetadata changes timestamps only for changed rows and returns saved values', () => {
  const harness = createHarness([
    pinRow('pin-a', { tags: 'alpha', icon: 'photo', status: '完了', updatedAt: 'old-a' }),
    pinRow('pin-b', { tags: 'beta', icon: 'food', status: '未対応', updatedAt: 'old-b' }),
    pinRow('pin-c', { tags: 'gamma', icon: 'photo', status: '完了', updatedAt: 'old-c' })
  ]);
  const result = plain(api.bulkUpdatePinMetadata(payload({
    ids: ['pin-a', 'pin-b', 'pin-c'], tagMode: 'none', icon: 'photo', status: '完了'
  })));

  assert.equal(result.updatedCount, 1);
  assert.equal(result.unchangedCount, 2);
  assert.equal(harness.rows[1][13], 'old-a');
  assert.notEqual(harness.rows[2][13], 'old-b');
  assert.equal(harness.rows[3][13], 'old-c');
  assert.equal(harness.rows[2][14], 'photo');
  assert.equal(harness.rows[2][10], '完了');
  assert.deepEqual(result.updates.map((update) => ({
    id: update.id,
    tags: update.tags,
    icon: update.icon,
    status: update.status,
    updatedAt: update.updatedAt
  })), [
    { id: 'pin-a', tags: ['alpha'], icon: 'photo', status: '完了', updatedAt: 'old-a' },
    { id: 'pin-b', tags: ['beta'], icon: 'photo', status: '完了', updatedAt: harness.rows[2][13] },
    { id: 'pin-c', tags: ['gamma'], icon: 'photo', status: '完了', updatedAt: 'old-c' }
  ]);
  assert.equal(harness.audit.writes.some((write) => write.method === 'setValue'), false);
  assert.equal(harness.audit.writes.every((write) => write.method === 'setValues' && write.numColumns === 1), true);
});

test('bulkUpdatePinStatus remains a compatible public API', () => {
  assert.equal(typeof api.bulkUpdatePinStatus, 'function');
  assert.equal(typeof api.bulkUpdatePinMetadata, 'function');
});

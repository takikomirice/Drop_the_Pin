const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');

const MAP_INFO_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明',
  '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧',
  '状態', 'タグ', 'イベント時刻', '更新時刻', 'アイコン', '音声ID'
];
const ROUTE_PINS_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const TEST_EDIT_TOKEN = 'test-edit-token';

function createSheet(rows) {
  return {
    rows: rows.map((row) => row.slice()),
    getLastRow() {
      return this.rows.length;
    },
    getLastColumn() {
      return this.rows.reduce((max, row) => Math.max(max, row.length), 0);
    },
    getDataRange() {
      return {
        getValues: () => this.rows.map((row) => row.slice())
      };
    },
    appendRow(row) {
      this.rows.push(row.slice());
    }
  };
}

function loadApi(mapRows, routePinRows = [ROUTE_PINS_HEADERS]) {
  const sheets = {
    map_info: createSheet(mapRows),
    route_pins: createSheet(routePinRows)
  };
  const context = {
    Logger: { log() {} },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    Utilities: {
      getUuid: () => 'new-pin-id',
      formatDate: () => '2026/04/29 12:34:56'
    },
    CacheService: {
      getScriptCache: () => ({
        get: (key) => (key === `EDIT_TOKEN_${TEST_EDIT_TOKEN}` ? '1' : null)
      })
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: (name) => sheets[name] || null,
        insertSheet: (name) => {
          sheets[name] = createSheet([]);
          return sheets[name];
        }
      })
    }
  };
  vm.runInNewContext(`${codeJs}
globalThis.__duplicatePinApi = {
  duplicatePin: duplicatePin,
  PinData: PinData
};`, context);
  return { api: context.__duplicatePinApi, sheets };
}

function sameRealm(value) {
  return JSON.parse(JSON.stringify(value));
}

function withEditToken(payload = {}) {
  return Object.assign({}, payload, { __editToken: TEST_EDIT_TOKEN });
}

function sourceRow(overrides = {}) {
  return [
    overrides.timestamp || '2026/04/01 10:00:00',
    Object.prototype.hasOwnProperty.call(overrides, 'title') ? overrides.title : '元のピン',
    overrides.description || '説明',
    Object.prototype.hasOwnProperty.call(overrides, 'lat') ? overrides.lat : 35.1,
    Object.prototype.hasOwnProperty.call(overrides, 'lng') ? overrides.lng : 139.2,
    overrides.color || '#43a047',
    overrides.fileId || 'drive-file-id',
    overrides.imageUrl || 'https://example.com/image.jpg',
    overrides.id || 'source-pin-id',
    overrides.links || 'https://example.com/a|not-url|https://example.com/b',
    overrides.status || '対応中',
    overrides.tags || 'alpha|beta',
    overrides.eventAt || '2026-04-29T12:30',
    overrides.updatedAt || '2026-04-02T00:00:00.000Z',
    overrides.icon || 'food',
    overrides.audioId || ''
  ];
}

test('duplicatePin creates an unplaced photo-less pin without route membership', () => {
  const routeRows = [
    ROUTE_PINS_HEADERS,
    ['route-a', 'source-pin-id', 0, '', '']
  ];
  const { api, sheets } = loadApi([
    MAP_INFO_HEADERS,
    sourceRow({ audioId: 'managed_audio_source_123' })
  ], routeRows);

  const result = api.duplicatePin(withEditToken({ sourcePinId: 'source-pin-id', mode: 'unplaced' }));

  assert.equal(result.ok, true);
  assert.equal(result.pin.id, 'new-pin-id');
  assert.equal(result.pin.title, '元のピン（コピー）');
  assert.equal(result.pin.description, '説明');
  assert.equal(result.pin.lat, null);
  assert.equal(result.pin.lng, null);
  assert.equal(result.pin.fileId, '');
  assert.equal(result.pin.imageUrl, '');
  assert.equal(result.pin.folderUrl, '');
  assert.equal(result.pin.hasAudio, false);
  assert.equal(Object.hasOwn(result.pin, 'audioId'), false);
  assert.deepEqual(sameRealm(result.pin.links), ['https://example.com/a', 'https://example.com/b']);
  assert.deepEqual(sameRealm(result.pin.tags), ['alpha', 'beta']);
  assert.equal(result.pin.status, '対応中');
  assert.equal(result.pin.color, '#43a047');
  assert.equal(result.pin.icon, 'food');
  assert.equal(result.pin.eventAt, '2026-04-29T12:30');
  assert.equal(result.pin.timestamp, '2026/04/29 12:34:56');
  assert.equal(result.pin.updatedAt, '2026/04/29 12:34:56');
  assert.deepEqual(sheets.route_pins.rows, routeRows);

  const appended = sheets.map_info.rows[2];
  assert.equal(appended[3], '');
  assert.equal(appended[4], '');
  assert.equal(appended[6], '');
  assert.equal(appended[7], '');
  assert.equal(appended[8], 'new-pin-id');
  assert.equal(appended[15], '', 'duplicate must never copy the source audio relation');
});

test('duplicatePin supports same location and point modes', () => {
  let loaded = loadApi([MAP_INFO_HEADERS, sourceRow()]);

  let result = loaded.api.duplicatePin(withEditToken({ sourcePinId: 'source-pin-id', mode: 'same' }));
  assert.equal(result.ok, true);
  assert.equal(result.pin.lat, 35.1);
  assert.equal(result.pin.lng, 139.2);

  loaded = loadApi([MAP_INFO_HEADERS, sourceRow({ lat: '', lng: '' })]);
  result = loaded.api.duplicatePin(withEditToken({ sourcePinId: 'source-pin-id', mode: 'same' }));
  assert.equal(result.ok, true);
  assert.equal(result.pin.lat, null);
  assert.equal(result.pin.lng, null);

  loaded = loadApi([MAP_INFO_HEADERS, sourceRow()]);
  result = loaded.api.duplicatePin(withEditToken({ sourcePinId: 'source-pin-id', mode: 'point', lat: 34.5, lng: 135.7 }));
  assert.equal(result.ok, true);
  assert.equal(result.pin.lat, 34.5);
  assert.equal(result.pin.lng, 135.7);
});

test('duplicatePin validates source and point location', () => {
  const { api } = loadApi([MAP_INFO_HEADERS, sourceRow()]);

  assert.deepEqual(sameRealm(api.duplicatePin(withEditToken({ sourcePinId: 'missing', mode: 'unplaced' }))), {
    ok: false,
    error: 'pin_not_found'
  });
  assert.deepEqual(sameRealm(api.duplicatePin(withEditToken({ sourcePinId: 'source-pin-id', mode: 'point', lat: 'x', lng: 135 }))), {
    ok: false,
    error: 'invalid_location'
  });
});

test('duplicatePin uses untitled copy label and keeps title under input limit', () => {
  const longTitle = 'あ'.repeat(80);
  let loaded = loadApi([MAP_INFO_HEADERS, sourceRow({ title: '' })]);

  let result = loaded.api.duplicatePin(withEditToken({ sourcePinId: 'source-pin-id', mode: 'unplaced' }));
  assert.equal(result.ok, true);
  assert.equal(result.pin.title, '無題（コピー）');

  loaded = loadApi([MAP_INFO_HEADERS, sourceRow({ title: longTitle })]);
  result = loaded.api.duplicatePin(withEditToken({ sourcePinId: 'source-pin-id', mode: 'unplaced' }));
  assert.equal(result.ok, true);
  assert.equal(result.pin.title.endsWith('（コピー）'), true);
  assert.equal(result.pin.title.length <= 80, true);
});

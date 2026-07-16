const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

const EDIT_TOKEN = 'shared-budget-edit-token';
const SHARE_HEADERS = [
  'createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors',
  'routeIds', 'routeTargetsJson'
];
const MAP_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色', 'ファイルID',
  '画像URL', 'ID', '参考URL一覧', '状態', 'タグ', 'イベント時刻', '更新時刻', 'アイコン'
];
const ROUTE_HEADERS = [
  'routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId',
  'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'
];
const ROUTE_PIN_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];
const TRACK_HEADERS = [
  'trackId', 'name', 'description', 'color', 'sourceType', 'sourceName', 'activeRevision',
  'payloadHash', 'segmentCount', 'pointCount', 'distanceMeters', 'minElevation', 'maxElevation',
  'startTime', 'endTime', 'boundsJson', 'createdAt', 'updatedAt', 'orderIndex', 'visible',
  'lineStyle', 'lineWidth'
];
const TRACK_SEGMENT_HEADERS = [
  'trackId', 'revisionId', 'segmentIndex', 'chunkIndex', 'pointsJson', 'pointCount',
  'createdAt', 'updatedAt'
];
const TOO_LARGE_CODE = 'SHARED_PAYLOAD_TOO_LARGE';
const CREATE_TOO_LARGE = {
  ok: false,
  errorCode: TOO_LARGE_CODE,
  error: '共有対象が大きすぎます。ピンまたはルートを減らしてください。'
};
const READ_TOO_LARGE = {
  ok: false,
  error: 'shared_payload_too_large',
  errorCode: TOO_LARGE_CODE
};

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function createSheet(rows, name, rangeReads) {
  return {
    rows: rows.map((row) => row.slice()),
    getLastRow() { return this.rows.length; },
    getLastColumn() { return this.rows.reduce((max, row) => Math.max(max, row.length), 0); },
    getDataRange() { return { getValues: () => this.rows.map((row) => row.slice()) }; },
    getRange(row, column, numRows = 1, numColumns = 1) {
      const sheet = this;
      const range = {
        getValue() { return (sheet.rows[row - 1] || [])[column - 1]; },
        getValues() {
          rangeReads.push({ sheet: name, row, column, numRows, numColumns });
          return Array.from({ length: numRows }, (_, offset) => {
            const source = sheet.rows[row - 1 + offset] || [];
            return source.slice(column - 1, column - 1 + numColumns);
          });
        },
        setValue(value) {
          if (!sheet.rows[row - 1]) sheet.rows[row - 1] = [];
          sheet.rows[row - 1][column - 1] = value;
          return range;
        },
        setValues(values) {
          values.forEach((valuesRow, rowOffset) => {
            if (!sheet.rows[row - 1 + rowOffset]) sheet.rows[row - 1 + rowOffset] = [];
            valuesRow.forEach((value, columnOffset) => {
              sheet.rows[row - 1 + rowOffset][column - 1 + columnOffset] = value;
            });
          });
          return range;
        },
        setBackground() { return range; },
        setFontColor() { return range; },
        setFontWeight() { return range; }
      };
      return range;
    },
    appendRow(row) { this.rows.push(row.slice()); },
    setFrozenRows() {}
  };
}

function pinRow(id, { title = id, description = `${id} description`, lat = 35, lng = 139 } = {}) {
  return ['', title, description, lat, lng, '#1e88e5', '', '', id, '', '', '', '', '', 'default'];
}

function routeRow(id, orderIndex) {
  return [id, `Route ${id}`, '#1e88e5', 'straight', false, '', '', '', '', orderIndex, true, true, true, 'solid'];
}

function storedShareRow(token, targets, routeIds = '__share_no_routes__') {
  return ['now', 'Budget share', token, '', 'or', true, '', '', routeIds, JSON.stringify({ v: 1, targets })];
}

function trackFixture(id, sourceType, pointCount, orderIndex) {
  const revision = `revision-${id}`;
  const metadata = [
    id, `Track ${id}`, '', '#2196f3', sourceType, `${id}.${sourceType}`, revision,
    'a'.repeat(64), 1, pointCount, 0, '', '', '', '',
    JSON.stringify({ south: 35, west: 139, north: 35.01, east: 139.01 }),
    '', '', orderIndex, true, 'solid', 4
  ];
  const segments = [];
  for (let start = 0, chunkIndex = 0; start < pointCount; start += 500, chunkIndex += 1) {
    const chunkLength = Math.min(500, pointCount - start);
    const points = Array.from({ length: chunkLength }, (_, offset) => [
      35 + ((start + offset) % 100) / 10000,
      139 + ((start + offset) % 100) / 10000,
      null,
      ''
    ]);
    segments.push([id, revision, 0, chunkIndex, JSON.stringify(points), chunkLength, '', '']);
  }
  return { metadata, segments };
}

function loadApi(options = {}) {
  const rangeReads = [];
  const sheets = {
    share_links: createSheet(options.shareRows || [SHARE_HEADERS], 'share_links', rangeReads),
    map_info: createSheet(options.mapRows || [MAP_HEADERS], 'map_info', rangeReads),
    routes: createSheet(options.routeRows || [ROUTE_HEADERS], 'routes', rangeReads),
    route_pins: createSheet(options.routePinRows || [ROUTE_PIN_HEADERS], 'route_pins', rangeReads),
    route_cache: createSheet(options.routeCacheRows || [ROUTE_CACHE_HEADERS], 'route_cache', rangeReads),
    tracks: createSheet(options.trackMetadataRows || [TRACK_HEADERS], 'tracks', rangeReads),
    track_segments: createSheet(options.trackSegmentRows || [TRACK_SEGMENT_HEADERS], 'track_segments', rangeReads)
  };
  const spreadsheet = {
    getSheetByName(name) { return sheets[name] || null; },
    insertSheet(name) {
      sheets[name] = createSheet([], name, rangeReads);
      return sheets[name];
    }
  };
  const context = {
    Logger: { log() {} },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.example/exec' }) },
    Utilities: { getUuid: () => 'new-share-token' },
    CacheService: {
      getScriptCache: () => ({
        get: (key) => key === `EDIT_TOKEN_${EDIT_TOKEN}` ? '1' : null
      })
    },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet }
  };
  vm.runInNewContext(`${codeJs}
globalThis.__sharedBudgetReconstructedTrackIds = [];
const __originalTrackFromStoredRows = trackFromStoredRows_;
trackFromStoredRows_ = function(metadata, rows) {
  const track = __originalTrackFromStoredRows(metadata, rows);
  globalThis.__sharedBudgetReconstructedTrackIds.push(metadata.trackId);
  return track;
};
globalThis.__sharedBudgetApi = {
  createShareLink, getSharedViewData,
  validateSharedPayloadBudget_: typeof validateSharedPayloadBudget_ === 'function' ? validateSharedPayloadBudget_ : null,
  utf8JsonByteLength_: typeof utf8JsonByteLength_ === 'function' ? utf8JsonByteLength_ : null,
  buildSharedViewDataForLink_: typeof buildSharedViewDataForLink_ === 'function' ? buildSharedViewDataForLink_ : null,
  constants: {
    pins: typeof SHARED_PAYLOAD_MAX_PINS === 'number' ? SHARED_PAYLOAD_MAX_PINS : null,
    routes: typeof SHARED_PAYLOAD_MAX_ROUTES === 'number' ? SHARED_PAYLOAD_MAX_ROUTES : null,
    coordinates: typeof SHARED_PAYLOAD_MAX_COORDINATE_POINTS === 'number' ? SHARED_PAYLOAD_MAX_COORDINATE_POINTS : null,
    bytes: typeof SHARED_PAYLOAD_MAX_JSON_BYTES === 'number' ? SHARED_PAYLOAD_MAX_JSON_BYTES : null
  }
};`, context);
  return {
    api: context.__sharedBudgetApi,
    rangeReads,
    reconstructedTrackIds: context.__sharedBudgetReconstructedTrackIds,
    sheets
  };
}

function withEditToken(payload = {}) {
  return Object.assign({}, payload, { __editToken: EDIT_TOKEN });
}

function functionSource(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let cursor = open; cursor < source.length; cursor += 1) {
    if (source[cursor] === '{') depth += 1;
    if (source[cursor] === '}') {
      depth -= 1;
      if (depth === 0) return source.slice(start, cursor + 1);
    }
  }
  assert.fail(`Unclosed ${name}`);
}

test('named shared budget constants and validator accept each exact count and reject limit plus one', () => {
  const { api } = loadApi();
  assert.deepEqual(plain(api.constants), {
    pins: 2000,
    routes: 100,
    coordinates: 50000,
    bytes: 5 * 1024 * 1024
  });
  assert.equal(typeof api.validateSharedPayloadBudget_, 'function');

  for (const [field, limit] of [
    ['pinCount', api.constants.pins],
    ['routeCount', api.constants.routes],
    ['coordinateCount', api.constants.coordinates]
  ]) {
    const exact = { pinCount: 0, routeCount: 0, coordinateCount: 0 };
    exact[field] = limit;
    assert.equal(api.validateSharedPayloadBudget_(null, exact).ok, true, `${field} exact limit`);
    exact[field] = limit + 1;
    assert.equal(api.validateSharedPayloadBudget_(null, exact).ok, false, `${field} limit + 1`);
  }
});

test('UTF-8 JSON bytes use real multibyte size and accept exactly 5 MiB only', () => {
  const { api } = loadApi();
  assert.equal(typeof api.utf8JsonByteLength_, 'function');
  const japanese = '日本語😀';
  assert.equal(api.utf8JsonByteLength_(japanese), Buffer.byteLength(japanese, 'utf8'));

  const emptyDto = { ok: true, text: '' };
  const emptyBytes = Buffer.byteLength(JSON.stringify(emptyDto), 'utf8');
  const japaneseBytes = Buffer.byteLength(japanese, 'utf8');
  const exactDto = {
    ok: true,
    text: japanese + 'a'.repeat(api.constants.bytes - emptyBytes - japaneseBytes)
  };
  assert.equal(Buffer.byteLength(JSON.stringify(exactDto), 'utf8'), api.constants.bytes);
  const metrics = { pinCount: 0, routeCount: 0, coordinateCount: 0 };
  assert.equal(api.validateSharedPayloadBudget_(exactDto, metrics).ok, true);
  exactDto.text += 'a';
  assert.equal(api.validateSharedPayloadBudget_(exactDto, metrics).ok, false);
});

test('create validates an unselected-route payload before append and request data cannot override byte budget', () => {
  const hugeJapaneseDescription = '日'.repeat(2 * 1024 * 1024);
  const { api, sheets } = loadApi({
    mapRows: [MAP_HEADERS, pinRow('huge-pin', { description: hugeJapaneseDescription })]
  });

  const created = api.createShareLink(withEditToken({
    routeTargets: [],
    sharedPayloadMaxPins: Number.MAX_SAFE_INTEGER,
    sharedPayloadMaxRoutes: Number.MAX_SAFE_INTEGER,
    sharedPayloadMaxCoordinatePoints: Number.MAX_SAFE_INTEGER,
    sharedPayloadMaxJsonBytes: Number.MAX_SAFE_INTEGER
  }));

  assert.deepEqual(plain(created), CREATE_TOO_LARGE);
  assert.equal(sheets.share_links.rows.length, 1);

  sheets.share_links.rows.push(storedShareRow('oversized-read', []));
  assert.deepEqual(plain(api.getSharedViewData('oversized-read')), READ_TOO_LARGE);
});

test('2000 pins create and read successfully, then one added pin fails closed without partial DTO', () => {
  const mapRows = [MAP_HEADERS];
  for (let index = 0; index < 2000; index += 1) mapRows.push(pinRow(`pin-${index}`));
  const { api, sheets } = loadApi({ mapRows });

  const created = api.createShareLink(withEditToken({ routeTargets: [] }));
  assert.equal(created.ok, true);
  assert.equal(sheets.share_links.rows.length, 2);
  const exactRead = api.getSharedViewData(created.token);
  assert.equal(exactRead.ok, true);
  assert.equal(exactRead.pins.length, 2000);

  sheets.map_info.rows.push(pinRow('pin-2000'));
  const oversizedRead = api.getSharedViewData(created.token);
  assert.deepEqual(plain(oversizedRead), READ_TOO_LARGE);
  assert.deepEqual(Object.keys(plain(oversizedRead)).sort(), ['error', 'errorCode', 'ok']);
  assert.deepEqual(plain(api.createShareLink(withEditToken({ routeTargets: [] }))), CREATE_TOO_LARGE);
  assert.equal(sheets.share_links.rows.length, 2);
});

test('100 actually returned routes succeed while 101 reject creation without append', () => {
  function fixture(routeCount) {
    const routeRows = [ROUTE_HEADERS];
    const routePinRows = [ROUTE_PIN_HEADERS];
    const targets = [];
    for (let index = 0; index < routeCount; index += 1) {
      const id = `route-${index}`;
      routeRows.push(routeRow(id, index));
      routePinRows.push([id, 'shared-pin', 0, '', '']);
      targets.push({ type: 'pin-route', id });
    }
    return loadApi({
      mapRows: [MAP_HEADERS, pinRow('shared-pin')],
      routeRows,
      routePinRows
    });
  }

  const exact = fixture(100);
  const exactResult = exact.api.createShareLink(withEditToken({
    routeTargets: Array.from({ length: 100 }, (_, index) => ({ type: 'pin-route', id: `route-${index}` }))
  }));
  assert.equal(exactResult.ok, true);
  assert.equal(exact.api.getSharedViewData(exactResult.token).routes.length, 100);

  const oversized = fixture(101);
  const oversizedResult = oversized.api.createShareLink(withEditToken({
    routeTargets: Array.from({ length: 101 }, (_, index) => ({ type: 'pin-route', id: `route-${index}` }))
  }));
  assert.deepEqual(plain(oversizedResult), CREATE_TOO_LARGE);
  assert.equal(oversized.sheets.share_links.rows.length, 1);
});

test('shared track assembly reads geometry cells for selected tracks only', () => {
  const selected = trackFixture('selected-track', 'gpx', 2, 0);
  const unselected = trackFixture('unselected-track', 'geojson', 2, 1);
  const { api, rangeReads, reconstructedTrackIds } = loadApi({
    trackMetadataRows: [TRACK_HEADERS, selected.metadata, unselected.metadata],
    trackSegmentRows: [TRACK_SEGMENT_HEADERS].concat(selected.segments, unselected.segments)
  });

  const result = api.createShareLink(withEditToken({
    routeTargets: [{ type: 'gpx-route', id: 'selected-track' }]
  }));

  assert.equal(result.ok, true);
  assert.deepEqual(plain(reconstructedTrackIds), ['selected-track']);
  const geometryReads = rangeReads.filter((entry) => entry.sheet === 'track_segments'
    && entry.row >= 2
    && entry.column <= 5
    && entry.column + entry.numColumns - 1 >= 5);
  assert.deepEqual(geometryReads, [{
    sheet: 'track_segments', row: 2, column: 5, numRows: 1, numColumns: 1
  }]);
});

test('one public pin plus 49999 track points succeeds without double-counting pin-route pinIds', () => {
  const fixtures = [
    trackFixture('track-a', 'gpx', 20000, 0),
    trackFixture('track-b', 'geojson', 20000, 1),
    trackFixture('track-c', 'gpx', 9999, 2)
  ];
  const targets = [
    { type: 'pin-route', id: 'pin-route-one' },
    { type: 'gpx-route', id: 'track-a' },
    { type: 'geojson-route', id: 'track-b' },
    { type: 'gpx-route', id: 'track-c' }
  ];
  const { api, reconstructedTrackIds } = loadApi({
    mapRows: [MAP_HEADERS, pinRow('shared-pin')],
    routeRows: [ROUTE_HEADERS, routeRow('pin-route-one', 0)],
    routePinRows: [ROUTE_PIN_HEADERS, ['pin-route-one', 'shared-pin', 0, '', '']],
    trackMetadataRows: [TRACK_HEADERS].concat(fixtures.map((item) => item.metadata)),
    trackSegmentRows: [TRACK_SEGMENT_HEADERS].concat(fixtures.flatMap((item) => item.segments))
  });

  const result = api.createShareLink(withEditToken({ routeTargets: targets }));

  assert.equal(result.ok, true);
  assert.deepEqual(plain(reconstructedTrackIds), ['track-a', 'track-b', 'track-c']);
});

test('coordinate overflow stops before rebuilding the crossing and remaining selected tracks', () => {
  const fixtures = [
    trackFixture('track-a', 'gpx', 20000, 0),
    trackFixture('track-b', 'geojson', 20000, 1),
    trackFixture('track-c', 'gpx', 10001, 2),
    trackFixture('track-d', 'geojson', 1, 3)
  ];
  const targets = [
    { type: 'gpx-route', id: 'track-a' },
    { type: 'geojson-route', id: 'track-b' },
    { type: 'gpx-route', id: 'track-c' },
    { type: 'geojson-route', id: 'track-d' }
  ];
  const { api, reconstructedTrackIds } = loadApi({
    shareRows: [SHARE_HEADERS, storedShareRow('coordinate-overflow', targets)],
    trackMetadataRows: [TRACK_HEADERS].concat(fixtures.map((item) => item.metadata)),
    trackSegmentRows: [TRACK_SEGMENT_HEADERS].concat(fixtures.flatMap((item) => item.segments))
  });

  const result = api.getSharedViewData('coordinate-overflow');

  assert.deepEqual(plain(result), READ_TOO_LARGE);
  assert.deepEqual(plain(reconstructedTrackIds), ['track-a', 'track-b']);
});

test('create and read use the same DTO builder and clients show the dedicated safe message', () => {
  const createSource = functionSource(codeJs, 'createShareLink');
  const readSource = functionSource(codeJs, 'getSharedViewData');
  assert.match(createSource, /buildSharedViewDataForLink_/);
  assert.match(readSource, /buildSharedViewDataForLink_/);

  const safeMessage = '共有データが大きすぎるため表示できません。共有対象を減らしてください。';
  assert.match(indexHtml, /SHARED_PAYLOAD_TOO_LARGE/);
  assert.ok(indexHtml.includes(safeMessage));
  assert.match(sharedHtml, /shared_payload_too_large/);
  assert.match(sharedHtml, /SHARED_PAYLOAD_TOO_LARGE/);
  assert.ok(sharedHtml.includes(safeMessage));
});

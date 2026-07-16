const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

const SHARE_HEADERS = ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'];
const ROUTE_HEADERS = ['routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'];
const ROUTE_PIN_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];
const MAP_HEADERS = ['タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色', 'ファイルID', '画像URL', 'ID', '参考URL一覧', '状態', 'タグ', 'イベント時刻', '更新時刻', 'アイコン'];
const TRACK_HEADERS = ['trackId', 'name', 'description', 'color', 'sourceType', 'sourceName', 'activeRevision', 'payloadHash', 'segmentCount', 'pointCount', 'distanceMeters', 'minElevation', 'maxElevation', 'startTime', 'endTime', 'boundsJson', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'lineStyle', 'lineWidth'];
const TRACK_SEGMENT_HEADERS = ['trackId', 'revisionId', 'segmentIndex', 'chunkIndex', 'pointsJson', 'pointCount', 'createdAt', 'updatedAt'];
const EDIT_TOKEN = 'share-route-targets-edit-token';

function createSheet(rows) {
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

function pinRow(id, title = id) {
  return ['', title, `${title} description`, 35, 139, '#1e88e5', '', '', id, '', '', 'alpha', '', '', 'default'];
}

function trackRows(id, sourceType, orderIndex = 0) {
  const revision = `rev-${id}-${sourceType}`;
  const metadata = [
    id, `${sourceType.toUpperCase()} ${id}`, `${sourceType} description`, sourceType === 'gpx' ? '#2196f3' : '#9c27b0',
    sourceType, `${id}.${sourceType === 'gpx' ? 'gpx' : 'geojson'}`, revision, 'a'.repeat(64), 1, 2, 100,
    10, 20, '2026-01-01T00:00:00.000Z', '2026-01-01T01:00:00.000Z',
    JSON.stringify({ south: 35, west: 139, north: 36, east: 140 }),
    '2026-01-01T00:00:00.000Z', '2026-01-01T00:00:00.000Z', orderIndex, true, 'solid', 4
  ];
  const segment = [id, revision, 0, 0, JSON.stringify([[35, 139, 10, '2026-01-01T00:00:00.000Z'], [36, 140, 20, '2026-01-01T01:00:00.000Z']]), 2, '', ''];
  return { metadata, segment };
}

function loadApi(options = {}) {
  const gpx = trackRows('same', 'gpx', 0);
  const geo = trackRows('geo-1', 'geojson', 1);
  const hidden = trackRows('not-selected', 'gpx', 2);
  const sheets = {
    share_links: createSheet(options.shareRows || [SHARE_HEADERS]),
    routes: createSheet(options.routeRows || [ROUTE_HEADERS, ['same', 'Pin same', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']]),
    route_pins: createSheet(options.routePinRows || [ROUTE_PIN_HEADERS]),
    route_cache: createSheet(options.routeCacheRows || [ROUTE_CACHE_HEADERS]),
    map_info: createSheet(options.mapRows || [MAP_HEADERS]),
    tracks: createSheet(options.trackMetadataRows || [TRACK_HEADERS, gpx.metadata, geo.metadata, hidden.metadata]),
    track_segments: createSheet(options.trackSegmentRows || [TRACK_SEGMENT_HEADERS, gpx.segment, geo.segment, hidden.segment])
  };
  const spreadsheet = {
    getSheetByName: (name) => sheets[name] || null,
    insertSheet: (name) => {
      sheets[name] = createSheet([]);
      return sheets[name];
    }
  };
  const context = {
    Logger: { log() {} },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.example/exec' }) },
    Utilities: { getUuid: () => 'new-share-token' },
    CacheService: { getScriptCache: () => ({ get: (key) => key === `EDIT_TOKEN_${EDIT_TOKEN}` ? '1' : null }) },
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet }
  };
  vm.runInNewContext(`${codeJs}
globalThis.__shareTargetsApi = {
  ensureShareLinksSheet_, shareRowToLink_, normalizeShareRouteTargets_,
  serializeShareRouteTargets_, parseShareRouteTargetsJson_, createShareLink,
  listShareLinks, getSharedViewData, getSharedRoadRouteCache_
};`, context);
  return { api: context.__shareTargetsApi, sheets, spreadsheet };
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function withEditToken(payload) {
  return Object.assign({}, payload, { __editToken: EDIT_TOKEN });
}

function targetJson(targets) {
  return JSON.stringify({ v: 1, targets });
}

test('routeTargetsJson is appended by header name without moving legacy or extension columns', () => {
  const headers = SHARE_HEADERS.concat(['extensionColumn']);
  const row = ['now', 'Legacy', 'legacy-token', '', 'or', true, '', '', '', 'keep-me'];
  const { api, sheets, spreadsheet } = loadApi({ shareRows: [headers, row] });

  api.ensureShareLinksSheet_(spreadsheet);

  assert.deepEqual(plain(sheets.share_links.rows[0]), headers.concat(['routeTargetsJson']));
  assert.equal(sheets.share_links.rows[1][9], 'keep-me');
  assert.equal(sheets.share_links.rows[1][10], undefined);
});

test('typed targets use a versioned canonical JSON format and type plus NUL plus id identity', () => {
  const { api } = loadApi();
  const targets = api.normalizeShareRouteTargets_([
    { type: 'pin-route', id: ' same ' },
    { type: 'gpx-route', id: 'same' },
    { type: 'pin-route', id: 'same' },
    { type: 'geojson-route', id: 'geo-1' },
    { type: 'gpx-route', id: 'missing' }
  ]);

  assert.deepEqual(plain(targets), [
    { type: 'pin-route', id: 'same' },
    { type: 'gpx-route', id: 'same' },
    { type: 'geojson-route', id: 'geo-1' }
  ]);
  const serialized = api.serializeShareRouteTargets_(targets);
  assert.equal(serialized, targetJson(plain(targets)));
  assert.deepEqual(plain(api.parseShareRouteTargetsJson_(serialized)), { state: 'valid', targets: plain(targets) });
});

test('legacy all, none and selected pin routes preserve their old scope and never publish tracks', () => {
  const base = {
    routeRows: [ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid'],
      ['route-b', 'Route B', '#43a047', 'straight', false, '', '', '', '', 1, true, true, true, 'solid']],
    routePinRows: [ROUTE_PIN_HEADERS, ['route-a', 'p1', 0, '', ''], ['route-b', 'p2', 0, '', '']],
    mapRows: [MAP_HEADERS, pinRow('p1'), pinRow('p2')]
  };
  for (const [routeIds, expected] of [['', ['route-a', 'route-b']], ['__share_no_routes__', []], ['route-b', ['route-b']]]) {
    const shareRows = [SHARE_HEADERS.concat(['routeTargetsJson']), ['now', 'Legacy', `legacy-${routeIds}`, '', 'or', true, '', '', routeIds, '']];
    const { api } = loadApi(Object.assign({}, base, { shareRows }));
    const result = api.getSharedViewData(`legacy-${routeIds}`);
    assert.equal(result.ok, true);
    assert.deepEqual(plain(result.routeGroups.map((route) => route.id)), expected);
    assert.deepEqual(plain(result.routes.map((route) => [route.type, route.id])), expected.map((id) => ['pin-route', id]));
    assert.equal(result.routes.some((route) => route.type !== 'pin-route'), false);
  }
});

test('new links publish selected pin, GPX and GeoJSON routes with collision-safe typed IDs only', () => {
  const headers = SHARE_HEADERS.concat(['routeTargetsJson']);
  const targets = [
    { type: 'pin-route', id: 'same' },
    { type: 'gpx-route', id: 'same' },
    { type: 'geojson-route', id: 'geo-1' }
  ];
  const { api } = loadApi({
    shareRows: [headers, ['now', 'Mixed', 'mixed-token', '', 'or', true, '', '', 'same', targetJson(targets)]],
    routeRows: [ROUTE_HEADERS, ['same', 'Pin same', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']],
    routePinRows: [ROUTE_PIN_HEADERS, ['same', 'p1', 0, '', '']],
    mapRows: [MAP_HEADERS, pinRow('p1')]
  });

  const result = api.getSharedViewData('mixed-token');

  assert.equal(result.ok, true);
  assert.deepEqual(plain(result.routes.map((route) => [route.type, route.id])), [
    ['pin-route', 'same'], ['gpx-route', 'same'], ['geojson-route', 'geo-1']
  ]);
  assert.deepEqual(plain(result.routeGroups.map((route) => route.id)), ['same']);
  assert.equal(JSON.stringify(result).includes('not-selected'), false);
});

test('typed shared routes filter formal pin and track order instead of following target selection order', () => {
  const trackA = trackRows('track-a', 'geojson', 8);
  const trackB = trackRows('track-b', 'gpx', 2);
  const targets = [
    { type: 'geojson-route', id: 'track-a' },
    { type: 'pin-route', id: 'pin-a' },
    { type: 'gpx-route', id: 'track-b' },
    { type: 'pin-route', id: 'pin-b' }
  ];
  const { api } = loadApi({
    shareRows: [SHARE_HEADERS.concat(['routeTargetsJson']), [
      'now', 'Ordered', 'ordered-token', '', 'or', true, '', '', 'pin-a|pin-b', targetJson(targets)
    ]],
    routeRows: [
      ROUTE_HEADERS,
      ['pin-a', 'Pin A', '#1e88e5', 'straight', false, '', '', '', '', 7, true, true, true, 'solid'],
      ['pin-b', 'Pin B', '#43a047', 'straight', false, '', '', '', '', 1, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PIN_HEADERS,
      ['pin-a', 'photo-a', 0, '', ''],
      ['pin-b', 'photo-b', 0, '', '']
    ],
    mapRows: [MAP_HEADERS, pinRow('photo-a'), pinRow('photo-b')],
    trackMetadataRows: [TRACK_HEADERS, trackA.metadata, trackB.metadata],
    trackSegmentRows: [TRACK_SEGMENT_HEADERS, trackA.segment, trackB.segment]
  });

  const result = api.getSharedViewData('ordered-token');

  assert.equal(result.ok, true);
  assert.deepEqual(plain(result.routes.map((route) => [route.type, route.id])), [
    ['pin-route', 'pin-b'],
    ['pin-route', 'pin-a'],
    ['gpx-route', 'track-b'],
    ['geojson-route', 'track-a']
  ]);
});

test('new targets empty means no routes and malformed JSON fails closed without legacy fallback', () => {
  const headers = SHARE_HEADERS.concat(['routeTargetsJson']);
  for (const [json, token] of [[targetJson([]), 'empty-targets'], ['{"v":1,"targets":[', 'broken-targets']]) {
    const { api } = loadApi({
      shareRows: [headers, ['now', 'Closed', token, '', 'or', true, '', '', '', json]],
      routeRows: [ROUTE_HEADERS, ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']],
      routePinRows: [ROUTE_PIN_HEADERS, ['route-a', 'p1', 0, '', '']],
      mapRows: [MAP_HEADERS, pinRow('p1')]
    });
    const result = api.getSharedViewData(token);
    assert.equal(result.ok, true);
    assert.equal(result.noRoutes, true);
    assert.deepEqual(plain(result.routeGroups), []);
    assert.deepEqual(plain(result.routes), []);
  }
});

test('track public DTO contains display geometry only and omits private metadata', () => {
  const targets = [{ type: 'gpx-route', id: 'same' }];
  const { api } = loadApi({
    shareRows: [SHARE_HEADERS.concat(['routeTargetsJson']), ['now', 'GPX', 'gpx-token', '', 'or', true, '', '', '__share_no_routes__', targetJson(targets)]]
  });

  const result = api.getSharedViewData('gpx-token');
  const route = plain(result.routes[0]);

  assert.deepEqual(Object.keys(route).sort(), [
    'bounds', 'color', 'description', 'distanceMeters', 'id', 'lineStyle', 'lineWidth', 'name', 'segments', 'type', 'visible'
  ]);
  assert.deepEqual(Object.keys(route.segments[0][0]).sort(), ['lat', 'lng']);
  assert.equal(route.lineWidth, 4);
  for (const privateKey of ['sourceName', 'revisionId', 'activeRevision', 'startTime', 'endTime', 'minElevation', 'maxElevation', 'createdAt', 'updatedAt', 'orderIndex']) {
    assert.equal(JSON.stringify(route).includes(privateKey), false, `${privateKey} must not be public`);
  }
});

test('new link creation stores canonical targets and only selected pin routes in legacy routeIds', () => {
  const { api, sheets, spreadsheet } = loadApi({
    shareRows: [SHARE_HEADERS],
    routeRows: [ROUTE_HEADERS, ['same', 'Pin same', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']]
  });
  api.ensureShareLinksSheet_(spreadsheet);
  const targets = [{ type: 'pin-route', id: 'same' }, { type: 'gpx-route', id: 'same' }, { type: 'geojson-route', id: 'geo-1' }];

  const result = api.createShareLink(withEditToken({ label: 'New', routeTargets: targets }));

  const header = sheets.share_links.rows[0];
  const saved = sheets.share_links.rows[1];
  assert.equal(saved[header.indexOf('routeIds')], 'same');
  assert.equal(saved[header.indexOf('routeTargetsJson')], targetJson(targets));
  assert.deepEqual(plain(result.shareLink.routeTargets), targets);
  assert.deepEqual(plain(api.listShareLinks(withEditToken()).items[0].routeTargets), targets);
});

test('invalid and disabled share links return no route or track payload', () => {
  const { api } = loadApi({
    shareRows: [SHARE_HEADERS.concat(['routeTargetsJson']), ['now', 'Disabled', 'disabled', '', 'or', false, '', '', '', targetJson([{ type: 'gpx-route', id: 'same' }])]]
  });
  assert.deepEqual(plain(api.getSharedViewData('missing')), { ok: false, error: 'invalid_share_link' });
  assert.deepEqual(plain(api.getSharedViewData('disabled')), { ok: false, error: 'revoked_share_link' });
});

test('road cache authorization follows typed pin-route targets and rejects tracks with the same ID', () => {
  const cacheKey = 'osrm|road|false|p1:35.00000,139.00000>p2:36.00000,140.00000';
  const common = {
    routeRows: [ROUTE_HEADERS, ['same', 'Road', '#1e88e5', 'road', false, '', '', '', '', 0, true, true, true, 'solid']],
    routePinRows: [ROUTE_PIN_HEADERS, ['same', 'p1', 0, '', ''], ['same', 'p2', 1, '', '']],
    mapRows: [MAP_HEADERS, pinRow('p1'), ['', 'p2', '', 36, 140, '#1e88e5', '', '', 'p2', '', '', 'alpha', '', '', 'default']],
    routeCacheRows: [ROUTE_CACHE_HEADERS, [cacheKey, 'same', '[[35,139],[36,140]]', 'osrm', '2026-01-01T00:00:00.000Z', '']]
  };
  const headers = SHARE_HEADERS.concat(['routeTargetsJson']);
  const denied = loadApi(Object.assign({}, common, { shareRows: [headers, ['now', 'Track only', 'track-only', '', 'or', true, '', '', '__share_no_routes__', targetJson([{ type: 'gpx-route', id: 'same' }])]] })).api;
  const allowed = loadApi(Object.assign({}, common, { shareRows: [headers, ['now', 'Pin route', 'pin-route', '', 'or', true, '', '', 'same', targetJson([{ type: 'pin-route', id: 'same' }])]] })).api;

  assert.deepEqual(plain(denied.getSharedRoadRouteCache_('track-only', 'same')), { ok: false });
  assert.equal(allowed.getSharedRoadRouteCache_('pin-route', 'same').ok, true);
});

test('share manager and shared viewer expose unified typed routes without edit affordances', () => {
  assert.match(indexHtml, /routeTargets:\s*\[\]/);
  assert.match(indexHtml, /pin-route[\s\S]*gpx-route[\s\S]*geojson-route/);
  assert.match(indexHtml, /ルート選択は線の表示対象。ピンの絞り込みはタグと色/);
  assert.match(indexHtml, /routeTargets:\s*routeTargets/);

  for (const id of ['shared-topbar', 'shared-display-name', 'shared-app-shell', 'shared-side-panel', 'shared-pin-list', 'shared-route-list', 'shared-mobile-sheet-tabs']) {
    assert.match(sharedHtml, new RegExp(`id=["']${id}["']`));
  }
  assert.match(sharedHtml, /Array\.isArray\(result\.routes\)[\s\S]*result\.routeGroups/);
  assert.doesNotMatch(sharedHtml, /Sortable|draggable\s*=\s*true|edit-token|share-open-btn/);
});

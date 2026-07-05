const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');

function createSheet(rows) {
  const sheet = {
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
    getRange(row, column, numRows = 1, numColumns = 1) {
      const range = {
        getValues: () => {
          const values = [];
          for (let r = 0; r < numRows; r += 1) {
            const source = this.rows[row - 1 + r] || [];
            values.push(source.slice(column - 1, column - 1 + numColumns));
          }
          return values;
        },
        setValues: (values) => {
          for (let r = 0; r < numRows; r += 1) {
            const targetRow = row - 1 + r;
            if (!this.rows[targetRow]) this.rows[targetRow] = [];
            for (let c = 0; c < numColumns; c += 1) {
              this.rows[targetRow][column - 1 + c] = values[r][c];
            }
          }
          return this;
        },
        setValue: (value) => {
          if (!this.rows[row - 1]) this.rows[row - 1] = [];
          this.rows[row - 1][column - 1] = value;
          return range;
        },
        setBackground: () => range,
        setFontColor: () => range,
        setFontWeight: () => range
      };
      return range;
    },
    appendRow(row) {
      this.rows.push(row.slice());
    },
    setFrozenRows() {}
  };
  return sheet;
}

const MAP_INFO_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色', 'ファイルID',
  '画像URL', 'ID', '参考URL一覧', '状態', 'タグ', 'イベント時刻', '更新時刻', 'アイコン'
];
const ROUTE_HEADERS = ['routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'];
const ROUTE_PIN_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];
const TEST_EDIT_TOKEN = 'test-edit-token';

function pinRow(id, title, color, tags, lat = 35, lng = 139) {
  return ['', title, '', lat, lng, color, '', '', id, '', '', tags, '', '', 'default'];
}

function loadApi({ shareRows, routeRows = [ROUTE_HEADERS], routePinRows = [ROUTE_PIN_HEADERS], mapRows = [MAP_INFO_HEADERS], routeCacheRows = [ROUTE_CACHE_HEADERS] }) {
  const sheets = {
    share_links: createSheet(shareRows),
    routes: createSheet(routeRows),
    route_pins: createSheet(routePinRows),
    route_cache: createSheet(routeCacheRows),
    map_info: createSheet(mapRows)
  };
  const context = {
    Logger: { log() {} },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.example/exec' }) },
    Utilities: { getUuid: () => 'uuid-token' },
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
  const script = `${codeJs}
globalThis.__shareRouteIdsApi = {
  SHARE_LINKS_HEADERS: SHARE_LINKS_HEADERS,
  normalizeShareRouteIds_: normalizeShareRouteIds_,
  serializeShareRouteIds_: serializeShareRouteIds_,
  deserializeShareRouteIds_: deserializeShareRouteIds_,
  ensureShareLinksSheet_: ensureShareLinksSheet_,
  shareRowToLink_: shareRowToLink_,
  createShareLink: createShareLink,
  listShareLinks: listShareLinks,
  getSharedViewData: getSharedViewData
  ,
  getSharedRoadRouteCache_: getSharedRoadRouteCache_
};`;
  vm.runInNewContext(script, context);
  return { api: context.__shareRouteIdsApi, sheets };
}

function sameRealm(value) {
  return JSON.parse(JSON.stringify(value));
}

function withEditToken(payload = {}) {
  return Object.assign({}, payload, { __editToken: TEST_EDIT_TOKEN });
}

test('ensureShareLinksSheet_ appends missing routeIds without moving existing colors data', () => {
  const existingHeaders = ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors'];
  const existingRow = ['2026-04-01T00:00:00.000Z', 'Share', 'tok-1', '', 'or', true, '', '#e53935|#1e88e5'];
  const { api, sheets } = loadApi({
    shareRows: [existingHeaders, existingRow],
    routeRows: [['routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle']]
  });

  api.ensureShareLinksSheet_({
    getSheetByName: (name) => (name === 'share_links' ? sheets.share_links : null),
    insertSheet: () => assert.fail('existing share_links sheet should be reused')
  });

  assert.deepEqual(
    sameRealm(sheets.share_links.rows[0]),
    ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds']
  );
  assert.equal(sheets.share_links.rows[1][7], '#e53935|#1e88e5');
  assert.equal(sheets.share_links.rows[1][8], undefined);
  assert.deepEqual(sameRealm(api.shareRowToLink_(sheets.share_links.rows[1]).routeIds), []);
});

test('createShareLink normalizes and saves existing routeIds while list and shared view return them', () => {
  const { api, sheets } = loadApi({
    shareRows: [['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors']],
    routeRows: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid'],
      ['route-b', 'Route B', '#43a047', 'straight', false, '', '', '', '', 1, true, true, true, 'solid']
    ]
  });

  assert.deepEqual(sameRealm(api.normalizeShareRouteIds_([' route-a ', 'missing-route', 'route-a', '', null, 'route-b'])), ['route-a', 'route-b']);
  assert.equal(api.serializeShareRouteIds_(['route-b', ' route-a ', 'route-b']), 'route-b|route-a');
  assert.deepEqual(sameRealm(api.deserializeShareRouteIds_(' route-b | |route-a|route-b ')), ['route-b', 'route-a']);

  const created = api.createShareLink(withEditToken({
    label: 'Selected routes',
    tags: ['alpha'],
    tagMode: 'and',
    colors: ['#E53935'],
    routeIds: [' route-a ', 'missing-route', 'route-a', 'route-b']
  }));

  assert.deepEqual(sameRealm(created.shareLink.routeIds), ['route-a', 'route-b']);
  assert.equal(sheets.share_links.rows[1][7], '#e53935');
  assert.equal(sheets.share_links.rows[1][8], 'route-a|route-b');

  const listed = api.listShareLinks(withEditToken());
  assert.deepEqual(sameRealm(listed.items[0].routeIds), ['route-a', 'route-b']);

  const shared = api.getSharedViewData('uuid-token');
  assert.deepEqual(sameRealm(shared.shareLink.routeIds), ['route-a', 'route-b']);
  assert.deepEqual(sameRealm(shared.allowedRouteIds), ['route-a', 'route-b']);
});

test('getSharedViewData uses selected routeIds only for route groups and keeps matching pins', () => {
  const { api } = loadApi({
    shareRows: [
      ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'],
      ['2026-04-01T00:00:00.000Z', 'Route A only', 'tok-route-a', 'alpha', 'or', true, '', '#e53935', 'route-a|missing-route']
    ],
    routeRows: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid'],
      ['route-b', 'Route B', '#43a047', 'straight', false, '', '', '', '', 1, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PIN_HEADERS,
      ['route-a', 'p1', 0, '', ''],
      ['route-a', 'p2', 1, '', ''],
      ['route-b', 'p2', 0, '', ''],
      ['route-b', 'p3', 1, '', '']
    ],
    mapRows: [
      MAP_INFO_HEADERS,
      pinRow('p1', 'Route A pin 1', '#e53935', 'alpha'),
      pinRow('p2', 'Shared route pin', '#e53935', 'alpha'),
      pinRow('p3', 'Route B only', '#e53935', 'alpha'),
      pinRow('p4', 'Orphan', '#e53935', 'alpha'),
      pinRow('p5', 'Wrong tag', '#e53935', 'beta')
    ]
  });

  const shared = api.getSharedViewData('tok-route-a');

  assert.equal(shared.ok, true);
  assert.deepEqual(sameRealm(shared.pins.map((pin) => pin.id)), ['p1', 'p2', 'p3', 'p4']);
  assert.deepEqual(sameRealm(shared.routeGroups.map((group) => group.routeId)), ['route-a']);
  assert.deepEqual(sameRealm(shared.routeGroups[0].pinIds), ['p1', 'p2']);
});

test('getSharedViewData keeps legacy all-route behavior when routeIds is empty', () => {
  const { api } = loadApi({
    shareRows: [
      ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'],
      ['2026-04-01T00:00:00.000Z', 'All routes', 'tok-all', 'alpha', 'or', true, '', '#e53935', '']
    ],
    routeRows: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid'],
      ['route-b', 'Route B', '#43a047', 'straight', false, '', '', '', '', 1, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PIN_HEADERS,
      ['route-a', 'p1', 0, '', ''],
      ['route-a', 'p2', 1, '', ''],
      ['route-b', 'p3', 0, '', '']
    ],
    mapRows: [
      MAP_INFO_HEADERS,
      pinRow('p1', 'Route A pin 1', '#e53935', 'alpha'),
      pinRow('p2', 'Route A pin 2', '#e53935', 'alpha'),
      pinRow('p3', 'Route B pin', '#e53935', 'alpha'),
      pinRow('p4', 'Orphan', '#e53935', 'alpha')
    ]
  });

  const shared = api.getSharedViewData('tok-all');

  assert.equal(shared.ok, true);
  assert.deepEqual(sameRealm(shared.pins.map((pin) => pin.id)), ['p1', 'p2', 'p3', 'p4']);
  assert.deepEqual(sameRealm(shared.routeGroups.map((group) => group.routeId)), ['route-a', 'route-b']);
});

test('createShareLink preserves explicit no-route selection and shared view returns pins without route groups', () => {
  const noRoutes = '__share_no_routes__';
  const { api, sheets } = loadApi({
    shareRows: [['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds']],
    routeRows: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PIN_HEADERS,
      ['route-a', 'p1', 0, '', '']
    ],
    mapRows: [
      MAP_INFO_HEADERS,
      pinRow('p1', 'Route A pin', '#e53935', 'alpha'),
      pinRow('p2', 'Orphan pin', '#e53935', 'alpha'),
      pinRow('p3', 'Wrong tag', '#e53935', 'beta'),
      pinRow('p4', 'Wrong color', '#1e88e5', 'alpha')
    ]
  });

  const created = api.createShareLink(withEditToken({
    label: 'No routes',
    tags: ['alpha'],
    colors: ['#e53935'],
    routeIds: [noRoutes]
  }));

  assert.deepEqual(sameRealm(created.shareLink.routeIds), [noRoutes]);
  assert.equal(sheets.share_links.rows[1][8], noRoutes);

  const listed = api.listShareLinks(withEditToken());
  assert.deepEqual(sameRealm(listed.items[0].routeIds), [noRoutes]);

  const shared = api.getSharedViewData('uuid-token');
  assert.equal(shared.ok, true);
  assert.deepEqual(sameRealm(shared.shareLink.routeIds), [noRoutes]);
  assert.deepEqual(sameRealm(shared.allowedRouteIds), []);
  assert.equal(shared.noRoutes, true);
  assert.deepEqual(sameRealm(shared.pins.map((pin) => pin.id)), ['p1', 'p2']);
  assert.deepEqual(sameRealm(shared.routeGroups), []);
});

test('getSharedViewData applies tag and color to pins while selected routes limit route groups only', () => {
  const { api } = loadApi({
    shareRows: [
      ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'],
      ['2026-04-01T00:00:00.000Z', 'Selected routes', 'tok-selected', 'alpha', 'or', true, '', '#e53935', 'route-a|route-b']
    ],
    routeRows: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid'],
      ['route-b', 'Route B', '#43a047', 'straight', false, '', '', '', '', 1, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PIN_HEADERS,
      ['route-a', 'p1', 0, '', ''],
      ['route-a', 'p2', 1, '', ''],
      ['route-b', 'p2', 0, '', ''],
      ['route-b', 'p3', 1, '', ''],
      ['route-b', 'p5', 2, '', '']
    ],
    mapRows: [
      MAP_INFO_HEADERS,
      pinRow('p1', 'Route A matching', '#e53935', 'alpha'),
      pinRow('p2', 'Both routes matching', '#e53935', 'alpha'),
      pinRow('p3', 'Route B matching', '#e53935', 'alpha'),
      pinRow('p4', 'Orphan matching', '#e53935', 'alpha'),
      pinRow('p5', 'Wrong color', '#1e88e5', 'alpha')
    ]
  });

  const shared = api.getSharedViewData('tok-selected');

  assert.equal(shared.ok, true);
  assert.deepEqual(sameRealm(shared.pins.map((pin) => pin.id)), ['p1', 'p2', 'p3', 'p4']);
  assert.deepEqual(sameRealm(shared.routeGroups.map((group) => [group.routeId, group.pinIds])), [
    ['route-a', ['p1', 'p2']],
    ['route-b', ['p2', 'p3']]
  ]);
});

test('getSharedViewData returns ok with empty pins when selected routes have no matching pins', () => {
  const { api } = loadApi({
    shareRows: [
      ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'],
      ['2026-04-01T00:00:00.000Z', 'No matches', 'tok-empty', 'missing-tag', 'or', true, '', '#e53935', 'route-a']
    ],
    routeRows: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'straight', false, '', '', '', '', 0, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PIN_HEADERS,
      ['route-a', 'p1', 0, '', '']
    ],
    mapRows: [
      MAP_INFO_HEADERS,
      pinRow('p1', 'Route A pin', '#e53935', 'alpha')
    ]
  });

  const shared = api.getSharedViewData('tok-empty');

  assert.equal(shared.ok, true);
  assert.deepEqual(sameRealm(shared.pins), []);
  assert.deepEqual(sameRealm(shared.routeGroups), []);
});

test('getSharedRoadRouteCache_ rejects non-selected routes and uses selected route cache', () => {
  const routeACacheKey = 'osrm|road|false|p1:35.00000,139.00000>p2:36.00000,140.00000';
  const routeBCacheKey = 'osrm|road|false|p2:36.00000,140.00000>p3:37.00000,141.00000';
  const { api } = loadApi({
    shareRows: [
      ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'],
      ['2026-04-01T00:00:00.000Z', 'Route A cache', 'tok-cache', 'alpha', 'or', true, '', '#e53935', 'route-a']
    ],
    routeRows: [
      ROUTE_HEADERS,
      ['route-a', 'Route A', '#1e88e5', 'road', false, '', '', '', '', 0, true, true, true, 'solid'],
      ['route-b', 'Route B', '#43a047', 'road', false, '', '', '', '', 1, true, true, true, 'solid']
    ],
    routePinRows: [
      ROUTE_PIN_HEADERS,
      ['route-a', 'p1', 0, '', ''],
      ['route-a', 'p2', 1, '', ''],
      ['route-b', 'p2', 0, '', ''],
      ['route-b', 'p3', 1, '', '']
    ],
    mapRows: [
      MAP_INFO_HEADERS,
      pinRow('p1', 'Route A 1', '#e53935', 'alpha', 35, 139),
      pinRow('p2', 'Route A 2', '#e53935', 'alpha', 36, 140),
      pinRow('p3', 'Route B 2', '#e53935', 'alpha', 37, 141)
    ],
    routeCacheRows: [
      ROUTE_CACHE_HEADERS,
      [routeACacheKey, 'route-a', '[[35,139],[36,140]]', 'osrm', '2026-04-02T00:00:00.000Z', ''],
      [routeBCacheKey, 'route-b', '[[36,140],[37,141]]', 'osrm', '2026-04-02T00:00:00.000Z', '']
    ]
  });

  assert.deepEqual(
    sameRealm(api.getSharedRoadRouteCache_('tok-cache', 'route-a')),
    { ok: true, routeId: 'route-a', coords: [[35, 139], [36, 140]] }
  );
  assert.deepEqual(sameRealm(api.getSharedRoadRouteCache_('tok-cache', 'route-b')), { ok: false });
});

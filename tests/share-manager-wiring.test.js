const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function extractFunction(source, name) {
  const functionStart = source.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `Expected function ${name} to exist`);
  const start = source.slice(Math.max(0, functionStart - 6), functionStart) === 'async '
    ? functionStart - 6
    : functionStart;
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  assert.fail(`Could not parse function ${name}`);
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function previewHarness() {
  const state = {
    shareMode: false,
    shareManager: {
      previewing: true,
      routeTargets: [],
      routeIds: ['same'],
      tags: ['keep'],
      colors: ['#e53935'],
      tagMode: 'or'
    },
    listSort: 'updated-desc',
    routeGroups: [{
      id: 'same', routeId: 'same', name: 'Pin same', color: '#1e88e5',
      visible: true, showNumbers: true, pinIds: ['pin-1']
    }],
    tracks: [
      { trackId: 'same', name: 'GPX same', sourceType: 'gpx', visible: true },
      { trackId: 'same', name: 'GeoJSON same', sourceType: 'geojson', visible: true }
    ],
    pins: [
      { id: 'pin-1', tags: ['keep'], color: '#e53935' },
      { id: 'pin-2', tags: ['other'], color: '#43a047' }
    ]
  };
  let previewUpdates = 0;
  const context = {
    Set,
    Array,
    String,
    Object,
    state,
    getRouteId: (group) => String(group && (group.routeId || group.id) || ''),
    getTrackId: (track) => String(track && (track.trackId || track.id) || ''),
    getRouteDisplayVisible: (group) => group && group.visible !== false,
    getRoutePinIdsForDisplay: (group) => Array.isArray(group && group.pinIds) ? group.pinIds.slice() : [],
    safeColor: (color) => color,
    getBasePins: () => state.pins,
    matchesSearch: () => true,
    matchesStatusFilter: () => true,
    matchesTagFilter(pin, tags) {
      return !tags.length || tags.some((tag) => (pin.tags || []).includes(tag));
    },
    matchesColorFilter(pin, colors) {
      return !colors.length || colors.includes(pin.color);
    },
    matchesIconFilter: () => true,
    sortPins: (pins) => pins,
    updateShareFilterPreview() { previewUpdates += 1; }
  };
  const names = [
    'getShareRouteTargetKey',
    'getShareRouteTargetsForState',
    'getSharePreviewRouteTargetKeySet',
    'isSharePreviewRouteTargetSelected',
    'getShareRouteTargetTypeForKind',
    'getRouteGroupsForLayerRender',
    'getRouteNumberDisplayForPin',
    'getTracksForLayerRender',
    'getUnifiedRouteEntries',
    'getUnifiedRouteEntriesForPanel',
    'getShareRouteTargetOptions',
    'selectAllShareRouteTargets',
    'getActivePinFilterState',
    'filterPinsByState',
    'getFilteredPins'
  ];
  vm.runInNewContext(
    `${names.map((name) => extractFunction(indexHtml, name)).join('\n')}
globalThis.__api = { ${names.join(', ')} };`,
    context
  );
  return { api: context.__api, state, previewUpdates: () => previewUpdates };
}

function selectedKinds(api) {
  return api.getUnifiedRouteEntriesForPanel().map((entry) => `${entry.kind}:${entry.id}`);
}

test('typed preview keeps empty selection route-free while tag and color matching pins remain visible', () => {
  const { api } = previewHarness();

  assert.equal(api.getShareRouteTargetKey('gpx-route', 'same'), 'gpx-route\0same');
  assert.deepEqual(plain(api.getRouteGroupsForLayerRender()), []);
  assert.deepEqual(plain(api.getTracksForLayerRender()), []);
  assert.deepEqual(plain(selectedKinds(api)), []);
  assert.equal(api.getRouteNumberDisplayForPin('pin-1'), null);
  assert.deepEqual(plain(api.getFilteredPins().map((pin) => pin.id)), ['pin-1']);
});

test('all and partial typed selections synchronize pin routes, imported routes, list, and pin numbers', () => {
  const { api, state, previewUpdates } = previewHarness();

  api.selectAllShareRouteTargets();
  assert.deepEqual(plain(api.getShareRouteTargetsForState()), [
    { type: 'pin-route', id: 'same' },
    { type: 'gpx-route', id: 'same' },
    { type: 'geojson-route', id: 'same' }
  ]);
  assert.equal(api.getRouteGroupsForLayerRender().length, 1);
  assert.equal(api.getTracksForLayerRender().length, 2);
  assert.deepEqual(plain(selectedKinds(api)), ['pin:same', 'gpx:same', 'geojson:same']);
  assert.equal(api.getRouteNumberDisplayForPin('pin-1').number, 1);
  assert.equal(previewUpdates(), 1);

  const cases = [
    ['pin-route', 1, [], ['pin:same'], 1],
    ['gpx-route', 0, ['gpx'], ['gpx:same'], null],
    ['geojson-route', 0, ['geojson'], ['geojson:same'], null]
  ];
  cases.forEach(([type, pinRouteCount, trackTypes, entries, pinNumber]) => {
    state.shareManager.routeTargets = [{ type, id: 'same' }];
    assert.equal(api.getRouteGroupsForLayerRender().length, pinRouteCount, type);
    assert.deepEqual(plain(api.getTracksForLayerRender().map((track) => track.sourceType)), trackTypes, type);
    assert.deepEqual(plain(selectedKinds(api)), entries, type);
    const display = api.getRouteNumberDisplayForPin('pin-1');
    assert.equal(display ? display.number : null, pinNumber, type);
  });
});

test('new share opens with no routes and only the all-select action selects every typed target', () => {
  assert.match(indexHtml, /id="share-close"[^>]*data-overlay-initial-focus/);
  const state = {
    listTagFilter: [], listColorFilter: [], listTagMode: 'or',
    shareManager: {
      tags: [], colors: [], routeIds: [], routeTargets: [{ type: 'pin-route', id: 'old' }],
      tagMode: 'or', links: [], loading: false, previewing: false,
      sessionGeneration: 0, listRequestId: 0
    }
  };
  let loads = 0;
  let previewUpdates = 0;
  let trackRenders = 0;
  const openOrder = [];
  const labelInput = { value: 'old', focus() {} };
  const options = [
    { type: 'pin-route', id: 'same' },
    { type: 'gpx-route', id: 'same' },
    { type: 'geojson-route', id: 'same' }
  ];
  const context = {
    Set, Array, String, Object, state,
    document: { getElementById: () => labelInput },
    getShareRouteTargetOptions: () => options,
    renderShareLinks() {}, renderPins() {}, renderSidePanel() {},
    renderTrackLayers() { trackRenders += 1; },
    openOverlay() { openOrder.push('open'); }, setTimeout() {},
    loadShareLinks() { loads += 1; openOrder.push('load'); },
    updateShareFilterPreview() { previewUpdates += 1; }
  };
  const names = [
    'advanceShareManagerSession', 'openShareDialog', 'getShareRouteTargetKey',
    'getShareRouteTargetsForState', 'selectAllShareRouteTargets'
  ];
  vm.runInNewContext(
    `${names.map((name) => extractFunction(indexHtml, name)).join('\n')}
globalThis.__api = { openShareDialog, getShareRouteTargetsForState, selectAllShareRouteTargets };`,
    context
  );

  context.__api.openShareDialog(null);
  assert.deepEqual(plain(state.shareManager.routeTargets), []);
  assert.equal(state.shareManager.sessionGeneration, 1);
  assert.equal(loads, 1);
  assert.deepEqual(openOrder, ['load', 'open']);
  assert.equal(trackRenders, 1);

  context.__api.selectAllShareRouteTargets();
  assert.deepEqual(plain(context.__api.getShareRouteTargetsForState()), options);
  assert.equal(previewUpdates, 1);

  state.shareManager.routeTargets = [{ type: 'gpx-route', id: 'same' }];
  assert.deepEqual(plain(context.__api.getShareRouteTargetsForState()), [
    { type: 'gpx-route', id: 'same' }
  ]);
  assert.match(indexHtml, /const routeTargets = getShareRouteTargetsForState\(\)\.map/);
});

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function staleHarness(responses) {
  const state = {
    listTagFilter: [], listColorFilter: [], listTagMode: 'or',
    shareManager: {
      tags: [], colors: [], routeIds: [], routeTargets: [], tagMode: 'or',
      links: [], loading: false, previewing: false,
      sessionGeneration: 0, listRequestId: 0
    }
  };
  const alerts = [];
  let callIndex = 0;
  const labelInput = { value: '', focus() {} };
  const context = {
    Promise, Array, String, Object, state,
    document: { getElementById: () => labelInput },
    withEditToken: (payload) => payload,
    withGAS(method) {
      assert.equal(method, 'listShareLinks');
      return responses[callIndex++].promise;
    },
    alert(message) { alerts.push(message); },
    renderShareLinks() {}, renderShareFilterUi() {}, renderPins() {}, renderSidePanel() {},
    renderTrackLayers() {}, openOverlay() {}, closeOverlay() {}, setTimeout() {},
    getShareRouteTargetOptions: () => []
  };
  const names = [
    'advanceShareManagerSession',
    'beginShareManagerAsyncRequest',
    'isShareManagerAsyncRequestCurrent',
    'loadShareLinks',
    'openShareDialog',
    'closeShareDialog'
  ];
  vm.runInNewContext(
    `${names.map((name) => extractFunction(indexHtml, name)).join('\n')}
globalThis.__api = { loadShareLinks, openShareDialog, closeShareDialog };`,
    context
  );
  return { api: context.__api, state, alerts };
}

async function drain() {
  await new Promise((resolve) => setImmediate(resolve));
}

test('close and reopen reject the old response and stale finally cannot clear current loading', async () => {
  const oldResponse = deferred();
  const currentResponse = deferred();
  const { api, state } = staleHarness([oldResponse, currentResponse]);

  api.openShareDialog(null);
  api.closeShareDialog();
  api.openShareDialog(null);
  assert.equal(state.shareManager.loading, true);

  oldResponse.resolve({ ok: true, items: [{ token: 'old' }] });
  await drain();
  assert.equal(state.shareManager.loading, true);
  assert.deepEqual(plain(state.shareManager.links), []);

  currentResponse.resolve({ ok: true, items: [{ token: 'current' }] });
  await drain();
  assert.equal(state.shareManager.loading, false);
  assert.deepEqual(plain(state.shareManager.links), [{ token: 'current' }]);
});

test('mutation-style reload request wins when list responses resolve in reverse order', async () => {
  const oldResponse = deferred();
  const reloadResponse = deferred();
  const { api, state } = staleHarness([oldResponse, reloadResponse]);
  state.shareManager.sessionGeneration = 1;

  const oldLoad = api.loadShareLinks();
  const mutationReload = api.loadShareLinks();
  reloadResponse.resolve({ ok: true, items: [{ token: 'after-mutation' }] });
  await mutationReload;
  oldResponse.resolve({ ok: true, items: [{ token: 'before-mutation' }] });
  await oldLoad;

  assert.deepEqual(plain(state.shareManager.links), [{ token: 'after-mutation' }]);
  assert.equal(state.shareManager.loading, false);
});

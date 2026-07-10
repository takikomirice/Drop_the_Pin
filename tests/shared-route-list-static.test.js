const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

function assertIncludes(source, needle, message) {
  assert.ok(source.includes(needle), message || `Expected source to include ${needle}`);
}

function functionBody(name) {
  const match = sharedHtml.match(new RegExp(`function ${name}\\(\\) \\{([\\s\\S]*?)\\n    \\}`));
  assert.ok(match, `Expected function ${name} to exist`);
  return match[1];
}

function sourceFunctionBody(source, name) {
  const index = source.indexOf(`function ${name}`);
  assert.notEqual(index, -1, `Expected function ${name} to exist`);
  const openIndex = source.indexOf('{', index);
  assert.notEqual(openIndex, -1, `Expected function ${name} to have a body`);
  let depth = 0;
  for (let i = openIndex; i < source.length; i += 1) {
    const char = source[i];
    if (char === '{') depth += 1;
    if (char === '}') depth -= 1;
    if (depth === 0) return source.slice(openIndex + 1, i);
  }
  assert.fail(`Could not parse function ${name}`);
}

function loadCodeTestApi(routeCacheRows, extraSource = '') {
  const context = {
    Logger: { log() {} },
    __routeCacheRows: routeCacheRows
  };
  const script = `${codeJs}
openDataSpreadsheet_ = function() {
  return {
    getSheetByName: function(name) {
      if (name !== ROUTE_CACHE_SHEET_NAME) return null;
      return {
        getLastRow: function() { return globalThis.__routeCacheRows.length; },
        getDataRange: function() {
          return {
            getValues: function() { return globalThis.__routeCacheRows; }
          };
        }
      };
    }
  };
};
${extraSource}
globalThis.__routeCacheTestApi = {
  routeCacheRowToEntry_: routeCacheRowToEntry_,
  readLatestRouteCacheEntryByCacheKey_: readLatestRouteCacheEntryByCacheKey_,
  getSharedRoutePinIdsForDisplay_: getSharedRoutePinIdsForDisplay_,
  buildSharedRoadRouteCacheKey_: buildSharedRoadRouteCacheKey_,
  getSharedRoadRouteCache_: getSharedRoadRouteCache_
};`;
  vm.runInNewContext(script, context);
  return context.__routeCacheTestApi;
}

function sameRealm(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadHtmlRouteDisplayFunction(source, name, extraDependencies = '') {
  const script = `${extraDependencies}
${sourceFunctionBody(source, name).trim().startsWith('function ')
    ? sourceFunctionBody(source, name)
    : `function ${name}${source.slice(source.indexOf('(', source.indexOf(`function ${name}`)), source.indexOf('{', source.indexOf(`function ${name}`)))} {${sourceFunctionBody(source, name)}}`}
globalThis.__routeDisplayFn = ${name};`;
  const context = {};
  vm.runInNewContext(script, context);
  return context.__routeDisplayFn;
}

test('shared panel is split into route and pin sections', () => {
  assertIncludes(sharedHtml, '<section id="shared-route-section">');
  assertIncludes(sharedHtml, '<h3 class="shared-section-title">ルート</h3>');
  assertIncludes(sharedHtml, '<div id="shared-route-list"></div>');
  assertIncludes(sharedHtml, '<section id="shared-pin-section">');
  assertIncludes(sharedHtml, '<h3 class="shared-section-title">ピン</h3>');
  assertIncludes(sharedHtml, '<div id="shared-list"></div>');
});

test('route list styles are present without edit controls', () => {
  [
    '.shared-section-title',
    '.shared-route-card',
    '.shared-route-card-header',
    '.shared-route-color-dot',
    '.shared-route-name',
    '.shared-route-display-mode',
    '.shared-route-display-mode.is-off',
    '.shared-route-meta',
    '.shared-route-pin-list',
    '.shared-route-pin-item',
    '.shared-route-pin-num',
    '.shared-route-pin-title'
  ].forEach((selector) => assertIncludes(sharedHtml, selector));

  assert.equal(/shared-route-(?:edit|delete|rename|drag|reorder)/.test(sharedHtml), false);
});

test('route edit settings do not expose start or end pin controls', () => {
  const overlayStart = indexHtml.indexOf('<div id="route-edit-overlay"');
  const overlayEnd = indexHtml.indexOf('<div id="share-overlay"', overlayStart);
  assert.notEqual(overlayStart, -1, 'Expected route edit overlay markup');
  assert.notEqual(overlayEnd, -1, 'Expected next overlay after route edit overlay');
  const overlayMarkup = indexHtml.slice(overlayStart, overlayEnd);

  assert.equal(overlayMarkup.includes('route-edit-endpoints'), false);
  assert.equal(overlayMarkup.includes('始点'), false);
  assert.equal(overlayMarkup.includes('終点'), false);
  assert.equal(indexHtml.includes('function buildRoutePinSelect'), false);

  const renderBody = sourceFunctionBody(indexHtml, 'renderRouteEditOverlay');
  assert.equal(renderBody.includes('route-edit-endpoints'), false);
  assert.equal(renderBody.includes('buildRoutePinSelect'), false);
  assert.equal(renderBody.includes('startPinId'), false);
});

test('route display order follows pinIds and ignores saved start and end pins', () => {
  const normalizeDependency = `function normalizeRoutePinIds(pinIds) {${sourceFunctionBody(indexHtml, 'normalizeRoutePinIds')}}`;
  const getIndexDisplayPinIds = loadHtmlRouteDisplayFunction(
    indexHtml,
    'getRoutePinIdsForDisplay',
    normalizeDependency
  );
  const getSharedDisplayPinIds = loadHtmlRouteDisplayFunction(sharedHtml, 'getSharedRoutePinIdsForDisplay');
  const group = {
    pinIds: [' p2 ', 'p3', 'p1', 'p2'],
    startPinId: 'p1',
    endPinId: 'p3',
    closed: false
  };
  const filtered = new Set(['p1', 'p3']);

  assert.deepEqual(sameRealm(getIndexDisplayPinIds(group)), ['p2', 'p3', 'p1']);
  assert.deepEqual(sameRealm(getIndexDisplayPinIds(group, filtered)), ['p3', 'p1']);
  assert.deepEqual(sameRealm(getSharedDisplayPinIds(group)), ['p2', 'p3', 'p1']);
  assert.deepEqual(sameRealm(getSharedDisplayPinIds(group, filtered)), ['p3', 'p1']);
});

test('route visibility is local state and original group visibility is not mutated', () => {
  assertIncludes(sharedHtml, 'routeVisibility: {}');
  assertIncludes(sharedHtml, 'state.routeVisibility[group.routeId] = group.visible !== false;');
  assertIncludes(sharedHtml, 'function isSharedRouteVisible(group)');
  assertIncludes(sharedHtml, 'state.routeVisibility[group.routeId] !== false');
  assert.equal(/group\.visible\s*=(?!=)/.test(sharedHtml), false);
});

test('phase 2 and phase 3 route visibility checks use local route visibility', () => {
  assertIncludes(sharedHtml, 'if (!isSharedRouteVisible(group) || group.showNumbers === false) continue;');
  assertIncludes(sharedHtml, 'if (!isSharedRouteVisible(group) || group.showLine === false) return;');
});

test('route list helpers and renderer implement expected interactions', () => {
  [
    'function getSharedRouteVisiblePins(group)',
    'function getSharedRouteDistanceMeters(group)',
    'function getSharedRouteMetaText(group, routePins)',
    'function fitSharedMapToRoute(group)',
    'function renderSharedRouteList()',
    'function getSharedRouteDisplayState(routeId)',
    'function cycleSharedRouteDisplayMode(routeId)',
    "routeSection.style.display = 'none'",
    'cycleSharedRouteDisplayMode(routeId);',
    'renderSharedRouteList();',
    'renderSharedPins();',
    'renderSharedMap();',
    'openSharedDetail(pin);',
    'highlightLinkedMarker(pin.id);'
  ].forEach((needle) => assertIncludes(sharedHtml, needle));
});

test('shared route card meta omits state labels and marks hidden distance as straight', () => {
  const body = sourceFunctionBody(sharedHtml, 'getSharedRouteMetaText');

  assertIncludes(body, "routePins.length + '地点・'");
  assertIncludes(body, "getSharedRouteDisplayState(getSharedRouteId(group)) === 'hidden'");
  assertIncludes(body, "'（直線距離）'");
  assert.equal(/道路|直線(?!距離)|非表示/.test(body), false);

  const renderBody = functionBody('renderSharedRouteList');
  assertIncludes(renderBody, 'meta.textContent = getSharedRouteMetaText(group, routePins);');
});

test('shared view initializes and renders the route list', () => {
  assertIncludes(sharedHtml, 'state.routeGroups.forEach(function(group)');
  assertIncludes(sharedHtml, 'renderSharedRouteList();');
});

test('shared route display state and controls are declared without orphan pin controls', () => {
  assert.equal(sharedHtml.includes('showOrphanPins'), false);
  assertIncludes(sharedHtml, 'initialRouteVisibility: {}');
  assertIncludes(sharedHtml, '<div class="shared-route-section-header">');
  assert.equal(sharedHtml.includes('id="shared-orphan-toggle"'), false);
  assert.equal(sharedHtml.includes('ルート外も表示'), false);
  assertIncludes(sharedHtml, '<button id="shared-reset-routes" class="shared-control-btn" type="button" title="ルート表示を初期状態に戻します">ルートも初期化</button>');
});

test('shared route membership never removes pins client-side', () => {
  assertIncludes(sharedHtml, 'function isPinAllowedByRoutes(pin)');
  const body = sourceFunctionBody(sharedHtml, 'isPinAllowedByRoutes');
  assertIncludes(body, 'return true;');
  assert.equal(body.includes('state.routeGroups'), false);
  assert.equal(body.includes('hasSharedRouteRestriction()'), false);
  assert.equal(body.includes('showOrphanPins'), false);
});

test('filtered shared pins use only search, tag, and color filters before sorting', () => {
  const body = functionBody('getFilteredSharedPins');
  assertIncludes(
    body,
    'return matchesSharedSearch(pin) && matchesSharedTags(pin) && matchesSharedColors(pin);'
  );
  assert.equal(body.includes('isPinAllowedByRoutes'), false);
});

test('route state reset restores route display state only', () => {
  const body = functionBody('resetSharedRouteState');
  assertIncludes(sharedHtml, 'function resetSharedRouteState()');
  assertIncludes(sharedHtml, 'state.routeVisibility = Object.assign({}, state.initialRouteVisibility);');
  assertIncludes(sharedHtml, 'state.routeDisplayMode = Object.assign({}, state.initialRouteDisplayMode);');
  assertIncludes(sharedHtml, 'renderSharedRouteList();');
  assertIncludes(sharedHtml, 'renderSharedPins();');
  assertIncludes(sharedHtml, 'renderSharedMap();');
  assertIncludes(sharedHtml, 'renderSharedRouteLayers();');
  assert.equal(body.includes('showOrphanPins'), false);
  assert.equal(/state\.activeTags\s*=/.test(body), false);
  assert.equal(/state\.listQuery\s*=/.test(body), false);
});

test('route list hides route-only controls when empty without orphan toggle state', () => {
  const body = functionBody('renderSharedRouteList');
  assert.equal(body.includes('orphanToggle'), false);
  assertIncludes(sharedHtml, "resetRoutesButton.style.display = 'none';");
});

test('shared route-restricted links select route cards and lines without client-side pin rejection', () => {
  assertIncludes(sharedHtml, 'allowedRouteIds: []');
  assertIncludes(sharedHtml, "const SHARED_ROUTE_NONE_SENTINEL = '__share_no_routes__';");
  assertIncludes(sharedHtml, 'function isSharedRouteSelectionNone(routeIds)');
  assertIncludes(sharedHtml, 'function hasSharedRouteRestriction()');
  assertIncludes(sharedHtml, 'if (isSharedRouteSelectionNone(state.allowedRouteIds)) return false;');
  assertIncludes(sharedHtml, 'state.allowedRouteIds = normalizeSharedAllowedRouteIds(result.allowedRouteIds);');
  assert.equal(sharedHtml.includes('if (!groups || !groups.length) return !hasSharedRouteRestriction();'), false);
  assert.equal(sharedHtml.includes('if (hasSharedRouteRestriction()) return false;'), false);
  assert.equal(sharedHtml.includes('orphanToggle'), false);
});

test('route visibility changes rerender pins and map without orphan toggle listener', () => {
  assertIncludes(sharedHtml, 'state.initialRouteVisibility = Object.assign({}, state.routeVisibility);');
  assert.equal(sharedHtml.includes('state.showOrphanPins = !state.showOrphanPins;'), false);
  assert.equal(sharedHtml.includes("document.getElementById('shared-orphan-toggle').addEventListener('click', function()"), false);
  assertIncludes(sharedHtml, "document.getElementById('shared-reset-routes').addEventListener('click', function()");
});

test('Code.js exposes a shared road route cache reader without cache writes', () => {
  assertIncludes(codeJs, 'function getSharedRoadRouteCache(data, routeId)');
  assertIncludes(codeJs, 'function getSharedRoadRouteCache_(token, routeId)');
  assertIncludes(codeJs, 'getSharedRoadRouteCache: getSharedRoadRouteCache');

  const body = sourceFunctionBody(codeJs, 'getSharedRoadRouteCache_');
  assertIncludes(body, 'getShareLinkByToken_(token)');
  assertIncludes(body, '!shareLink.enabled || shareLink.revokedAt');
  assertIncludes(body, 'if (!routeId) return { ok: false };');
  assertIncludes(body, 'if (!isShareRouteIdAllowed_(routeId, shareLink.routeIds)) return { ok: false };');
  assertIncludes(body, 'allowedPinIdSet');
  assertIncludes(body, 'const allRouteGroups = getRouteGroups();');
  assertIncludes(body, 'getSharedPinsForShareLink_(shareLink, allRouteGroups)');
  assertIncludes(body, 'rawGroup');
  assertIncludes(body, '!isRouteClosedToAllowedPins_(rawGroup, allowedPinIdSet)');
  assertIncludes(body, 'getSharedRouteGroups_(sharedPins, allRouteGroups)');
  assertIncludes(body, "group.routeMode !== 'road'");
  assertIncludes(body, 'const expectedCacheKey = buildSharedRoadRouteCacheKey_(group, pinById, SHARED_ROAD_ROUTE_CACHE_PROVIDER);');
  assertIncludes(body, 'readLatestRouteCacheEntryByCacheKey_(expectedCacheKey)');
  assertIncludes(body, 'entry.routeId !== routeId');
  assertIncludes(body, 'return { ok: true, routeId: routeId, coords: entry.coords };');
  assert.equal(body.includes('readLatestRouteCacheEntryForRoute_(routeId)'), false);
  assert.equal(/putRouteCache|appendRow|setValues|deleteRow|setValue/.test(body), false);
});

test('Code.js validates route_cache rows against the shared route cache key', () => {
  assertIncludes(codeJs, "const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];");
  assertIncludes(codeJs, "const SHARED_ROAD_ROUTE_CACHE_PROVIDER = 'osrm';");
  assertIncludes(codeJs, 'function getRouteCacheSheetForRead_()');
  assertIncludes(codeJs, 'return getRequiredSheet_(ROUTE_CACHE_SHEET_NAME);');
  assertIncludes(codeJs, 'function readLatestRouteCacheEntryByCacheKey_(cacheKey)');
  assertIncludes(codeJs, 'function buildSharedRoadRouteCacheKey_(group, pinById, provider)');
  assertIncludes(codeJs, 'roundSharedRouteCacheCoord_(entry.latLng[0])');
  assertIncludes(codeJs, 'encodeURIComponent(entry.pinId)');

  const rowBody = sourceFunctionBody(codeJs, 'routeCacheRowToEntry_');
  assertIncludes(rowBody, 'const normalizedCoords = normalizeRouteCacheCoords_(coords);');
  assertIncludes(rowBody, 'if (normalizedCoords.length < 2) return null;');

  const readBody = sourceFunctionBody(codeJs, 'readLatestRouteCacheEntryByCacheKey_');
  assertIncludes(readBody, 'normalizeRouteCacheKey_(rows[i][0]) !== normalizedCacheKey');
  assert.equal(readBody.includes('normalizeRouteId_(rows[i][1])'), false);
});

test('Code.js shared route display and road cache key follow pinIds order', () => {
  const api = loadCodeTestApi([]);
  const group = {
    routeId: 'route-1',
    routeMode: 'road',
    closed: false,
    startPinId: 'p1',
    endPinId: 'p3',
    pinIds: ['p2', 'p3', 'p1', 'p2']
  };
  const pinById = {
    p1: { id: 'p1', lat: 1.1, lng: 2.2 },
    p2: { id: 'p2', lat: 3.3, lng: 4.4 },
    p3: { id: 'p3', lat: 5.5, lng: 6.6 }
  };

  assert.deepEqual(sameRealm(api.getSharedRoutePinIdsForDisplay_(group)), ['p2', 'p3', 'p1']);
  assert.equal(
    api.buildSharedRoadRouteCacheKey_(group, pinById, 'osrm'),
    'osrm|road|false|p2:3.30000,4.40000>p3:5.50000,6.60000>p1:1.10000,2.20000'
  );
});

test('Code.js reads the latest valid route_cache entry by exact cacheKey', () => {
  const cacheKey = 'osrm|road|false|p1:1.10000,2.20000>p2:3.30000,4.40000';
  const api = loadCodeTestApi([
    ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'],
    ['other-key', 'route-latest-by-id', '[[9,9],[8,8]]', 'osrm', '2026-04-03T00:00:00.000Z', ''],
    [cacheKey, 'route-old', '[[1,2],[3,4]]', 'osrm', '2026-04-01T00:00:00.000Z', ''],
    [cacheKey, 'route-new', '[{"lat":5,"lng":6},[7,8],["bad",9]]', 'osrm', '2026-04-02T00:00:00.000Z', '']
  ]);

  const entry = api.readLatestRouteCacheEntryByCacheKey_(cacheKey);
  assert.equal(entry.routeId, 'route-new');
  assert.deepEqual(sameRealm(entry.coords), [[5, 6], [7, 8]]);
  assert.equal(api.routeCacheRowToEntry_([cacheKey, 'route-invalid', '[[1,2]]', 'osrm', '2026-04-04T00:00:00.000Z', '']), null);
});

test('getSharedRoadRouteCache_ uses expected cacheKey instead of latest routeId row', () => {
  const expectedCacheKey = 'osrm|road|false|p1:1.10000,2.20000>p2:3.30000,4.40000';
  const unrelatedLatestKey = 'osrm|road|false|p1:9.00000,9.00000>p2:8.00000,8.00000';
  const api = loadCodeTestApi([
    ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'],
    [unrelatedLatestKey, 'route-1', '[[90,90],[80,80]]', 'osrm', '2026-04-03T00:00:00.000Z', ''],
    [expectedCacheKey, 'route-1', '[[10,20],[30,40]]', 'osrm', '2026-04-01T00:00:00.000Z', '']
  ], `
const __sharedPins = [
  { id: 'p1', lat: 1.1, lng: 2.2 },
  { id: 'p2', lat: 3.3, lng: 4.4 }
];
const __sharedGroup = {
  routeId: 'route-1',
  routeMode: 'road',
  closed: false,
  startPinId: 'p1',
  endPinId: 'p2',
  pinIds: ['p1', 'p2']
};
getShareLinkByToken_ = function(token) {
  return token === 'share-token' ? { enabled: true, revokedAt: '' } : null;
};
getSharedPinsForShareLink_ = function() { return __sharedPins; };
getRouteGroups = function() { return [__sharedGroup]; };
getSharedRouteGroups_ = function() { return [__sharedGroup]; };
`);

  const result = api.getSharedRoadRouteCache_('share-token', 'route-1');
  assert.deepEqual(sameRealm(result), { ok: true, routeId: 'route-1', coords: [[10, 20], [30, 40]] });
});

test('shared.html draws cached road routes through GAS only and falls back to straight lines', () => {
  assertIncludes(sharedHtml, 'roadRouteCache: {}');
  assertIncludes(sharedHtml, 'let sharedRouteRenderVersion = 0;');
  assertIncludes(sharedHtml, 'function drawSharedRoadRouteLine(layerGroup, latLngs, group)');
  assertIncludes(sharedHtml, 'weight: 4');
  assertIncludes(sharedHtml, 'opacity: 0.86');
  assertIncludes(sharedHtml, "lineCap: 'round'");
  assertIncludes(sharedHtml, "lineJoin: 'round'");
  assertIncludes(sharedHtml, 'interactive: false');
  assertIncludes(sharedHtml, "attachSharedRouteHitPolyline(layerGroup, latLngs, group, '道路');");
  assertIncludes(sharedHtml, 'drawSharedStraightRouteLine(layerGroup, latLngs, group);');
  assertIncludes(sharedHtml, "withGAS('getSharedRoadRouteCache', { token: state.token, routeId: routeId })");
  assertIncludes(sharedHtml, 'if (version !== sharedRouteRenderVersion) return;');
  assertIncludes(sharedHtml, 'state.roadRouteCache[routeId] = { ok: true, coords: coords };');
});

test('shared.html keeps route popup method labels explicit and avoids road API calls', () => {
  assertIncludes(sharedHtml, 'function buildSharedRouteInfoPopupHtml(group, latLngs, methodLabel)');
  assertIncludes(sharedHtml, "const routeMethodLabel = String(methodLabel || '直線');");
  assertIncludes(sharedHtml, "attachSharedRouteHitPolyline(layerGroup, latLngs, group, '直線');");
  assertIncludes(sharedHtml, "attachSharedRouteHitPolyline(layerGroup, latLngs, group, '道路');");
  assert.equal(/OSRM|router\.project-osrm|OpenRouteService|Directions API/.test(sharedHtml), false);
});

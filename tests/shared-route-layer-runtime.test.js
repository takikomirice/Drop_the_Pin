const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');

function functionSource(name) {
  const start = sharedHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = sharedHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < sharedHtml.length; index += 1) {
    if (sharedHtml[index] === '{') depth += 1;
    if (sharedHtml[index] === '}') depth -= 1;
    if (depth === 0) return sharedHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse function ${name}`);
}

function trackRoute(type, id, segments = [[
  { lat: 35.0, lng: 139.0 },
  { lat: 35.1, lng: 139.1 }
]]) {
  return {
    id,
    type,
    name: `${type}-${id}`,
    color: '#1e88e5',
    lineWidth: 4,
    lineStyle: 'solid',
    distanceMeters: 1200,
    segments
  };
}

function pinRoute(id) {
  return { routeId: id, color: '#43a047', showLine: true, routeMode: 'straight' };
}

function createHarness({
  routeGroups = [],
  trackRoutes = [],
  hiddenTrackKeys = [],
  roadRequest = () => Promise.resolve(null)
} = {}) {
  const polylineCalls = [];
  const pinDraws = [];
  const roadDraws = [];
  const removedLayers = [];
  const sharedRouteLayers = Object.create(null);
  const mapLayers = new Set();
  let clearCalls = 0;

  function createLayerGroup(key) {
    return {
      key,
      lines: [],
      clearCount: 0,
      clearLayers() {
        this.clearCount += 1;
        this.lines.length = 0;
      }
    };
  }

  const sharedMap = {
    hasLayer(layer) { return mapLayers.has(layer); },
    removeLayer(layer) {
      removedLayers.push(layer.key);
      mapLayers.delete(layer);
    }
  };

  const context = {
    sharedRouteRenderVersion: 0,
    sharedRouteLayers,
    sharedMap,
    getSharedRouteGroups: () => routeGroups.slice(),
    getSharedTrackRoutes: () => trackRoutes.slice(),
    clearSharedRouteLayers() {
      clearCalls += 1;
      Object.keys(sharedRouteLayers).forEach((key) => {
        const layer = sharedRouteLayers[key];
        layer.clearLayers();
        if (sharedMap.hasLayer(layer)) sharedMap.removeLayer(layer);
        delete sharedRouteLayers[key];
      });
    },
    getSharedRouteId: (group) => String(group && group.routeId || ''),
    getSharedTypedRouteKey(route) {
      return (route.type === 'gpx-route' ? 'gpx:' : 'geojson:') + route.id;
    },
    getSharedRouteLayerGroup(key) {
      if (!sharedRouteLayers[key]) sharedRouteLayers[key] = createLayerGroup(key);
      mapLayers.add(sharedRouteLayers[key]);
      return sharedRouteLayers[key];
    },
    isSharedRouteVisible: (group) => group.visible !== false,
    getSharedRouteLatLngsForDisplay: () => [[35.0, 139.0], [35.1, 139.1]],
    getSharedRouteEffectiveMode: (group) => group.routeMode === 'road' ? 'road' : 'straight',
    drawSharedStraightRouteLine(layerGroup, latLngs, group) {
      pinDraws.push(group.routeId);
      layerGroup.lines.push({ kind: 'pin', latLngs });
    },
    drawSharedRoadRouteLine(layerGroup, latLngs, group) {
      roadDraws.push(group.routeId);
      layerGroup.lines.push({ kind: 'road', latLngs });
    },
    getSharedCachedRoadRouteCoords: () => null,
    requestSharedRoadRouteCache: roadRequest,
    renderSharedRouteList() {},
    state: { roadRouteCache: {} },
    isSharedTypedRouteVisible(route) {
      return !hiddenTrackKeys.includes(context.getSharedTypedRouteKey(route));
    },
    safeRouteColor: (color) => color,
    getSharedRouteLineDashArray: () => null,
    escHtml: (value) => String(value),
    formatSharedDistance: (value) => `${value}m`,
    console: { warn() {} },
    L: {
      polyline(latLngs, options) {
        const line = {
          latLngs,
          options,
          popup: '',
          addTo(layerGroup) {
            layerGroup.lines.push(this);
            return this;
          },
          bindPopup(markup) {
            this.popup = markup;
            return this;
          }
        };
        polylineCalls.push(line);
        return line;
      }
    }
  };

  const render = vm.runInNewContext(`(${functionSource('renderSharedRouteLayers')})`, context);

  return {
    render,
    context,
    polylineCalls,
    pinDraws,
    roadDraws,
    removedLayers,
    sharedRouteLayers,
    seedStaleLayer(key = 'stale') {
      const layer = createLayerGroup(key);
      layer.lines.push({ kind: 'stale' });
      sharedRouteLayers[key] = layer;
      mapLayers.add(layer);
      return layer;
    },
    clearCalls: () => clearCalls
  };
}

test('track-only GPX renders a polyline into its route layer', () => {
  const route = trackRoute('gpx-route', 'gpx-1');
  route.lineWidth = 9;
  const harness = createHarness({ trackRoutes: [route] });

  harness.render();

  assert.equal(harness.polylineCalls.length, 1);
  assert.equal(harness.polylineCalls[0].options.weight, 4);
  assert.equal(harness.sharedRouteLayers['gpx:gpx-1'].lines.length, 1);
});

test('track-only GeoJSON renders a polyline', () => {
  const harness = createHarness({ trackRoutes: [trackRoute('geojson-route', 'geo-1')] });

  harness.render();

  assert.equal(harness.polylineCalls.length, 1);
  assert.equal(harness.sharedRouteLayers['geojson:geo-1'].lines.length, 1);
});

test('mixed pin and GPX routes render independently and remove inactive layers', () => {
  const harness = createHarness({
    routeGroups: [pinRoute('pin-1')],
    trackRoutes: [trackRoute('gpx-route', 'gpx-1')]
  });
  const staleLayer = harness.seedStaleLayer();

  harness.render();

  assert.deepEqual(harness.pinDraws, ['pin-1']);
  assert.equal(harness.polylineCalls.length, 1);
  assert.equal(harness.sharedRouteLayers['pin-1'].lines.length, 1);
  assert.equal(harness.sharedRouteLayers['gpx:gpx-1'].lines.length, 1);
  assert.equal(staleLayer.clearCount, 1);
  assert.deepEqual(harness.removedLayers, ['stale']);
  assert.equal(harness.sharedRouteLayers.stale, undefined);
});

test('no routes clears all route layers without creating a polyline', () => {
  const harness = createHarness();
  const staleLayer = harness.seedStaleLayer();

  harness.render();

  assert.equal(harness.clearCalls(), 1);
  assert.equal(staleLayer.clearCount, 1);
  assert.equal(harness.polylineCalls.length, 0);
  assert.equal(Object.keys(harness.sharedRouteLayers).length, 0);
});

test('hidden tracks and segments shorter than two points do not render polylines', () => {
  const harness = createHarness({
    trackRoutes: [
      trackRoute('gpx-route', 'hidden'),
      trackRoute('geojson-route', 'short', [[{ lat: 35.0, lng: 139.0 }]])
    ],
    hiddenTrackKeys: ['gpx:hidden']
  });

  harness.render();

  assert.equal(harness.polylineCalls.length, 0);
});

test('a pending road result cannot draw after the route becomes hidden and rendering advances', async () => {
  const gate = (() => {
    let resolve;
    const promise = new Promise((resolvePromise) => { resolve = resolvePromise; });
    return { promise, resolve };
  })();
  const group = { ...pinRoute('road-1'), routeMode: 'road', visible: true };
  const harness = createHarness({
    routeGroups: [group],
    roadRequest: () => gate.promise
  });

  harness.render();
  assert.deepEqual(harness.pinDraws, ['road-1']);
  group.visible = false;
  harness.render();
  gate.resolve([[35, 139], [35.1, 139.1]]);
  await gate.promise;
  await Promise.resolve();
  await Promise.resolve();

  assert.deepEqual(harness.roadDraws, []);
  assert.equal(harness.context.state.roadRouteCache['road-1'], undefined);
});

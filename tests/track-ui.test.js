const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

function assertIncludes(source, needle) {
  assert.ok(source.includes(needle), `Expected source to include ${needle}`);
}

function sourceFunctionBody(source, name) {
  const index = source.indexOf(`function ${name}(`);
  assert.notEqual(index, -1, `Expected function ${name}`);
  const openIndex = source.indexOf('{', index);
  let depth = 0;
  for (let i = openIndex; i < source.length; i += 1) {
    if (source[i] === '{') depth += 1;
    if (source[i] === '}') depth -= 1;
    if (depth === 0) return source.slice(openIndex + 1, i);
  }
  assert.fail(`Could not parse ${name}`);
}

function loadTrackRendering() {
  const start = indexHtml.indexOf('function getTrackId(');
  const end = indexHtml.indexOf('\n    function createClientRouteId(', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const polylineCalls = [];
  const groups = [];
  const map = {
    hasLayer(layer) { return layer.added === true; },
    setViewCalls: [], fitBoundsCalls: [],
    setView(latLng, zoom) { this.setViewCalls.push([latLng, zoom]); },
    fitBounds(bounds, options) { this.fitBoundsCalls.push([bounds, options]); }
  };
  const L = {
    layerGroup() {
      const group = { layers: [], added: false, removed: false, clearCount: 0,
        addTo() { this.added = true; return this; },
        clearLayers() { this.layers = []; this.clearCount += 1; },
        remove() { this.removed = true; }
      };
      groups.push(group);
      return group;
    },
    polyline(latLngs, options) {
      const line = { latLngs, options, addTo(group) { group.layers.push(this); return this; } };
      polylineCalls.push(line);
      return line;
    },
    latLngBounds(points) { return { points }; }
  };
  const context = {
    state: { shareMode: false, tracks: [], trackLayers: {}, trackVisibilityOverrides: {} }, map, L,
    getShareRouteTargetTypeForKind(kind) {
      return kind === 'gpx' ? 'gpx-route' : (kind === 'geojson' ? 'geojson-route' : '');
    },
    isSharePreviewRouteTargetSelected() { return true; },
    document: {
      createElement() { throw new Error('DOM should not be used while rendering map layers'); },
      getElementById() { return null; }
    },
    console
  };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nglobalThis.__api = { renderTrackLayers, fitTrackBounds, getTrackDisplayVisible, toggleTrackVisibility };`, context);
  return { api: context.__api, state: context.state, map, polylineCalls, groups };
}

test('track state, separate panel and read-only load path are present without shared view exposure', () => {
  assertIncludes(indexHtml, 'tracks: [],');
  assertIncludes(indexHtml, 'trackLayers: Object.create(null),');
  assertIncludes(indexHtml, 'trackVisibilityOverrides: Object.create(null),');
  assertIncludes(indexHtml, 'id="side-tracks"');
  assertIncludes(indexHtml, 'id="track-list"');
  assertIncludes(indexHtml, "withGASNoArg('getTracks')");
  assertIncludes(indexHtml, "console.error('TRACKS_LOAD_FAILED')");
  assert.equal(sharedHtml.includes('getTracks'), false);
  assert.equal(sharedHtml.includes('trackLayers'), false);
  assert.equal(sharedHtml.includes('side-tracks'), false);
});

test('Leaflet renders each track segment separately with lat-lng order and style options', () => {
  const { api, state, polylineCalls, groups } = loadTrackRendering();
  state.tracks = [{
    id: 'track-a', trackId: 'track-a', color: '#2196f3', visible: true, lineStyle: 'dotted', lineWidth: 6,
    segments: [
      { index: 0, points: [{ lat: 35, lng: 139 }, { lat: 35.1, lng: 139.1 }] },
      { index: 1, points: [{ lat: 36, lng: 140 }] }
    ]
  }];

  api.renderTrackLayers();

  assert.equal(polylineCalls.length, 2);
  assert.deepEqual(JSON.parse(JSON.stringify(polylineCalls[0].latLngs)), [[35, 139], [35.1, 139.1]]);
  assert.equal(polylineCalls[0].options.color, '#2196f3');
  assert.equal(polylineCalls[0].options.weight, 4);
  assert.equal(polylineCalls[0].options.dashArray, '1 8');
  assert.equal(polylineCalls[0].options.lineCap, 'round');
  assert.equal(polylineCalls[0].options.lineJoin, 'round');
  assert.equal(polylineCalls[0].options.pane, 'trackPane');
  assert.equal(groups[0].layers.length, 2);

  state.trackVisibilityOverrides['track-a'] = false;
  api.renderTrackLayers();
  assert.equal(groups[0].layers.length, 0);
  assert.equal(Object.prototype.hasOwnProperty.call(state, 'routeLayers'), false);
});

test('track layer rerender removes stale layers and fit uses setView for one point or fitBounds for many', () => {
  const { api, state, map, groups } = loadTrackRendering();
  const one = { trackId: 'one', visible: true, segments: [{ points: [{ lat: 35, lng: 139 }] }] };
  const many = { trackId: 'many', visible: true, segments: [{ points: [{ lat: 35, lng: 139 }, { lat: 36, lng: 140 }] }] };
  state.tracks = [one];
  api.renderTrackLayers();
  state.tracks = [many];
  api.renderTrackLayers();
  assert.equal(groups[0].removed, true);

  api.fitTrackBounds(one);
  api.fitTrackBounds(many);
  assert.deepEqual(JSON.parse(JSON.stringify(map.setViewCalls[0])), [[35, 139], 16]);
  assert.equal(map.fitBoundsCalls.length, 1);
  assert.equal(map.fitBoundsCalls[0][1].maxZoom, 16);
});

test('saved hidden state can be locally overridden and dashed or solid styles stay distinct', () => {
  const { api, state, polylineCalls } = loadTrackRendering();
  const track = {
    trackId: 'hidden', visible: false, color: '#2196f3', lineStyle: 'dashed', lineWidth: 4,
    segments: [{ points: [{ lat: 35, lng: 139 }, { lat: 36, lng: 140 }] }]
  };
  state.tracks = [track];
  api.renderTrackLayers();
  assert.equal(polylineCalls.length, 0);
  state.trackVisibilityOverrides.hidden = true;
  api.renderTrackLayers();
  assert.equal(polylineCalls[0].options.dashArray, '10 8');
  track.lineStyle = 'solid';
  api.renderTrackLayers();
  assert.equal(polylineCalls[1].options.dashArray, undefined);
});

test('toggling one track does not rebuild unrelated high-volume track layers', () => {
  const { api, state, groups, polylineCalls } = loadTrackRendering();
  state.tracks = ['a', 'b'].map((trackId, index) => ({
    trackId, visible: true, color: '#2196f3', lineStyle: 'solid', lineWidth: 4,
    segments: [{ points: [{ lat: 35 + index, lng: 139 }, { lat: 35.1 + index, lng: 139.1 }] }]
  }));
  api.renderTrackLayers();
  const unrelatedClearCount = groups[1].clearCount;
  const lineCount = polylineCalls.length;
  api.toggleTrackVisibility('a');
  assert.equal(groups[0].clearCount, 2);
  assert.equal(groups[1].clearCount, unrelatedClearCount);
  assert.equal(polylineCalls.length, lineCount);
});

test('track list uses textContent and exposes visibility and fit while unified cards own edit-only deletion', () => {
  const body = sourceFunctionBody(indexHtml, 'buildTrackItem');
  assertIncludes(body, 'name.textContent');
  assertIncludes(body, 'meta.textContent');
  assertIncludes(body, 'source.textContent');
  assertIncludes(body, 'toggleTrackVisibility');
  assertIncludes(body, 'fitTrackBounds');
  assert.equal(body.includes('deleteTrack'), false);
  assert.equal(body.includes('editTrack'), false);
  const unified = sourceFunctionBody(indexHtml, 'buildUnifiedTrackRouteItem');
  assertIncludes(unified, 'canEditRouteControls()');
  assertIncludes(unified, 'deleteTrackFromUi');
  assertIncludes(unified, '削除');
  assert.equal(sharedHtml.includes('deleteTrackFromUi'), false);
  assertIncludes(sourceFunctionBody(indexHtml, 'renderTrackPanel'), 'トラックはまだありません。');
});

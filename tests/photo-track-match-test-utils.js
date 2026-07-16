const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function loadCore() {
  const start = indexHtml.indexOf('const TrackGeometryCore =');
  const end = indexHtml.indexOf('\n    const GeoJsonTrackInterchangeCore =', start);
  assert.notEqual(start, -1, 'Expected TrackGeometryCore');
  assert.notEqual(end, -1, 'Expected track core boundary');
  const context = {
    PIN_COLORS: [{ hex: '#2196f3' }],
    Date,
    Math
  };
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__core = typeof PhotoTrackMatchCore === "undefined" ? null : PhotoTrackMatchCore;',
    context
  );
  assert.ok(context.__core, 'Expected PhotoTrackMatchCore');
  return context.__core;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function options(overrides = {}) {
  return { utcOffsetMinutes: 0, ...overrides };
}

function point(lat, lng, time, elevation = null) {
  return { lat, lng, elevation, time };
}

function track(segments, overrides = {}) {
  return {
    trackId: 'track-1',
    revisionId: 'rev-1',
    name: 'Track',
    description: '',
    color: '#2196f3',
    sourceType: 'gpx',
    sourceName: 'track.gpx',
    segments: segments.map((points) => ({ points })),
    ...overrides
  };
}

function photo(capturedAt, overrides = {}) {
  return { id: 'item-1', capturedAt, lat: null, lng: null, ...overrides };
}

module.exports = { loadCore, options, photo, plain, point, track };

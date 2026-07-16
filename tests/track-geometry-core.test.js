const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function loadCore() {
  const start = indexHtml.indexOf('const TrackGeometryCore =');
  assert.notEqual(start, -1, 'TrackGeometryCore should exist');
  const end = indexHtml.indexOf('\n    const GeoJsonTrackInterchangeCore =', start);
  assert.notEqual(end, -1, 'TrackGeometryCore should be declared before state');
  const context = { PIN_COLORS: [
    '#e53935', '#e91e63', '#9c27b0', '#3f51b5', '#2196f3', '#00bcd4', '#009688',
    '#4caf50', '#8bc34a', '#ffeb3b', '#ff9800', '#ff5722', '#795548', '#607d8b', '#212121'
  ].map((hex) => ({ hex })) };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nglobalThis.__core = TrackGeometryCore;`, context);
  return context.__core;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

test('TrackGeometryCore normalizes without mutating input and computes segment-aware summary', () => {
  const core = loadCore();
  const input = {
    id: 'track-a',
    trackId: 'track-a',
    revisionId: 'rev-a',
    name: ' Ridge ',
    description: '',
    color: '#2196f3',
    sourceType: 'gpx',
    sourceName: 'ridge.gpx',
    segments: [
      { index: 7, points: [
        { lat: 35, lng: 139, elevation: 100, time: '2026-07-11T01:00:00+09:00', ignored: true },
        { lat: 35.001, lng: 139.001, elevation: null, time: '' }
      ] },
      { index: 8, points: [
        { lat: 36, lng: 140, elevation: 300, time: '2026-07-11T00:30:00.000Z' }
      ] }
    ],
    arbitrary: 'drop-me'
  };
  const before = JSON.stringify(input);

  const track = core.normalizeTrack(input);

  assert.equal(JSON.stringify(input), before);
  assert.equal(track.id, 'track-a');
  assert.equal(track.trackId, 'track-a');
  assert.equal(track.name, 'Ridge');
  assert.equal(track.segmentCount, 2);
  assert.equal(track.pointCount, 3);
  assert.ok(track.distanceMeters > 140 && track.distanceMeters < 150, 'only adjacent points in the first segment contribute');
  assert.equal(track.minElevation, 100);
  assert.equal(track.maxElevation, 300);
  assert.equal(track.startTime, '2026-07-10T16:00:00.000Z');
  assert.equal(track.endTime, '2026-07-11T00:30:00.000Z');
  assert.deepEqual(plain(track.bounds), { south: 35, west: 139, north: 36, east: 140 });
  assert.deepEqual(Object.keys(track.segments[0].points[0]).sort(), ['elevation', 'lat', 'lng', 'time']);
  assert.equal(Object.prototype.hasOwnProperty.call(track, 'arbitrary'), false);
});

test('TrackGeometryCore rejects numeric strings, invalid coordinates, elevation and RFC3339 times', () => {
  const core = loadCore();
  const base = { trackId: 't', revisionId: 'r', name: 'T', color: '#2196f3', sourceType: 'manual', segments: [] };
  const invalidPoints = [
    { lat: '35', lng: 139, elevation: null, time: '' },
    { lat: 91, lng: 139, elevation: null, time: '' },
    { lat: 35, lng: -181, elevation: null, time: '' },
    { lat: 35, lng: 139, elevation: Infinity, time: '' },
    { lat: 35, lng: 139, elevation: '', time: '' },
    { lat: 35, lng: 139, elevation: null, time: null },
    { lat: 35, lng: 139, elevation: null },
    { lat: 35, lng: 139, elevation: null, time: '2026-02-30T00:00:00Z' },
    { lat: 35, lng: 139, elevation: null, time: '2026-07-11T00:00:00' },
    { lat: 35, lng: 139, elevation: null, time: ' 2026-07-11T00:00:00Z ' }
  ];
  invalidPoints.forEach((point) => {
    assert.throws(() => core.normalizeTrack(Object.assign({}, base, { segments: [{ points: [point] }] })));
  });
});

test('TrackGeometryCore keeps RFC3339 calendar, millisecond and offset boundaries aligned with storage', () => {
  const core = loadCore();
  assert.equal(core.normalizeTime('0001-01-01T00:00:00Z'), '0001-01-01T00:00:00.000Z');
  assert.equal(core.normalizeTime('2000-02-29T23:59:59.123456789+14:00'), '2000-02-29T09:59:59.123Z');
  assert.equal(core.normalizeTime('2000-02-29T00:00:00.1-14:00'), '2000-02-29T14:00:00.100Z');
  [
    '0000-01-01T00:00:00Z',
    '1900-02-29T00:00:00Z',
    '2000-02-29T00:00:00.1234567890Z',
    '2000-02-29T00:00:00+14:01',
    '2000-02-29T00:00:00+23:59',
    '2000-02-29T00:00:00+24:00'
  ].forEach((value) => assert.throws(() => core.normalizeTime(value)));

  const negativeZero = core.normalizeTrack({
    trackId: 'track-zero', revisionId: 'rev-zero', name: 'Zero', color: '#2196f3',
    sourceType: 'geojson', segments: [{ points: [
      { lat: -0, lng: -0, elevation: -0, time: '' }
    ] }]
  }).segments[0].points[0];
  assert.equal(Object.is(negativeZero.lat, -0), true);
  assert.equal(Object.is(negativeZero.lng, -0), true);
  assert.equal(Object.is(negativeZero.elevation, -0), true);
});

test('TrackGeometryCore enforces server identifier and source metadata boundaries', () => {
  const core = loadCore();
  const base = { trackId: 't', id: 't', revisionId: 'r', name: 'T', description: '', color: '#2196f3', sourceType: 'manual', sourceName: '', segments: [] };
  [
    Object.assign({}, base, { trackId: 1, id: 1 }),
    Object.assign({}, base, { trackId: '=formula', id: '=formula' }),
    Object.assign({}, base, { revisionId: 'r'.repeat(129) }),
    Object.assign({}, base, { sourceName: 'C:/track.gpx' }),
    Object.assign({}, base, { trackId: 'a', id: 'b' }),
    Object.assign({}, base, { color: ['#2196f3'] }),
    Object.assign({}, base, { sourceType: ['manual'] }),
    Object.assign({}, base, { lineStyle: ['solid'] })
  ].forEach((input) => assert.throws(() => core.normalizeTrack(input)));
});

test('TrackGeometryCore clone and persistable projections deep-copy only contract properties', () => {
  const core = loadCore();
  const normalized = core.normalizeTrack({
    trackId: 'track-a', revisionId: 'rev-a', name: 'Track', description: 'D', color: '#2196f3',
    sourceType: 'geojson', sourceName: 'track.geojson', visible: false, lineStyle: 'dashed', lineWidth: 7,
    orderIndex: 3, segments: [{ points: [{ lat: 35, lng: 139, elevation: null, time: '' }] }]
  });
  const cloned = core.cloneTrack(normalized);
  const persistable = core.toPersistableTrack(Object.assign({ extra: true }, normalized));

  assert.equal(normalized.lineWidth, 4);
  assert.equal(persistable.lineWidth, 4);

  cloned.segments[0].points[0].lat = 0;
  assert.equal(normalized.segments[0].points[0].lat, 35);
  assert.equal(persistable.id, undefined);
  assert.equal(persistable.extra, undefined);
  assert.equal(persistable.createdAt, undefined);
  assert.equal(persistable.pointCount, undefined);
  assert.equal(persistable.trackId, 'track-a');
  assert.equal(persistable.revisionId, 'rev-a');
  assert.deepEqual(plain(persistable.segments), [{ index: 0, points: [{ lat: 35, lng: 139, elevation: null, time: '' }] }]);
});

test('TrackGeometryCore handles all-null elevation, all-empty time and empty geometry', () => {
  const core = loadCore();
  const summary = core.computeSummary([
    { index: 0, points: [{ lat: -90, lng: -180, elevation: null, time: '' }, { lat: 90, lng: 180, elevation: null, time: '' }] }
  ]);
  assert.equal(summary.minElevation, null);
  assert.equal(summary.maxElevation, null);
  assert.equal(summary.startTime, '');
  assert.equal(summary.endTime, '');
  assert.deepEqual(plain(summary.bounds), { south: -90, west: -180, north: 90, east: 180 });
  assert.deepEqual(plain(core.computeSummary([]).bounds), null);
});

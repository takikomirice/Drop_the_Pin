const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadCore() {
  const start = indexHtml.indexOf('const TrackGeometryCore =');
  const end = indexHtml.indexOf('\n    const GeoJsonTrackInterchangeCore =', start);
  assert.notEqual(start, -1, 'Expected TrackGeometryCore');
  assert.notEqual(end, -1, 'Expected GeoJSON core boundary');
  const context = {
    PIN_COLORS: [{ hex: '#e53935' }],
    Date,
    Math,
    Set
  };
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__trackBatch = typeof TrackBatchTransformCore === "undefined" '
      + '? null : TrackBatchTransformCore;',
    context
  );
  assert.ok(context.__trackBatch, 'Expected TrackBatchTransformCore');
  return context.__trackBatch;
}

function point(index, time = '') {
  return {
    lat: 35 + index / 1000,
    lng: 139,
    elevation: index,
    time
  };
}

test('shared transform chooses the smallest time interval and preserves segment endpoints', () => {
  const core = loadCore();
  const start = Date.parse('2026-01-01T00:00:00Z');
  const points = Array.from({ length: 7 }, (_, index) => (
    point(index, new Date(start + index * 1000).toISOString())
  ));
  const result = core.transform([{ index: 0, points }], {
    maxPoints: 4,
    interruptionMs: 4 * 60 * 60 * 1000,
    compressionIntervals: [1, 2, 3, 4, 5]
  });
  assert.deepEqual(plain(result.parts[0][0].points.map((value) => value.time)), [
    '2026-01-01T00:00:00.000Z',
    '2026-01-01T00:00:02.000Z',
    '2026-01-01T00:00:04.000Z',
    '2026-01-01T00:00:06.000Z'
  ]);
  assert.deepEqual(plain(result.compressionIntervals), [2]);
  assert.equal(result.timeCompressedPointCount, 3);
  assert.equal(result.overflowCompressedPointCount, 0);
  assert.equal(result.compressedPointCount, 3);
});

test('shared transform splits only at the configured interruption and partitions overflow in order', () => {
  const start = Date.parse('2026-01-01T23:59:59Z');
  const times = [
    start,
    start + 2000,
    start + 4 * 60 * 60 * 1000 + 2000
  ];
  const result = loadCore().transform([{
    index: 8,
    points: times.map((time, index) => point(index, new Date(time).toISOString()))
  }], {
    maxPoints: 1,
    interruptionMs: 4 * 60 * 60 * 1000,
    compressionIntervals: []
  });
  assert.equal(result.interruptionCount, 1);
  assert.deepEqual(
    plain(result.parts.map((part) => part.flatMap((segment) => (
      segment.points.map((value) => value.lat)
    )))),
    [[35], [35.001], [35.002]]
  );
});

test('shared transform accounts for a valid overflow reducer without mutating the input', () => {
  const segments = [{ index: 4, points: [point(0), point(1), point(2)] }];
  const originalPoints = segments[0].points;
  const result = loadCore().transform(segments, {
    maxPoints: 2,
    interruptionMs: 1000,
    compressionIntervals: [],
    reduceOverflow(value) {
      return [{
        index: 0,
        points: [value[0].points[0], value[0].points[2]]
      }];
    }
  });
  assert.equal(result.overflowCompressedPointCount, 1);
  assert.equal(result.compressedPointCount, 1);
  assert.deepEqual(plain(result.parts[0][0].points.map((value) => value.lat)), [35, 35.002]);
  assert.equal(segments[0].points, originalPoints);
  assert.equal(segments[0].index, 4);
});

test('shared part names preserve a bounded suffix', () => {
  const core = loadCore();
  assert.equal(core.partName('walk', 0, 1), 'walk');
  const name = core.partName('x'.repeat(100), 0, 2);
  assert.equal(name.length, 100);
  assert.match(name, /\(1\/2\)$/);
});

test('shared summary aggregation combines numeric bounds elevation and time fields', () => {
  const summary = loadCore().aggregateSummaries([
    { summary: {
      segmentCount: 1,
      pointCount: 2,
      distanceMeters: 12.5,
      minElevation: 3,
      maxElevation: 8,
      startTime: '2026-01-01T00:00:00.000Z',
      endTime: '2026-01-01T00:00:02.000Z',
      bounds: { south: 34, west: 138, north: 35, east: 139 }
    } },
    { summary: {
      segmentCount: 2,
      pointCount: 4,
      distanceMeters: 7.5,
      minElevation: -2,
      maxElevation: 6,
      startTime: '',
      endTime: '2026-01-01T00:00:05.000Z',
      bounds: { south: 33, west: 137, north: 36, east: 140 }
    } }
  ]);
  assert.deepEqual(plain(summary), {
    segmentCount: 3,
    pointCount: 6,
    distanceMeters: 20,
    minElevation: -2,
    maxElevation: 8,
    startTime: '2026-01-01T00:00:00.000Z',
    endTime: '2026-01-01T00:00:05.000Z',
    bounds: { south: 33, west: 137, north: 36, east: 140 }
  });
});

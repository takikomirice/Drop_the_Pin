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
  const context = { PIN_COLORS: [{ hex: '#e53935' }], Date, Math, Set };
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__shape = typeof TrackShapeSimplificationCore === "undefined" '
      + '? null : TrackShapeSimplificationCore;',
    context
  );
  assert.ok(context.__shape, 'Expected TrackShapeSimplificationCore');
  return context.__shape;
}

function point(lat, lng, options = {}) {
  return {
    lat,
    lng,
    elevation: options.elevation ?? null,
    time: options.time || ''
  };
}

test('shape reduction removes collinear density before sharp corners', () => {
  const result = loadCore().reduce([{
    index: 0,
    points: [
      point(0, 0), point(0, 0.25), point(0, 0.5),
      point(1, 0.5), point(1, 0.75), point(1, 1)
    ]
  }], 4);
  assert.deepEqual(plain(result[0].points.map((value) => [value.lat, value.lng])), [
    [0, 0], [0, 0.5], [1, 0.5], [1, 1]
  ]);
});

test('shape reduction protects endpoints timed points and global elevation extrema', () => {
  const timed = '2026-01-01T00:00:00.000Z';
  const points = [
    point(0, 0),
    point(0, 1, { elevation: 5 }),
    point(0, 2, { elevation: -10 }),
    point(0, 3, { time: timed }),
    point(0, 4, { elevation: 50 }),
    point(0, 5)
  ];
  const result = loadCore().reduce([{ index: 0, points }], 5);
  assert.deepEqual(plain(result[0].points), plain([
    points[0], points[2], points[3], points[4], points[5]
  ]));
});

test('shape reduction preserves segment and point order with deterministic ties', () => {
  const segments = [
    { index: 9, points: [point(0, 0), point(0, 1), point(0, 2)] },
    { index: 4, points: [point(1, 0), point(1, 1), point(1, 2)] }
  ];
  const core = loadCore();
  const first = core.reduce(segments, 5);
  const second = core.reduce(segments, 5);
  assert.deepEqual(plain(first), plain(second));
  assert.deepEqual(plain(first.map((segment) => segment.index)), [0, 1]);
  assert.deepEqual(
    plain(first.flatMap((segment) => segment.points.map((value) => value.lat))),
    [0, 0, 1, 1, 1]
  );
});

test('shape reduction returns every protected point when protection exceeds the target', () => {
  const points = [
    point(0, 0, { time: '2026-01-01T00:00:00.000Z' }),
    point(0, 1, { time: '2026-01-01T00:00:01.000Z' }),
    point(0, 2, { time: '2026-01-01T00:00:02.000Z' })
  ];
  assert.deepEqual(
    plain(loadCore().reduce([{ index: 0, points }], 2)[0].points),
    plain(points)
  );
});

test('shape reduction treats an antimeridian crossing as a local move', () => {
  const values = [
    point(0, 179.8),
    point(0.1, 179.9),
    point(0.2, -180),
    point(1, -180),
    point(1, -179.9)
  ];
  const result = loadCore().reduce([{ index: 0, points: values }], 4);
  assert.deepEqual(
    plain(result[0].points.map((value) => [value.lat, value.lng])),
    [[0, 179.8], [0.2, -180], [1, -180], [1, -179.9]]
  );
});

test('shape reduction is immutable and returns the exact requested count when possible', () => {
  const points = [
    point(0, 0), point(0, 1), point(0, 2), point(1, 2), point(1, 3)
  ];
  const segments = [{ index: 7, points }];
  const result = loadCore().reduce(segments, 3);
  assert.equal(result[0].points.length, 3);
  assert.notEqual(result, segments);
  assert.notEqual(result[0], segments[0]);
  assert.notEqual(result[0].points, points);
  assert.equal(segments[0].index, 7);
  assert.equal(segments[0].points.length, 5);
});

test('shape reduction safely clones input for invalid targets', () => {
  const segments = [{ index: 3, points: [point(0, 0), point(0, 1)] }];
  [0, -1, 1.5, '1'].forEach((value) => {
    const result = loadCore().reduce(segments, value);
    assert.deepEqual(plain(result), [{ index: 0, points: plain(segments[0].points) }]);
    assert.notEqual(result[0].points, segments[0].points);
  });
});

const assert = require('node:assert/strict');
const test = require('node:test');
const { loadCore, options, photo, plain, point, track } = require('./photo-track-match-test-utils');

test('timeline indexes exact points edges gaps anomalies overlap stats and safe warnings immutably', () => {
  const core = loadCore();
  const input = track([
    [
      point(35, 139, '2026-07-11T00:00:00Z', 10),
      point(35.1, 139.1, '', null),
      point(35.2, 139.2, '2026-07-11T00:00:00Z', 20),
      point(35.3, 139.3, '2026-07-10T23:59:00Z', 30),
      point(35.4, 139.4, '2026-07-11T00:20:00Z', 40)
    ],
    [point(36, 140, '2026-07-11T00:10:00Z'), point(36.1, 140.1, '2026-07-11T00:11:00Z')]
  ]);
  const before = JSON.stringify(input);
  const timeline = core.buildTrackTimeline(input, options());
  assert.equal(JSON.stringify(input), before);
  assert.deepEqual(plain(timeline.stats), {
    segmentCount: 2,
    pointCount: 7,
    timedPointCount: 6,
    untimedPointCount: 1,
    exactTimestampCount: 5,
    interpolationEdgeCount: 1,
    largeGapEdgeCount: 1,
    duplicateTimestampEdgeCount: 0,
    nonMonotonicEdgeCount: 1,
    firstTime: '2026-07-10T23:59:00.000Z',
    lastTime: '2026-07-11T00:20:00.000Z'
  });
  assert.deepEqual(plain(timeline.warnings), [
    { code: 'TRACK_MATCH_POINTS_WITHOUT_TIME', count: 1 },
    { code: 'TRACK_MATCH_LARGE_GAPS', count: 1 },
    { code: 'TRACK_MATCH_NON_MONOTONIC_EDGES', count: 1 },
    { code: 'TRACK_MATCH_OVERLAPPING_SEGMENTS', count: 1 }
  ]);
});

test('exact match returns safe point fields and merges only identical duplicates', () => {
  const core = loadCore();
  const identical = track([[
    point(35, 139, '2026-07-11T00:00:00Z', 100),
    point(35, 139, '2026-07-11T00:00:00Z', 100)
  ]]);
  const result = core.matchPhotos(identical, [photo('2026-07-11T00:00', { filename: 'secret.jpg', previewUrl: 'blob:secret' })], options());
  assert.deepEqual(plain(result.results[0]), {
    itemId: 'item-1', status: 'matched-exact', trackId: 'track-1', revisionId: 'rev-1',
    photoTimeUtc: '2026-07-11T00:00:00.000Z', lat: 35, lng: 139, elevation: 100,
    segmentIndex: 0, pointIndex: 0, timeDeltaSeconds: 0
  });
  assert.equal(JSON.stringify(result).includes('secret'), false);

  const different = track([[point(35, 139, '2026-07-11T00:00:00Z'), point(36, 140, '2026-07-11T00:00:00Z')]]);
  assert.equal(core.matchPhotos(different, [photo('2026-07-11T00:00')], options()).results[0].status, 'ambiguous');
});

test('interpolation returns ratios coordinates elevation and preserves point order', () => {
  const core = loadCore();
  const input = track([[
    point(35, 139, '2026-07-11T00:00:00Z', 100),
    point(39, 143, '2026-07-11T00:04:00Z', 200)
  ]]);
  const result = core.matchPhotos(input, [
    photo('2026-07-11T00:01', { id: 'quarter' }),
    photo('2026-07-11T00:02', { id: 'half' }),
    photo('2026-07-11T00:03', { id: 'three-quarter' })
  ], options());
  assert.deepEqual(result.results.map((entry) => entry.ratio), [0.25, 0.5, 0.75]);
  assert.deepEqual(plain(result.results[1]), {
    itemId: 'half', status: 'matched-interpolated', trackId: 'track-1', revisionId: 'rev-1',
    photoTimeUtc: '2026-07-11T00:02:00.000Z', lat: 37, lng: 141, elevation: 150,
    segmentIndex: 0, fromPointIndex: 0, toPointIndex: 1, ratio: 0.5, gapSeconds: 240,
    timeDeltaSeconds: 0
  });
  assert.equal(JSON.stringify(input), JSON.stringify(track([[
    point(35, 139, '2026-07-11T00:00:00Z', 100), point(39, 143, '2026-07-11T00:04:00Z', 200)
  ]])));
});

test('interpolation does not bridge segment boundaries and nulls elevation unless both endpoints have it', () => {
  const core = loadCore();
  const separated = track([
    [point(35, 139, '2026-07-11T00:00:00Z', 100)],
    [point(36, 140, '2026-07-11T00:02:00Z', 200)]
  ]);
  assert.equal(core.matchPhotos(separated, [photo('2026-07-11T00:01')], options()).results[0].status, 'ambiguous');
  const oneNull = track([[point(35, 139, '2026-07-11T00:00:00Z', 100), point(36, 140, '2026-07-11T00:02:00Z', null)]]);
  assert.equal(core.matchPhotos(oneNull, [photo('2026-07-11T00:01')], options()).results[0].elevation, null);
});

test('longitude interpolation follows the shortest antimeridian path deterministically', () => {
  const core = loadCore();
  const match = (from, to) => core.matchPhotos(
    track([[point(0, from, '2026-07-11T00:00:00Z'), point(0, to, '2026-07-11T00:02:00Z')]]),
    [photo('2026-07-11T00:01')], options()
  ).results[0].lng;
  assert.equal(match(179, -179), 180);
  assert.equal(match(-179, 179), -180);
  assert.equal(match(170, -10), 80);
  assert.ok(match(179.9, -179.9) >= -180 && match(179.9, -179.9) <= 180);
});

test('large gaps include no coordinates and multiple simultaneous blocked gaps are ambiguous', () => {
  const core = loadCore();
  const oneGap = core.matchPhotos(track([[
    point(35, 139, '2026-07-11T00:00:00Z'), point(36, 140, '2026-07-11T00:05:01Z')
  ]]), [photo('2026-07-11T00:02:00')], options()).results[0];
  assert.deepEqual(plain(oneGap), {
    itemId: 'item-1', status: 'gap-too-large', trackId: 'track-1', revisionId: 'rev-1',
    photoTimeUtc: '2026-07-11T00:02:00.000Z', gapSeconds: 301,
    segmentIndex: 0, fromPointIndex: 0, toPointIndex: 1
  });
  assert.equal('lat' in oneGap, false);
  const twoGaps = track([
    [point(35, 139, '2026-07-11T00:00:00Z'), point(36, 140, '2026-07-11T00:10:00Z')],
    [point(37, 141, '2026-07-11T00:01:00Z'), point(38, 142, '2026-07-11T00:11:00Z')]
  ]);
  assert.equal(core.matchPhotos(twoGaps, [photo('2026-07-11T00:05')], options()).results[0].status, 'ambiguous');
});

test('300 second edge interpolates while non-monotonic and duplicate edges never interpolate', () => {
  const core = loadCore();
  const boundary = track([[point(0, 0, '2026-07-11T00:00:00Z'), point(1, 1, '2026-07-11T00:05:00Z')]]);
  assert.equal(core.matchPhotos(boundary, [photo('2026-07-11T00:02')], options()).results[0].status, 'matched-interpolated');
  const duplicate = track([[point(0, 0, '2026-07-11T00:00:00Z'), point(0, 0, '2026-07-11T00:00:00Z')]]);
  const timeline = core.buildTrackTimeline(duplicate, options());
  assert.equal(timeline.stats.duplicateTimestampEdgeCount, 1);
  assert.deepEqual(plain(timeline.warnings), [{ code: 'TRACK_MATCH_DUPLICATE_TIMESTAMPS', count: 1 }]);
});

test('overlapping segment interpolation candidates are ambiguous', () => {
  const core = loadCore();
  const input = track([
    [point(0, 0, '2026-07-11T00:00:00Z'), point(1, 1, '2026-07-11T00:02:00Z')],
    [point(10, 10, '2026-07-11T00:00:30Z'), point(11, 11, '2026-07-11T00:02:30Z')]
  ]);
  assert.equal(core.matchPhotos(input, [photo('2026-07-11T00:01')], options()).results[0].status, 'ambiguous');
});

test('an exact point and an interpolation candidate in overlapping segments are ambiguous', () => {
  const core = loadCore();
  const input = track([
    [point(0, 0, '2026-07-11T00:01:00Z')],
    [point(10, 10, '2026-07-11T00:00:00Z'), point(11, 11, '2026-07-11T00:02:00Z')]
  ]);
  assert.equal(core.matchPhotos(input, [photo('2026-07-11T00:01')], options()).results[0].status, 'ambiguous');
});

test('outside range reports nearest side while endpoint tolerance can match unique endpoints', () => {
  const core = loadCore();
  const input = track([[point(35, 139, '2026-07-11T00:00:00Z'), point(36, 140, '2026-07-11T00:01:00Z')]]);
  const before = core.matchPhotos(input, [photo('2026-07-10T23:59:58')], options()).results[0];
  assert.equal(before.status, 'outside-track-range');
  assert.equal(before.side, 'before');
  assert.equal(before.nearestTimeDeltaSeconds, 2);
  const after = core.matchPhotos(input, [photo('2026-07-11T00:01:03')], options()).results[0];
  assert.equal(after.side, 'after');
  const endpoint = core.matchPhotos(input, [photo('2026-07-10T23:59:58')], options({ endpointToleranceSeconds: 2 })).results[0];
  assert.equal(endpoint.status, 'matched-endpoint');
  assert.equal(endpoint.lat, 35);
  assert.equal(endpoint.timeDeltaSeconds, 2);
});

test('equidistant endpoints at different positions are ambiguous', () => {
  const core = loadCore();
  const input = track([
    [point(35, 139, '2026-07-11T00:00:00Z')],
    [point(36, 140, '2026-07-11T00:00:00Z')]
  ]);
  assert.equal(core.matchPhotos(input, [photo('2026-07-10T23:59:59')], options({ endpointToleranceSeconds: 1 })).results[0].status, 'ambiguous');
});

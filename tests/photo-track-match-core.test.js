const assert = require('node:assert/strict');
const test = require('node:test');
const { loadCore, options, photo, plain, point, track } = require('./photo-track-match-test-utils');

test('matchPhotos reports fixed options counts timeline stats and warnings from result statuses', () => {
  const core = loadCore();
  const input = track([[point(35, 139, '2026-07-11T00:00:00Z'), point(36, 140, '2026-07-11T00:01:00Z')]]);
  const result = core.matchPhotos(input, [
    photo('2026-07-11T00:00', { id: 'exact' }),
    photo('2026-07-11T00:00:30', { id: 'interpolated' }),
    photo('2026-07-10T23:59:59', { id: 'endpoint' }),
    photo('2026-07-11T00:00', { id: 'gps', lat: 1, lng: 2 }),
    photo('', { id: 'missing' }),
    photo('invalid', { id: 'invalid-time' }),
    photo('2026-07-11T00:02', { id: 'outside' }),
    photo('2026-07-11T00:00', { id: '', lat: null, lng: null })
  ], options({ endpointToleranceSeconds: 1 }));
  assert.deepEqual(plain(result.options), {
    utcOffsetMinutes: 0, clockCorrectionSeconds: 0,
    maxInterpolationGapSeconds: 300, endpointToleranceSeconds: 1
  });
  assert.deepEqual(plain(result.counts), {
    total: 8, matched: 3, exact: 1, interpolated: 1, endpoint: 1,
    skippedExistingGps: 1, missingTime: 1, invalidTime: 1, outsideRange: 1,
    gapTooLarge: 0, ambiguous: 0, invalidInput: 1
  });
  assert.equal(result.trackId, 'track-1');
  assert.equal(result.timelineStats.timedPointCount, 2);
  assert.deepEqual(plain(result.warnings), []);
});

test('matchPhotos permits zero and twenty photos and rejects twenty-one without truncation', () => {
  const core = loadCore();
  const input = track([[point(35, 139, '2026-07-11T00:00:00Z')]]);
  assert.equal(core.matchPhotos(input, [], options()).results.length, 0);
  assert.equal(core.matchPhotos(input, Array.from({ length: 20 }, (_, index) => photo('2026-07-11T00:00', { id: `item-${index}` })), options()).results.length, 20);
  assert.throws(() => core.matchPhotos(input, Array.from({ length: 21 }, (_, index) => photo('2026-07-11T00:00', { id: `item-${index}` })), options()));
});

test('toItemPatches returns matched entries only in order with a deep safe whitelist', () => {
  const core = loadCore();
  const batch = core.matchPhotos(
    track([[point(35, 139, '2026-07-11T00:00:00Z'), point(36, 140, '2026-07-11T00:02:00Z')]]),
    [photo('2026-07-11T00:01', { id: 'matched', metadataStatus: 'no-gps', runtime: { file: {} }, previewUrl: 'blob:x' }), photo('', { id: 'skip' })],
    options()
  );
  const before = JSON.stringify(batch);
  const patches = core.toItemPatches(batch);
  assert.deepEqual(plain(patches), [{
    itemId: 'matched',
    patch: { lat: 35.5, lng: 139.5 },
    match: {
      status: 'matched-interpolated', trackId: 'track-1', revisionId: 'rev-1',
      photoTimeUtc: '2026-07-11T00:01:00.000Z', segmentIndex: 0,
      fromPointIndex: 0, toPointIndex: 1, ratio: 0.5, gapSeconds: 120, timeDeltaSeconds: 0
    }
  }]);
  assert.equal(JSON.stringify(batch), before);
  patches[0].patch.lat = 0;
  patches[0].match.status = 'changed';
  assert.equal(batch.results[0].lat, 35.5);
  assert.equal(batch.results[0].status, 'matched-interpolated');
  assert.equal(JSON.stringify(patches).includes('metadataStatus'), false);
  assert.equal(JSON.stringify(patches).includes('runtime'), false);
  assert.equal(JSON.stringify(patches).includes('previewUrl'), false);
});

test('batch results remain isolated across photos and do not mutate photos or options', () => {
  const core = loadCore();
  const inputPhotos = [photo('2026-07-11T00:00', { id: 'one' }), photo('2026-07-11T00:00', { id: 'two' })];
  const inputOptions = options();
  const beforePhotos = JSON.stringify(inputPhotos);
  const beforeOptions = JSON.stringify(inputOptions);
  const result = core.matchPhotos(track([[point(35, 139, '2026-07-11T00:00:00Z')]]), inputPhotos, inputOptions);
  result.results[0].lat = 0;
  assert.equal(result.results[1].lat, 35);
  assert.equal(JSON.stringify(inputPhotos), beforePhotos);
  assert.equal(JSON.stringify(inputOptions), beforeOptions);
});

test('safe outputs never spread arbitrary photo or track metadata', () => {
  const core = loadCore();
  const inputTrack = track([[point(35, 139, '2026-07-11T00:00:00Z')]], {
    title: 'track-secret', xml: '<gpx/>', arbitrary: { token: 'secret' }
  });
  const inputPhoto = photo('2026-07-11T00:00', {
    filename: 'photo-secret.jpg', title: 'photo-secret', previewUrl: 'blob:secret', File: { secret: true }
  });
  const serialized = JSON.stringify(core.matchPhotos(inputTrack, [inputPhoto], options()));
  ['track-secret', '<gpx/>', 'token', 'photo-secret', 'blob:secret', 'File'].forEach((value) => assert.equal(serialized.includes(value), false));
});

const assert = require('node:assert/strict');
const test = require('node:test');
const { loadCore, options, photo, plain, point, track } = require('./photo-track-match-test-utils');

test('normalizeOptions requires a finite integer fixed UTC offset and normalizes defaults', () => {
  const core = loadCore();
  assert.deepEqual(plain(core.normalizeOptions({ utcOffsetMinutes: 540 })), {
    utcOffsetMinutes: 540,
    clockCorrectionSeconds: 0,
    maxInterpolationGapSeconds: 300,
    endpointToleranceSeconds: 0
  });
  [-840, 840, -480].forEach((value) => assert.equal(core.normalizeOptions({ utcOffsetMinutes: value }).utcOffsetMinutes, value));
  [undefined, '540', NaN, Infinity, -841, 841, 1.5].forEach((value) => {
    const input = value === undefined ? {} : { utcOffsetMinutes: value };
    assert.throws(() => core.normalizeOptions(input));
  });
});

test('normalizeOptions validates clock correction gap and endpoint boundaries without mutating input', () => {
  const core = loadCore();
  const input = { utcOffsetMinutes: 0, clockCorrectionSeconds: 300, maxInterpolationGapSeconds: 1, endpointToleranceSeconds: 3600 };
  const before = JSON.stringify(input);
  assert.deepEqual(plain(core.normalizeOptions(input)), input);
  assert.equal(JSON.stringify(input), before);
  [-86400, 86400, -120, 0].forEach((value) => assert.equal(core.normalizeOptions(options({ clockCorrectionSeconds: value })).clockCorrectionSeconds, value));
  [86401, -86401, '1', NaN, Infinity, 1.5].forEach((value) => assert.throws(() => core.normalizeOptions(options({ clockCorrectionSeconds: value }))));
  [0, 86401, '300', 1.5].forEach((value) => assert.throws(() => core.normalizeOptions(options({ maxInterpolationGapSeconds: value }))));
  [-1, 3601, '0', 1.5].forEach((value) => assert.throws(() => core.normalizeOptions(options({ endpointToleranceSeconds: value }))));
  assert.equal(core.normalizeOptions(options({ maxInterpolationGapSeconds: 86400 })).maxInterpolationGapSeconds, 86400);
  assert.equal(core.normalizeOptions(options({ endpointToleranceSeconds: 0 })).endpointToleranceSeconds, 0);
});

test('parsePhotoWallTime accepts minute second and one to three fractional digits', () => {
  const core = loadCore();
  const cases = [
    ['2026-07-11T10:00', '2026-07-11T01:00:00.000Z'],
    ['2026-07-11T10:00:01', '2026-07-11T01:00:01.000Z'],
    ['2026-07-11T10:00:01.1', '2026-07-11T01:00:01.100Z'],
    ['2026-07-11T10:00:01.12', '2026-07-11T01:00:01.120Z'],
    ['2026-07-11T10:00:01.123', '2026-07-11T01:00:01.123Z'],
    ['2000-02-29T00:00', '2000-02-28T15:00:00.000Z']
  ];
  cases.forEach(([value, expected]) => assert.equal(core.parsePhotoWallTime(value, options({ utcOffsetMinutes: 540 })), expected));
  assert.equal(core.parsePhotoWallTime('2026-07-11T10:00', options({ utcOffsetMinutes: -480 })), '2026-07-11T18:00:00.000Z');
});

test('parsePhotoWallTime applies camera correction after UTC conversion with documented sign', () => {
  const core = loadCore();
  assert.equal(core.parsePhotoWallTime('2026-07-11T10:00', options({ utcOffsetMinutes: 540, clockCorrectionSeconds: 300 })), '2026-07-11T01:05:00.000Z');
  assert.equal(core.parsePhotoWallTime('2026-07-11T10:00', options({ utcOffsetMinutes: 540, clockCorrectionSeconds: -120 })), '2026-07-11T00:58:00.000Z');
});

test('parsePhotoWallTime rejects ambiguous timezone calendar and whitespace forms', () => {
  const core = loadCore();
  [
    '', '2026-07-11', '2026/07/11 10:00', '2026:07:11 10:00',
    '2026-07-11T10:00Z', '2026-07-11T10:00+09:00', '2026-07-11T10:00:00.1234',
    '0000-01-01T00:00', '1900-02-29T00:00', '2026-02-29T00:00', '2026-04-31T00:00',
    '2026-07-11T24:00', '2026-07-11T10:60', '2026-07-11T10:00:60',
    ' 2026-07-11T10:00', '2026-07-11T10:00 '
  ].forEach((value) => assert.throws(() => core.parsePhotoWallTime(value, options())));
  assert.throws(() => core.parsePhotoWallTime(20260711, options()));
});

test('photo input classifies existing GPS missing time invalid time and invalid shape safely', () => {
  const core = loadCore();
  const timeline = core.buildTrackTimeline(track([[point(35, 139, '2026-07-11T00:00:00Z')]]), options());
  assert.equal(core.matchPhoto(timeline, photo('invalid', { lat: 35, lng: 139 }), options()).status, 'skipped-existing-gps');
  assert.equal(core.matchPhoto(timeline, photo(''), options()).status, 'photo-time-missing');
  assert.equal(core.matchPhoto(timeline, photo('invalid'), options()).status, 'photo-time-invalid');
  [
    photo('2026-07-11T00:00', { lat: 35, lng: null }),
    photo('2026-07-11T00:00', { lat: '35', lng: '139' }),
    photo('2026-07-11T00:00', { lat: 91, lng: 139 }),
    photo('2026-07-11T00:00', { id: '' }),
    photo('2026-07-11T00:00', { id: 'x'.repeat(129) }),
    photo('2026-07-11T00:00', { id: '__proto__' })
  ].forEach((value) => assert.equal(core.matchPhoto(timeline, value, options()).status, 'photo-input-invalid'));
  const inherited = Object.create({ id: 'inherited', capturedAt: '2026-07-11T00:00' });
  inherited.lat = null;
  inherited.lng = null;
  assert.equal(core.matchPhoto(timeline, inherited, options()).status, 'photo-input-invalid');
});

test('track without timed points returns track-time-unavailable', () => {
  const core = loadCore();
  const timeline = core.buildTrackTimeline(track([[point(35, 139, '')]]), options());
  assert.equal(core.matchPhoto(timeline, photo('2026-07-11T00:00'), options()).status, 'track-time-unavailable');
});

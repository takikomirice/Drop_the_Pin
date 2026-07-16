const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');

function loadClientNormalizer() {
  const start = indexHtml.indexOf('const TrackGeometryCore =');
  const end = indexHtml.indexOf('\n    const GeoJsonTrackInterchangeCore =', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = { Date, Math, PIN_COLORS: [{ hex: '#e53935' }] };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nglobalThis.__normalize = TrackGeometryCore.normalizeTime;`, context);
  return context.__normalize;
}

function loadServerNormalizer() {
  const start = codeJs.indexOf('function normalizeTrackTime_(');
  const end = codeJs.indexOf('\nfunction normalizeTrackPoint_(', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = {
    Date,
    trackStorageError_: (code, message, retryable) => Object.assign(new Error(message), { code, retryable })
  };
  vm.runInNewContext(`${codeJs.slice(start, end)}\nglobalThis.__normalize = normalizeTrackTime_;`, context);
  return context.__normalize;
}

function normalizers() {
  return [loadClientNormalizer(), loadServerNormalizer()];
}

test('client and server normalize the accepted GPX time contract identically', () => {
  const vectors = [
    ['2026-07-11T01:02:03Z', '2026-07-11T01:02:03.000Z'],
    ['2026-07-11T01:02:03.1Z', '2026-07-11T01:02:03.100Z'],
    ['2026-07-11T01:02:03.12Z', '2026-07-11T01:02:03.120Z'],
    ['2026-07-11T01:02:03.123Z', '2026-07-11T01:02:03.123Z'],
    ['2026-07-11T01:02:03.123456789Z', '2026-07-11T01:02:03.123Z'],
    ['2026-07-11T01:02:03+09:00', '2026-07-10T16:02:03.000Z'],
    ['2026-07-11T01:02:03.987654321-14:00', '2026-07-11T15:02:03.987Z'],
    ['2026-07-11T01:02:03+14:00', '2026-07-10T11:02:03.000Z'],
    ['0001-01-01T00:00:00Z', '0001-01-01T00:00:00.000Z'],
    ['2000-02-29T23:59:59.999999999Z', '2000-02-29T23:59:59.999Z']
  ];
  normalizers().forEach((normalize) => {
    vectors.forEach(([input, expected]) => assert.equal(normalize(input), expected, input));
  });
});

test('client and server reject the same invalid GPX time vectors', () => {
  const invalid = [
    '2026-07-11T01:02:03',
    '2026-07-11T01:02:03.Z',
    '2026-07-11T01:02:03.1234567890Z',
    '2026-07-11T01:02:60Z',
    '2026-07-11T24:00:00Z',
    '2026-07-11T01:02:03+14:01',
    '2026-07-11T01:02:03-14:01',
    '2026-07-11T01:02:03+15:00',
    '2026-07-11T01:02:03+23:59',
    '0000-01-01T00:00:00Z',
    '1900-02-29T00:00:00Z',
    '2026-02-30T00:00:00Z',
    ' 2026-07-11T01:02:03Z',
    '2026-07-11T01:02:03Zx'
  ];
  normalizers().forEach((normalize) => {
    invalid.forEach((input) => assert.throws(() => normalize(input), undefined, input));
  });
});

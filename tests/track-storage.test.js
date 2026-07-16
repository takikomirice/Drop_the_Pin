const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const TEST_EDIT_TOKEN = 'track-edit-token';
const TRACKS_HEADERS = [
  'trackId', 'name', 'description', 'color', 'sourceType', 'sourceName', 'activeRevision',
  'payloadHash', 'segmentCount', 'pointCount', 'distanceMeters', 'minElevation', 'maxElevation',
  'startTime', 'endTime', 'boundsJson', 'createdAt', 'updatedAt', 'orderIndex', 'visible',
  'lineStyle', 'lineWidth'
];
const TRACK_SEGMENTS_HEADERS = [
  'trackId', 'revisionId', 'segmentIndex', 'chunkIndex', 'pointsJson', 'pointCount', 'createdAt', 'updatedAt'
];

function cloneRows(rows) {
  return rows.map((row) => row.slice());
}

function makeSheet(name, rows, audit) {
  const sheet = {
    name,
    rows: cloneRows(rows),
    formulas: rows.map((row) => row.map(() => '')),
    maxRows: Math.max(1, rows.length),
    setCalls: 0,
    failSetCall: 0,
    getLastRow() {
      for (let i = this.rows.length - 1; i >= 0; i -= 1) {
        if ((this.rows[i] || []).some((value) => value !== '' && value != null)) return i + 1;
      }
      return 0;
    },
    getLastColumn() { return this.rows.reduce((max, row) => Math.max(max, row.length), 0); },
    getMaxRows() { return this.maxRows; },
    insertRowsAfter(_position, count) { this.maxRows += count; },
    getDataRange() { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); },
    getRange(row, column, numRows = 1, numColumns = 1) {
      if (typeof row === 'string') throw new Error(`Unsupported A1 range in track harness: ${row}`);
      const range = {
        getValue: () => ((sheet.rows[row - 1] || [])[column - 1] ?? ''),
        getValues: () => Array.from({ length: numRows }, (_, r) =>
          Array.from({ length: numColumns }, (_, c) => ((sheet.rows[row - 1 + r] || [])[column - 1 + c] ?? ''))),
        getFormulas: () => Array.from({ length: numRows }, (_, r) =>
          Array.from({ length: numColumns }, (_, c) => ((sheet.formulas[row - 1 + r] || [])[column - 1 + c] ?? ''))),
        setValue(value) { return this.setValues([[value]]); },
        setValues(values) {
          sheet.setCalls += 1;
          audit.writes.push({ sheet: name, row, column, values: cloneRows(values), lockHeld: audit.lockHeld });
          if (sheet.failSetCall && sheet.setCalls === sheet.failSetCall) throw new Error(`simulated ${name} write failure`);
          values.forEach((valuesRow, r) => {
            while (sheet.rows.length <= row - 1 + r) sheet.rows.push([]);
            while (sheet.formulas.length <= row - 1 + r) sheet.formulas.push([]);
            valuesRow.forEach((value, c) => {
              sheet.rows[row - 1 + r][column - 1 + c] = value;
              sheet.formulas[row - 1 + r][column - 1 + c] = typeof value === 'string' && value.startsWith('=') ? value : '';
            });
          });
          sheet.maxRows = Math.max(sheet.maxRows, row - 1 + values.length);
          return this;
        },
        setBackground() { return this; }, setFontColor() { return this; }, setFontWeight() { return this; }
      };
      return range;
    },
    setFrozenRows() {}
  };
  return sheet;
}

function loadApi(options = {}) {
  const audit = { spreadsheetCalls: 0, insertCalls: 0, lockCalls: 0, lockHeld: false, flushCalls: 0, writes: [], driveCalls: 0, propertyCalls: 0 };
  const properties = new Map();
  const sheets = {};
  Object.entries(options.rows || {
    tracks: [TRACKS_HEADERS],
    track_segments: [TRACK_SEGMENTS_HEADERS]
  }).forEach(([name, rows]) => { sheets[name] = makeSheet(name, rows, audit); });
  const spreadsheet = {
    getSheetByName(name) { return sheets[name] || null; },
    insertSheet(name) { audit.insertCalls += 1; sheets[name] = makeSheet(name, [], audit); return sheets[name]; }
  };
  const context = {
    Logger: { log() {} },
    CacheService: { getScriptCache: () => ({ get: (key) => key === `EDIT_TOKEN_${TEST_EDIT_TOKEN}` ? '1' : null }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty(key) { audit.propertyCalls += 1; return properties.has(key) ? properties.get(key) : null; },
      setProperty(key, value) { audit.propertyCalls += 1; properties.set(key, String(value)); },
      deleteProperty(key) { audit.propertyCalls += 1; properties.delete(key); },
      getProperties() { audit.propertyCalls += 1; return Object.fromEntries(properties); }
    }) },
    LockService: { getScriptLock: () => ({
      tryLock() {
        audit.lockCalls += 1;
        audit.lockHeld = options.lockAvailable !== false;
        return audit.lockHeld;
      },
      releaseLock() { audit.lockHeld = false; }
    }) },
    SpreadsheetApp: {
      getActiveSpreadsheet() { audit.spreadsheetCalls += 1; return spreadsheet; },
      flush() { audit.flushCalls += 1; if (options.failFlushCall === audit.flushCalls) throw new Error('simulated flush failure'); }
    },
    DriveApp: new Proxy({}, { get() { audit.driveCalls += 1; throw new Error('Drive must not be called'); } }),
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest(_algorithm, value) { return Array.from(crypto.createHash('sha256').update(String(value)).digest()).map((byte) => byte > 127 ? byte - 256 : byte); },
      getUuid() { throw new Error('track API must not generate IDs'); }
    }
  };
  vm.runInNewContext(`${codeJs}\nglobalThis.__trackApi = {
    normalizeTrackBundle_, normalizeTrackPoint_, computeTrackSummary_, chunkTrackSegments_, hashTrackPayload_,
    saveTrackBundle, getTracks, deleteTrack,
    updateTrackDisplaySettings: typeof updateTrackDisplaySettings === 'function' ? updateTrackDisplaySettings : null,
    updateTracksOrder: typeof updateTracksOrder === 'function' ? updateTracksOrder : null,
    ensureHeaderSheet_, TRACKS_HEADERS, TRACK_SEGMENTS_HEADERS
  };`, context);
  return { api: context.__trackApi, sheets, audit, spreadsheet, properties };
}

function plain(value) { return JSON.parse(JSON.stringify(value)); }
function withToken(payload) { return Object.assign({}, payload, { __editToken: TEST_EDIT_TOKEN }); }
function point(lat, lng, elevation = null, time = '') { return { lat, lng, elevation, time }; }
function bundle(overrides = {}) {
  return Object.assign({
    trackId: 'track-a', revisionId: 'rev-a', name: 'Track A', description: '', color: '#2196f3',
    sourceType: 'gpx', sourceName: 'track.gpx', visible: true, lineStyle: 'solid', lineWidth: 4,
    segments: [{ index: 0, points: [point(35, 139, 100, '2026-07-11T01:02:03+09:00'), point(35.001, 139.001)] }]
  }, overrides);
}

test('track constants, fixed headers and setupSheet registration are present', () => {
  assert.ok(codeJs.includes("const TRACKS_SHEET_NAME = 'tracks';"));
  assert.ok(codeJs.includes("const TRACK_SEGMENTS_SHEET_NAME = 'track_segments';"));
  assert.ok(codeJs.includes('const MAX_TRACKS = 100;'));
  assert.ok(codeJs.includes('const MAX_TRACK_SEGMENTS = 200;'));
  assert.ok(codeJs.includes('const MAX_TRACK_POINTS = 20000;'));
  assert.ok(codeJs.includes('const TRACK_POINTS_PER_CHUNK = 500;'));
  assert.ok(codeJs.includes('const TRACK_POINTS_JSON_MAX_LENGTH = 40000;'));
  assert.ok(codeJs.includes('ensureHeaderSheet_(ss, TRACKS_SHEET_NAME, TRACKS_HEADERS);'));
  assert.ok(codeJs.includes('ensureHeaderSheet_(ss, TRACK_SEGMENTS_SHEET_NAME, TRACK_SEGMENTS_HEADERS);'));
});

test('track setup helper fills missing headers idempotently without changing existing data or extension columns', () => {
  const partialHeaders = TRACKS_HEADERS.slice();
  partialHeaders[21] = '';
  partialHeaders.push('customColumn');
  const existingRow = ['track-existing', 'name'];
  existingRow[22] = '=CUSTOM()';
  const { api, sheets, spreadsheet, audit } = loadApi({ rows: {
    tracks: [partialHeaders, existingRow], track_segments: [TRACK_SEGMENTS_HEADERS]
  } });
  api.ensureHeaderSheet_(spreadsheet, 'tracks', api.TRACKS_HEADERS);
  api.ensureHeaderSheet_(spreadsheet, 'tracks', api.TRACKS_HEADERS);
  assert.equal(sheets.tracks.rows[0][21], 'lineWidth');
  assert.equal(sheets.tracks.rows[0][22], 'customColumn');
  assert.equal(sheets.tracks.rows[1][0], 'track-existing');
  assert.equal(sheets.tracks.rows[1][22], '=CUSTOM()');
  assert.equal(audit.insertCalls, 0);
});

test('server normalization is strict, immutable and recomputes summary', () => {
  const { api } = loadApi();
  const input = bundle({ pointCount: 999, distanceMeters: 999, createdAt: 'forged', extra: true });
  const before = JSON.stringify(input);
  const normalized = api.normalizeTrackBundle_(input);
  assert.equal(JSON.stringify(input), before);
  assert.equal(normalized.pointCount, 2);
  assert.equal(normalized.segmentCount, 1);
  assert.ok(normalized.distanceMeters > 140 && normalized.distanceMeters < 150);
  assert.equal(normalized.startTime, '2026-07-10T16:02:03.000Z');
  assert.equal(normalized.endTime, '2026-07-10T16:02:03.000Z');
  assert.equal(normalized.createdAt, undefined);
  assert.equal(normalized.extra, undefined);
  assert.throws(() => api.normalizeTrackPoint_(point('35', 139)), /coordinate/i);
  assert.throws(() => api.normalizeTrackPoint_(point(91, 139)), /coordinate/i);
  assert.throws(() => api.normalizeTrackPoint_(point(35, 139, '', '')), /elevation/i);
  assert.throws(() => api.normalizeTrackPoint_({ lat: 35, lng: 139, elevation: null, time: null }), /time/i);
  assert.throws(() => api.normalizeTrackPoint_(point(35, 139, null, ' 2026-07-11T00:00:00Z ')), /time/i);
  assert.throws(() => api.normalizeTrackPoint_(point(35, 139, null, '2026-02-30T00:00:00Z')), /time/i);
  assert.throws(() => api.normalizeTrackBundle_(bundle({ trackId: '=bad' })), /trackId/i);
  assert.throws(() => api.normalizeTrackBundle_(bundle({ color: '#1e88e5' })), /color/i);
});

test('server track time and negative-zero normalization matches client boundary behavior', () => {
  const { api } = loadApi();
  const valid = api.normalizeTrackPoint_(point(
    -0, -0, -0, '2000-02-29T23:59:59.123456789+14:00'
  ));
  assert.equal(valid.time, '2000-02-29T09:59:59.123Z');
  assert.equal(Object.is(valid.lat, -0), true);
  assert.equal(Object.is(valid.lng, -0), true);
  assert.equal(Object.is(valid.elevation, -0), true);
  assert.equal(api.normalizeTrackPoint_(point(0, 0, null, '0001-01-01T00:00:00Z')).time,
    '0001-01-01T00:00:00.000Z');
  [
    '0000-01-01T00:00:00Z',
    '1900-02-29T00:00:00Z',
    '2000-02-29T00:00:00.1234567890Z',
    '2000-02-29T00:00:00+14:01',
    '2000-02-29T00:00:00+23:59',
    '2000-02-29T00:00:00+24:00'
  ].forEach((value) => assert.throws(() => api.normalizeTrackPoint_(point(0, 0, null, value))));
});

test('chunking preserves segment and point order while enforcing both limits', () => {
  const { api } = loadApi();
  const points = Array.from({ length: 501 }, (_, index) => point(35 + index / 100000, 139 + index / 100000, index, ''));
  const normalized = api.normalizeTrackBundle_(bundle({ segments: [{ points }, { points: [point(36, 140)] }] }));
  const rows = plain(api.chunkTrackSegments_(normalized));
  assert.deepEqual(rows.map((row) => [row.segmentIndex, row.chunkIndex, row.pointCount]), [[0, 0, 500], [0, 1, 1], [1, 0, 1]]);
  rows.forEach((row) => assert.ok(row.pointsJson.length <= 40000));
  assert.deepEqual(JSON.parse(rows[0].pointsJson)[0], [35, 139, 0, '']);
  assert.deepEqual(JSON.parse(rows[1].pointsJson)[0].slice(0, 2), [35.005, 139.005]);
});

test('normalization enforces segment and point maxima and hash ignores client summaries and order', () => {
  const { api } = loadApi();
  const maxSegments = Array.from({ length: 200 }, (_, index) => ({ points: [point(35, 139 + index / 1000)] }));
  assert.equal(api.normalizeTrackBundle_(bundle({ segments: maxSegments })).segmentCount, 200);
  assert.throws(() => api.normalizeTrackBundle_(bundle({ segments: maxSegments.concat([{ points: [point(35, 139)] }]) })), /segment/i);
  const tooManyPoints = Array.from({ length: 20001 }, () => point(35, 139));
  assert.throws(() => api.normalizeTrackBundle_(bundle({ segments: [{ points: tooManyPoints }] })), /point/i);
  const left = api.normalizeTrackBundle_(bundle({ orderIndex: 1, pointCount: 999, bounds: { south: 0 } }));
  const right = api.normalizeTrackBundle_(bundle({ orderIndex: 9, pointCount: 1, bounds: { south: 90 } }));
  assert.equal(api.hashTrackPayload_(left), api.hashTrackPayload_(right));
  assert.notEqual(api.hashTrackPayload_(left), api.hashTrackPayload_(api.normalizeTrackBundle_(bundle({ name: 'Different' }))));
});

test('save, replay, conflict, new revision, formula codec and read are isolated from routes and Drive', () => {
  const routeRows = [['routeId'], ['route-a']];
  const { api, sheets, audit } = loadApi({ rows: {
    tracks: [TRACKS_HEADERS], track_segments: [TRACK_SEGMENTS_HEADERS], routes: routeRows, route_pins: [['routeId']], route_cache: [['cacheKey']], map_info: [['ID']]
  } });
  const dangerous = bundle({ name: '=SUM(1,1)', description: "'keep", sourceName: '\u200Bdtp-sheet:v1:v:name.gpx' });
  const first = plain(api.saveTrackBundle(withToken(dangerous)));
  assert.equal(first.ok, true);
  assert.equal(first.deduplicated, false);
  assert.equal(first.track.name, '=SUM(1,1)');
  assert.notEqual(sheets.tracks.rows[1][1], '=SUM(1,1)');
  assert.equal(sheets.tracks.rows[1][2], "'keep");
  const writesBeforeReplay = audit.writes.length;
  const replay = plain(api.saveTrackBundle(withToken(dangerous)));
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(audit.writes.length, writesBeforeReplay);

  const conflict = plain(api.saveTrackBundle(withToken(bundle({ name: 'Changed' }))));
  assert.deepEqual(conflict, { ok: false, error: '同じリビジョンに異なる内容が送信されました。', errorCode: 'TRACK_REVISION_PAYLOAD_CONFLICT', retryable: false });

  const next = plain(api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'Revision B' }))));
  assert.equal(next.ok, true);
  assert.equal(sheets.track_segments.rows.filter((row) => row[0] === 'track-a' && row[1] === 'rev-a').length, 0);
  const read = plain(api.getTracks());
  assert.equal(read.ok, true);
  assert.equal(read.tracks.length, 1);
  assert.equal(read.tracks[0].revisionId, 'rev-b');
  assert.equal(Object.prototype.hasOwnProperty.call(read.tracks[0], 'payloadHash'), false);
  assert.deepEqual(sheets.routes.rows, routeRows);
  assert.equal(audit.driveCalls, 0);
});

test('metadata write interruption keeps old active revision and retry converges', () => {
  const { api, sheets } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  sheets.tracks.failSetCall = sheets.tracks.setCalls + 1;
  const failed = plain(api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' }))));
  assert.equal(failed.ok, false);
  assert.equal(plain(api.getTracks()).tracks[0].revisionId, 'rev-a');
  sheets.tracks.failSetCall = 0;
  const retried = plain(api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' }))));
  assert.equal(retried.ok, true);
  assert.equal(plain(api.getTracks()).tracks[0].revisionId, 'rev-b');
});

test('interrupted first save stays invisible and retry replaces staged rows into one active track', () => {
  const { api, sheets } = loadApi();
  sheets.tracks.failSetCall = 1;
  const failed = plain(api.saveTrackBundle(withToken(bundle())));
  assert.equal(failed.ok, false);
  assert.deepEqual(plain(api.getTracks()).tracks, []);
  assert.ok(sheets.track_segments.rows.some((row) => row[1] === 'rev-a'));
  sheets.tracks.failSetCall = 0;
  const retried = plain(api.saveTrackBundle(withToken(bundle())));
  assert.equal(retried.ok, true);
  assert.equal(plain(api.getTracks()).tracks.length, 1);
  assert.equal(sheets.track_segments.rows.filter((row) => row[0] === 'track-a' && row[1] === 'rev-a').length, 1);
});

test('staged revision journal rejects a changed payload before replacing interrupted rows', () => {
  const { api, sheets } = loadApi();
  sheets.tracks.failSetCall = 1;
  assert.equal(plain(api.saveTrackBundle(withToken(bundle()))).ok, false);
  sheets.tracks.failSetCall = 0;
  const conflict = plain(api.saveTrackBundle(withToken(bundle({ name: 'Changed after interruption' }))));
  assert.equal(conflict.errorCode, 'TRACK_REVISION_PAYLOAD_CONFLICT');
  assert.equal(conflict.retryable, false);
  assert.deepEqual(plain(api.getTracks()).tracks, []);
});

test('segment write and flush boundary failures preserve active visibility and converge on retry', () => {
  const segmentFailure = loadApi();
  segmentFailure.sheets.track_segments.failSetCall = 1;
  assert.equal(plain(segmentFailure.api.saveTrackBundle(withToken(bundle()))).ok, false);
  assert.deepEqual(plain(segmentFailure.api.getTracks()).tracks, []);
  segmentFailure.sheets.track_segments.failSetCall = 0;
  assert.equal(plain(segmentFailure.api.saveTrackBundle(withToken(bundle()))).ok, true);
  assert.equal(segmentFailure.properties.size, 0);

  const segmentFlushFailure = loadApi({ failFlushCall: 1 });
  assert.equal(plain(segmentFlushFailure.api.saveTrackBundle(withToken(bundle()))).ok, false);
  assert.deepEqual(plain(segmentFlushFailure.api.getTracks()).tracks, []);
  assert.equal(plain(segmentFlushFailure.api.saveTrackBundle(withToken(bundle()))).ok, true);
  assert.equal(segmentFlushFailure.properties.size, 0);

  const metadataFlushFailure = loadApi({ failFlushCall: 2 });
  assert.equal(plain(metadataFlushFailure.api.saveTrackBundle(withToken(bundle()))).ok, false);
  assert.equal(plain(metadataFlushFailure.api.getTracks()).tracks[0].revisionId, 'rev-a');
  assert.equal(plain(metadataFlushFailure.api.saveTrackBundle(withToken(bundle()))).deduplicated, true);
  assert.equal(metadataFlushFailure.properties.size, 0);

  const cleanupFlushFailure = loadApi({ failFlushCall: 5 });
  assert.equal(cleanupFlushFailure.api.saveTrackBundle(withToken(bundle())).ok, true);
  assert.equal(plain(cleanupFlushFailure.api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' })))).ok, false);
  assert.equal(plain(cleanupFlushFailure.api.getTracks()).tracks[0].revisionId, 'rev-b');
  assert.equal(plain(cleanupFlushFailure.api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' })))).deduplicated, true);
  assert.equal([...cleanupFlushFailure.properties.keys()].filter((key) => key.startsWith('TRACK_STAGE_V1_')).length, 0);
  assert.equal([...cleanupFlushFailure.properties.keys()].filter((key) => key.startsWith('TRACK_RETIRED_REVISION_V1_')).length, 1);
});

test('cleanup interruption exposes new active revision and replay safely removes stale rows', () => {
  const { api, sheets } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  sheets.track_segments.failSetCall = sheets.track_segments.setCalls + 2;
  const failed = plain(api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' }))));
  assert.equal(failed.ok, false);
  assert.equal(plain(api.getTracks()).tracks[0].revisionId, 'rev-b');
  sheets.track_segments.failSetCall = 0;
  const replay = plain(api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' }))));
  assert.equal(replay.ok, true);
  assert.equal(replay.deduplicated, true);
  assert.equal(sheets.track_segments.rows.some((row) => row[1] === 'rev-a'), false);
});

test('a retired revision cannot roll back the active revision even with its original payload', () => {
  const { api } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  assert.equal(api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' }))).ok, true);

  const sameOldRevision = plain(api.saveTrackBundle(withToken(bundle())));
  assert.equal(sameOldRevision.errorCode, 'TRACK_REVISION_PAYLOAD_CONFLICT');
  assert.equal(sameOldRevision.retryable, false);
  const changedOldRevision = plain(api.saveTrackBundle(withToken(bundle({ name: 'Changed old A' }))));
  assert.equal(changedOldRevision.errorCode, 'TRACK_REVISION_PAYLOAD_CONFLICT');
  assert.equal(plain(api.getTracks()).tracks[0].revisionId, 'rev-b');
});

test('getTracks excludes only corrupted tracks and returns safe warnings', () => {
  const { api, sheets } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  assert.equal(api.saveTrackBundle(withToken(bundle({ trackId: 'track-b', revisionId: 'rev-b', name: 'B', orderIndex: 0 }))).ok, true);
  const badRow = sheets.track_segments.rows.find((row) => row[0] === 'track-a');
  badRow[5] = 999;
  const result = plain(api.getTracks());
  assert.deepEqual(result.tracks.map((track) => track.trackId), ['track-b']);
  assert.deepEqual(result.warnings, [{ code: 'TRACK_SEGMENTS_CORRUPTED', trackId: 'track-a' }]);
  assert.equal(JSON.stringify(result.warnings).includes('pointsJson'), false);
});

test('getTracks classifies an invalid stored track identifier as metadata corruption', () => {
  const { api, sheets } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  sheets.tracks.rows[1][0] = '=invalid';
  const result = plain(api.getTracks());
  assert.deepEqual(result.tracks, []);
  assert.deepEqual(result.warnings, [{ code: 'TRACK_METADATA_CORRUPTED', trackId: '' }]);
});

test('getTracks bounds corrupt chunk indexes and sizes before reconstruction work', () => {
  const { api, sheets } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  const segmentRow = sheets.track_segments.rows.find((row) => row[0] === 'track-a');
  segmentRow[2] = 1000000000;
  let result = plain(api.getTracks());
  assert.deepEqual(result.warnings, [{ code: 'TRACK_SEGMENTS_CORRUPTED', trackId: 'track-a' }]);
  segmentRow[2] = 0;
  segmentRow[5] = 501;
  result = plain(api.getTracks());
  assert.deepEqual(result.warnings, [{ code: 'TRACK_SEGMENTS_CORRUPTED', trackId: 'track-a' }]);
});

test('getTracks omits duplicate metadata rows and strictly rejects malformed stored metadata', () => {
  const { api, sheets } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  sheets.tracks.rows.push(sheets.tracks.rows[1].slice());
  let result = plain(api.getTracks());
  assert.deepEqual(result.tracks, []);
  assert.deepEqual(result.warnings, [{ code: 'TRACK_METADATA_CORRUPTED', trackId: 'track-a' }]);

  sheets.tracks.rows[2][19] = 'not-a-boolean';
  result = plain(api.getTracks());
  assert.deepEqual(result.tracks, []);
  assert.deepEqual(result.warnings, [{ code: 'TRACK_METADATA_CORRUPTED', trackId: 'track-a' }]);

  sheets.tracks.rows.pop();
  sheets.tracks.rows[1][19] = 'not-a-boolean';
  result = plain(api.getTracks());
  assert.deepEqual(result.tracks, []);
  assert.deepEqual(result.warnings, [{ code: 'TRACK_METADATA_CORRUPTED', trackId: 'track-a' }]);
});

test('prototype-like valid track IDs save, count and read without object-key collisions', () => {
  const { api } = loadApi();
  ['__proto__', 'constructor', 'toString'].forEach((trackId, index) => {
    assert.equal(api.saveTrackBundle(withToken(bundle({ trackId, revisionId: `rev-${index}`, name: trackId }))).ok, true);
  });
  const result = plain(api.getTracks());
  assert.deepEqual(result.tracks.map((track) => track.trackId), ['__proto__', 'constructor', 'toString']);
  assert.deepEqual(result.warnings, []);
});

test('omitted orderIndex appends tracks and replay preserves order with a stable read tie-break', () => {
  const { api, sheets } = loadApi();
  const first = plain(api.saveTrackBundle(withToken(bundle({
    trackId: 'track-b', revisionId: 'rev-b', name: 'B'
  }))));
  const second = plain(api.saveTrackBundle(withToken(bundle({
    trackId: 'track-c', revisionId: 'rev-c', name: 'C'
  }))));
  assert.equal(first.track.orderIndex, 0);
  assert.equal(second.track.orderIndex, 1);

  assert.equal(plain(api.saveTrackBundle(withToken(bundle({
    trackId: 'track-a', revisionId: 'rev-a', name: 'A', orderIndex: 0
  })))).ok, true);
  const read = plain(api.getTracks());
  assert.deepEqual(read.tracks.map((track) => track.trackId), ['track-a', 'track-b', 'track-c']);

  const replay = plain(api.saveTrackBundle(withToken(bundle({
    trackId: 'track-c', revisionId: 'rev-c', name: 'C'
  }))));
  assert.equal(replay.deduplicated, true);
  assert.equal(replay.track.orderIndex, 1);
});

test('updateTracksOrder persists GPX and GeoJSON order and survives a fresh getTracks read', () => {
  const { api, sheets, audit } = loadApi();
  assert.equal(typeof api.updateTracksOrder, 'function');
  assert.equal(api.saveTrackBundle(withToken(bundle({
    trackId: 'track-a', revisionId: 'rev-a', sourceType: 'gpx', sourceName: 'a.gpx'
  }))).ok, true);
  assert.equal(api.saveTrackBundle(withToken(bundle({
    trackId: 'track-b', revisionId: 'rev-b', sourceType: 'geojson', sourceName: 'b.geojson'
  }))).ok, true);
  audit.writes.length = 0;

  const result = plain(api.updateTracksOrder(withToken({ orderedIds: ['track-b', 'track-a'] })));

  assert.deepEqual(result.tracks.map((track) => [track.trackId, track.orderIndex]), [
    ['track-b', 0], ['track-a', 1]
  ]);
  assert.deepEqual(plain(api.getTracks()).tracks.map((track) => track.trackId), ['track-b', 'track-a']);
  assert.equal(sheets.tracks.rows.find((row) => row[0] === 'track-a')[18], 1);
  assert.equal(sheets.tracks.rows.find((row) => row[0] === 'track-b')[18], 0);
  assert.equal(audit.writes.length, 1);
  assert.equal(audit.writes[0].column, 19);
  assert.equal(audit.writes[0].lockHeld, true);
});

test('updateTracksOrder rejects duplicate and missing track IDs without writing', () => {
  const { api, audit } = loadApi();
  assert.equal(typeof api.updateTracksOrder, 'function');
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  audit.writes.length = 0;

  let result = plain(api.updateTracksOrder(withToken({ orderedIds: ['track-a', 'track-a'] })));
  assert.equal(result.errorCode, 'INVALID_TRACK_PAYLOAD');
  assert.equal(audit.writes.length, 0);

  result = plain(api.updateTracksOrder(withToken({ orderedIds: ['missing'] })));
  assert.equal(result.errorCode, 'TRACK_NOT_FOUND');
  assert.equal(audit.writes.length, 0);
});

test('updateTrackDisplaySettings changes only editable metadata and preserves formulas and extensions', () => {
  const { api, sheets, audit } = loadApi();
  assert.equal(typeof api.updateTrackDisplaySettings, 'function');
  assert.equal(api.saveTrackBundle(withToken(bundle({
    name: 'Original', description: 'before', lineWidth: 8, orderIndex: 3
  }))).ok, true);
  const row = sheets.tracks.rows[1];
  const formulas = sheets.tracks.formulas[1];
  row.push('cached extension');
  formulas.push('=EXTENSION()');
  const preserved = row.slice();
  audit.writes.length = 0;

  const result = plain(api.updateTrackDisplaySettings(withToken({
    trackId: 'track-a', name: '=Renamed', description: '+updated', color: '#e53935',
    visible: false, lineStyle: 'dashed', lineWidth: 10,
    segments: [{ index: 99, points: [] }], orderIndex: 0
  })));

  assert.equal(result.ok, true);
  assert.deepEqual(result.track, {
    trackId: 'track-a', name: '=Renamed', description: '+updated', color: '#e53935',
    visible: false, lineStyle: 'dashed', lineWidth: 4, updatedAt: result.track.updatedAt
  });
  assert.match(result.track.updatedAt, /^\d{4}-\d{2}-\d{2}T/);
  assert.notEqual(row[1], '=Renamed');
  assert.notEqual(row[2], '+updated');
  assert.equal(row[3], '#e53935');
  assert.equal(row[19], false);
  assert.equal(row[20], 'dashed');
  assert.equal(row[21], 4);
  for (const index of [0, 4, 5, 6, 7, 8, 9, 10, 16, 18]) {
    assert.equal(row[index], preserved[index], `column ${index + 1}`);
  }
  assert.equal(row[22], '=EXTENSION()');
  assert.equal(sheets.tracks.formulas[1][22], '=EXTENSION()');
  assert.equal(sheets.track_segments.rows.length, 2);
  const read = plain(api.getTracks()).tracks[0];
  assert.equal(read.name, '=Renamed');
  assert.equal(read.description, '+updated');
  assert.equal(audit.writes.length, 1);
  assert.equal(audit.writes[0].sheet, 'tracks');
  assert.equal(audit.writes[0].row, 2);
  assert.equal(audit.writes[0].column, 1);
  assert.equal(audit.writes[0].lockHeld, true);
});

test('updateTrackDisplaySettings validates before writing and classifies missing, busy, and write failures', () => {
  const missingToken = loadApi();
  assert.throws(() => missingToken.api.updateTrackDisplaySettings({
    trackId: 'track-a', name: 'A', description: '', color: '#2196f3', visible: true, lineStyle: 'solid'
  }), /編集権限/);
  assert.equal(missingToken.audit.spreadsheetCalls, 0);

  const setup = loadApi();
  assert.equal(setup.api.saveTrackBundle(withToken(bundle())).ok, true);
  setup.audit.writes.length = 0;
  const valid = { trackId: 'track-a', name: 'A', description: '', color: '#2196f3', visible: true, lineStyle: 'solid' };
  for (const patch of [
    { name: '' }, { name: 'x'.repeat(101) }, { description: 'x'.repeat(401) },
    { color: '#ffffff' }, { visible: 'true' }, { lineStyle: 'doubled' }
  ]) {
    const result = plain(setup.api.updateTrackDisplaySettings(withToken(Object.assign({}, valid, patch))));
    assert.equal(result.errorCode, 'INVALID_TRACK_PAYLOAD');
  }
  assert.equal(setup.audit.writes.length, 0);

  let result = plain(setup.api.updateTrackDisplaySettings(withToken(Object.assign({}, valid, { trackId: 'missing' }))));
  assert.equal(result.errorCode, 'TRACK_NOT_FOUND');
  assert.equal(setup.audit.writes.length, 0);

  const busy = loadApi({ lockAvailable: false });
  result = plain(busy.api.updateTrackDisplaySettings(withToken(valid)));
  assert.equal(result.errorCode, 'TRACK_STORAGE_BUSY');
  assert.equal(result.retryable, true);

  const failed = loadApi();
  assert.equal(failed.api.saveTrackBundle(withToken(bundle())).ok, true);
  failed.sheets.tracks.failSetCall = failed.sheets.tracks.setCalls + 1;
  result = plain(failed.api.updateTrackDisplaySettings(withToken(valid)));
  assert.equal(result.errorCode, 'TRACK_METADATA_UPDATE_FAILED');
  assert.equal(result.retryable, true);
});

test('deleteTrack is token-first, idempotent and removes every revision without creating sheets', () => {
  const missing = loadApi({ rows: {} });
  assert.throws(() => missing.api.saveTrackBundle(bundle()), /編集権限/);
  assert.equal(missing.audit.spreadsheetCalls, 0);
  assert.equal(missing.audit.lockCalls, 0);
  assert.equal(missing.audit.propertyCalls, 0);
  assert.equal(missing.audit.insertCalls, 0);
  assert.equal(plain(missing.api.getTracks()).errorCode, 'TRACK_SHEETS_MISSING');
  assert.equal(plain(missing.api.saveTrackBundle(withToken(bundle()))).errorCode, 'TRACK_SHEETS_MISSING');
  assert.equal(plain(missing.api.deleteTrack(withToken({ trackId: 'track-a' }))).errorCode, 'TRACK_SHEETS_MISSING');
  assert.equal(missing.audit.insertCalls, 0);

  const { api, sheets, audit, properties } = loadApi();
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);
  assert.equal(api.saveTrackBundle(withToken(bundle({ revisionId: 'rev-b', name: 'B' }))).ok, true);
  assert.ok(properties.size > 0);
  sheets.track_segments.rows.push(['track-a', 'orphan-rev', 0, 0, '[[35,139,null,""]]', 1, '', '']);
  const deleted = plain(api.deleteTrack(withToken({ trackId: 'track-a' })));
  const again = plain(api.deleteTrack(withToken({ trackId: 'track-a' })));
  assert.deepEqual(deleted, { ok: true, deleted: true, removedShareReferences: 0 });
  assert.deepEqual(again, { ok: true, deleted: false, removedShareReferences: 0 });
  assert.equal(sheets.track_segments.rows.some((row) => row[0] === 'track-a'), false);
  assert.equal(properties.size, 0);
  assert.equal(audit.driveCalls, 0);
});

test('deleteTrack removes only matching typed track targets from valid share JSON', () => {
  const shareHeaders = [
    'createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt',
    'colors', 'routeIds', 'routeTargetsJson'
  ];
  const validTargets = {
    v: 1,
    targets: [
      { type: 'pin-route', id: 'track-a' },
      { type: 'gpx-route', id: 'track-a' },
      { type: 'geojson-route', id: 'track-a' },
      { type: 'gpx-route', id: 'track-b' }
    ]
  };
  const brokenJson = '{"v":1,"targets":[';
  const { api, sheets } = loadApi({ rows: {
    tracks: [TRACKS_HEADERS],
    track_segments: [TRACK_SEGMENTS_HEADERS],
    share_links: [
      shareHeaders,
      ['', 'Valid', 'valid', '', 'or', true, '', '', 'legacy-a|legacy-b', JSON.stringify(validTargets)],
      ['', 'Broken', 'broken', '', 'or', true, '', '', 'legacy-b', brokenJson]
    ]
  } });
  assert.equal(api.saveTrackBundle(withToken(bundle())).ok, true);

  const result = plain(api.deleteTrack(withToken({ trackId: 'track-a' })));

  assert.deepEqual(result, { ok: true, deleted: true, removedShareReferences: 2 });
  assert.deepEqual(JSON.parse(sheets.share_links.rows[1][9]), {
    v: 1,
    targets: [
      { type: 'pin-route', id: 'track-a' },
      { type: 'gpx-route', id: 'track-b' }
    ]
  });
  assert.equal(sheets.share_links.rows[1][8], 'legacy-a|legacy-b');
  assert.equal(sheets.share_links.rows[2][9], brokenJson);
  assert.equal(sheets.share_links.rows[2][8], 'legacy-b');
});

test('deleteTrack distinguishes busy, metadata, segment, and share reference failures', () => {
  const busy = loadApi({ lockAvailable: false });
  assert.equal(plain(busy.api.deleteTrack(withToken({ trackId: 'track-a' }))).errorCode, 'TRACK_STORAGE_BUSY');

  const metadataFailure = loadApi();
  assert.equal(metadataFailure.api.saveTrackBundle(withToken(bundle())).ok, true);
  metadataFailure.sheets.tracks.failSetCall = metadataFailure.sheets.tracks.setCalls + 1;
  const metadataResult = plain(metadataFailure.api.deleteTrack(withToken({ trackId: 'track-a' })));
  assert.equal(metadataResult.errorCode, 'TRACK_METADATA_DELETE_FAILED');
  assert.equal(metadataFailure.sheets.tracks.rows.some((row) => row[0] === 'track-a'), true);

  const segmentFailure = loadApi();
  assert.equal(segmentFailure.api.saveTrackBundle(withToken(bundle())).ok, true);
  segmentFailure.sheets.track_segments.failSetCall = segmentFailure.sheets.track_segments.setCalls + 1;
  const segmentResult = plain(segmentFailure.api.deleteTrack(withToken({ trackId: 'track-a' })));
  assert.equal(segmentResult.errorCode, 'TRACK_SEGMENTS_DELETE_FAILED');
  assert.equal(segmentFailure.sheets.track_segments.rows.some((row) => row[0] === 'track-a'), true);

  const shareHeaders = [
    'createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt',
    'colors', 'routeIds', 'routeTargetsJson'
  ];
  const shareFailure = loadApi({ rows: {
    tracks: [TRACKS_HEADERS],
    track_segments: [TRACK_SEGMENTS_HEADERS],
    share_links: [shareHeaders, ['', '', '', '', '', true, '', '', '', JSON.stringify({
      v: 1, targets: [{ type: 'gpx-route', id: 'track-a' }]
    })]]
  } });
  assert.equal(shareFailure.api.saveTrackBundle(withToken(bundle())).ok, true);
  shareFailure.sheets.share_links.failSetCall = shareFailure.sheets.share_links.setCalls + 1;
  const shareResult = plain(shareFailure.api.deleteTrack(withToken({ trackId: 'track-a' })));
  assert.equal(shareResult.errorCode, 'TRACK_SHARE_REFERENCES_DELETE_FAILED');
  assert.equal(shareResult.serverDeleted, true);
  assert.equal(shareFailure.sheets.tracks.rows.some((row) => row[0] === 'track-a'), false);
  assert.equal(shareFailure.sheets.track_segments.rows.some((row) => row[0] === 'track-a'), false);
});

test('saveTrackBundle enforces the 100-track limit before any sheet write', () => {
  const rows = [TRACKS_HEADERS];
  for (let index = 0; index < 100; index += 1) rows.push([`existing-${index}`]);
  const { api, audit } = loadApi({ rows: { tracks: rows, track_segments: [TRACK_SEGMENTS_HEADERS] } });
  const result = plain(api.saveTrackBundle(withToken(bundle({ trackId: 'track-new' }))));
  assert.equal(result.errorCode, 'TRACK_LIMIT_EXCEEDED');
  assert.equal(audit.writes.length, 0);
});

test('track storage fixes line width at 4 while accepting legacy stored widths', () => {
  const { api, sheets } = loadApi();
  const saved = plain(api.saveTrackBundle(withToken(bundle({ lineWidth: 9 }))));
  assert.equal(saved.ok, true);
  assert.equal(saved.track.lineWidth, 4);
  assert.equal(sheets.tracks.rows[1][21], 4);

  sheets.tracks.rows[1][21] = 7;
  const legacyRead = plain(api.getTracks());
  assert.equal(legacyRead.ok, true);
  assert.equal(legacyRead.tracks[0].lineWidth, 4);
});

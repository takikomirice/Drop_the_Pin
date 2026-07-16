const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadModules() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const GeoJsonTrackInterchangeCore =', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = {
    PIN_COLORS: [{ hex: '#2196f3', label: '青' }],
    PIN_ICONS: [{ id: 'photo', label: '写真' }],
    PIN_STATUSES: ['未対応'],
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    Date,
    Math,
    Map,
    Set
  };
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__jobCore = ImportJobCore;\n'
      + 'globalThis.__previewCore = typeof PhotoTrackMatchPreviewCore === "undefined" ? null : PhotoTrackMatchPreviewCore;\n'
      + 'globalThis.__matchCore = PhotoTrackMatchCore;\n'
      + 'globalThis.__geometry = TrackGeometryCore;',
    context
  );
  return {
    jobCore: context.__jobCore,
    previewCore: context.__previewCore,
    matchCore: context.__matchCore,
    geometry: context.__geometry
  };
}

function job(jobCore) {
  const fileA = { name: 'a.jpg' };
  const fileB = { name: 'b.jpg' };
  const value = jobCore.createJob({
    id: 'job-1',
    sourceType: 'multi-photo',
    createdAt: '2026-07-12T00:00:00.000Z',
    items: [
      {
        id: 'a', sourceType: 'photo', sourceRef: 'a.jpg', title: 'GPSあり',
        description: 'keep', lat: 35, lng: 139, capturedAt: '2026-07-12T09:00:00',
        tags: ['旅'], links: ['https://example.com'], color: '#2196f3', icon: 'photo',
        metadataStatus: 'gps', conversionStatus: 'not-needed', uploadStatus: 'queued',
        runtime: { originalFile: fileA, uploadFile: fileA, previewUrl: 'blob:a' }
      },
      {
        id: 'b', sourceType: 'photo', sourceRef: 'b.jpg', title: 'GPSなし',
        lat: null, lng: null, capturedAt: '2026-07-12T09:05:00',
        tags: [], links: [], color: '#2196f3', icon: 'photo', uploadStatus: 'queued',
        runtime: { originalFile: fileB, uploadFile: fileB, previewUrl: 'blob:b' }
      }
    ]
  });
  return { value, fileA, fileB };
}

function track(revisionId = 'rev-1', offset = 0, overrides = {}) {
  return {
    trackId: 'track-1', revisionId, name: '<img src=x onerror=1>', description: '',
    color: '#2196f3', sourceType: 'gpx', sourceName: 'private.gpx',
    segments: [{ points: [
      { lat: 10 + offset, lng: 20 + offset, elevation: null, time: '2026-07-12T00:04:00.000Z' },
      { lat: 20 + offset, lng: 30 + offset, elevation: null, time: '2026-07-12T00:06:00.000Z' }
    ] }],
    ...overrides
  };
}

test('location patches change only lat/lng and preserve identity, order, runtime, and edited fields immutably', () => {
  const { jobCore } = loadModules();
  const source = job(jobCore);
  const before = plain(source.value);
  const next = jobCore.applyLocationPatches(source.value, [
    { itemId: 'b', patch: { lat: 15, lng: 25 } }
  ]);

  assert.notEqual(next, source.value);
  assert.deepEqual(plain(source.value), before);
  assert.equal(next.id, source.value.id);
  assert.deepEqual(next.items.map((item) => item.id), ['a', 'b']);
  assert.equal(next.items[0].lat, 35);
  assert.equal(next.items[0].lng, 139);
  assert.equal(next.items[1].lat, 15);
  assert.equal(next.items[1].lng, 25);
  assert.equal(next.items[0].runtime.originalFile, source.fileA);
  assert.equal(next.items[1].runtime.uploadFile, source.fileB);
  assert.equal(next.items[1].runtime.previewUrl, 'blob:b');
  assert.equal(next.items[0].title, 'GPSあり');
  assert.equal(next.items[0].description, 'keep');
  assert.notEqual(next.items[0], source.value.items[0]);
  assert.notEqual(next.items[1], source.value.items[1]);
});

test('location patches reject duplicate, unknown, partial, string, non-finite, and out-of-range coordinates', () => {
  const { jobCore } = loadModules();
  const source = job(jobCore).value;
  const invalid = [
    [{ itemId: 'b', patch: { lat: 1, lng: 2 } }, { itemId: 'b', patch: { lat: 3, lng: 4 } }],
    [{ itemId: 'missing', patch: { lat: 1, lng: 2 } }],
    [{ itemId: 'b', patch: { lat: 1 } }],
    [{ itemId: 'b', patch: { lat: '1', lng: 2 } }],
    [{ itemId: 'b', patch: { lat: NaN, lng: 2 } }],
    [{ itemId: 'b', patch: { lat: 91, lng: 2 } }],
    [{ itemId: 'b', patch: { lat: 1, lng: 181 } }]
  ];
  invalid.forEach((patches) => assert.throws(() => jobCore.applyLocationPatches(source, patches)));
});

test('location patches require own safe string item ids and clone an empty patch immutably', () => {
  const { jobCore } = loadModules();
  const source = job(jobCore).value;
  const inherited = Object.create({ itemId: 'b' });
  inherited.patch = { lat: 1, lng: 2 };
  [
    [{ itemId: 1, patch: { lat: 1, lng: 2 } }],
    [inherited],
    [{ itemId: '__proto__', patch: { lat: 1, lng: 2 } }],
    [{ itemId: 'constructor', patch: { lat: 1, lng: 2 } }],
    [{ itemId: 'prototype', patch: { lat: 1, lng: 2 } }]
  ].forEach((patches) => assert.throws(
    () => jobCore.applyLocationPatches(source, patches),
    (error) => error.code === 'INVALID_IMPORT_LOCATION_PATCH'
  ));

  const cloned = jobCore.applyLocationPatches(source, []);
  assert.notEqual(cloned, source);
  assert.notEqual(cloned.items, source.items);
  assert.deepEqual(plain(cloned), plain(source));
  assert.notEqual(cloned.items[0].tags, source.items[0].tags);
  assert.notEqual(cloned.items[0].links, source.items[0].links);
  assert.equal(cloned.items[0].runtime.originalFile, source.items[0].runtime.originalFile);
});

test('match execution is non-mutating, selects only matches, and explicit apply protects original GPS', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  assert.ok(previewCore);
  const source = job(jobCore);
  const tracks = [track()];
  const state = {};
  const controller = previewCore.create({
    state,
    getTracks: () => tracks,
    matchCore,
    trackGeometry: geometry,
    importJobCore: jobCore,
    timezoneOffset: () => -540
  });
  controller.initialize(source.value);
  controller.setTrack('track-1', source.value);
  const before = plain(source.value);
  const result = controller.run(source.value);

  assert.deepEqual(plain(source.value), before);
  assert.equal(result.counts.skippedExistingGps, 1);
  assert.equal(result.counts.matched, 1);
  assert.deepEqual(Array.from(state.selectedItemIds), ['b']);
  const applied = controller.apply(source.value);
  assert.equal(applied.items[0].lat, 35);
  assert.equal(applied.items[0].lng, 139);
  assert.equal(applied.items[1].lat, 15);
  assert.equal(applied.items[1].lng, 25);
  assert.equal(state.appliedByItemId.b.trackId, 'track-1');
  assert.equal(state.appliedByItemId.b.revisionId, 'rev-1');
  assert.equal(Object.getPrototypeOf(state.appliedByItemId), null);
  assert.equal(Object.getPrototypeOf(state.originalCoordinatesByItemId), null);
  assert.equal(JSON.stringify(state).includes('private.gpx'), false);
  assert.equal(JSON.stringify(state).includes('blob:b'), false);
});

test('Preview initialization rejects duplicate and dangerous item ids before snapshotting coordinates', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  function controller() {
    return previewCore.create({
      state: {}, getTracks: () => [track()], matchCore, trackGeometry: geometry,
      importJobCore: jobCore, timezoneOffset: () => -540
    });
  }
  const dangerous = jobCore.createJob({
    id: 'dangerous', sourceType: 'multi-photo',
    items: [{ id: '__proto__', title: 'danger', lat: null, lng: null,
      color: '#2196f3', icon: 'photo', uploadStatus: 'queued' }]
  });
  assert.throws(() => controller().initialize(dangerous),
    (error) => error.code === 'TRACK_MATCH_PHOTOS_INVALID');

  const source = job(jobCore).value;
  const duplicate = { ...source, items: [source.items[0], { ...source.items[1], id: 'a' }] };
  assert.throws(() => controller().initialize(duplicate),
    (error) => error.code === 'TRACK_MATCH_PHOTOS_INVALID');
});

test('matcher-owned coordinates can be replaced after re-match but manual coordinates cannot', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const tracks = [track()];
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => tracks, matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  const source = job(jobCore).value;
  controller.initialize(source);
  controller.setTrack('track-1', source);
  controller.run(source);
  const first = controller.apply(source);
  controller.setOption('utcOffsetText', '+09:01', first);
  assert.equal(state.stale, true);
  controller.run(first);
  const replaced = controller.apply(first);
  assert.notEqual(replaced.items[1].lat, first.items[1].lat);

  const manual = jobCore.updateDraftItem(replaced, 'b', { lat: 40, lng: 140 });
  controller.onDraftChange(replaced, manual, { itemId: 'b', field: 'lat' });
  assert.equal(Object.prototype.hasOwnProperty.call(state.appliedByItemId, 'b'), false);
  controller.run(manual);
  assert.equal(state.result.results.find((entry) => entry.itemId === 'b').status, 'skipped-existing-gps');
  assert.equal(manual.items[1].lat, 40);
  assert.equal(manual.items[1].lng, 140);

  const cleared = jobCore.updateDraftItem(manual, 'b', { lat: null, lng: null });
  controller.onDraftChange(manual, cleared, { itemId: 'b', field: 'lat' });
  controller.run(cleared);
  assert.equal(state.result.results.find((entry) => entry.itemId === 'b').status.startsWith('matched-'), true);
});

test('capturedAt, options, item changes, and current track revision make old results stale and non-applicable', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  let tracks = [track()];
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => tracks, matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  const source = job(jobCore).value;
  controller.initialize(source);
  controller.setTrack('track-1', source);
  controller.run(source);

  const changedTime = jobCore.updateDraftItem(source, 'b', { capturedAt: '2026-07-12T09:06:00' });
  controller.onDraftChange(source, changedTime, { itemId: 'b', field: 'capturedAt' });
  assert.equal(state.stale, true);
  assert.throws(() => controller.apply(changedTime));

  controller.run(changedTime);
  tracks = [track('rev-2', 5)];
  const model = controller.getViewModel(changedTime);
  assert.equal(model.stale, true);
  assert.throws(() => controller.apply(changedTime));

  controller.setTrack('track-1', changedTime);
  controller.run(changedTime);
  const removed = jobCore.removeDraftItem(changedTime, 'a', { revokeObjectURL() {} });
  controller.onDraftChange(changedTime, removed, { itemId: 'a', field: 'remove' });
  assert.equal(state.stale, true);
});

test('clear removes pending results but keeps applied coordinates and cleanup releases all match state', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [track()], matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  const source = job(jobCore).value;
  controller.initialize(source);
  controller.setTrack('track-1', source);
  controller.run(source);
  const applied = controller.apply(source);
  controller.clearResult();
  assert.equal(state.result, null);
  assert.deepEqual(Array.from(state.selectedItemIds), []);
  assert.equal(applied.items[1].lat, 15);
  assert.equal(state.appliedByItemId.b.trackId, 'track-1');
  controller.cleanup();
  assert.equal(state.trackId, '');
  assert.equal(state.result, null);
  assert.deepEqual(Array.from(state.selectedItemIds), []);
  assert.deepEqual(Object.keys(state.appliedByItemId), []);
  assert.deepEqual(Object.keys(state.originalCoordinatesByItemId), []);
});

test('invalid option input exposes only the fixed safe settings error and disables matching', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [track()], matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  const source = job(jobCore).value;
  controller.initialize(source);
  controller.setOption('utcOffsetText', '+14:01', source);
  const model = controller.getViewModel(source);
  assert.equal(model.canRun, false);
  assert.equal(model.errorCode, 'TRACK_MATCH_OPTIONS_INVALID');
  assert.equal(model.errorMessage, '時刻照合の設定を確認してください。');
  assert.equal(model.errorMessage.includes('+14:01'), false);
});

test('track candidates keep only normalized timed current revisions across supported source types', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const tracks = [
    track('rev-same'),
    track('rev-same'),
    track('rev-old', 0, { trackId: 'track-conflict' }),
    track('rev-new', 1, { trackId: 'track-conflict' }),
    track('rev-geo', 0, { trackId: 'track-geo', sourceType: 'geojson' }),
    track('rev-manual', 0, { trackId: 'track-manual', sourceType: 'manual' }),
    track('rev-empty', 0, {
      trackId: 'track-no-time',
      segments: [{ points: [{ lat: 1, lng: 2, elevation: null, time: '' }] }]
    }),
    track('rev-invalid', 0, { trackId: 'track-invalid', name: '' })
  ];
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => tracks, matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  const source = job(jobCore).value;
  controller.initialize(source);
  const candidates = controller.getViewModel(source).tracks;
  assert.deepEqual(plain(candidates.map((entry) => [entry.trackId, entry.revisionId])), [
    ['track-1', 'rev-same'], ['track-geo', 'rev-geo'], ['track-manual', 'rev-manual']
  ]);
  assert.equal(controller.getViewModel(source).candidateWarningCode, 'TRACK_MATCH_REVISION_CONFLICT');
  candidates.forEach((candidate) => {
    assert.equal(Object.prototype.hasOwnProperty.call(candidate, 'segments'), false);
    assert.equal(Object.prototype.hasOwnProperty.call(candidate, 'sourceName'), false);
  });
});

test('candidate normalization is cached across renders and cleanup for 20,000 points', () => {
  const { jobCore, previewCore, geometry } = loadModules();
  let normalizeCalls = 0;
  const trackedGeometry = {
    normalizeTrack(value) {
      normalizeCalls += 1;
      return geometry.normalizeTrack(value);
    }
  };
  const segments = Array.from({ length: 200 }, (_, segmentIndex) => ({
    points: Array.from({ length: 100 }, (_, pointIndex) => ({
      lat: 30 + segmentIndex / 1000,
      lng: 130 + pointIndex / 1000,
      elevation: null,
      time: new Date(Date.UTC(2026, 0, 1, 0, segmentIndex, pointIndex)).toISOString()
    }))
  }));
  const largeTrack = track('rev-large', 0, { segments });
  const source = jobCore.createJob({
    id: 'job-large', sourceType: 'multi-photo',
    items: Array.from({ length: 20 }, (_, index) => ({
      id: 'p' + index, sourceType: 'photo', title: 'P' + index,
      lat: null, lng: null, capturedAt: '2026-01-01T09:00:00',
      color: '#2196f3', icon: 'photo', uploadStatus: 'queued'
    }))
  });
  const fakeMatchCore = {
    matchPhotos(selectedTrack, photos) {
      return {
        trackId: selectedTrack.trackId, revisionId: selectedTrack.revisionId,
        counts: { total: photos.length }, warnings: [],
        results: photos.map((photo, index) => index === 0 ? {
          itemId: photo.id, status: 'matched-exact', trackId: selectedTrack.trackId,
          revisionId: selectedTrack.revisionId, photoTimeUtc: '2026-01-01T00:00:00.000Z',
          lat: 30, lng: 130, segmentIndex: 0, pointIndex: 0, timeDeltaSeconds: 0
        } : {
          itemId: photo.id, status: 'photo-time-missing', trackId: selectedTrack.trackId,
          revisionId: selectedTrack.revisionId
        })
      };
    }
  };
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [largeTrack], matchCore: fakeMatchCore, trackGeometry: trackedGeometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  controller.initialize(source);
  assert.equal(normalizeCalls, 1);
  controller.getViewModel(source);
  controller.getViewModel(source);
  assert.equal(normalizeCalls, 1);
  controller.run(source);
  controller.setSelected('p0', false);
  controller.getViewModel(source);
  assert.equal(normalizeCalls, 1);
  largeTrack.revisionId = 'rev-large-2';
  assert.equal(controller.getViewModel(source).stale, true);
  assert.equal(normalizeCalls, 2);
  controller.cleanup();
  controller.initialize(source);
  assert.equal(normalizeCalls, 3);
});

test('malformed match results are rejected and valid results are deeply whitelisted in job order', () => {
  const { jobCore, previewCore, geometry } = loadModules();
  const source = job(jobCore).value;
  const baseResults = [
    { itemId: 'b', status: 'matched-exact', trackId: 'track-1', revisionId: 'rev-1',
      photoTimeUtc: '2026-07-12T00:05:00.000Z', lat: 15, lng: 25, elevation: null,
      segmentIndex: 0, pointIndex: 0, timeDeltaSeconds: 0 },
    { itemId: 'a', status: 'skipped-existing-gps', trackId: 'track-1', revisionId: 'rev-1' }
  ];
  function result(results = baseResults) {
    return {
      trackId: 'track-1', revisionId: 'rev-1', results,
      counts: { total: 2 }, warnings: [], arbitrarySecret: 'do-not-retain'
    };
  }
  function createController(matchResult) {
    const state = {};
    const controller = previewCore.create({
      state, getTracks: () => [track()],
      matchCore: { matchPhotos() { return matchResult; } },
      trackGeometry: geometry, importJobCore: jobCore, timezoneOffset: () => -540
    });
    controller.initialize(source);
    return { controller, state };
  }

  const inherited = Object.create({ results: baseResults });
  inherited.trackId = 'track-1';
  inherited.revisionId = 'rev-1';
  inherited.counts = { total: 2 };
  [
    null,
    inherited,
    result([{ ...baseResults[0], status: 'matched-evil' }, baseResults[1]]),
    result([baseResults[0], baseResults[0]]),
    result([{ ...baseResults[0], itemId: 'missing' }, baseResults[1]]),
    result([{ ...baseResults[0], lat: 91 }, baseResults[1]]),
    { ...result(), revisionId: 'rev-other' }
  ].forEach((matchResult) => {
    const { controller, state } = createController(matchResult);
    assert.throws(() => controller.run(source), (error) => error.code === 'TRACK_MATCH_RESULT_INVALID');
    assert.equal(state.result, null);
    assert.equal(state.errorCode, 'TRACK_MATCH_RESULT_INVALID');
  });

  const valid = result(baseResults.map((entry) => ({ ...entry, arbitrarySecret: 'entry-secret' })));
  const { controller, state } = createController(valid);
  controller.run(source);
  assert.deepEqual(Array.from(state.result.results, (entry) => entry.itemId), ['a', 'b']);
  assert.equal(JSON.stringify(state.result).includes('secret'), false);
  assert.deepEqual(Array.from(state.selectedItemIds), ['b']);
});

test('manual coordinate events release ownership even for the same value and removed items release snapshots', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [track()], matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  const source = job(jobCore).value;
  controller.initialize(source);
  controller.run(source);
  const applied = controller.apply(source);
  controller.clearResult();
  const sameValue = jobCore.updateDraftItem(applied, 'b', { lat: applied.items[1].lat });
  controller.onDraftChange(applied, sameValue, { itemId: 'b', field: 'lat' });
  assert.equal(Object.prototype.hasOwnProperty.call(state.appliedByItemId, 'b'), false);

  const removed = jobCore.removeDraftItem(sameValue, 'b', { revokeObjectURL() {} });
  controller.onDraftChange(sameValue, removed, { itemId: 'b', field: 'remove' });
  assert.equal(Object.prototype.hasOwnProperty.call(state.originalCoordinatesByItemId, 'b'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(state.resultInputByItemId, 'b'), false);
  controller.cleanup();
  controller.cleanup();
  assert.deepEqual(Object.keys(state.originalCoordinatesByItemId), []);
});

test('partial GPS is invalid and never becomes an applicable match', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const source = job(jobCore).value;
  const partial = jobCore.updateDraftItem(source, 'b', { lat: null, lng: 25 });
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [track()], matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  controller.initialize(partial);
  const result = controller.run(partial);
  assert.equal(result.results.find((entry) => entry.itemId === 'b').status, 'photo-input-invalid');
  assert.deepEqual(Array.from(state.selectedItemIds), []);
  assert.equal(controller.getViewModel(partial).canApply, false);
});

test('an original valid GPS snapshot remains permanently protected after manual clearing', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const source = job(jobCore).value;
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [track()], matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  controller.initialize(source);
  const clearedGps = jobCore.updateDraftItem(source, 'a', {
    lat: null, lng: null, capturedAt: '2026-07-12T09:05:00'
  });
  controller.onDraftChange(source, clearedGps, { itemId: 'a', field: 'lat' });
  controller.run(clearedGps);
  controller.setSelected('b', false);
  assert.deepEqual(Array.from(state.selectedItemIds), ['a']);
  assert.throws(() => controller.apply(clearedGps),
    (error) => error.code === 'TRACK_MATCH_APPLY_BLOCKED');
  assert.deepEqual(plain(state.originalCoordinatesByItemId.a), { lat: 35, lng: 139 });
});

test('apply validates every selected patch before atomically changing the job or ownership state', () => {
  const { jobCore, previewCore, geometry } = loadModules();
  const source = job(jobCore).value;
  const injected = {
    trackId: 'track-1', revisionId: 'rev-1', counts: { total: 2 }, warnings: [],
    results: [
      { itemId: 'a', status: 'matched-exact', trackId: 'track-1', revisionId: 'rev-1',
        photoTimeUtc: '2026-07-12T00:00:00.000Z', lat: 10, lng: 20, segmentIndex: 0, pointIndex: 0, timeDeltaSeconds: 0 },
      { itemId: 'b', status: 'matched-exact', trackId: 'track-1', revisionId: 'rev-1',
        photoTimeUtc: '2026-07-12T00:05:00.000Z', lat: 15, lng: 25, segmentIndex: 0, pointIndex: 1, timeDeltaSeconds: 0 }
    ]
  };
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [track()], matchCore: { matchPhotos() { return injected; } },
    trackGeometry: geometry, importJobCore: jobCore, timezoneOffset: () => -540
  });
  controller.initialize(source);
  controller.run(source);
  const beforeJob = plain(source);
  const beforeOwnership = plain(state.appliedByItemId);
  assert.throws(() => controller.apply(source), (error) => error.code === 'TRACK_MATCH_APPLY_BLOCKED');
  assert.equal(state.errorCode, 'TRACK_MATCH_APPLY_BLOCKED');
  assert.equal(controller.getViewModel(source).errorMessage, '位置を適用できません。写真の座標を確認してください。');
  assert.deepEqual(plain(source), beforeJob);
  assert.deepEqual(plain(state.appliedByItemId), beforeOwnership);
  assert.deepEqual(Array.from(state.selectedItemIds), ['a', 'b']);
});

test('UTC minus zero is canonicalized in state and remains numeric at the Core boundary', () => {
  const { jobCore, previewCore, geometry } = loadModules();
  const source = job(jobCore).value;
  let receivedOptions = null;
  const controller = previewCore.create({
    state: {}, getTracks: () => [track()],
    matchCore: { matchPhotos(_track, photos, options) {
      receivedOptions = options;
      return {
        trackId: 'track-1', revisionId: 'rev-1', counts: { total: photos.length }, warnings: [],
        results: photos.map((photo) => ({
          itemId: photo.id, status: photo.id === 'a' ? 'skipped-existing-gps' : 'photo-time-missing',
          trackId: 'track-1', revisionId: 'rev-1'
        }))
      };
    } },
    trackGeometry: geometry, importJobCore: jobCore, timezoneOffset: () => -540
  });
  controller.initialize(source);
  controller.setOption('utcOffsetText', '-00:00', source);
  assert.equal(controller.getViewModel(source).utcOffsetText, '+00:00');
  controller.run(source);
  assert.equal(receivedOptions.utcOffsetMinutes, 0);
  assert.equal(typeof receivedOptions.utcOffsetMinutes, 'number');
});

test('matched photos are selected by default and explicit partial selection applies only checked items', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  const source = job(jobCore).value;
  const withSecondMatch = jobCore.addItem(source, {
    id: 'd', sourceType: 'photo', title: 'D', lat: null, lng: null,
    capturedAt: '2026-07-12T09:05:00', color: '#2196f3', icon: 'photo',
    uploadStatus: 'queued', runtime: { uploadFile: { name: 'd.jpg' }, previewUrl: 'blob:d' }
  });
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => [track()], matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  controller.initialize(withSecondMatch);
  controller.run(withSecondMatch);
  assert.deepEqual(Array.from(state.selectedItemIds), ['b', 'd']);
  controller.setSelected('d', false);
  assert.deepEqual(Array.from(state.selectedItemIds), ['b']);
  const applied = controller.apply(withSecondMatch);
  assert.deepEqual(plain(applied.items.map((item) => [item.id, item.lat, item.lng])), [
    ['a', 35, 139], ['b', 15, 25], ['d', null, null]
  ]);
});

test('item identity/order, effective coordinates, track disappearance, and Preview job replacement stale old results', () => {
  const { jobCore, previewCore, matchCore, geometry } = loadModules();
  let tracks = [track()];
  const state = {};
  const controller = previewCore.create({
    state, getTracks: () => tracks, matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  const source = job(jobCore).value;

  function rerun() {
    controller.cleanup();
    controller.initialize(source);
    controller.run(source);
    assert.equal(controller.getViewModel(source).canApply, true);
  }

  rerun();
  const reordered = { ...source, items: [source.items[1], source.items[0]] };
  controller.onDraftChange(source, reordered, { field: 'reorder' });
  assert.equal(controller.getViewModel(reordered).canApply, false);
  assert.throws(() => controller.apply(reordered), (error) => error.code === 'TRACK_MATCH_RESULT_STALE');

  rerun();
  const replacedId = {
    ...source,
    items: source.items.map((item) => item.id === 'b' ? { ...item, id: 'replacement' } : item)
  };
  controller.onDraftChange(source, replacedId, { field: 'id' });
  assert.equal(controller.getViewModel(replacedId).canApply, false);

  rerun();
  const manualCoordinate = jobCore.updateDraftItem(source, 'b', { lat: 40, lng: 140 });
  controller.onDraftChange(source, manualCoordinate, { itemId: 'b', field: 'lat' });
  assert.equal(controller.getViewModel(manualCoordinate).canApply, false);
  assert.throws(() => controller.apply(manualCoordinate), (error) => error.code === 'TRACK_MATCH_RESULT_STALE');

  rerun();
  tracks = [];
  const missingTrackModel = controller.getViewModel(source);
  assert.equal(missingTrackModel.canApply, false);
  assert.equal(missingTrackModel.canRun, false);
  assert.equal(missingTrackModel.trackId, '');
  assert.equal(missingTrackModel.trackRevisionId, '');
  assert.throws(() => controller.apply(source), (error) => error.code === 'TRACK_MATCH_RESULT_STALE');

  tracks = [track()];
  rerun();
  const replacementJob = { ...source, id: 'job-2' };
  controller.onDraftChange(source, replacementJob, { field: 'job' });
  assert.equal(controller.getViewModel(replacementJob).canApply, false);
  assert.throws(() => controller.apply(replacementJob), (error) => error.code === 'TRACK_MATCH_RESULT_STALE');
});

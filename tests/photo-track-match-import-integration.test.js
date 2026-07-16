const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function loadModules() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  const context = {
    console,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#2196f3', label: '青' }],
    PIN_ICONS: [{ id: 'photo', label: '写真' }],
    PIN_STATUSES: ['未対応'],
    URL: { revokeObjectURL() {} },
    Date,
    Math,
    Map,
    Set
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__jobCore = ImportJobCore;\n'
      + 'globalThis.__processor = ImportPhotoItemProcessor;\n'
      + 'globalThis.__workflow = MultiPhotoImportWorkflow;\n'
      + 'globalThis.__flow = ImportFlowController;\n'
      + 'globalThis.__previewCore = PhotoTrackMatchPreviewCore;\n'
      + 'globalThis.__matchCore = PhotoTrackMatchCore;\n'
      + 'globalThis.__geometry = TrackGeometryCore;',
    context
  );
  return {
    jobCore: context.__jobCore,
    processor: context.__processor,
    workflow: context.__workflow,
    flow: context.__flow,
    previewCore: context.__previewCore,
    matchCore: context.__matchCore,
    geometry: context.__geometry
  };
}

test('multi-photo workflow initializes the dedicated matcher, forwards it to Preview, and cleans it up', async () => {
  const { workflow } = loadModules();
  const audit = { initialized: [], cleaned: 0, opens: [], resets: 0 };
  const state = {
    builder: null, controller: null, job: null, resourceUrlApi: null,
    targetFolderId: '', preparing: false, registering: false,
    cancellingPreparation: false, requestToken: 0, trackMatch: {}
  };
  const matchController = {
    initialize(job) { audit.initialized.push(job); },
    cleanup() { audit.cleaned += 1; },
    getViewModel() { return { enabled: true }; }
  };
  const instance = workflow.create({
    state,
    builderApi: {
      create() {
        return {
          start() { return Promise.resolve({ id: 'job-1', sourceType: 'multi-photo', status: 'idle', items: [{ id: 'a' }] }); },
          cancel() {}, release() {}
        };
      },
      getResourceUrlApi() { return { revokeObjectURL() {} }; },
      rememberResourceUrlApi() {}
    },
    processorApi: { create() { return { processItem() {} }; } },
    flowApi: {
      create() {
        return { open(options) { audit.opens.push(options); }, isRunning() { return false; } };
      }
    },
    validator: { validateJob() {} },
    createTrackMatch() { return matchController; },
    resetTrackMatch() { audit.resets += 1; }
  });

  const job = await instance.start([{ name: 'a.jpg' }], {
    targetFolderId: 'folder', tags: [], color: '#2196f3', icon: 'photo', status: ''
  });
  assert.equal(audit.resets, 1);
  assert.deepEqual(audit.initialized, [job]);
  assert.equal(audit.opens[0].trackMatch, matchController);
  audit.opens[0].onClose({ job, discarded: true });
  assert.equal(audit.cleaned >= 1, true);
  assert.equal(state.job, null);
  assert.equal(instance.isBusy(), false);
});

test('ImportFlowController forwards only the dedicated trackMatch extension to ImportPreviewUI', () => {
  const openBlock = indexHtml.slice(
    indexHtml.indexOf('function open(options) {', indexHtml.indexOf('const ImportFlowController')),
    indexHtml.indexOf('function getJob()', indexHtml.indexOf('const ImportFlowController'))
  );
  assert.match(openBlock, /trackMatch:\s*openOptions\.trackMatch/);
  assert.doesNotMatch(openBlock, /renderExtraPanel|innerHTML/);
});

test('existing photo processor persists applied coordinates only and never sends match metadata', async () => {
  const { jobCore, processor, previewCore, matchCore, geometry } = loadModules();
  const uploadFiles = [{ name: 'a.jpg' }, { name: 'b.jpg' }, { name: 'c.jpg' }];
  const job = jobCore.createJob({
    id: 'job-1', sourceType: 'multi-photo', items: [
      { id: 'a', title: 'A', lat: 35, lng: 139, capturedAt: '2026-07-12T09:00:00', color: '#2196f3', icon: 'photo', runtime: { uploadFile: uploadFiles[0] } },
      { id: 'b', title: 'B', lat: null, lng: null, capturedAt: '2026-07-12T09:05:00', color: '#2196f3', icon: 'photo', runtime: { uploadFile: uploadFiles[1] } },
      { id: 'c', title: 'C', lat: null, lng: null, capturedAt: '', color: '#2196f3', icon: 'photo', runtime: { uploadFile: uploadFiles[2] } }
    ]
  });
  const track = {
    trackId: 'track-1', revisionId: 'rev-1', name: '登山記録', description: '',
    color: '#2196f3', sourceType: 'gpx', sourceName: 'track.gpx',
    segments: [{ points: [
      { lat: 10, lng: 20, elevation: null, time: '2026-07-12T00:04:00.000Z' },
      { lat: 20, lng: 30, elevation: null, time: '2026-07-12T00:06:00.000Z' }
    ] }]
  };
  const matchState = {};
  const matchController = previewCore.create({
    state: matchState,
    getTracks: () => [track],
    matchCore,
    trackGeometry: geometry,
    importJobCore: jobCore,
    timezoneOffset: () => -540
  });
  matchController.initialize(job);
  const matchResult = matchController.run(job);
  assert.equal(matchResult.results.find((entry) => entry.itemId === 'a').status, 'skipped-existing-gps');
  assert.equal(matchResult.results.find((entry) => entry.itemId === 'b').status, 'matched-interpolated');
  assert.equal(matchResult.results.find((entry) => entry.itemId === 'c').status, 'photo-time-missing');
  assert.deepEqual(Array.from(matchState.selectedItemIds), ['b']);
  const applied = matchController.apply(job);
  assert.equal(applied.items[0].lat, 35);
  assert.equal(applied.items[2].lat, null);
  const calls = [];
  const itemProcessor = processor.create({
    resizePhoto: async () => 'data:image/jpeg;base64,YWJj',
    callGAS(method, payload) {
      calls.push([method, payload]);
      return Promise.resolve({ ok: true, pin: { id: payload.itemId } });
    },
    withEditToken(payload) { return payload; },
    getTargetFolderId() { return 'folder'; }
  });
  for (const item of applied.items) {
    await itemProcessor.processItem(item, { jobId: applied.id, itemId: item.id, attempt: 1 });
  }

  assert.deepEqual(calls.map((call) => call[0]), [
    'saveImportPhotoItem', 'saveImportPhotoItem', 'saveImportPhotoItem'
  ]);
  assert.deepEqual(calls.map((call) => [call[1].itemId, call[1].lat, call[1].lng]), [
    ['a', 35, 139], ['b', 15, 25], ['c', null, null]
  ]);
  calls.forEach(([, payload]) => {
    assert.equal(Object.prototype.hasOwnProperty.call(payload, 'trackId'), false);
    assert.equal(Object.prototype.hasOwnProperty.call(payload, 'revisionId'), false);
    assert.equal(Object.prototype.hasOwnProperty.call(payload, 'match'), false);
  });
});

test('applied coordinates survive response loss and deduplicated photo retry without metadata or duplicate resources', async () => {
  const { jobCore, processor, previewCore, matchCore, geometry } = loadModules();
  const uploadFile = { name: 'matched.jpg' };
  const source = jobCore.createJob({
    id: 'job-response-loss', sourceType: 'multi-photo', items: [
      { id: 'matched', title: 'Matched', lat: null, lng: null,
        capturedAt: '2026-07-12T09:05:00', color: '#2196f3', icon: 'photo',
        runtime: { originalFile: uploadFile, uploadFile, previewUrl: 'blob:matched' } },
      { id: 'other', title: 'Other', lat: null, lng: null, capturedAt: '',
        color: '#2196f3', icon: 'photo', runtime: { previewUrl: 'blob:other' } }
    ]
  });
  const savedTrack = {
    trackId: 'track-1', revisionId: 'rev-1', name: '登山記録', description: '',
    color: '#2196f3', sourceType: 'gpx', sourceName: 'private.gpx',
    segments: [{ points: [
      { lat: 10, lng: 20, elevation: null, time: '2026-07-12T00:04:00.000Z' },
      { lat: 20, lng: 30, elevation: null, time: '2026-07-12T00:06:00.000Z' }
    ] }]
  };
  const matchController = previewCore.create({
    state: {}, getTracks: () => [savedTrack], matchCore, trackGeometry: geometry,
    importJobCore: jobCore, timezoneOffset: () => -540
  });
  matchController.initialize(source);
  matchController.run(source);
  const applied = matchController.apply(source);
  assert.deepEqual([applied.items[0].lat, applied.items[0].lng], [15, 25]);
  assert.deepEqual([applied.items[1].lat, applied.items[1].lng], [null, null]);

  const server = { driveFiles: 0, mapRows: 0, receipt: null, calls: [] };
  const itemProcessor = processor.create({
    resizePhoto: async () => 'data:image/jpeg;base64,YWJj',
    async callGAS(method, payload) {
      assert.equal(method, 'saveImportPhotoItem');
      server.calls.push({ ...payload });
      if (!server.receipt) {
        server.driveFiles += 1;
        server.mapRows += 1;
        server.receipt = { pin: { id: 'pin-1', lat: payload.lat, lng: payload.lng } };
        throw new Error('simulated response loss');
      }
      return { ok: true, pin: server.receipt.pin, deduplicated: true };
    },
    withEditToken(payload) { return { ...payload, __editToken: 'latest-token' }; },
    getTargetFolderId() { return 'folder'; }
  });

  await assert.rejects(() => itemProcessor.processItem(applied.items[0], {
    jobId: applied.id, itemId: 'matched', attempt: 1
  }), (error) => error.code === 'IMPORT_ITEM_SAVE_FAILED'
    && error.message === '写真を保存できませんでした。再試行してください。'
    && error.retryable === true
    && !/response loss/.test(error.message));
  const pin = await itemProcessor.processItem(applied.items[0], {
    jobId: applied.id, itemId: 'matched', attempt: 2
  });
  assert.deepEqual(pin, { id: 'pin-1', lat: 15, lng: 25 });
  assert.equal(server.driveFiles, 1);
  assert.equal(server.mapRows, 1);
  assert.equal(server.calls.length, 2);
  assert.deepEqual(server.calls.map((payload) => [payload.jobId, payload.itemId, payload.lat, payload.lng]), [
    ['job-response-loss', 'matched', 15, 25],
    ['job-response-loss', 'matched', 15, 25]
  ]);
  ['trackId', 'revisionId', 'ratio', 'photoTimeUtc', 'result',
    'utcOffsetMinutes', 'clockCorrectionSeconds'].forEach((key) => {
    server.calls.forEach((payload) => assert.equal(Object.prototype.hasOwnProperty.call(payload, key), false));
  });
  server.calls.forEach((payload) => assert.equal(payload.status, ''));

  const revoked = [];
  jobCore.releaseJobResources(applied, { revokeObjectURL(url) { revoked.push(url); } });
  jobCore.releaseJobResources(applied, { revokeObjectURL(url) { revoked.push(url); } });
  assert.deepEqual(revoked.sort(), ['blob:matched', 'blob:other']);
});

test('production wiring keeps match state inside multi-photo state and reuses current tracks without new routes', () => {
  const stateBlock = indexHtml.slice(indexHtml.indexOf('multiPhotoImport: {'), indexHtml.indexOf('csvInterchange: {'));
  assert.match(stateBlock, /trackMatch:/);
  assert.match(indexHtml, /createTrackMatch:[\s\S]*getTracks:\s*function\(\)\s*\{\s*return state\.tracks/);
  assert.equal((indexHtml.match(/saveImportPhotoItem/g) || []).length > 0, true);
  assert.doesNotMatch(indexHtml, /savePhotoTrackMatch|saveMatchedPhotoLocations/);
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function loadWorkflow() {
  const start = indexHtml.indexOf('const AudioPinImportWorkflow =');
  const end = indexHtml.indexOf('const DrivePhotoImportSourceCore =', start);
  assert.notEqual(start, -1, 'AudioPinImportWorkflow must exist');
  assert.notEqual(end, -1, 'DrivePhotoImportSourceCore boundary must exist');
  const context = { console, Blob, Uint8Array, Promise };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nglobalThis.__api = AudioPinImportWorkflow;`, context);
  return { api: context.__api, source: indexHtml.slice(start, end) };
}

function operation(overrides = {}) {
  return {
    mediaKind: 'audio', sourceKind: 'local', operationMode: 'create-pin',
    selectionLimit: 1, returnOverlayId: 'add-menu-overlay',
    jobId: 'audio-job-1', itemId: 'audio-item-1',
    idempotencyKey: 'audio-job-1:audio-item-1', sourceDriveFileId: '', status: 'active',
    ...overrides
  };
}

function result() {
  const blob = new Blob([new Uint8Array([0x49, 0x44, 0x33, 4, 0, 0, 0, 0, 0, 0])], { type: 'audio/mpeg' });
  return {
    blob, fileName: 'trimmed.mp3', mimeType: 'audio/mpeg', sizeBytes: blob.size,
    sourceDurationSeconds: 45, durationSeconds: 30, selectionStart: 0, selectionEnd: 30,
    sampleRate: 48000, bitrate: 192000, numberOfChannels: 1
  };
}

function draftFields(title) {
  return {
    title, description: 'memo', eventAt: '2026-07-22T10:20',
    color: '#e53935', icon: 'audio', status: '未対応', tags: ['field'],
    links: ['https://example.com']
  };
}

test('start keeps the editor Blob client-only and opens editable Preview with filename-stem title', async () => {
  const { api, source } = loadWorkflow();
  const calls = [];
  const forbidden = [];
  const workflow = api.create({
    callGAS: async (...args) => { calls.push(args); return { ok: true }; },
    openPreview: (draft) => calls.push(['preview', draft.title]),
    readAudioMetadata: () => forbidden.push('metadata'),
    readExifGps: () => forbidden.push('exif'),
    geocodeFilename: () => forbidden.push('geocoder')
  });
  const draft = workflow.start({
    operation: operation(), sourceFileName: 'Tokyo field.note.MP3', editorResult: result()
  });

  assert.equal(draft.title, 'Tokyo field.note');
  draft.title = '利用者が編集したタイトル';
  assert.deepEqual(calls, [['preview', 'Tokyo field.note']]);
  assert.deepEqual(forbidden, []);
  assert.equal(workflow.hasPendingDraft(), true);
  assert.doesNotMatch(source, /readAudioMetadata|readExifGps|geocodeFilename|EXIF/);
});

test('save requires an explicit map or unplaced choice and is idempotent in the client', async () => {
  const { api } = loadWorkflow();
  const gasCalls = [];
  const saved = [];
  const workflow = api.create({
    withEditToken: (payload) => ({ ...payload, __editToken: 'send-only' }),
    callGAS: async (method, payload) => {
      gasCalls.push([method, payload]);
      return { ok: true, pin: { id: 'pin-audio-1', title: payload.pin.title } };
    },
    onSaved: (pin) => saved.push(pin.id)
  });
  const draft = workflow.start({
    operation: operation(), sourceFileName: 'voice.wav', editorResult: result()
  });
  Object.assign(draft, draftFields('edited'));

  await assert.rejects(workflow.save(draft), (error) => error && error.code === 'AUDIO_LOCATION_REQUIRED');
  assert.equal(gasCalls.length, 0);
  workflow.setLocationChoice({ kind: 'unplaced' });
  const [first, second] = await Promise.all([workflow.save(draft), workflow.save(draft)]);
  const third = await workflow.save(draft);

  assert.equal(first.pin.id, 'pin-audio-1');
  assert.equal(second.pin.id, 'pin-audio-1');
  assert.equal(third.pin.id, 'pin-audio-1');
  assert.equal(gasCalls.length, 1);
  assert.equal(gasCalls[0][0], 'saveImportAudioItem');
  assert.equal(gasCalls[0][1].idempotencyKey, 'audio-job-1:audio-item-1');
  assert.equal(gasCalls[0][1].sourceKind, 'local');
  assert.equal(gasCalls[0][1].sourceDriveFileId, '');
  assert.equal(gasCalls[0][1].sourceFileName, 'voice.wav');
  assert.equal(gasCalls[0][1].audioMimeType, 'audio/mpeg');
  assert.match(gasCalls[0][1].audioBase64, /^[A-Za-z0-9+/]+={0,2}$/);
  assert.deepEqual(JSON.parse(JSON.stringify(gasCalls[0][1].pin)), {
    title: 'edited', description: 'memo', eventTime: '2026-07-22T10:20',
    lat: null, lng: null, color: '#e53935', icon: 'audio', status: '未対応',
    tags: ['field'], links: ['https://example.com']
  });
  assert.equal(gasCalls[0][1].__editToken, 'send-only');
  assert.deepEqual(saved, ['pin-audio-1']);
});

test('map choice is the only source of coordinates even when Drive filename looks like a place', async () => {
  const { api } = loadWorkflow();
  let payload;
  const workflow = api.create({
    callGAS: async (_method, value) => { payload = value; return { ok: true, pin: { id: 'pin-map' } }; }
  });
  const draft = workflow.start({
    operation: operation({ sourceKind: 'drive', sourceDriveFileId: 'drive_AAAAAAAAAAA' }),
    sourceFileName: '東京駅.m4a', editorResult: result()
  });
  Object.assign(draft, draftFields(draft.title));
  workflow.setLocationChoice({ kind: 'map', lat: 35.681236, lng: 139.767125 });
  await workflow.save(draft);

  assert.equal(payload.pin.lat, 35.681236);
  assert.equal(payload.pin.lng, 139.767125);
  assert.equal(payload.sourceDriveFileId, 'drive_AAAAAAAAAAA');
  assert.equal(payload.sourceFileName, '東京駅.m4a');
});

test('response failure retains draft/blob for same-id retry while cancel releases exactly once', async () => {
  const { api } = loadWorkflow();
  const payloads = [];
  let attempts = 0;
  let releases = 0;
  const workflow = api.create({
    callGAS: async (_method, payload) => {
      attempts += 1;
      payloads.push(payload);
      if (attempts === 1) return { ok: false, error: '一時的な失敗', retryable: true };
      return { ok: true, pin: { id: 'pin-retry' } };
    },
    releaseEditorResult: () => { releases += 1; }
  });
  const draft = workflow.start({ operation: operation(), sourceFileName: 'retry.mp3', editorResult: result() });
  Object.assign(draft, draftFields('retry'));
  workflow.setLocationChoice({ kind: 'unplaced' });
  await assert.rejects(workflow.save(draft));
  assert.equal(workflow.hasPendingDraft(), true);
  assert.equal(releases, 0);
  await workflow.retry();
  assert.equal(payloads.length, 2);
  assert.equal(payloads[0].idempotencyKey, payloads[1].idempotencyKey);
  assert.equal(releases, 1);
  assert.equal(workflow.cancel(), false);
});

test('cleanup warning remains a successful save and cancel before save releases client state', async () => {
  const { api } = loadWorkflow();
  const warnings = [];
  let releases = 0;
  const workflow = api.create({
    callGAS: async () => ({
      ok: true, pin: { id: 'pin-cleanup' }, cleanupRequired: true,
      warning: '元音声の整理を再試行してください。'
    }),
    onCleanupWarning: (message) => warnings.push(message),
    releaseEditorResult: () => { releases += 1; }
  });
  let draft = workflow.start({ operation: operation(), sourceFileName: 'cleanup.wav', editorResult: result() });
  Object.assign(draft, draftFields('cleanup'));
  workflow.setLocationChoice({ kind: 'unplaced' });
  const response = await workflow.save(draft);
  assert.equal(response.ok, true);
  assert.deepEqual(warnings, ['元音声の整理を再試行してください。']);
  assert.equal(releases, 1);

  draft = workflow.start({ operation: operation({ jobId: 'audio-job-2', idempotencyKey: 'audio-job-2:audio-item-1' }), sourceFileName: 'cancel.wav', editorResult: result() });
  assert.equal(workflow.cancel(), true);
  assert.equal(workflow.cancel(), false);
  assert.equal(releases, 2);
});

test('valid save commits before fallible observers and releases the editor result exactly once', async () => {
  const { api } = loadWorkflow();
  let gasCalls = 0;
  let releases = 0;
  let savedCalls = 0;
  let warningCalls = 0;
  const response = {
    ok: true,
    pin: {
      id: 'pin-observer-safe',
      get title() { throw new Error('nonessential pin getter failed'); }
    },
    cleanupRequired: true,
    get warning() { throw new Error('warning getter failed'); }
  };
  const workflow = api.create({
    callGAS: async () => { gasCalls += 1; return response; },
    onSaved(pin) {
      savedCalls += 1;
      void pin.title;
    },
    onCleanupWarning() {
      warningCalls += 1;
      throw new Error('warning observer failed');
    },
    releaseEditorResult() {
      releases += 1;
      throw new Error('release observer failed');
    }
  });
  const draft = workflow.start({ operation: operation(), sourceFileName: 'observer.mp3', editorResult: result() });
  Object.assign(draft, draftFields('observer'));
  workflow.setLocationChoice({ kind: 'unplaced' });

  const first = await workflow.save(draft);
  const replay = await workflow.save(draft);
  assert.equal(first, response);
  assert.equal(replay, response);
  assert.equal(gasCalls, 1);
  assert.equal(savedCalls, 1);
  assert.equal(warningCalls, 0, 'a throwing warning getter is isolated before callback dispatch');
  assert.equal(releases, 1);
  assert.equal(workflow.hasPendingDraft(), false);
});

test('a malformed pin id remains a genuine save failure and retains the draft for retry', async () => {
  const { api } = loadWorkflow();
  let releases = 0;
  const workflow = api.create({
    callGAS: async () => ({ ok: true, pin: { get id() { throw new Error('bad id'); } } }),
    releaseEditorResult() { releases += 1; }
  });
  const draft = workflow.start({ operation: operation(), sourceFileName: 'invalid.mp3', editorResult: result() });
  Object.assign(draft, draftFields('invalid'));
  workflow.setLocationChoice({ kind: 'unplaced' });

  await assert.rejects(workflow.save(draft));
  assert.equal(workflow.hasPendingDraft(), true);
  assert.equal(releases, 0);
});

test('committed save observes and catches rejected callback thenables without awaiting them', async () => {
  const { api } = loadWorkflow();
  const observed = [];
  let releaseThenReads = 0;
  function rejectedThenable(label) {
    return {
      then(_resolve, reject) {
        observed.push(label);
        reject(new Error(`${label} rejected`));
      }
    };
  }
  const hostileReleaseThenable = {};
  Object.defineProperty(hostileReleaseThenable, 'then', {
    get() {
      releaseThenReads += 1;
      throw new Error('hostile then getter');
    }
  });
  const response = {
    ok: true, pin: { id: 'pin-async-observers' }, cleanupRequired: true, warning: 'cleanup'
  };
  const workflow = api.create({
    callGAS: async () => response,
    onSaved() { return rejectedThenable('saved'); },
    onCleanupWarning() { return rejectedThenable('warning'); },
    releaseEditorResult() { return hostileReleaseThenable; }
  });
  const draft = workflow.start({ operation: operation(), sourceFileName: 'async.mp3', editorResult: result() });
  Object.assign(draft, draftFields('async'));
  workflow.setLocationChoice({ kind: 'unplaced' });

  const saved = await workflow.save(draft);
  await Promise.resolve();
  await Promise.resolve();

  assert.equal(saved, response);
  assert.deepEqual(observed.sort(), ['saved', 'warning']);
  assert.equal(releaseThenReads, 1);
  assert.equal(workflow.hasPendingDraft(), false);
  assert.equal(await workflow.save(draft), response);
});

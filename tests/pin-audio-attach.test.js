const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

function countId(id) {
  return (indexHtml.match(new RegExp(`id=["']${id}["']`, 'g')) || []).length;
}

function functionSource(source, name) {
  let start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  if (source.slice(Math.max(0, start - 6), start) === 'async ') start -= 6;
  const bodyStart = source.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < source.length; index += 1) {
    const character = source[index];
    if (quote) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'" || character === '`') {
      quote = character;
      continue;
    }
    if (character === '{') depth += 1;
    if (character === '}') {
      depth -= 1;
      if (depth === 0) return source.slice(start, index + 1);
    }
  }
  assert.fail(`Could not parse ${name}`);
}

function loadMediaImportShell() {
  const start = indexHtml.indexOf('const MediaImportShell =');
  const end = indexHtml.indexOf('const ImportAudioItemProcessor =', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = { console, Promise, Date, Math };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nthis.api = MediaImportShell;`, context);
  return context.api;
}

function loadAudioWorkflow() {
  const start = indexHtml.indexOf('const AudioPinImportWorkflow =');
  const end = indexHtml.indexOf('const DrivePhotoImportSourceCore =', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = { console, Blob, Uint8Array, Promise };
  vm.runInNewContext(`${indexHtml.slice(start, end)}\nthis.api = AudioPinImportWorkflow;`, context);
  return context.api;
}

function editorResult() {
  const blob = new Blob([
    new Uint8Array([0x49, 0x44, 0x33, 4, 0, 0, 0, 0, 0, 0])
  ], { type: 'audio/mpeg' });
  return {
    blob,
    fileName: 'trimmed.mp3',
    mimeType: 'audio/mpeg',
    sizeBytes: blob.size,
    sourceDurationSeconds: 45,
    durationSeconds: 30,
    selectionStart: 0,
    selectionEnd: 30,
    sampleRate: 48000,
    bitrate: 192000,
    numberOfChannels: 1
  };
}

function existingOperation(mode = 'attach-existing-pin') {
  return {
    mediaKind: 'audio',
    sourceKind: 'local',
    operationMode: mode,
    selectionLimit: 1,
    returnOverlayId: 'pin-audio-source-overlay',
    targetPinId: 'pin-existing-1',
    expectedUpdatedAt: '2026-07-22T10:00:00+09:00',
    jobId: 'audio-job-existing',
    itemId: 'audio-item-existing',
    idempotencyKey: 'audio-job-existing:audio-item-existing',
    sourceDriveFileId: '',
    status: 'active'
  };
}

test('editable pin detail exposes mutually exclusive audio actions and a single-file source chooser', () => {
  [
    'pin-detail-audio-actions', 'pin-detail-audio-add', 'pin-detail-audio-replace',
    'pin-detail-audio-delete', 'pin-audio-source-overlay', 'pin-audio-source-title',
    'pin-audio-source-target', 'pin-audio-source-local', 'pin-audio-source-drive',
    'pin-audio-source-cancel', 'pin-audio-source-error', 'pin-audio-attach-file-input'
  ].forEach((id) => assert.equal(countId(id), 1, id));

  const input = indexHtml.match(/<input[^>]+id="pin-audio-attach-file-input"[^>]*>/)[0];
  assert.match(input, /accept="[^"]*(?:\.m4a|audio\/mp4)[^"]*"/);
  assert.match(input, /(?:\.mp3|audio\/mpeg)/);
  assert.match(input, /(?:\.wav|audio\/wav)/);
  assert.doesNotMatch(input, /\bmultiple\b/);
  assert.equal(sharedHtml.includes('pin-detail-audio-add'), false);
  assert.equal(sharedHtml.includes('pin-audio-source-overlay'), false);
});

test('pin detail audio renderer gates add versus replace/delete by edit access and hasAudio', () => {
  const elements = Object.fromEntries([
    'pin-detail-audio-actions', 'pin-detail-audio-add', 'pin-detail-audio-replace',
    'pin-detail-audio-delete'
  ].map((id) => [id, { id, style: {}, disabled: false, onclick: null }]));
  let editable = true;
  const opened = [];
  const removed = [];
  const context = {
    document: { getElementById: (id) => elements[id] },
    canEdit: () => editable,
    openPinAudioSource: (...args) => opened.push(args),
    removeAudioFromPinDetail: (...args) => removed.push(args)
  };
  vm.runInNewContext(
    `${functionSource(indexHtml, 'renderPinDetailAudioActions')}; this.render = renderPinDetailAudioActions;`,
    context
  );
  const snapshot = { pinId: 'pin-1', expectedUpdatedAt: 'v1' };

  context.render({ id: 'pin-1', hasAudio: false }, snapshot);
  assert.equal(elements['pin-detail-audio-actions'].style.display, '');
  assert.equal(elements['pin-detail-audio-add'].style.display, '');
  assert.equal(elements['pin-detail-audio-replace'].style.display, 'none');
  assert.equal(elements['pin-detail-audio-delete'].style.display, 'none');
  elements['pin-detail-audio-add'].onclick();
  assert.equal(opened[0][1], 'attach-existing-pin');

  context.render({ id: 'pin-1', hasAudio: true }, snapshot);
  assert.equal(elements['pin-detail-audio-add'].style.display, 'none');
  assert.equal(elements['pin-detail-audio-replace'].style.display, '');
  assert.equal(elements['pin-detail-audio-delete'].style.display, '');
  elements['pin-detail-audio-replace'].onclick();
  elements['pin-detail-audio-delete'].onclick();
  assert.equal(opened[1][1], 'replace-existing-audio');
  assert.equal(removed.length, 1);

  editable = false;
  context.render({ id: 'pin-1', hasAudio: false }, snapshot);
  assert.equal(elements['pin-detail-audio-actions'].style.display, 'none');
  assert.equal(elements['pin-detail-audio-add'].onclick, null);
});

test('MediaImportShell snapshots existing-pin CAS fields into an immutable operation', () => {
  const shell = loadMediaImportShell().create({
    createId: (prefix) => `${prefix}-fixed`
  });
  const operation = shell.begin({
    mediaKind: 'audio', sourceKind: 'drive', operationMode: 'replace-existing-audio',
    targetPinId: 'pin-existing-1', expectedUpdatedAt: 'snapshot-v1',
    selectionLimit: 1, returnOverlayId: 'pin-audio-source-overlay'
  });

  assert.equal(operation.targetPinId, 'pin-existing-1');
  assert.equal(operation.expectedUpdatedAt, 'snapshot-v1');
  assert.equal(Object.isFrozen(operation), true);
});

test('existing-pin audio workflow skips location and sends only target CAS fields', async () => {
  const api = loadAudioWorkflow();
  const gasCalls = [];
  const previews = [];
  const saved = [];
  const inputOperation = existingOperation();
  const targetPin = {
    id: 'pin-existing-1', title: '元のタイトル', description: '元の説明',
    lat: 35, lng: 135, updatedAt: inputOperation.expectedUpdatedAt,
    color: '#123456', icon: 'default', status: '対応中', tags: ['keep'],
    links: ['https://example.com'], eventAt: '2026-07-22T09:00'
  };
  const workflow = api.create({
    withEditToken: (payload) => ({ ...payload, __editToken: 'token' }),
    openPreview: (draft, options) => previews.push({ draft: { ...draft }, options }),
    callGAS: async (method, payload) => {
      gasCalls.push([method, payload]);
      return {
        ok: true,
        pin: { ...targetPin, hasAudio: true, updatedAt: 'snapshot-v2' }
      };
    },
    onSaved: (pin) => saved.push(pin)
  });
  const draft = workflow.start({
    operation: inputOperation,
    targetPin,
    sourceFileName: 'voice.wav',
    editorResult: editorResult()
  });
  inputOperation.expectedUpdatedAt = 'mutated-after-open';
  targetPin.title = 'mutated title';

  assert.equal(previews[0].options.readOnlyFields, true);
  assert.equal(previews[0].options.operationMode, 'attach-existing-pin');
  assert.equal(draft.title, '元のタイトル');
  const response = await workflow.save(draft);

  assert.equal(response.pin.hasAudio, true);
  assert.equal(gasCalls.length, 1);
  assert.equal(gasCalls[0][0], 'saveImportAudioItem');
  assert.equal(gasCalls[0][1].operationMode, 'attach-existing-pin');
  assert.equal(gasCalls[0][1].targetPinId, 'pin-existing-1');
  assert.equal(gasCalls[0][1].expectedUpdatedAt, '2026-07-22T10:00:00+09:00');
  assert.equal(Object.hasOwn(gasCalls[0][1], 'pin'), false);
  assert.equal(Object.hasOwn(gasCalls[0][1], 'lat'), false);
  assert.equal(Object.hasOwn(gasCalls[0][1], 'lng'), false);
  assert.equal(saved.length, 1);
});

test('existing-pin source operation retains the detail snapshot for both local and Drive', () => {
  const begins = [];
  const state = {
    pinAudioAttach: {
      targetPinId: 'pin-1', expectedUpdatedAt: 'detail-v1',
      operationMode: 'replace-existing-audio'
    },
    audioImport: { requestToken: 0 }
  };
  const context = {
    state,
    MediaImportShell: {
      create() {
        return {
          begin(value) {
            begins.push(value);
            return Object.freeze({ ...value, jobId: 'job', itemId: 'item', idempotencyKey: 'job:item' });
          }
        };
      }
    }
  };
  vm.runInNewContext(
    `${functionSource(indexHtml, 'createAudioImportOperation')}; this.createOperation = createAudioImportOperation;`,
    context
  );

  context.createOperation('local');
  context.createOperation('drive');
  assert.deepEqual(begins.map((value) => ({
    sourceKind: value.sourceKind,
    operationMode: value.operationMode,
    targetPinId: value.targetPinId,
    expectedUpdatedAt: value.expectedUpdatedAt,
    selectionLimit: value.selectionLimit,
    returnOverlayId: value.returnOverlayId
  })), [
    {
      sourceKind: 'local', operationMode: 'replace-existing-audio',
      targetPinId: 'pin-1', expectedUpdatedAt: 'detail-v1', selectionLimit: 1,
      returnOverlayId: 'pin-audio-source-overlay'
    },
    {
      sourceKind: 'drive', operationMode: 'replace-existing-audio',
      targetPinId: 'pin-1', expectedUpdatedAt: 'detail-v1', selectionLimit: 1,
      returnOverlayId: 'pin-audio-source-overlay'
    }
  ]);
});

test('audio removal uses the detail-open CAS snapshot and updates client state without reload', async () => {
  const calls = [];
  const updated = [];
  const hints = [];
  const invalidated = [];
  const context = {
    requestAppConfirmation: async () => true,
    withEditToken: (payload) => ({ ...payload, __editToken: 'token' }),
    withGAS: async (method, payload) => {
      calls.push([method, payload]);
      return {
        ok: true,
        pin: { id: 'pin-1', title: 'keep', hasAudio: false, updatedAt: 'detail-v2' },
        cleanupRequired: true,
        warning: '後処理を再試行します。'
      };
    },
    upsertImportedPin: (pin) => updated.push(pin),
    showTransientHint: (message) => hints.push(message),
    showAppNotification: () => assert.fail('success must not show an error'),
    pinAudioPlayer: { invalidate: (pinId) => invalidated.push(pinId) },
    openPinDetail: () => {},
    getPinById: () => ({ id: 'pin-1', hasAudio: false })
  };
  vm.runInNewContext(
    `${functionSource(indexHtml, 'removeAudioFromPinDetail')}; this.removeAudio = removeAudioFromPinDetail;`,
    context
  );

  const result = await context.removeAudio(
    { id: 'pin-1', title: 'keep', hasAudio: true, updatedAt: 'mutated-live' },
    { pinId: 'pin-1', expectedUpdatedAt: 'detail-v1' }
  );
  assert.equal(result, true);
  assert.equal(calls[0][0], 'removePinAudio');
  assert.equal(calls[0][1].pinId, 'pin-1');
  assert.equal(calls[0][1].expectedUpdatedAt, 'detail-v1');
  assert.equal(updated[0].hasAudio, false);
  assert.deepEqual(invalidated, ['pin-1']);
  assert.deepEqual(hints, ['後処理を再試行します。']);
});

test('audio removal conflict never overwrites the pin and explicitly asks for reload', async () => {
  const updated = [];
  const notices = [];
  const context = {
    requestAppConfirmation: async () => true,
    withEditToken: (payload) => payload,
    withGAS: async () => ({
      ok: false, errorCode: 'PIN_AUDIO_CONFLICT',
      error: 'ピンが別の操作で更新されました。画面を再読み込みしてから再試行してください。'
    }),
    upsertImportedPin: (pin) => updated.push(pin),
    showTransientHint: () => {},
    showAppNotification: (notice) => notices.push(notice),
    openPinDetail: () => {},
    getPinById: () => ({ id: 'pin-1', hasAudio: true })
  };
  vm.runInNewContext(
    `${functionSource(indexHtml, 'removeAudioFromPinDetail')}; this.removeAudio = removeAudioFromPinDetail;`,
    context
  );

  const result = await context.removeAudio(
    { id: 'pin-1', hasAudio: true },
    { pinId: 'pin-1', expectedUpdatedAt: 'detail-v1' }
  );
  assert.equal(result, false);
  assert.deepEqual(updated, []);
  assert.match(`${notices[0].title}\n${notices[0].message}`, /再読み込み/);
});

test('pin delete warning names every managed attachment combination', () => {
  const context = {};
  vm.runInNewContext(
    `${functionSource(indexHtml, 'pinDeleteAttachmentMessage')}; this.messageFor = pinDeleteAttachmentMessage;`,
    context
  );

  assert.match(context.messageFor({ fileId: 'photo', hasAudio: true }), /写真.*音声|音声.*写真/);
  assert.match(context.messageFor({ fileId: 'photo', hasAudio: false }), /写真/);
  assert.doesNotMatch(context.messageFor({ fileId: 'photo', hasAudio: false }), /音声/);
  assert.match(context.messageFor({ fileId: '', hasAudio: true }), /音声/);
  assert.doesNotMatch(context.messageFor({ fileId: '', hasAudio: true }), /写真/);
  assert.doesNotMatch(context.messageFor({ fileId: '', hasAudio: false }), /写真|音声/);
});

test('client pin cloning preserves only the public hasAudio flag', () => {
  const context = {
    safeColor: (value) => value,
    normalizeIcon: (value) => value
  };
  vm.runInNewContext(
    `${functionSource(indexHtml, 'clonePin')}; this.clone = clonePin;`,
    context
  );

  const cloned = context.clone({
    id: 'pin-audio', title: 'Audio', color: '#123456', icon: 'default',
    hasAudio: true, audioId: 'server-secret-audio-id'
  });
  assert.equal(cloned.hasAudio, true);
  assert.equal(Object.hasOwn(cloned, 'audioId'), false);
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

class FakeClassList {
  constructor() { this.values = new Set(); }
  add(value) { this.values.add(value); }
  remove(value) { this.values.delete(value); }
  contains(value) { return this.values.has(value); }
  toggle(value, force) {
    const enabled = force == null ? !this.contains(value) : !!force;
    if (enabled) this.add(value); else this.remove(value);
    return enabled;
  }
}

class FakeElement {
  constructor(tagName, id) {
    this.tagName = String(tagName || 'div').toUpperCase();
    this.id = id || '';
    this.classList = new FakeClassList();
    this.children = [];
    this.parentNode = null;
    this.dataset = {};
    this.style = {};
    this.attributes = {};
    this.listeners = Object.create(null);
    this.disabled = false;
    this.value = '';
    this._textContent = '';
  }
  set textContent(value) {
    this._textContent = value == null ? '' : String(value);
    this.children = [];
  }
  get textContent() {
    return this._textContent + this.children.map((child) => child.textContent).join('');
  }
  set innerHTML(_value) { throw new Error('Phase 3 UI must not write innerHTML'); }
  get innerHTML() { return ''; }
  appendChild(child) {
    child.parentNode = this;
    this.children.push(child);
    return child;
  }
  replaceChildren(...children) {
    this.children = [];
    children.forEach((child) => this.appendChild(child));
  }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  removeAttribute(name) { delete this.attributes[name]; }
  addEventListener(type, listener) {
    if (!this.listeners[type]) this.listeners[type] = [];
    this.listeners[type].push(listener);
  }
  dispatch(type, target = this) {
    const event = { type, target, preventDefault() {} };
    (this.listeners[type] || []).forEach((listener) => listener(event));
  }
  closest(selector) {
    if (selector === '[data-import-item-id]' && this.dataset.importItemId) return this;
    return this.parentNode ? this.parentNode.closest(selector) : null;
  }
  focus() { this.focused = true; }
}

function createDocument() {
  const ids = [
    'import-preview-overlay', 'import-preview-sheet', 'import-preview-title', 'import-preview-source',
    'import-preview-job-status', 'import-preview-operation-note', 'import-preview-operation-error',
    'import-preview-completion-message', 'import-preview-presets', 'import-preview-preset-select',
    'import-preview-preset-apply-selected', 'import-preview-preset-apply-all',
    'import-preview-preset-status', 'import-preview-preset-error', 'import-preview-preset-retry',
    'import-preview-count-total', 'import-preview-count-waiting',
    'import-preview-count-processing', 'import-preview-count-succeeded',
    'import-preview-count-failed', 'import-preview-filter-all',
    'import-preview-filter-needs-review', 'import-preview-filter-processing',
    'import-preview-filter-succeeded', 'import-preview-filter-failed',
    'import-preview-list', 'import-preview-empty',
    'import-preview-editor', 'import-preview-photo-pane', 'import-preview-image-trigger', 'import-preview-image',
    'import-preview-location-summary', 'import-preview-item-status',
    'import-preview-item-error', 'import-preview-edit-title', 'import-preview-edit-description',
    'import-preview-edit-lat', 'import-preview-edit-lng', 'import-preview-time-field-label',
    'import-preview-edit-captured-at', 'import-preview-edit-links',
    'import-preview-edit-tags', 'import-preview-edit-color', 'import-preview-color-preview',
    'import-preview-edit-icon', 'import-preview-icon-preview', 'import-preview-edit-status',
    'import-preview-edit-metadata-status', 'import-preview-edit-conversion-status',
    'import-preview-delete', 'import-preview-primary', 'import-preview-cancel',
    'import-preview-resume', 'import-preview-retry', 'import-preview-close', 'import-preview-discard',
    'multi-photo-track-match-panel', 'multi-photo-track-select', 'multi-photo-track-utc-offset',
    'multi-photo-track-clock-correction', 'multi-photo-track-max-gap',
    'multi-photo-track-endpoint-tolerance', 'multi-photo-track-run',
    'multi-photo-track-status', 'multi-photo-track-error', 'multi-photo-track-counts',
    'multi-photo-track-warnings', 'multi-photo-track-results', 'multi-photo-track-apply',
    'multi-photo-track-clear'
  ];
  const elements = Object.create(null);
  ids.forEach((id) => {
    const tag = id === 'import-preview-image' ? 'img'
      : id === 'import-preview-image-trigger' ? 'button'
      : id === 'import-preview-preset-select' || id === 'import-preview-edit-status'
          || id === 'import-preview-edit-color' || id === 'import-preview-edit-icon' ? 'select'
        : id.startsWith('import-preview-edit-') ? 'input'
          : /(?:delete|primary|cancel|resume|retry|apply-selected|apply-all|close|discard)$/.test(id)
            ? 'button' : 'div';
    elements[id] = new FakeElement(tag, id);
  });
  elements['import-preview-overlay'].classList.add('sheet-overlay');
  const fields = {
    'import-preview-edit-title': 'title', 'import-preview-edit-description': 'description',
    'import-preview-edit-lat': 'lat', 'import-preview-edit-lng': 'lng',
    'import-preview-edit-captured-at': 'capturedAt', 'import-preview-edit-links': 'links',
    'import-preview-edit-tags': 'tags',
    'import-preview-edit-color': 'color', 'import-preview-edit-icon': 'icon',
    'import-preview-edit-status': 'status',
    'import-preview-edit-metadata-status': 'metadataStatus',
    'import-preview-edit-conversion-status': 'conversionStatus'
  };
  Object.entries(fields).forEach(([id, field]) => { elements[id].dataset.importField = field; });
  return {
    elements,
    document: {
      getElementById(id) { return elements[id] || null; },
      createElement(tagName) { return new FakeElement(tagName); }
    }
  };
}

function loadIntegration(overrides = {}) {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const dom = createDocument();
  let nextId = 0;
  const context = {
    console, document: dom.document, Number, Object, Array, String, Error, Set,
    Promise, Date, Math,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935', label: '赤' }],
    PIN_ICONS: [{ id: 'default', label: '標準' }, { id: 'photo', label: '写真' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    crypto: { randomUUID: () => `uuid-${++nextId}` },
    ...overrides
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__core = ImportJobCore;\n'
      + 'globalThis.__builder = MultiPhotoImportBuilder;\n'
      + 'globalThis.__ui = ImportPreviewUI;\n'
      + 'globalThis.__flow = ImportFlowController;\n'
      + 'globalThis.__processor = ImportPhotoItemProcessor;\n'
      + 'globalThis.__validator = ImportPhotoDraftValidator;\n'
      + 'globalThis.__workflow = MultiPhotoImportWorkflow;\n'
      + 'globalThis.__previewCore = PhotoTrackMatchPreviewCore;\n'
      + 'globalThis.__matchCore = PhotoTrackMatchCore;\n'
      + 'globalThis.__geometry = TrackGeometryCore;',
    context
  );
  return { context, elements: dom.elements };
}

function workflowState() {
  return {
    builder: null, controller: null, job: null, resourceUrlApi: null,
    targetFolderId: '', preparing: false, registering: false,
    cancellingPreparation: false, requestToken: 0
  };
}

async function flushMicrotasks(rounds = 12) {
  for (let index = 0; index < rounds; index += 1) await Promise.resolve();
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

test('Phase 3 retries only the response-lost item and closes with complete resource cleanup', async () => {
  const createdUrls = [];
  const revokedUrls = [];
  const preparationOrder = [];
  let preparing = 0;
  let maxPreparing = 0;
  const files = ['one.jpg', 'two.jpg', 'three.jpg'].map((name) => ({ name, type: 'image/jpeg' }));
  const pinsByKey = new Map();
  const calls = [];
  const upserted = new Map();
  const integration = loadIntegration({
    prepareMultiPhotoFile: async (file) => {
      preparationOrder.push(file.name);
      preparing += 1;
      maxPreparing = Math.max(maxPreparing, preparing);
      await Promise.resolve();
      preparing -= 1;
      return {
        originalFile: file, uploadFile: file, lat: null, lng: null,
        metadataStatus: 'no-gps', conversionStatus: 'not-needed'
      };
    },
    URL: {
      createObjectURL(file) {
        const url = `blob:${file.name}`;
        createdUrls.push(url);
        return url;
      },
      revokeObjectURL(url) { revokedUrls.push(url); }
    }
  });
  const state = workflowState();
  const workflow = integration.context.__workflow.create({
    state,
    builderApi: integration.context.__builder,
    processorApi: integration.context.__processor,
    flowApi: integration.context.__flow,
    validator: integration.context.__validator,
    resizePhoto: async () => 'data:image/jpeg;base64,/9j/',
    withEditToken(payload) { return payload; },
    callGAS(_method, payload) {
      calls.push({ key: payload.idempotencyKey, filename: payload.filename, folder: payload.targetFolderId });
      let pin = pinsByKey.get(payload.idempotencyKey);
      const deduplicated = !!pin;
      if (!pin) {
        pin = { id: `pin-${pinsByKey.size + 1}`, title: payload.title };
        pinsByKey.set(payload.idempotencyKey, pin);
      }
      if (payload.filename === 'two.jpg' && !deduplicated) {
        return Promise.reject(new Error('response lost after save'));
      }
      return Promise.resolve({ ok: true, deduplicated, pin });
    },
    onSaved(pin) { upserted.set(pin.id, pin); }
  });
  const snapshot = {
    tags: ['観察'], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder-snapshot'
  };

  const preparedJob = await workflow.start(files, snapshot);
  snapshot.targetFolderId = 'changed-after-start';
  assert.equal(maxPreparing, 2);
  assert.deepEqual(preparationOrder, ['one.jpg', 'two.jpg', 'three.jpg']);
  assert.deepEqual(preparedJob.items.map((item) => item.sourceRef), preparationOrder);
  assert.equal(workflow.isBusy(), true);

  const partial = await state.controller.start(preparedJob);
  assert.deepEqual(plain(partial.counts), { total: 3, succeeded: 2, failed: 1, processing: 0, waiting: 0 });
  assert.deepEqual(calls.map((call) => call.filename), ['one.jpg', 'two.jpg', 'three.jpg']);
  assert.ok(calls.every((call) => call.folder === 'folder-snapshot'));

  const completed = await state.controller.retryFailed();
  assert.deepEqual(plain(completed.counts), { total: 3, succeeded: 3, failed: 0, processing: 0, waiting: 0 });
  assert.deepEqual(calls.map((call) => call.filename), ['one.jpg', 'two.jpg', 'three.jpg', 'two.jpg']);
  assert.equal(calls[1].key, calls[3].key);
  assert.equal(new Set(calls.map((call) => call.key)).size, 3);
  assert.equal(pinsByKey.size, 3);
  assert.equal(upserted.size, 3);

  assert.equal(integration.context.__ui.close(), true);
  assert.deepEqual(revokedUrls.slice().sort(), createdUrls.slice().sort());
  assert.equal(new Set(revokedUrls).size, revokedUrls.length);
  completed.items.forEach((item) => {
    assert.equal(item.runtime.originalFile, null);
    assert.equal(item.runtime.uploadFile, null);
    assert.equal(item.runtime.previewUrl, '');
  });
  assert.equal(workflow.isBusy(), false);
  assert.equal(state.builder, null);
  assert.equal(state.controller, null);
  assert.equal(state.targetFolderId, '');

  const next = await workflow.start([{ name: 'next.jpg', type: 'image/jpeg' }], {
    tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'next-folder'
  });
  assert.equal(next.items.length, 1);
  integration.context.__ui.close();
  assert.equal(workflow.isBusy(), false);
});

test('Phase 8 applies in Preview then Flow retries response loss with one immutable matched job', async () => {
  const revokedUrls = [];
  const files = ['matched.jpg', 'other.jpg'].map((name) => ({ name, type: 'image/jpeg' }));
  const integration = loadIntegration({
    prepareMultiPhotoFile: async (file) => ({
      originalFile: file,
      uploadFile: file,
      lat: null,
      lng: null,
      capturedAt: file.name === 'matched.jpg' ? '2026-07-12T09:05:00' : '',
      metadataStatus: 'no-gps',
      conversionStatus: 'not-needed'
    }),
    URL: {
      createObjectURL(file) { return `blob:${file.name}`; },
      revokeObjectURL(url) { revokedUrls.push(url); }
    }
  });
  const state = { ...workflowState(), trackMatch: {} };
  const savedByKey = new Map();
  const calls = [];
  let driveCreates = 0;
  let mapWrites = 0;
  let matchController = null;
  const savedTrack = {
    trackId: 'track-1', revisionId: 'rev-1', name: '登山記録', description: '',
    color: '#e53935', sourceType: 'gpx', sourceName: 'private.gpx',
    segments: [{ points: [
      { lat: 10, lng: 20, elevation: null, time: '2026-07-12T00:04:00.000Z' },
      { lat: 20, lng: 30, elevation: null, time: '2026-07-12T00:06:00.000Z' }
    ] }]
  };
  const workflow = integration.context.__workflow.create({
    state,
    builderApi: integration.context.__builder,
    processorApi: integration.context.__processor,
    flowApi: integration.context.__flow,
    validator: integration.context.__validator,
    resizePhoto: async () => 'data:image/jpeg;base64,/9j/',
    withEditToken(payload) { return { ...payload, __editToken: 'send-time-token' }; },
    callGAS(method, payload) {
      assert.equal(method, 'saveImportPhotoItem');
      calls.push({ ...payload });
      let pin = savedByKey.get(payload.idempotencyKey);
      const deduplicated = !!pin;
      if (!pin) {
        driveCreates += 1;
        mapWrites += 1;
        pin = { id: `pin-${savedByKey.size + 1}`, lat: payload.lat, lng: payload.lng };
        savedByKey.set(payload.idempotencyKey, pin);
      }
      if (payload.filename === 'matched.jpg' && !deduplicated) {
        return Promise.reject(new Error('response lost after committed save'));
      }
      return Promise.resolve({ ok: true, deduplicated, pin });
    },
    resetTrackMatch() {
      integration.context.__previewCore.resetState(state.trackMatch, () => -540);
    },
    createTrackMatch() {
      matchController = integration.context.__previewCore.create({
        state: state.trackMatch,
        getTracks: () => [savedTrack],
        matchCore: integration.context.__matchCore,
        trackGeometry: integration.context.__geometry,
        importJobCore: integration.context.__core,
        timezoneOffset: () => -540
      });
      return matchController;
    }
  });

  const prepared = await workflow.start(files, {
    tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder'
  });
  assert.ok(matchController);
  integration.elements['multi-photo-track-run'].dispatch('click');
  integration.elements['multi-photo-track-apply'].dispatch('click');

  const applied = integration.context.__ui.getJob();
  assert.equal(applied, state.controller.getJob());
  assert.equal(applied, state.job);
  assert.deepEqual(plain(applied.items.map((item) => [item.sourceRef, item.lat, item.lng])), [
    ['matched.jpg', 15, 25], ['other.jpg', null, null]
  ]);
  assert.notEqual(applied, prepared);

  const partial = await state.controller.start(applied);
  assert.deepEqual(plain(partial.counts), {
    total: 2, succeeded: 1, failed: 1, processing: 0, waiting: 0
  });
  assert.equal(integration.context.__ui.getJob(), state.controller.getJob());
  assert.equal(state.job, state.controller.getJob());
  const completed = await state.controller.retryFailed();
  assert.deepEqual(plain(completed.counts), {
    total: 2, succeeded: 2, failed: 0, processing: 0, waiting: 0
  });
  assert.deepEqual(plain(completed.items.map((item) => [item.sourceRef, item.lat, item.lng])), [
    ['matched.jpg', 15, 25], ['other.jpg', null, null]
  ]);
  assert.equal(driveCreates, 2);
  assert.equal(mapWrites, 2);
  assert.equal(savedByKey.size, 2);
  const matchedCalls = calls.filter((payload) => payload.filename === 'matched.jpg');
  assert.equal(matchedCalls.length, 2);
  assert.deepEqual(matchedCalls.map((payload) => [payload.jobId, payload.itemId, payload.lat, payload.lng]), [
    [matchedCalls[0].jobId, matchedCalls[0].itemId, 15, 25],
    [matchedCalls[0].jobId, matchedCalls[0].itemId, 15, 25]
  ]);
  calls.forEach((payload) => {
    ['trackId', 'revisionId', 'ratio', 'photoTimeUtc', 'result',
      'utcOffsetMinutes', 'clockCorrectionSeconds'].forEach((key) => {
      assert.equal(Object.prototype.hasOwnProperty.call(payload, key), false);
    });
  });

  assert.equal(integration.context.__ui.close(), true);
  assert.deepEqual(revokedUrls.slice().sort(), ['blob:matched.jpg', 'blob:other.jpg']);
  assert.equal(new Set(revokedUrls).size, revokedUrls.length);
  assert.equal(workflow.isBusy(), false);
});

test('Phase 3 preparation cancellation starts only two of four items and releases every created URL', async () => {
  const started = [];
  const revoked = [];
  let workflow;
  let cancelledCallback = 0;
  const integration = loadIntegration({
    prepareMultiPhotoFile: async (file) => {
      started.push(file.name);
      return {
        originalFile: file, uploadFile: file, lat: null, lng: null,
        metadataStatus: 'no-gps', conversionStatus: 'not-needed'
      };
    },
    URL: {
      createObjectURL(file) {
        const url = `blob:${file.name}`;
        if (workflow) workflow.cancelPreparation();
        return url;
      },
      revokeObjectURL(url) { revoked.push(url); }
    }
  });
  const state = workflowState();
  workflow = integration.context.__workflow.create({
    state,
    builderApi: integration.context.__builder,
    processorApi: integration.context.__processor,
    flowApi: integration.context.__flow,
    validator: integration.context.__validator,
    resizePhoto: async () => 'data:image/jpeg;base64,/9j/',
    callGAS() { throw new Error('registration must not start'); },
    withEditToken(payload) { return payload; },
    onPreparationCancelled() { cancelledCallback += 1; }
  });

  const result = await workflow.start(
    ['one.jpg', 'two.jpg', 'three.jpg', 'four.jpg'].map((name) => ({ name, type: 'image/jpeg' })),
    { tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder' }
  );

  assert.equal(result, null);
  assert.deepEqual(started, ['one.jpg', 'two.jpg']);
  assert.equal(cancelledCallback, 1);
  assert.deepEqual(revoked, ['blob:one.jpg']);
  assert.equal(new Set(revoked).size, revoked.length);
  assert.equal(workflow.isBusy(), false);
  assert.equal(state.builder, null);
  assert.equal(state.job, null);
  await flushMicrotasks();
  assert.deepEqual(started, ['one.jpg', 'two.jpg']);
});

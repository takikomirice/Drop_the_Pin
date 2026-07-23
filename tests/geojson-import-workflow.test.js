const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(__dirname, '..', 'shared.html'), 'utf8');
const readme = fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8');

function sourceFunction(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let cursor = open; cursor < source.length; cursor += 1) {
    if (source[cursor] === '{') depth += 1;
    if (source[cursor] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, cursor + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

function createDocument() {
  const elements = new Map();
  function element(id = '') {
    if (elements.has(id)) return elements.get(id);
    const classes = new Set();
    const value = {
      id, style: {}, dataset: {}, attributes: {}, children: [], value: '', disabled: false,
      textContent: '', label: '', title: '',
      classList: {
        add(name) { classes.add(name); }, remove(name) { classes.delete(name); },
        contains(name) { return classes.has(name); },
        toggle(name, force) { if (force) classes.add(name); else classes.delete(name); }
      },
      addEventListener() {},
      appendChild(child) { this.children.push(child); return child; },
      replaceChildren(...children) { this.children = children; },
      setAttribute(name, content) { this.attributes[name] = String(content); },
      removeAttribute(name) { delete this.attributes[name]; },
      closest() { return null; }
    };
    elements.set(id, value);
    return value;
  }
  return {
    getElementById(id) { return element(id); },
    createElement(tag) { return element(`${tag}-${elements.size + 1}`); }
  };
}

function loadModules() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  const context = {
    console, Promise, Date, Math, Set, document: createDocument(), Blob,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935', label: '赤' }],
    PIN_ICONS: [{ id: 'default', label: '標準' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    URL,
    crypto: { randomUUID: (() => { let id = 0; return () => `uuid-${++id}`; })() }
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__modules = { core: ImportJobCore, geo: GeoJsonPinInterchangeCore, '
      + 'flow: ImportFlowController, preview: ImportPreviewUI, processor: ImportPinItemProcessor, '
      + 'validator: ImportPinDraftValidator, presets: InputPresetApplyCore };',
    context
  );
  return context.__modules;
}

function plain(value) { return JSON.parse(JSON.stringify(value)); }

function deferred() {
  let resolve;
  const promise = new Promise((done) => { resolve = done; });
  return { promise, resolve };
}

async function flushMicrotasks(rounds = 8) {
  for (let index = 0; index < rounds; index += 1) await Promise.resolve();
}

function geoJson(features) {
  return JSON.stringify({ type: 'FeatureCollection', dropThePinSchemaVersion: 1, features });
}

function feature(title, sourceId, geometry = null, properties = {}) {
  return {
    type: 'Feature',
    id: sourceId ? `feature-${sourceId}` : undefined,
    geometry,
    properties: { sourceId, title, ...properties }
  };
}

test('GeoJSON 3 valid plus 1 invalid Feature uses the existing pin flow and retries only response loss', async () => {
  const modules = loadModules();
  let job = modules.geo.buildImportJob(geoJson([
    feature('One', 'legacy-1', { type: 'Point', coordinates: [139, 35] }, { eventAt: '2026-07-11T10:00' }),
    feature('Two', 'legacy-1', null, { eventAt: '2026-07-11T11:00' }),
    feature('', 'legacy-bad'),
    feature('Three', 'legacy-3', null, { eventAt: '2026-07-11T12:00' })
  ]), {
    jobId: 'geojson-job', createdAt: '2026-07-11T00:00:00Z',
    generateId: (() => { let id = 0; return () => `item-${++id}`; })()
  });
  const invalid = job.items.find((item) => item.uploadStatus === 'failed');
  assert.equal(invalid.errorCode, 'GEOJSON_FEATURE_TITLE_REQUIRED');
  assert.equal(invalid.retryable, false);
  assert.equal(invalid.attempts, 0);
  assert.throws(() => modules.validator.validateJob(job), (error) => error.code === 'IMPORT_DRAFT_PREPARATION_FAILED');
  job = modules.core.removeDraftItem(job, invalid.id, null);
  job = modules.core.updateDraftItem(job, job.items[0].id, { description: 'Preview edited' });
  const presetApplied = modules.presets.apply(job.items[0], {
    presetId: 'preset-1', name: 'GeoJSON preset', enabled: true, orderIndex: 0,
    tagsMode: 'set', tags: ['preset-tag'], colorMode: 'set', color: '#e53935',
    iconMode: 'set', icon: 'default', statusMode: 'set', status: '未対応'
  });
  job = modules.core.updateDraftItem(job, job.items[0].id, {
    tags: presetApplied.tags, color: presetApplied.color,
    icon: presetApplied.icon, status: presetApplied.status
  });
  assert.equal(modules.validator.validateJob(job), true);

  const calls = [];
  const pinsByKey = new Map();
  const upserted = new Map([['legacy-1', { id: 'legacy-1', title: 'Existing' }]]);
  let active = 0;
  let maxActive = 0;
  const processor = modules.processor.create({
    withEditToken(payload) { return { ...payload, __editToken: 'token' }; },
    async callGAS(method, payload) {
      assert.equal(method, 'saveImportPinItem');
      calls.push(payload);
      active += 1;
      maxActive = Math.max(maxActive, active);
      await Promise.resolve();
      let pin = pinsByKey.get(payload.idempotencyKey);
      const deduplicated = !!pin;
      if (!pin) {
        pin = { id: `pin-${pinsByKey.size + 1}`, title: payload.title };
        pinsByKey.set(payload.idempotencyKey, pin);
      }
      active -= 1;
      if (payload.title === 'Two' && !deduplicated) throw new Error('response lost');
      return { ok: true, deduplicated, pin };
    },
    onSaved(pin) { upserted.set(pin.id, pin); }
  });
  const controller = modules.flow.create({
    processItem: processor.processItem,
    concurrency: 2,
    validateJob: modules.validator.validateJob
  });
  const geoState = { loading: false, requestToken: 1, job, controller, registering: false };
  controller.open({
    job,
    title: 'GeoJSONインポートを確認',
    sourceLabel: `GeoJSON ${job.items.length}件`,
    closePolicy: 'discard-only',
    timeFieldLabel: 'イベント時刻',
    completionItemLabel: 'ピン',
    onClose() {
      geoState.requestToken += 1;
      geoState.job = null;
      geoState.controller = null;
      geoState.registering = false;
    }
  });
  const partial = await controller.start(job);
  assert.deepEqual(plain(partial.counts), { total: 3, succeeded: 2, failed: 1, processing: 0, waiting: 0 });
  assert.equal(maxActive, 2);
  assert.deepEqual(calls.map((payload) => payload.title).sort(), ['One', 'Three', 'Two']);
  for (const payload of calls) {
    for (const excluded of ['sourceId', 'sourceRef', 'runtime', 'base64', 'targetFolderId', 'fileId', 'imageUrl']) {
      assert.equal(Object.prototype.hasOwnProperty.call(payload, excluded), false, excluded);
    }
  }
  assert.equal(calls.find((payload) => payload.title === 'One').description, 'Preview edited');
  assert.deepEqual(plain(calls.find((payload) => payload.title === 'One').tags), ['preset-tag']);
  assert.equal(calls.find((payload) => payload.title === 'One').lat, 35);
  assert.equal(calls.find((payload) => payload.title === 'One').lng, 139);

  const completed = await controller.retryFailed();
  assert.deepEqual(plain(completed.counts), { total: 3, succeeded: 3, failed: 0, processing: 0, waiting: 0 });
  assert.equal(calls.length, 4);
  assert.equal(calls.find((payload) => payload.title === 'Two').idempotencyKey, calls[3].idempotencyKey);
  assert.equal(new Set(calls.map((payload) => payload.idempotencyKey)).size, 3);
  assert.equal(pinsByKey.size, 3);
  assert.equal(upserted.size, 4);
  assert.equal(upserted.get('legacy-1').title, 'Existing');
  assert.equal(modules.preview.close(), true);
  assert.equal(geoState.job, null);
  assert.equal(geoState.controller, null);
  assert.equal(geoState.registering, false);

  const nextJob = modules.geo.buildImportJob(geoJson([feature('Next', 'legacy-1')]), {
    jobId: 'next-job', generateId: () => 'next-item'
  });
  assert.equal(nextJob.items.length, 1);
  assert.equal(nextJob.items[0].title, 'Next');
  assert.notEqual(nextJob.items[0].id, 'legacy-1');
});

test('GeoJSON cancellation waits for two in-flight saves and resume starts only remaining queued items', async () => {
  const modules = loadModules();
  const job = modules.geo.buildImportJob(geoJson(['One', 'Two', 'Three', 'Four'].map((title) => (
    feature(title, title.toLowerCase())
  ))), {
    jobId: 'cancel-job',
    generateId: (() => { let id = 0; return () => `item-${++id}`; })()
  });
  const gates = [deferred(), deferred()];
  const calls = [];
  let active = 0;
  let maxActive = 0;
  const processor = modules.processor.create({
    withEditToken(payload) { return payload; },
    callGAS(_method, payload) {
      calls.push(payload);
      active += 1;
      maxActive = Math.max(maxActive, active);
      const gate = calls.length <= 2 ? gates[calls.length - 1].promise : Promise.resolve();
      return gate.then(() => {
        active -= 1;
        return { ok: true, pin: { id: `pin-${payload.itemId}` } };
      });
    }
  });
  const controller = modules.flow.create({
    processItem: processor.processItem,
    concurrency: 2,
    validateJob: modules.validator.validateJob
  });
  controller.open({ job, closePolicy: 'discard-only' });
  const running = controller.start(job);
  await flushMicrotasks();
  assert.equal(calls.length, 2);
  assert.equal(maxActive, 2);
  controller.cancel();
  gates.forEach((gate) => gate.resolve());
  const cancelled = await running;
  assert.equal(cancelled.status, 'cancelled');
  assert.deepEqual(plain(cancelled.counts), { total: 4, succeeded: 2, failed: 0, processing: 0, waiting: 2 });
  assert.equal(calls.length, 2);

  const completed = await controller.resume();
  assert.equal(completed.status, 'completed');
  assert.deepEqual(plain(completed.counts), { total: 4, succeeded: 4, failed: 0, processing: 0, waiting: 0 });
  assert.deepEqual(calls.map((payload) => payload.itemId), ['item-1', 'item-2', 'item-3', 'item-4']);
  assert.equal(new Set(calls.map((payload) => payload.idempotencyKey)).size, 4);
});

test('production GeoJSON preview and mutual exclusion reuse existing flow boundaries', () => {
  assert.match(indexHtml, /function openGeoJsonImportPreview\(job\)/);
  assert.match(indexHtml, /title:\s*'GeoJSON取込を確認'/);
  assert.match(indexHtml, /sourceLabel:\s*'GeoJSON '\s*\+\s*job\.items\.length\s*\+\s*'件'/);
  assert.match(indexHtml, /closePolicy:\s*'discard-only'/);
  assert.match(indexHtml, /timeFieldLabel:\s*'イベント時刻'/);
  assert.match(indexHtml, /completionItemLabel:\s*'ピン'/);
  assert.match(indexHtml, /function openGeoJsonImportPreview[\s\S]{0,2200}ImportPinItemProcessor\.create/);
  assert.match(indexHtml, /function openGeoJsonImportPreview[\s\S]{0,2200}ImportFlowController\.create/);
  assert.match(indexHtml, /function openGeoJsonImportPreview[\s\S]{0,2200}concurrency:\s*2/);
  assert.match(indexHtml, /function openGeoJsonImportPreview[\s\S]{0,2200}validateJob:\s*ImportPinDraftValidator\.validateJob/);
  assert.match(indexHtml, /function openGeoJsonImportPreview[\s\S]{0,2200}onSaved:\s*upsertImportedPin/);
  assert.doesNotMatch(indexHtml, /function openGeoJsonImportPreview[\s\S]{0,2200}hidePrimaryAction/);

  assert.match(indexHtml, /function isGeoJsonImportBusy\(\)/);
  assert.match(indexHtml, /function isProductionImportBusy\(\)[\s\S]{0,220}isGeoJsonImportBusy\(\)/);
  assert.match(indexHtml, /canStartImport:\s*function\(\)[\s\S]{0,180}!isMultiPhotoImportBusyState\(\)[\s\S]{0,100}!isGeoJsonImportBusy\(\)[\s\S]{0,100}!isTrackImportBusy\(\)/);
  assert.match(indexHtml, /canStartImport:\s*function\(\)[\s\S]{0,180}!isCsvImportBusy\(\)[\s\S]{0,100}!isMultiPhotoImportBusyState\(\)[\s\S]{0,100}!isTrackImportBusy\(\)/);
  assert.match(indexHtml, /function canStartMultiPhotoImport\(\)[\s\S]{0,700}isGeoJsonImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'hasPendingMutationWork'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'initializeApp'), /beforeunload[\s\S]*hasPendingMutationWork\(\)/);
  assert.match(indexHtml, /closeOverlay\('data-overlay',\s*\{\s*preserveGeoJsonInterchange:\s*true,\s*restoreFocus:\s*false\s*\}\)/);
  assert.match(indexHtml, /geoJsonInterchangeController\.invalidate\(\)/);
  assert.doesNotMatch(sharedHtml, /saveImportPinItem|GeoJsonPinInterchangeCore|geojson-file-input/i);
});

test('GeoJSON preview handoff atomically clears ownership for every initialization failure', () => {
  const scenarios = ['processor', 'controller', 'settings-close', 'preview-open'];
  scenarios.forEach((failure) => {
    const job = { id: 'geojson-job', items: [{ id: 'geojson-item' }] };
    const state = {
      geoJsonInterchange: {
        loading: true, requestToken: 7, job, controller: null, registering: true
      }
    };
    const flowController = {
      open() {
        if (failure === 'preview-open') throw new Error('preview render failed with secret details');
        return job;
      }
    };
    const context = {
      state,
      isCsvImportBusy: () => false,
      isMultiPhotoImportBusyState: () => false,
      isTrackImportBusy: () => false,
      ImportPinItemProcessor: {
        create() {
          if (failure === 'processor') throw new Error('processor failed');
          return { processItem() {} };
        }
      },
      ImportPinDraftValidator: { validateJob() {} },
      ImportFlowController: {
        create() {
          if (failure === 'controller') throw new Error('controller failed');
          return flowController;
        }
      },
      ImportPreviewUI: { close() { return true; } },
      closeOverlay() {
        if (failure === 'settings-close') throw new Error('settings close failed');
      },
      withGAS() {}, withEditToken(value) { return value; }, upsertImportedPin() {},
      renderGeoJsonInterchangeBusy() {}, reloadInputPresets() {}, ensureInputPresetsLoaded() {},
      getEnabledInputPresets() { return []; }
    };
    const helper = indexHtml.includes('function resetInterchangePreviewOwnership(')
      ? `${sourceFunction(indexHtml, 'resetInterchangePreviewOwnership')}\n`
      : '';
    vm.runInNewContext(
      `${helper}${sourceFunction(indexHtml, 'openGeoJsonImportPreview')}\nthis.openGeoJson = openGeoJsonImportPreview;`,
      context
    );
    assert.throws(() => context.openGeoJson(job));
    assert.equal(state.geoJsonInterchange.job, null, failure);
    assert.equal(state.geoJsonInterchange.controller, null, failure);
    assert.equal(state.geoJsonInterchange.registering, false, failure);
    assert.equal(state.geoJsonInterchange.loading, false, failure);
    assert.equal(state.geoJsonInterchange.requestToken, 8, failure);
  });
});

test('README summarizes GeoJSON Point format, limits, import behavior, and verification', () => {
  assert.match(readme, /GeoJSONピン[^\n]*FeatureCollection[^\n]*Point[^\n]*geometry: null[^\n]*2MB[^\n]*20 Feature/);
  assert.match(readme, /sourceId[^\n]*既存ピンを上書きせず[^\n]*新規ピン/);
  assert.match(readme, /CSV／GeoJSON Pointの1件／20件取込[^\n]*上限超過[^\n]*不正項目削除/);
  assert.match(readme, /Nodeテスト[^\n]*Chromiumテスト[^\n]*Apps Script[^\n]*実Drive[^\n]*未実施/);
});

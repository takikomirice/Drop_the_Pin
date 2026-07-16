const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

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
    console, Promise, Date, Math, Set, document: createDocument(),
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
      + 'globalThis.__modules = { core: ImportJobCore, csv: CsvPinInterchangeCore, '
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

test('CSV 3 valid plus 1 invalid row registers only valid rows and retries response loss idempotently', async () => {
  const modules = loadModules();
  const csv = [
    'schemaVersion,sourceId,title,description,lat,lng,color,icon,status,tags,eventAt,links',
    '1,legacy-1,One,,35,139,#e53935,default,,[],2026-07-11T10:00,[]',
    '1,legacy-1,Two,,,,#e53935,default,,[],2026-07-11T11:00,[]',
    '1,legacy-bad,,,,,#e53935,default,,[],,[]',
    '1,legacy-3,Three,,,,#e53935,default,,[],2026-07-11T12:00,[]'
  ].join('\n');
  let job = modules.csv.buildImportJob(csv, {
    jobId: 'csv-job', createdAt: '2026-07-11T00:00:00Z',
    generateId: (() => { let id = 0; return () => `item-${++id}`; })()
  });
  const invalid = job.items.find((item) => item.uploadStatus === 'failed');
  assert.equal(invalid.retryable, false);
  assert.equal(invalid.attempts, 0);
  assert.throws(() => modules.validator.validateJob(job), (error) => error.code === 'IMPORT_DRAFT_PREPARATION_FAILED');
  job = modules.core.removeDraftItem(job, invalid.id, null);
  job = modules.core.updateDraftItem(job, job.items[0].id, { description: 'Preview edited' });
  const presetApplied = modules.presets.apply(job.items[0], {
    presetId: 'preset-1', name: 'CSV preset', enabled: true, orderIndex: 0,
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
  const csvState = { loading: false, requestToken: 1, job, controller, registering: false };
  controller.open({
    job,
    closePolicy: 'discard-only',
    completionItemLabel: 'ピン',
    onClose() {
      csvState.requestToken += 1;
      csvState.job = null;
      csvState.controller = null;
      csvState.registering = false;
    }
  });
  const partial = await controller.start(job);
  assert.deepEqual(plain(partial.counts), { total: 3, succeeded: 2, failed: 1, processing: 0, waiting: 0 });
  assert.equal(maxActive, 2);
  assert.deepEqual(calls.map((payload) => payload.title).sort(), ['One', 'Three', 'Two']);
  assert.ok(calls.every((payload) => !Object.prototype.hasOwnProperty.call(payload, 'sourceId')));
  assert.ok(calls.every((payload) => !Object.prototype.hasOwnProperty.call(payload, 'sourceRef')));
  assert.ok(calls.every((payload) => !Object.prototype.hasOwnProperty.call(payload, 'base64')));
  assert.ok(calls.every((payload) => !Object.prototype.hasOwnProperty.call(payload, 'targetFolderId')));
  assert.equal(calls.find((payload) => payload.title === 'One').description, 'Preview edited');
  assert.deepEqual(plain(calls.find((payload) => payload.title === 'One').tags), ['preset-tag']);

  const completed = await controller.retryFailed();
  assert.deepEqual(plain(completed.counts), { total: 3, succeeded: 3, failed: 0, processing: 0, waiting: 0 });
  assert.equal(calls.length, 4);
  assert.equal(calls.find((payload) => payload.title === 'Two').idempotencyKey, calls[3].idempotencyKey);
  assert.equal(new Set(calls.map((payload) => payload.idempotencyKey)).size, 3);
  assert.equal(pinsByKey.size, 3);
  assert.equal(upserted.size, 4);
  assert.equal(upserted.get('legacy-1').title, 'Existing');
  assert.equal(modules.preview.close(), true);
  assert.equal(csvState.job, null);
  assert.equal(csvState.controller, null);
  assert.equal(csvState.registering, false);
  const nextJob = modules.csv.buildImportJob('schemaVersion,title\n1,Next', {
    jobId: 'next-job', generateId: () => 'next-item'
  });
  assert.equal(nextJob.items.length, 1);
  assert.equal(nextJob.items[0].title, 'Next');
});

test('CSV cancellation waits for two in-flight items and resume saves only the remaining queued items', async () => {
  const modules = loadModules();
  const items = ['One', 'Two', 'Three', 'Four'].map((title, index) => ({
    id: `item-${index + 1}`, sourceType: 'csv', sourceRef: `CSV ${index + 2}行目`,
    title, description: '', lat: null, lng: null, capturedAt: '', tags: [], links: [],
    color: '#e53935', icon: 'default', status: '', uploadStatus: 'queued', attempts: 0,
    runtime: { originalFile: null, uploadFile: null, previewUrl: '' }
  }));
  const job = modules.core.createJob({ id: 'cancel-job', sourceType: 'csv', items });
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

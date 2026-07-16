const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

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
    this.src = '';
    this._textContent = '';
  }
  set textContent(value) {
    this._textContent = value == null ? '' : String(value);
    this.children = [];
  }
  get textContent() {
    return this._textContent + this.children.map((child) => child.textContent).join('');
  }
  set innerHTML(_value) { throw new Error('Import flow UI must not write innerHTML'); }
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
  setAttribute(name, value) {
    this.attributes[name] = String(value);
    if (name === 'src') this.src = String(value);
  }
  removeAttribute(name) {
    delete this.attributes[name];
    if (name === 'src') this.src = '';
  }
  addEventListener(type, listener) {
    if (!this.listeners[type]) this.listeners[type] = [];
    this.listeners[type].push(listener);
  }
  listenerCount(type) { return (this.listeners[type] || []).length; }
  dispatch(type, target) {
    const event = { type, target: target || this, preventDefault() {} };
    (this.listeners[type] || []).forEach((listener) => listener(event));
  }
  closest(selector) {
    if (selector === '[data-import-item-id]' && this.dataset.importItemId) return this;
    return this.parentNode ? this.parentNode.closest(selector) : null;
  }
  focus() {}
}

function createDocumentHarness() {
  const ids = [
    'import-preview-overlay', 'import-preview-sheet', 'import-preview-title', 'import-preview-source',
    'import-preview-job-status', 'import-preview-operation-note', 'import-preview-operation-error',
    'import-preview-completion-message',
    'import-preview-presets', 'import-preview-preset-select', 'import-preview-preset-apply-selected',
    'import-preview-preset-apply-all', 'import-preview-preset-status',
    'import-preview-preset-error', 'import-preview-preset-retry',
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
    'import-preview-resume', 'import-preview-retry', 'import-preview-close',
    'import-preview-discard',
    'multi-photo-track-match-panel', 'multi-photo-track-select', 'multi-photo-track-utc-offset',
    'multi-photo-track-clock-correction', 'multi-photo-track-max-gap',
    'multi-photo-track-endpoint-tolerance', 'multi-photo-track-run',
    'multi-photo-track-status', 'multi-photo-track-error', 'multi-photo-track-counts',
    'multi-photo-track-warnings', 'multi-photo-track-results', 'multi-photo-track-apply',
    'multi-photo-track-clear'
  ];
  const elements = Object.create(null);
  ids.forEach((id) => {
    const tagName = id === 'import-preview-image' ? 'img'
      : id === 'import-preview-image-trigger' ? 'button'
      : id === 'import-preview-preset-select' || id === 'import-preview-edit-status'
          || id === 'import-preview-edit-color' || id === 'import-preview-edit-icon' ? 'select'
        : id.startsWith('import-preview-edit-') ? 'input'
          : /(?:delete|primary|cancel|resume|retry|apply-selected|apply-all|close|discard)$/.test(id) ? 'button'
          : 'div';
    elements[id] = new FakeElement(tagName, id);
  });
  elements['import-preview-overlay'].classList.add('sheet-overlay');
  const fields = {
    'import-preview-edit-title': 'title',
    'import-preview-edit-description': 'description',
    'import-preview-edit-lat': 'lat',
    'import-preview-edit-lng': 'lng',
    'import-preview-edit-captured-at': 'capturedAt',
    'import-preview-edit-links': 'links',
    'import-preview-edit-tags': 'tags',
    'import-preview-edit-color': 'color',
    'import-preview-edit-icon': 'icon',
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

function loadModules(options = {}) {
  const start = indexHtml.indexOf('const ImportJobCore = (function() {');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const harness = createDocumentHarness();
  const context = {
    document: harness.document,
    URL: options.urlApi || { revokeObjectURL() {} },
    console,
    Number,
    Object,
    Array,
    String,
    Error,
    Date,
    Promise
  };
  context.PIN_COLORS = [{ hex: '#e53935', label: '赤' }];
  context.PIN_ICONS = [{ id: 'default', label: '標準' }, { id: 'photo', label: '写真' }];
  context.PIN_STATUSES = ['未対応', '対応中', '完了', '保留'];
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__core = ImportJobCore;\n'
      + 'globalThis.__queue = ImportQueueRunner;\n'
      + 'globalThis.__ui = ImportPreviewUI;\n'
      + 'globalThis.__flow = ImportFlowController;\n'
      + 'globalThis.__processor = ImportPhotoItemProcessor;\n'
      + 'globalThis.__validator = ImportPhotoDraftValidator;',
    context
  );
  return {
    core: context.__core,
    queue: context.__queue,
    ui: context.__ui,
    flow: context.__flow,
    processor: context.__processor,
    validator: context.__validator,
    elements: harness.elements
  };
}

function createDeferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

async function flushMicrotasks(rounds = 10) {
  for (let index = 0; index < rounds; index += 1) await Promise.resolve();
}

function deterministicNow() {
  let tick = 0;
  return () => `flow-time-${++tick}`;
}

function createIdleJob(core, id, items) {
  return core.createJob({
    id,
    sourceType: 'test',
    items: items.map((item, index) => ({
      id: item.id || `${id}-item-${index + 1}`,
      sourceType: 'test',
      ...item
    }))
  });
}

function jobWithState(core, id, status, uploadStatuses) {
  const idle = createIdleJob(core, id, uploadStatuses.map((uploadStatus) => ({ uploadStatus })));
  const items = idle.items;
  return Object.assign({}, idle, {
    status,
    cancelRequested: status === 'cancelled',
    counts: core.getCounts({ items })
  });
}

test('preview initialization failure rolls back Flow ownership and permits a later open', () => {
  const modules = loadModules();
  const job = createIdleJob(modules.core, 'initialization-failure', [{ title: 'one' }]);
  const controller = modules.flow.create({ processItem: async () => ({ ok: true }) });
  const originalOpen = modules.ui.open;
  modules.ui.open = function() { throw new Error('preview render failed'); };

  assert.throws(() => controller.open({ job }), /preview render failed/);
  assert.equal(controller.getJob(), null);

  modules.ui.open = originalOpen;
  assert.equal(controller.open({ job }), job);
  assert.equal(controller.getJob(), job);
  assert.equal(modules.ui.close({ discard: true }), true);
});

test('flow preview DOM exposes job status, operation error, and four execution buttons', () => {
  const ids = [
    'import-preview-job-status', 'import-preview-operation-note', 'import-preview-operation-error',
    'import-preview-primary', 'import-preview-cancel', 'import-preview-resume', 'import-preview-retry'
  ];
  ids.forEach((id) => assert.match(indexHtml, new RegExp(`id=["']${id}["']`)));
});

test('idle start moves through ready and running to completed while UI redraws', async () => {
  const { core, flow, elements } = loadModules();
  const changes = [];
  const settled = [];
  const controller = flow.create({
    processItem() {},
    now: deterministicNow(),
    onJobChange(job, event) { changes.push({ status: job.status, type: event.type }); },
    onSettled(job) { settled.push(job.status); }
  });
  const idle = createIdleJob(core, 'complete', [{}, {}]);

  controller.open({ job: idle, title: '確認', sourceLabel: 'CSV' });
  assert.equal(elements['import-preview-job-status'].textContent, '準備中');
  assert.equal(elements['import-preview-primary'].disabled, false);

  const finalJob = await controller.start();

  assert.equal(finalJob.status, 'completed');
  assert.equal(controller.getJob(), finalJob);
  assert.equal(elements['import-preview-job-status'].textContent, '完了');
  assert.equal(elements['import-preview-count-succeeded'].textContent, '2');
  assert.ok(changes.some((entry) => entry.status === 'ready'));
  assert.ok(changes.some((entry) => entry.status === 'running'));
  assert.ok(changes.some((entry) => entry.status === 'completed'));
  assert.deepEqual(settled, ['completed']);
  assert.equal(controller.isRunning(), false);
});

test('draft edits and removals stay synchronized with the controller before start', async () => {
  const { core, flow, elements } = loadModules();
  const processed = [];
  const controller = flow.create({
    processItem(item) { processed.push({ id: item.id, title: item.title }); },
    now: deterministicNow()
  });
  controller.open({
    job: createIdleJob(core, 'draft-sync', [
      { id: 'draft-sync-keep', title: 'before edit' },
      { id: 'draft-sync-remove', title: 'remove me' }
    ])
  });

  elements['import-preview-edit-title'].value = 'after edit';
  elements['import-preview-editor'].dispatch('change', elements['import-preview-edit-title']);
  elements['import-preview-list'].dispatch(
    'click',
    elements['import-preview-list'].children[1]
  );
  elements['import-preview-delete'].dispatch('click');

  assert.equal(controller.getJob().items.length, 1);
  assert.equal(controller.getJob().items[0].title, 'after edit');
  const finalJob = await controller.start();
  assert.deepEqual(processed, [{ id: 'draft-sync-keep', title: 'after edit' }]);
  assert.equal(finalJob.counts.total, 1);
});

test('preview buttons follow idle, ready, running, cancelled, and completed states', () => {
  const { core, ui, elements } = loadModules();
  const callbacks = {
    onPrimaryAction() {}, onCancelAction() {}, onResumeAction() {}, onRetryAction() {}
  };

  const idle = createIdleJob(core, 'idle-ui', [{}]);
  ui.open({ job: idle, ...callbacks });
  assert.equal(elements['import-preview-job-status'].textContent, '準備中');
  assert.equal(elements['import-preview-primary'].style.display, '');
  assert.equal(elements['import-preview-primary'].disabled, false);
  assert.equal(elements['import-preview-cancel'].style.display, 'none');

  const ready = core.readyJob(idle);
  ui.setJob(ready);
  assert.equal(elements['import-preview-job-status'].textContent, '開始待ち');
  assert.equal(elements['import-preview-primary'].disabled, false);
  assert.equal(elements['import-preview-delete'].disabled, true);

  const running = jobWithState(core, 'running-ui', 'running', ['processing', 'queued']);
  ui.setJob(running);
  assert.equal(elements['import-preview-job-status'].textContent, '処理中');
  assert.equal(elements['import-preview-cancel'].style.display, '');
  assert.equal(elements['import-preview-cancel'].disabled, false);
  assert.equal(elements['import-preview-primary'].disabled, true);
  assert.equal(elements['import-preview-close'].disabled, true);
  assert.equal(elements['import-preview-discard'].disabled, true);

  const cancelling = jobWithState(core, 'cancelling-ui', 'cancelled', ['processing', 'queued']);
  ui.setJob(cancelling);
  assert.equal(elements['import-preview-job-status'].textContent, 'キャンセル中');
  assert.match(elements['import-preview-operation-note'].textContent, /処理中項目の結果確定を待っています/);
  assert.equal(elements['import-preview-resume'].disabled, true);
  assert.equal(elements['import-preview-close'].disabled, true);

  const cancelled = jobWithState(core, 'cancelled-ui', 'cancelled', ['succeeded', 'queued']);
  ui.setJob(cancelled);
  assert.equal(elements['import-preview-job-status'].textContent, 'キャンセル済み');
  assert.equal(elements['import-preview-resume'].style.display, '');
  assert.equal(elements['import-preview-resume'].disabled, false);
  assert.equal(elements['import-preview-close'].disabled, false);

  const partial = jobWithState(core, 'partial-ui', 'completed', ['succeeded', 'failed']);
  ui.setJob(partial);
  assert.equal(elements['import-preview-job-status'].textContent, '一部失敗');
  assert.equal(elements['import-preview-retry'].style.display, '');
  assert.equal(elements['import-preview-retry'].disabled, false);

  const completed = jobWithState(core, 'completed-ui', 'completed', ['succeeded', 'succeeded']);
  ui.setJob(completed);
  assert.equal(elements['import-preview-job-status'].textContent, '完了');
  assert.equal(elements['import-preview-retry'].style.display, 'none');
  assert.equal(elements['import-preview-close'].disabled, false);
});

test('rapid start button clicks create only one queue execution', async () => {
  const { core, flow, elements } = loadModules();
  const gate = createDeferred();
  let processCalls = 0;
  let settledCalls = 0;
  const controller = flow.create({
    processItem() { processCalls += 1; return gate.promise; },
    now: deterministicNow(),
    onSettled() { settledCalls += 1; }
  });
  controller.open({ job: createIdleJob(core, 'double-start', [{}]) });

  elements['import-preview-primary'].dispatch('click');
  elements['import-preview-primary'].dispatch('click');
  await flushMicrotasks();

  assert.equal(processCalls, 1);
  assert.equal(controller.isRunning(), true);
  assert.equal(elements['import-preview-delete'].disabled, true);
  assert.equal(elements['import-preview-close'].disabled, true);
  assert.equal(elements['import-preview-discard'].disabled, true);

  gate.resolve();
  await flushMicrotasks();
  assert.equal(settledCalls, 1);
});

test('ready notifications cannot reenter open or start before the operation is reserved', async () => {
  const { core, flow } = loadModules();
  const calls = [];
  const first = createIdleJob(core, 'ready-reentry-first', [{}]);
  const second = createIdleJob(core, 'ready-reentry-second', [{}]);
  let controller;
  let reentered = false;
  let openError = null;
  let nestedStartResult = null;
  controller = flow.create({
    processItem(item) { calls.push(item.id); },
    now: deterministicNow(),
    onJobChange(_job, event) {
      if (event.type !== 'job-ready' || reentered) return;
      reentered = true;
      try {
        controller.open({ job: second });
      } catch (error) {
        openError = error;
      }
      nestedStartResult = controller.start().then(
        () => ({ resolved: true }),
        (error) => ({ error })
      );
    }
  });
  controller.open({ job: first });

  const finalJob = await controller.start();
  const nested = await nestedStartResult;

  assert.equal(finalJob.id, 'ready-reentry-first');
  assert.equal(openError && openError.code, 'IMPORT_FLOW_RUNNING');
  assert.equal(nested.error && nested.error.code, 'IMPORT_FLOW_OPERATION_RUNNING');
  assert.deepEqual(calls, ['ready-reentry-first-item-1']);
});

test('cancel clicks are idempotent and enable resume only after in-flight settlement', async () => {
  const { core, flow, elements } = loadModules();
  const firstGate = createDeferred();
  const calls = [];
  const controller = flow.create({
    processItem(item) {
      calls.push(item.id);
      if (item.id === 'cancel-flow-item-1' && calls.length === 1) return firstGate.promise;
      return Promise.resolve();
    },
    now: deterministicNow()
  });
  controller.open({ job: createIdleJob(core, 'cancel-flow', [{}, {}, {}]) });
  const firstRun = controller.start();
  await flushMicrotasks();

  elements['import-preview-cancel'].dispatch('click');
  elements['import-preview-cancel'].dispatch('click');
  assert.equal(controller.getJob().status, 'cancelled');
  assert.equal(controller.getJob().counts.processing, 1);
  assert.equal(elements['import-preview-job-status'].textContent, 'キャンセル中');
  assert.equal(elements['import-preview-resume'].disabled, true);
  assert.equal(elements['import-preview-close'].disabled, true);

  firstGate.resolve();
  const cancelled = await firstRun;
  assert.equal(cancelled.status, 'cancelled');
  assert.deepEqual(calls, ['cancel-flow-item-1']);
  assert.equal(elements['import-preview-job-status'].textContent, 'キャンセル済み');
  assert.equal(elements['import-preview-resume'].disabled, false);

  const completed = await controller.resume();
  assert.equal(completed.status, 'completed');
  assert.deepEqual(calls, ['cancel-flow-item-1', 'cancel-flow-item-2', 'cancel-flow-item-3']);
  assert.equal(completed.items[0].attempts, 1);
  assert.equal(completed.items[1].attempts, 1);
});

test('resume sends only queued items and explicit retry later sends the failed item', async () => {
  const { core, flow } = loadModules();
  const gates = new Map();
  const firstRunCalls = [];
  let resumed = false;
  const allCalls = [];
  const validatedStatuses = [];
  const controller = flow.create({
    concurrency: 2,
    processItem(item) {
      allCalls.push(item.id);
      if (resumed) return Promise.resolve();
      firstRunCalls.push(item.id);
      const gate = createDeferred();
      gates.set(item.id, gate);
      return gate.promise;
    },
    validateJob(job) { validatedStatuses.push(job.status); },
    now: deterministicNow()
  });
  controller.open({ job: createIdleJob(core, 'resume-flow', [{}, {}, {}]) });
  const firstRun = controller.start();
  await flushMicrotasks();
  controller.cancel();
  gates.get('resume-flow-item-1').resolve();
  gates.get('resume-flow-item-2').reject(new Error('retry me'));
  const cancelled = await firstRun;
  assert.equal(cancelled.counts.succeeded, 1);
  assert.equal(cancelled.counts.failed, 1);
  assert.equal(cancelled.counts.waiting, 1);

  resumed = true;
  const partial = await controller.resume();
  assert.equal(partial.status, 'completed');
  assert.equal(partial.counts.failed, 1);
  assert.deepEqual(firstRunCalls, ['resume-flow-item-1', 'resume-flow-item-2']);
  assert.deepEqual(allCalls, [
    'resume-flow-item-1', 'resume-flow-item-2',
    'resume-flow-item-3'
  ]);
  assert.equal(partial.items[0].attempts, 1);
  assert.equal(partial.items[1].attempts, 1);
  assert.equal(partial.items[2].attempts, 1);

  const completed = await controller.retryFailed();
  assert.equal(completed.status, 'completed');
  assert.equal(completed.counts.failed, 0);
  assert.deepEqual(allCalls, [
    'resume-flow-item-1', 'resume-flow-item-2',
    'resume-flow-item-3', 'resume-flow-item-2'
  ]);
  assert.equal(completed.items[1].attempts, 2);
  assert.deepEqual(validatedStatuses, ['idle', 'ready', 'ready']);
});

test('cancelled failures without queued work expose retry instead of resume', async () => {
  const { core, flow, elements } = loadModules();
  const gate = createDeferred();
  let calls = 0;
  const controller = flow.create({
    processItem() {
      calls += 1;
      return calls === 1 ? gate.promise : Promise.resolve();
    },
    now: deterministicNow()
  });
  controller.open({ job: createIdleJob(core, 'cancelled-failure', [{}]) });
  const firstRun = controller.start();
  await flushMicrotasks();
  controller.cancel();
  gate.reject(Object.assign(new Error('response lost'), { retryable: true }));
  const cancelled = await firstRun;
  assert.equal(cancelled.status, 'cancelled');
  assert.equal(elements['import-preview-resume'].style.display, 'none');
  assert.equal(elements['import-preview-retry'].style.display, '');
  assert.equal(elements['import-preview-retry'].disabled, false);

  const completed = await controller.retryFailed();
  assert.equal(completed.status, 'completed');
  assert.equal(calls, 2);
});

test('retryFailed processes only failures and reports partial then complete status', async () => {
  const { core, flow, elements } = loadModules();
  const calls = Object.create(null);
  let settledCalls = 0;
  const validatedStatuses = [];
  const controller = flow.create({
    processItem(item) {
      calls[item.id] = (calls[item.id] || 0) + 1;
      if (item.id === 'retry-flow-item-1' && calls[item.id] === 1) {
        return Promise.reject(new Error('first failure'));
      }
      return Promise.resolve();
    },
    validateJob(job) { validatedStatuses.push(job.status); },
    now: deterministicNow(),
    onSettled() { settledCalls += 1; }
  });
  controller.open({ job: createIdleJob(core, 'retry-flow', [{}, {}]) });

  const partial = await controller.start();
  assert.equal(partial.status, 'completed');
  assert.equal(elements['import-preview-job-status'].textContent, '一部失敗');
  assert.equal(elements['import-preview-retry'].disabled, false);

  const completed = await controller.retryFailed();
  assert.equal(completed.status, 'completed');
  assert.equal(elements['import-preview-job-status'].textContent, '完了');
  assert.deepEqual(plain(calls), { 'retry-flow-item-1': 2, 'retry-flow-item-2': 1 });
  assert.equal(completed.items[0].attempts, 2);
  assert.equal(completed.items[1].attempts, 1);
  assert.equal(settledCalls, 2);
  assert.deepEqual(validatedStatuses, ['idle', 'ready']);

  const same = await controller.retryFailed();
  assert.equal(same, completed);
  assert.equal(settledCalls, 2);
});

test('operation errors use textContent and clear on the next normal operation', async () => {
  let failClock = true;
  const { core, flow, elements } = loadModules();
  const gate = createDeferred();
  const controller = flow.create({
    processItem() { return gate.promise; },
    now() {
      if (failClock) throw new Error('<img src=x onerror=alert(1)>');
      return 'safe-time';
    }
  });
  const ready = core.readyJob(createIdleJob(core, 'error-flow', [{}]));
  controller.open({ job: ready });

  await assert.rejects(controller.start());
  assert.equal(elements['import-preview-operation-error'].textContent, '<img src=x onerror=alert(1)>');
  assert.equal(elements['import-preview-operation-error'].style.display, 'block');

  failClock = false;
  const running = controller.start();
  await flushMicrotasks();
  assert.equal(elements['import-preview-operation-error'].textContent, '');
  assert.equal(elements['import-preview-operation-error'].style.display, 'none');
  gate.resolve();
  await running;
});

test('resume is rejected while a cancelled job still has processing items', async () => {
  const { core, flow, elements } = loadModules();
  const gate = createDeferred();
  const controller = flow.create({
    processItem() { return gate.promise; },
    now: deterministicNow()
  });
  controller.open({ job: createIdleJob(core, 'resume-blocked', [{}, {}]) });
  const running = controller.start();
  await flushMicrotasks();
  controller.cancel();

  await assert.rejects(
    controller.resume(),
    (error) => error.code === 'IMPORT_FLOW_OPERATION_RUNNING'
      || error.code === 'IMPORT_JOB_PROCESSING_ITEMS_REMAIN'
  );
  assert.notEqual(elements['import-preview-operation-error'].textContent, '');
  gate.resolve();
  await running;
});

test('callback exceptions do not break UI or queue and onSettled fires once per run', async () => {
  const { core, flow, elements } = loadModules();
  let changeCalls = 0;
  let settledCalls = 0;
  const controller = flow.create({
    processItem() {},
    now: deterministicNow(),
    onJobChange() { changeCalls += 1; throw new Error('change callback'); },
    onSettled() { settledCalls += 1; throw new Error('settled callback'); }
  });
  controller.open({ job: createIdleJob(core, 'callbacks', [{}]) });

  const finalJob = await controller.start();

  assert.equal(finalJob.status, 'completed');
  assert.ok(changeCalls >= 4);
  assert.equal(settledCalls, 1);
  assert.equal(elements['import-preview-job-status'].textContent, '完了');
});

test('async callback rejections are observed without interrupting the workflow', async () => {
  const { core, flow } = loadModules();
  let callbackCalls = 0;
  let rejectionHandlers = 0;
  function rejectedThenable() {
    callbackCalls += 1;
    return {
      catch(handler) {
        rejectionHandlers += 1;
        handler(new Error('async observer failure'));
      }
    };
  }
  const controller = flow.create({
    processItem() {},
    now: deterministicNow(),
    onJobChange() { return rejectedThenable(); },
    onSettled() { return rejectedThenable(); }
  });
  controller.open({ job: createIdleJob(core, 'async-callbacks', [{}]) });

  const finalJob = await controller.start();

  assert.equal(finalJob.status, 'completed');
  assert.equal(rejectionHandlers, callbackCalls);
  assert.ok(rejectionHandlers > 1);
});

test('normal close retains resources while discard close releases them and controller references', async () => {
  const revoked = [];
  const { core, flow, ui } = loadModules({
    urlApi: { revokeObjectURL(url) { revoked.push(url); } }
  });
  const job = createIdleJob(core, 'close-flow', [{
    runtime: {
      originalFile: new Blob(['original']),
      uploadFile: new Blob(['upload']),
      previewUrl: 'blob:close-flow'
    }
  }]);
  const closeEvents = [];
  const controller = flow.create({ processItem() {}, now: deterministicNow() });
  controller.open({ job, onClose(info) { closeEvents.push(info.discarded); } });

  assert.equal(ui.close(), true);
  assert.equal(controller.getJob(), job);
  assert.equal(job.items[0].runtime.previewUrl, 'blob:close-flow');
  assert.deepEqual(revoked, []);

  controller.open({ job, onClose(info) { closeEvents.push(info.discarded); } });
  assert.equal(ui.close({ discard: true }), true);
  assert.equal(controller.getJob(), null);
  assert.equal(job.items[0].runtime.previewUrl, '');
  assert.deepEqual(revoked, ['blob:close-flow']);
  assert.deepEqual(closeEvents, [false, true]);
});

test('close and discard stay blocked until job-settled notifications and the run Promise finish', async () => {
  const revoked = [];
  const harness = loadModules({ urlApi: { revokeObjectURL(url) { revoked.push(url); } } });
  const { core, flow, ui } = harness;
  const closeAttempts = [];
  let controller = null;
  controller = flow.create({
    processItem() {},
    now: deterministicNow(),
    onJobChange(_job, event) {
      if (event.type === 'item-succeeded' || event.type === 'job-settled') {
        closeAttempts.push(ui.close({ discard: true }));
      }
    }
  });
  const job = createIdleJob(core, 'settled-close', [{
    runtime: { previewUrl: 'blob:settled-close' }
  }]);
  controller.open({ job });

  const finalJob = await controller.start();

  assert.deepEqual(closeAttempts, [false, false]);
  assert.equal(controller.getJob(), finalJob);
  assert.equal(finalJob.items[0].runtime.previewUrl, 'blob:settled-close');
  assert.equal(ui.close({ discard: true }), true);
  assert.equal(controller.getJob(), null);
  assert.deepEqual(revoked, ['blob:settled-close']);
});

test('opening another job while a run is active is rejected without changing the UI job', async () => {
  const { core, flow, ui } = loadModules();
  const gate = createDeferred();
  const controller = flow.create({
    processItem() { return gate.promise; },
    now: deterministicNow()
  });
  const first = createIdleJob(core, 'open-first', [{}]);
  const second = createIdleJob(core, 'open-second', [{}]);
  controller.open({ job: first });
  const running = controller.start();
  await flushMicrotasks();

  assert.throws(
    () => controller.open({ job: second }),
    (error) => error.code === 'IMPORT_FLOW_RUNNING'
  );
  assert.equal(controller.getJob().id, 'open-first');
  assert.equal(ui.getJob().id, 'open-first');

  gate.resolve();
  await running;
});

test('flow controller creates one runner and exposes the expected public API', () => {
  const { flow } = loadModules();
  const controller = flow.create({ processItem() {} });
  for (const method of ['open', 'start', 'cancel', 'resume', 'retryFailed', 'getJob', 'isRunning']) {
    assert.equal(typeof controller[method], 'function', method);
  }
  assert.throws(
    () => flow.create({}),
    (error) => error.code === 'IMPORT_QUEUE_PROCESSOR_REQUIRED'
  );
});

test('optional draft validation runs before ready state, resize, or processing', async () => {
  const { core, flow, elements } = loadModules();
  let validations = 0;
  let processCalls = 0;
  const controller = flow.create({
    processItem() { processCalls += 1; },
    validateJob(job) {
      validations += 1;
      assert.equal(job.status, 'idle');
      const error = new Error('photo.jpg: タイトルを入力してください。');
      error.code = 'IMPORT_DRAFT_TITLE_REQUIRED';
      throw error;
    }
  });
  const idle = createIdleJob(core, 'invalid-draft', [{ title: '' }]);
  controller.open({ job: idle });

  await assert.rejects(controller.start(), (error) => error.code === 'IMPORT_DRAFT_TITLE_REQUIRED');

  assert.equal(validations, 1);
  assert.equal(processCalls, 0);
  assert.equal(controller.getJob().status, 'idle');
  assert.equal(elements['import-preview-operation-error'].textContent, 'photo.jpg: タイトルを入力してください。');
});

test('production Processor and Flow retry only the failed item with stable folder and idempotency key', async () => {
  const { core, flow, processor, validator } = loadModules();
  const payloads = [];
  const savedPins = new Map();
  let secondAttempts = 0;
  let currentFormFolder = 'folder-snapshot';
  const targetFolderIdSnapshot = currentFormFolder;
  const itemProcessor = processor.create({
    resizePhoto: async (file) => `jpeg:${file.name}`,
    withEditToken: (payload) => ({ ...payload, editToken: 'token' }),
    getTargetFolderId: () => targetFolderIdSnapshot,
    callGAS: async (_method, payload) => {
      payloads.push(payload);
      if (payload.itemId === 'integrated-item-2') {
        secondAttempts += 1;
        if (secondAttempts === 1) {
          return { ok: false, errorCode: 'IMPORT_ITEM_LEASE_LOST', error: 'lease lost', retryable: true };
        }
        return { ok: true, deduplicated: true, pin: { id: 'pin-2', title: 'second' } };
      }
      return { ok: true, deduplicated: false, pin: { id: 'pin-1', title: 'first' } };
    },
    onSaved(pin) { savedPins.set(pin.id, pin); }
  });
  const controller = flow.create({
    concurrency: 2,
    processItem: itemProcessor.processItem,
    validateJob: validator.validateJob,
    now: deterministicNow()
  });
  const common = {
    title: '写真', color: '#e53935', icon: 'photo', status: '', tags: [],
    lat: null, lng: null
  };
  const job = createIdleJob(core, 'integrated', [
    { ...common, id: 'integrated-item-1', runtime: { uploadFile: { name: 'one.jpg' } } },
    { ...common, id: 'integrated-item-2', runtime: { uploadFile: { name: 'two.jpg' } } }
  ]);
  controller.open({ job });
  currentFormFolder = 'folder-changed-after-snapshot';

  const partial = await controller.start();
  assert.equal(partial.counts.succeeded, 1);
  assert.equal(partial.counts.failed, 1);
  const completed = await controller.retryFailed();

  assert.equal(completed.counts.succeeded, 2);
  assert.equal(payloads.length, 3);
  assert.deepEqual(payloads.map((payload) => payload.targetFolderId), [
    'folder-snapshot', 'folder-snapshot', 'folder-snapshot'
  ]);
  const secondPayloads = payloads.filter((payload) => payload.itemId === 'integrated-item-2');
  assert.equal(secondPayloads.length, 2);
  assert.equal(secondPayloads[0].idempotencyKey, 'integrated:integrated-item-2');
  assert.equal(secondPayloads[1].idempotencyKey, secondPayloads[0].idempotencyKey);
  assert.deepEqual(Array.from(savedPins.keys()).sort(), ['pin-1', 'pin-2']);
});

test('flow open forwards production resource, preset, and close settings without changing defaults', () => {
  const { core, flow, ui } = loadModules();
  const resourceUrlApi = { revokeObjectURL() {} };
  const loadInputPresets = () => Promise.resolve();
  const getInputPresets = () => [];
  const captured = [];
  const originalOpen = ui.open;
  ui.open = function(options) {
    captured.push(options);
    return originalOpen(options);
  };
  const controller = flow.create({ processItem() {} });
  controller.open({
    job: createIdleJob(core, 'forwarded', [{}]),
    resourceUrlApi,
    loadInputPresets,
    getInputPresets,
    closePolicy: 'discard-only'
  });
  assert.equal(captured[0].resourceUrlApi, resourceUrlApi);
  assert.equal(captured[0].loadInputPresets, loadInputPresets);
  assert.equal(captured[0].getInputPresets, getInputPresets);
  assert.equal(captured[0].closePolicy, 'discard-only');
  ui.close({ discard: true });

  controller.open({ job: createIdleJob(core, 'default-open', [{}]) });
  assert.equal(captured[1].closePolicy, undefined);
  assert.equal(captured[1].resourceUrlApi, undefined);
});

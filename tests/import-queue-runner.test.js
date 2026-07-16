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

function loadImportModules(extra = {}) {
  const start = indexHtml.indexOf('const ImportJobCore = (function() {');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1, 'Expected ImportJobCore');
  assert.notEqual(end, -1, 'Expected import modules before application state');
  const context = {
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'default' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    URL,
    Date,
    ...extra
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__core = ImportJobCore;\n'
      + 'globalThis.__queue = ImportQueueRunner;',
    context
  );
  return { core: context.__core, queue: context.__queue };
}

function createReadyJob(core, id, items) {
  return core.readyJob(core.createJob({
    id,
    sourceType: 'test',
    items: items.map((item, index) => ({
      id: item.id || `${id}-item-${index + 1}`,
      sourceType: item.sourceType || 'test',
      ...item
    }))
  }));
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

async function flushMicrotasks(rounds = 8) {
  for (let index = 0; index < rounds; index += 1) await Promise.resolve();
}

function deterministicNow() {
  let tick = 0;
  return () => `time-${++tick}`;
}

test('runner requires processItem and accepts only concurrency 1 or 2', () => {
  const { queue } = loadImportModules();

  assert.throws(
    () => queue.create({}),
    (error) => error.code === 'IMPORT_QUEUE_PROCESSOR_REQUIRED'
  );
  for (const concurrency of [0, 3, -1, 1.5, '1']) {
    assert.throws(
      () => queue.create({ concurrency, processItem() {} }),
      (error) => error.code === 'INVALID_IMPORT_QUEUE_CONCURRENCY'
    );
  }
  assert.doesNotThrow(() => queue.create({ concurrency: 1, processItem() {} }));
  assert.doesNotThrow(() => queue.create({ concurrency: 2, processItem() {} }));
});

test('start accepts only ready jobs', async () => {
  const { core, queue } = loadImportModules();
  const runner = queue.create({ processItem() {} });
  const idle = core.createJob({ id: 'idle', items: [{ id: 'idle-item' }] });
  const completed = Object.assign({}, idle, { status: 'completed' });

  for (const invalid of [null, idle, completed]) {
    await assert.rejects(
      runner.start(invalid),
      (error) => error.code === 'INVALID_IMPORT_QUEUE_JOB'
    );
  }

  const ready = core.readyJob(idle);
  const finalJob = await runner.start(ready);
  assert.equal(finalJob.status, 'completed');
  assert.match(finalJob.startedAt, /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$/);
});

test('default concurrency is one and replenishes three queued items sequentially', async () => {
  const { core, queue } = loadImportModules();
  const gates = new Map();
  const started = [];
  let active = 0;
  let maximumActive = 0;
  const runner = queue.create({
    processItem(item) {
      active += 1;
      maximumActive = Math.max(maximumActive, active);
      started.push(item.id);
      const gate = createDeferred();
      gates.set(item.id, gate);
      return gate.promise.finally(() => { active -= 1; });
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'serial', [{}, {}, {}]);

  const resultPromise = runner.start(ready);
  await flushMicrotasks();
  assert.deepEqual(started, ['serial-item-1']);
  assert.equal(maximumActive, 1);

  gates.get('serial-item-1').resolve();
  await flushMicrotasks();
  assert.deepEqual(started, ['serial-item-1', 'serial-item-2']);
  assert.equal(maximumActive, 1);

  gates.get('serial-item-2').resolve();
  await flushMicrotasks();
  assert.deepEqual(started, ['serial-item-1', 'serial-item-2', 'serial-item-3']);
  gates.get('serial-item-3').resolve();

  const finalJob = await resultPromise;
  assert.equal(finalJob.status, 'completed');
  assert.equal(finalJob.counts.succeeded, 3);
  assert.equal(maximumActive, 1);
});

test('concurrency two never exceeds two and refills open slots', async () => {
  const { core, queue } = loadImportModules();
  const gates = new Map();
  const started = [];
  let active = 0;
  let maximumActive = 0;
  const runner = queue.create({
    concurrency: 2,
    processItem(item) {
      active += 1;
      maximumActive = Math.max(maximumActive, active);
      started.push(item.id);
      const gate = createDeferred();
      gates.set(item.id, gate);
      return gate.promise.finally(() => { active -= 1; });
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'parallel', [{}, {}, {}, {}]);

  const resultPromise = runner.start(ready);
  await flushMicrotasks();
  assert.deepEqual(started, ['parallel-item-1', 'parallel-item-2']);
  assert.equal(maximumActive, 2);

  gates.get('parallel-item-1').resolve();
  await flushMicrotasks();
  assert.deepEqual(started, ['parallel-item-1', 'parallel-item-2', 'parallel-item-3']);
  assert.equal(active, 2);

  gates.get('parallel-item-2').resolve();
  gates.get('parallel-item-3').resolve();
  await flushMicrotasks();
  assert.deepEqual(started, [
    'parallel-item-1', 'parallel-item-2', 'parallel-item-3', 'parallel-item-4'
  ]);
  assert.ok(maximumActive <= 2);
  gates.get('parallel-item-4').resolve();

  const finalJob = await resultPromise;
  assert.equal(finalJob.status, 'completed');
  assert.equal(finalJob.counts.succeeded, 4);
  assert.equal(maximumActive, 2);
});

test('concurrency two registers a full 20-item job without exceeding two in flight', async () => {
  const { core, queue } = loadImportModules();
  let active = 0;
  let maximumActive = 0;
  let calls = 0;
  const runner = queue.create({
    concurrency: 2,
    async processItem() {
      calls += 1;
      active += 1;
      maximumActive = Math.max(maximumActive, active);
      await new Promise((resolve) => setImmediate(resolve));
      active -= 1;
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'twenty', Array.from({ length: 20 }, () => ({})));

  const finalJob = await runner.start(ready);

  assert.equal(calls, 20);
  assert.equal(maximumActive, 2);
  assert.equal(finalJob.status, 'completed');
  assert.equal(finalJob.counts.succeeded, 20);
});

test('sync throws and promise rejections become failures without stopping remaining items', async () => {
  const { core, queue } = loadImportModules();
  const calls = [];
  const runner = queue.create({
    concurrency: 2,
    processItem(item) {
      calls.push(item.id);
      if (item.id.endsWith('-1')) throw new Error('sync failure');
      if (item.id.endsWith('-2')) return Promise.reject(new Error('async failure'));
      return Promise.resolve('ignored result');
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'mixed', [{}, {}, {}]);

  const finalJob = await runner.start(ready);

  assert.equal(finalJob.status, 'completed');
  assert.equal(finalJob.counts.failed, 2);
  assert.equal(finalJob.counts.succeeded, 1);
  assert.deepEqual(calls.sort(), ['mixed-item-1', 'mixed-item-2', 'mixed-item-3']);
  assert.equal(finalJob.items[0].error, 'sync failure');
  assert.equal(finalJob.items[1].error, 'async failure');
  assert.equal(Object.hasOwn(finalJob.items[2], 'result'), false);
});

test('all-success jobs resolve completed and preserve runtime resources', async () => {
  const { core, queue } = loadImportModules();
  const originalFile = new Blob(['original'], { type: 'image/jpeg' });
  const uploadFile = new Blob(['upload'], { type: 'image/jpeg' });
  let processorItem = null;
  const runner = queue.create({
    processItem(item) {
      processorItem = item;
      assert.equal(item.runtime.originalFile, originalFile);
      assert.equal(item.runtime.uploadFile, uploadFile);
      item.title = 'mutated externally';
      item.tags.push('mutated');
      item.links.push('https://mutated.example');
      item.runtime.previewUrl = 'javascript:mutated';
      return { arbitrary: 'must be ignored' };
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'resource', [{
    title: 'original title',
    tags: ['safe'],
    links: ['https://safe.example'],
    runtime: { originalFile, uploadFile, previewUrl: 'blob:preview' }
  }]);

  const finalJob = await runner.start(ready);

  assert.notEqual(processorItem, finalJob.items[0]);
  assert.equal(finalJob.status, 'completed');
  assert.equal(finalJob.items[0].title, 'original title');
  assert.deepEqual(plain(finalJob.items[0].tags), ['safe']);
  assert.deepEqual(plain(finalJob.items[0].links), ['https://safe.example']);
  assert.equal(finalJob.items[0].runtime.originalFile, originalFile);
  assert.equal(finalJob.items[0].runtime.uploadFile, uploadFile);
  assert.equal(finalJob.items[0].runtime.previewUrl, 'blob:preview');
  assert.equal(Object.hasOwn(finalJob.items[0], 'arbitrary'), false);
});

test('succeeded items are skipped and attempts increment on each processing start', async () => {
  const { core, queue } = loadImportModules();
  const processed = [];
  const contexts = [];
  const runner = queue.create({
    processItem(item, context) {
      processed.push(item.id);
      contexts.push({ ...context });
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'attempts', [
    { id: 'already-done', uploadStatus: 'succeeded', attempts: 1 },
    { id: 'queued-again', uploadStatus: 'queued', attempts: 2 }
  ]);

  const finalJob = await runner.start(ready);

  assert.deepEqual(processed, ['queued-again']);
  assert.deepEqual(contexts, [{ jobId: 'attempts', itemId: 'queued-again', attempt: 3 }]);
  assert.equal(finalJob.items[0].attempts, 1);
  assert.equal(finalJob.items[1].attempts, 3);
  assert.equal(finalJob.items[1].uploadStatus, 'succeeded');
});

test('a running runner rejects double start and can be reused after settlement', async () => {
  const { core, queue } = loadImportModules();
  const firstGate = createDeferred();
  const processed = [];
  const runner = queue.create({
    processItem(item) {
      processed.push(item.id);
      return item.id === 'first-item-1' ? firstGate.promise : Promise.resolve();
    },
    now: deterministicNow()
  });
  const first = createReadyJob(core, 'first', [{}]);
  const second = createReadyJob(core, 'second', [{}]);

  const firstPromise = runner.start(first);
  await flushMicrotasks();
  assert.equal(runner.isRunning(), true);
  await assert.rejects(
    runner.start(second),
    (error) => error.code === 'IMPORT_QUEUE_ALREADY_RUNNING'
  );

  firstGate.resolve();
  const firstFinal = await firstPromise;
  assert.equal(firstFinal.status, 'completed');
  assert.equal(runner.isRunning(), false);

  const secondFinal = await runner.start(second);
  assert.equal(secondFinal.status, 'completed');
  assert.deepEqual(processed, ['first-item-1', 'second-item-1']);
  assert.equal(runner.getJob(), secondFinal);
});

test('state notification exceptions are isolated and event order includes item ids', async () => {
  const { core, queue } = loadImportModules();
  const events = [];
  const runner = queue.create({
    processItem() {},
    onStateChange(job, event) {
      events.push({ type: event.type, itemId: event.itemId || null, status: job.status });
      throw new Error('listener failure');
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'events', [{ id: 'event-item' }]);

  const finalJob = await runner.start(ready);

  assert.equal(finalJob.status, 'completed');
  assert.deepEqual(events, [
    { type: 'job-started', itemId: null, status: 'running' },
    { type: 'item-started', itemId: 'event-item', status: 'running' },
    { type: 'item-succeeded', itemId: 'event-item', status: 'completed' },
    { type: 'job-settled', itemId: null, status: 'completed' }
  ]);
});

test('failed notifications identify the item without exposing errors in the event', async () => {
  const { core, queue } = loadImportModules();
  const events = [];
  const runner = queue.create({
    processItem() { throw new Error('private failure detail'); },
    onStateChange(_job, event) { events.push({ ...event }); },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'failed-event', [{ id: 'failed-item' }]);

  await runner.start(ready);

  const failedEvent = events.find((event) => event.type === 'item-failed');
  assert.deepEqual(failedEvent, { type: 'item-failed', itemId: 'failed-item' });
});

test('cancel is idempotent, starts no new items, and accepts an in-flight success', async () => {
  const { core, queue } = loadImportModules();
  const gate = createDeferred();
  const started = [];
  const events = [];
  const runner = queue.create({
    processItem(item) {
      started.push(item.id);
      return gate.promise;
    },
    onStateChange(_job, event) { events.push({ ...event }); },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'cancel', [{}, {}, {}]);

  const resultPromise = runner.start(ready);
  await flushMicrotasks();
  assert.deepEqual(started, ['cancel-item-1']);
  const firstCancel = runner.cancel();
  const secondCancel = runner.cancel();
  assert.equal(firstCancel.status, 'cancelled');
  assert.equal(secondCancel, firstCancel);
  assert.equal(events.filter((event) => event.type === 'cancel-requested').length, 1);

  gate.resolve();
  const finalJob = await resultPromise;

  assert.equal(finalJob.status, 'cancelled');
  assert.equal(finalJob.counts.succeeded, 1);
  assert.equal(finalJob.counts.waiting, 2);
  assert.equal(finalJob.counts.processing, 0);
  assert.deepEqual(started, ['cancel-item-1']);
  assert.deepEqual(events.map((event) => event.type), [
    'job-started', 'item-started', 'cancel-requested', 'item-succeeded', 'job-settled'
  ]);
  assert.equal(runner.isRunning(), false);

  const before = plain(finalJob);
  assert.equal(runner.cancel(), finalJob);
  assert.deepEqual(plain(finalJob), before);
});

test('cancel with concurrency two settles all in-flight results and leaves queued work untouched', async () => {
  const { core, queue } = loadImportModules();
  const gates = new Map();
  const started = [];
  const runner = queue.create({
    concurrency: 2,
    processItem(item) {
      started.push(item.id);
      const gate = createDeferred();
      gates.set(item.id, gate);
      return gate.promise;
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'cancel-parallel', [{}, {}, {}]);

  const resultPromise = runner.start(ready);
  await flushMicrotasks();
  assert.deepEqual(started, ['cancel-parallel-item-1', 'cancel-parallel-item-2']);
  runner.cancel();
  gates.get('cancel-parallel-item-1').resolve();
  gates.get('cancel-parallel-item-2').reject(new Error('in-flight failure'));

  const finalJob = await resultPromise;

  assert.equal(finalJob.status, 'cancelled');
  assert.equal(finalJob.counts.succeeded, 1);
  assert.equal(finalJob.counts.failed, 1);
  assert.equal(finalJob.counts.waiting, 1);
  assert.equal(finalJob.items[2].uploadStatus, 'queued');
  assert.deepEqual(started, ['cancel-parallel-item-1', 'cancel-parallel-item-2']);
});

test('cancel requested by job-started notification resolves without starting an item', async () => {
  const { core, queue } = loadImportModules();
  const started = [];
  let runner = null;
  runner = queue.create({
    processItem(item) { started.push(item.id); },
    onStateChange(_job, event) {
      if (event.type === 'job-started') runner.cancel();
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'cancel-before-item', [{}, {}]);

  const finalJob = await runner.start(ready);

  assert.equal(finalJob.status, 'cancelled');
  assert.equal(finalJob.counts.waiting, 2);
  assert.equal(finalJob.counts.processing, 0);
  assert.deepEqual(started, []);
  assert.equal(runner.isRunning(), false);
});

test('a cancelled job can be resumed and runs only its remaining items', async () => {
  const { core, queue } = loadImportModules();
  const firstGate = createDeferred();
  const calls = [];
  const runner = queue.create({
    processItem(item) {
      calls.push(item.id);
      if (item.id === 'resume-item-1' && calls.length === 1) return firstGate.promise;
      return Promise.resolve();
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'resume', [{}, {}, {}]);

  const firstRun = runner.start(ready);
  await flushMicrotasks();
  runner.cancel();
  firstGate.resolve();
  const cancelled = await firstRun;
  assert.equal(cancelled.status, 'cancelled');

  const resumed = core.resumeCancelledJob(cancelled);
  const completed = await runner.start(resumed);

  assert.equal(completed.status, 'completed');
  assert.deepEqual(calls, ['resume-item-1', 'resume-item-2', 'resume-item-3']);
  assert.equal(completed.items[0].attempts, 1);
  assert.equal(completed.items[1].attempts, 1);
  assert.equal(completed.items[2].attempts, 1);
});

test('retryFailedItems runs only failures and increments their attempts', async () => {
  const { core, queue } = loadImportModules();
  const callCounts = Object.create(null);
  const runner = queue.create({
    processItem(item) {
      callCounts[item.id] = (callCounts[item.id] || 0) + 1;
      if (item.id === 'retry-failed' && callCounts[item.id] === 1) {
        return Promise.reject(new Error('first attempt failed'));
      }
      return Promise.resolve();
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'retry', [
    { id: 'retry-failed' },
    { id: 'retry-succeeded' }
  ]);

  const firstFinal = await runner.start(ready);
  assert.equal(firstFinal.status, 'completed');
  assert.equal(firstFinal.counts.failed, 1);
  assert.equal(firstFinal.counts.succeeded, 1);

  const retryReady = core.retryFailedItems(firstFinal);
  const retryFinal = await runner.start(retryReady);

  assert.equal(retryFinal.status, 'completed');
  assert.deepEqual(plain(callCounts), { 'retry-failed': 2, 'retry-succeeded': 1 });
  assert.equal(retryFinal.items[0].attempts, 2);
  assert.equal(retryFinal.items[1].attempts, 1);
  assert.equal(retryFinal.counts.succeeded, 2);
});

test('runner preserves safe error code and retryability without retaining the exception', async () => {
  const { core, queue } = loadImportModules();
  const runner = queue.create({
    processItem() {
      const error = new Error('入力内容を確認してください。');
      error.code = 'INVALID_IMPORT_PAYLOAD';
      error.retryable = false;
      error.stack = 'secret stack trace';
      return Promise.reject(error);
    },
    now: deterministicNow()
  });

  const finalJob = await runner.start(createReadyJob(core, 'coded-failure', [{}]));
  const failed = finalJob.items[0];
  assert.equal(failed.error, '入力内容を確認してください。');
  assert.equal(failed.errorCode, 'INVALID_IMPORT_PAYLOAD');
  assert.equal(failed.retryable, false);
  assert.equal(JSON.stringify(core.toPersistableItem(failed)).includes('secret stack trace'), false);
  assert.equal(core.retryFailedItems(finalJob), finalJob);
});

test('simultaneous concurrency-two completions do not lose either job update', async () => {
  const { core, queue } = loadImportModules();
  const gates = [];
  const runner = queue.create({
    concurrency: 2,
    processItem() {
      const gate = createDeferred();
      gates.push(gate);
      return gate.promise;
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'simultaneous', [{}, {}]);

  const resultPromise = runner.start(ready);
  await flushMicrotasks();
  assert.equal(gates.length, 2);
  gates[0].resolve();
  gates[1].resolve();

  const finalJob = await resultPromise;

  assert.equal(finalJob.status, 'completed');
  assert.equal(finalJob.counts.succeeded, 2);
  assert.equal(finalJob.counts.processing, 0);
  assert.deepEqual(finalJob.items.map((item) => item.uploadStatus), ['succeeded', 'succeeded']);
});

test('unstringifiable rejection cannot end the run before another worker settles', async () => {
  const { core, queue } = loadImportModules();
  const secondGate = createDeferred();
  const started = [];
  const runner = queue.create({
    concurrency: 2,
    processItem(item) {
      started.push(item.id);
      if (item.id === 'hostile-item-1') return Promise.reject(Object.create(null));
      if (item.id === 'hostile-item-2') return secondGate.promise;
      return Promise.resolve();
    },
    now: deterministicNow()
  });
  const ready = createReadyJob(core, 'hostile', [{}, {}, {}]);
  let settled = false;
  const observed = runner.start(ready).then(
    (job) => ({ job }),
    (error) => ({ error })
  );
  observed.then(() => { settled = true; });

  await flushMicrotasks();
  const settledBeforeSecondWorker = settled;
  const runningBeforeSecondWorker = runner.isRunning();
  secondGate.resolve();
  const outcome = await observed;

  assert.equal(settledBeforeSecondWorker, false);
  assert.equal(runningBeforeSecondWorker, true);
  assert.equal(outcome.error, undefined);
  assert.equal(outcome.job.status, 'completed');
  assert.equal(outcome.job.counts.failed, 1);
  assert.equal(outcome.job.counts.succeeded, 2);
  assert.equal(outcome.job.items[0].error, 'Import item failed.');
  assert.deepEqual(started, ['hostile-item-1', 'hostile-item-2', 'hostile-item-3']);
});

test('runner stays decoupled from preview UI, console logging, and resource disposal', () => {
  const start = indexHtml.indexOf('const ImportQueueRunner = (function() {');
  const end = indexHtml.indexOf('const ImportPreviewUI = (function() {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const source = indexHtml.slice(start, end);

  assert.equal(source.includes('ImportPreviewUI'), false);
  assert.equal(source.includes('console.'), false);
  assert.equal(source.includes('releaseJobResources'), false);
  assert.equal(source.includes('releaseItemResources'), false);
});

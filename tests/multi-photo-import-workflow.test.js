const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

function loadWorkflow() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = {
    console,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'default' }, { id: 'photo' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    URL: { createObjectURL() {}, revokeObjectURL() {} },
    crypto: { randomUUID: () => 'uuid' }
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__workflow = typeof MultiPhotoImportWorkflow === "undefined" ? null : MultiPhotoImportWorkflow;',
    context
  );
  return context.__workflow;
}

function createHarness(overrides = {}) {
  const workflowApi = loadWorkflow();
  assert.ok(workflowApi);
  const audit = { builderConfigs: [], starts: [], processorConfigs: [], flowConfigs: [], opens: [], remembered: [] };
  const state = {
    builder: null, controller: null, job: null, resourceUrlApi: null,
    targetFolderId: '', preparing: false, registering: false,
    cancellingPreparation: false, requestToken: 0
  };
  const resourceUrlApi = { revokeObjectURL() {} };
  let builderSession;
  const config = {
    state,
    builderApi: {
      create(builderConfig) {
        audit.builderConfigs.push(builderConfig);
        builderSession = {
          cancelCalls: 0,
          start(files, defaults) {
            const fileList = Array.from(files);
            audit.starts.push({ files: fileList, defaults });
            return Promise.resolve({
              id: 'job-1',
              status: 'idle',
              items: fileList.map((file, index) => ({
                id: 'item-' + (index + 1),
                runtime: { uploadFile: file }
              }))
            });
          },
          cancel() { this.cancelCalls += 1; },
          release() {}
        };
        return builderSession;
      },
      getResourceUrlApi() { return resourceUrlApi; },
      rememberResourceUrlApi(job, api) { audit.remembered.push([job, api]); }
    },
    processorApi: {
      create(processorConfig) {
        audit.processorConfigs.push(processorConfig);
        return { processItem() {} };
      }
    },
    flowApi: {
      create(flowConfig) {
        audit.flowConfigs.push(flowConfig);
        return {
          open(openOptions) { audit.opens.push(openOptions); },
          isRunning() { return false; }
        };
      }
    },
    validator: { validateJob() { return true; } },
    resizePhoto() {}, callGAS() {}, withEditToken(value) { return value; },
    loadInputPresets() { return Promise.resolve(); }, getInputPresets() { return []; },
    ...overrides
  };
  return { workflow: workflowApi.create(config), audit, state, resourceUrlApi, getBuilder: () => builderSession };
}

test('workflow keeps two-way registration for a lightweight prepared photo batch', async () => {
  const { workflow, audit, state, resourceUrlApi } = createHarness();
  const snapshot = {
    tags: ['観察'], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder-old'
  };
  const startPromise = workflow.start([
    { name: 'one.jpg', size: 2 * 1024 * 1024 },
    { name: 'two.jpg', size: 3 * 1024 * 1024 }
  ], snapshot);
  snapshot.tags.push('late');
  snapshot.targetFolderId = 'folder-new';
  const job = await startPromise;

  assert.equal(job.id, 'job-1');
  assert.equal(audit.builderConfigs[0].concurrency, 2);
  assert.deepEqual(JSON.parse(JSON.stringify(audit.starts[0].defaults)), {
    tags: ['観察'], color: '#e53935', icon: 'photo', status: ''
  });
  assert.equal(audit.processorConfigs[0].getTargetFolderId(), 'folder-old');
  assert.equal(audit.flowConfigs[0].concurrency, 2);
  assert.equal(audit.flowConfigs[0].processItem instanceof Function, true);
  assert.equal(audit.flowConfigs[0].validateJob instanceof Function, true);
  assert.equal(audit.opens[0].resourceUrlApi, resourceUrlApi);
  assert.equal(audit.opens[0].closePolicy, 'discard-only');
  assert.equal(audit.opens[0].title, '複数写真を確認');
  assert.equal(audit.opens[0].sourceLabel, '写真 2件');
  assert.equal(state.targetFolderId, 'folder-old');

  const nextJob = { id: 'job-1', status: 'running', items: job.items, counts: { processing: 1 } };
  audit.flowConfigs[0].onJobChange(nextJob, { type: 'job-running' });
  assert.equal(state.job, nextJob);
  assert.equal(state.registering, true);
  assert.deepEqual(audit.remembered.at(-1), [nextJob, resourceUrlApi]);

  audit.opens[0].onClose({ job: nextJob, discarded: true });
  assert.equal(state.builder, null);
  assert.equal(state.controller, null);
  assert.equal(state.job, null);
  assert.equal(state.targetFolderId, '');
  assert.equal(workflow.isBusy(), false);
});

test('workflow serializes registration for one large photo or a large prepared batch', async () => {
  const snapshot = {
    tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder'
  };
  const largePhoto = createHarness();
  await largePhoto.workflow.start([
    { name: 'large.jpg', size: 12 * 1024 * 1024 }
  ], snapshot);
  assert.equal(largePhoto.audit.flowConfigs[0].concurrency, 1);

  const largeBatch = createHarness();
  await largeBatch.workflow.start(Array.from({ length: 9 }, (_, index) => ({
    name: 'photo-' + index + '.jpg',
    size: 8 * 1024 * 1024
  })), snapshot);
  assert.equal(largeBatch.audit.flowConfigs[0].concurrency, 1);
});

test('preparation cancellation is idempotent, waits for Builder settlement, and ignores late progress', async () => {
  const gate = deferred();
  const progress = [];
  let cancelled = 0;
  const harness = createHarness({
    onProgress(value) { progress.push(value.eventType); },
    onPreparationCancelled() { cancelled += 1; }
  });
  const originalCreate = harness.workflow;
  // Replace the next Builder start through the captured session after start has synchronously created it.
  const promise = originalCreate.start([{ name: 'slow.jpg' }], {
    tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder'
  });
  const builder = harness.getBuilder();
  // The default fake already resolved; use a second harness with a deferred builder for the actual cancellation path.
  await promise;
  harness.audit.opens[0].onClose({ job: harness.state.job, discarded: true });

  let delayedConfig;
  let delayedSession;
  const delayed = createHarness({
    builderApi: {
      create(config) {
        delayedConfig = config;
        delayedSession = { cancelCalls: 0, start: () => gate.promise, cancel() { this.cancelCalls += 1; }, release() {} };
        return delayedSession;
      },
      getResourceUrlApi() { return { revokeObjectURL() {} }; },
      rememberResourceUrlApi() {}
    },
    onProgress(value) { progress.push(value.eventType); },
    onPreparationCancelled() { cancelled += 1; }
  });
  const running = delayed.workflow.start([{ name: 'slow.jpg' }], {
    tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder'
  });
  assert.equal(delayed.workflow.cancelPreparation(), true);
  assert.equal(delayed.workflow.cancelPreparation(), false);
  assert.equal(delayedSession.cancelCalls, 1);
  assert.equal(delayed.state.cancellingPreparation, true);
  const error = new Error('cancelled');
  error.code = 'MULTI_PHOTO_CANCELLED';
  gate.reject(error);
  assert.equal(await running, null);
  assert.equal(cancelled, 1);
  assert.equal(delayed.workflow.isBusy(), false);
  const before = progress.length;
  delayedConfig.onProgress({ eventType: 'late' });
  assert.equal(progress.length, before);
  assert.equal(builder.cancelCalls, 0);
});

test('fatal preparation error is sanitized and releases the session without opening Flow', async () => {
  const failures = [];
  let released = 0;
  const harness = createHarness({
    builderApi: {
      create() {
        return {
          start: () => Promise.reject(new Error('secret stack and image bytes')),
          cancel() {}, release() { released += 1; }
        };
      },
      getResourceUrlApi() { return null; }, rememberResourceUrlApi() {}
    },
    onPreparationError(message) { failures.push(message); }
  });
  await assert.rejects(
    harness.workflow.start([{ name: 'bad.jpg' }], {
      tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder'
    }),
    (error) => error.code === 'MULTI_PHOTO_PREPARATION_FAILED'
  );
  assert.deepEqual(failures, ['写真の準備に失敗しました。もう一度選択してください。']);
  assert.equal(released, 1);
  assert.equal(harness.audit.opens.length, 0);
  assert.equal(harness.workflow.isBusy(), false);
});

test('missing folder snapshot is reported safely without creating a Builder session', async () => {
  const failures = [];
  const harness = createHarness({
    onPreparationError(message) { failures.push(message); }
  });

  await assert.rejects(
    harness.workflow.start([{ name: 'photo.jpg' }], {
      tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: ''
    }),
    (error) => error.code === 'MULTI_PHOTO_TARGET_FOLDER_REQUIRED'
  );

  assert.deepEqual(failures, ['保存先フォルダを選択してください。']);
  assert.equal(harness.audit.builderConfigs.length, 0);
  assert.equal(harness.workflow.isBusy(), false);
});

test('synchronous Builder creation failures are sanitized and clear busy state', async () => {
  const failures = [];
  const harness = createHarness({
    builderApi: {
      create() { throw new Error('secret Builder initialization details'); },
      getResourceUrlApi() { return null; },
      rememberResourceUrlApi() {}
    },
    onPreparationError(message) { failures.push(message); }
  });

  await assert.rejects(
    harness.workflow.start([{ name: 'photo.jpg' }], {
      tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder'
    }),
    (error) => error.code === 'MULTI_PHOTO_PREPARATION_FAILED'
  );

  assert.deepEqual(failures, ['写真の準備に失敗しました。もう一度選択してください。']);
  assert.equal(harness.workflow.isBusy(), false);
});

test('Flow wiring failure releases resources already transferred into the prepared job', async () => {
  const revoked = [];
  const resourceUrlApi = { revokeObjectURL(url) { revoked.push(url); } };
  const uploadFile = { name: 'photo.jpg' };
  const preparedJob = {
    id: 'job-1', status: 'idle',
    items: [{
      id: 'item-1', error: null,
      runtime: { originalFile: uploadFile, uploadFile, previewUrl: 'blob:prepared' }
    }]
  };
  const failures = [];
  const harness = createHarness({
    builderApi: {
      create() {
        return { start: () => Promise.resolve(preparedJob), cancel() {}, release() {} };
      },
      getResourceUrlApi() { return resourceUrlApi; },
      rememberResourceUrlApi() {}
    },
    flowApi: {
      create() { throw new Error('preview initialization failed'); }
    },
    onPreparationError(message) { failures.push(message); }
  });

  await assert.rejects(
    harness.workflow.start([uploadFile], {
      tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'folder'
    }),
    (error) => error.code === 'MULTI_PHOTO_PREPARATION_FAILED'
  );

  assert.deepEqual(revoked, ['blob:prepared']);
  assert.equal(preparedJob.items[0].runtime.originalFile, null);
  assert.equal(preparedJob.items[0].runtime.uploadFile, null);
  assert.equal(preparedJob.items[0].runtime.previewUrl, '');
  assert.deepEqual(failures, ['写真の準備に失敗しました。もう一度選択してください。']);
  assert.equal(harness.workflow.isBusy(), false);
});

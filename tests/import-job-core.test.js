const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function loadImportJobCore(extra = {}) {
  const start = indexHtml.indexOf('const ImportJobCore = (function() {');
  assert.notEqual(start, -1, 'Expected ImportJobCore to exist');
  const end = indexHtml.indexOf('\n    const state = {', start);
  assert.notEqual(end, -1, 'Expected ImportJobCore to precede application state');
  const context = {
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'default' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    URL,
    Date,
    ...extra
  };
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\nglobalThis.__importJobCore = ImportJobCore;`,
    context
  );
  return context.__importJobCore;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function createItems(api, count, overrides = {}) {
  return Array.from({ length: count }, (_, index) => api.createItem({
    id: `item-${index + 1}`,
    sourceType: 'photo',
    sourceRef: `photo-${index + 1}.jpg`,
    ...overrides
  }));
}

test('ImportItem starts with serializable defaults and isolated runtime handles', () => {
  const api = loadImportJobCore();

  assert.deepEqual(plain(api.createItem({
    id: 'item-1',
    sourceType: 'photo',
    sourceRef: 'IMG_0001.jpg'
  })), {
    id: 'item-1',
    sourceType: 'photo',
    sourceRef: 'IMG_0001.jpg',
    title: '',
    description: '',
    lat: null,
    lng: null,
    capturedAt: '',
    tags: [],
    links: [],
    color: '',
    icon: '',
    status: '',
    metadataStatus: 'idle',
    conversionStatus: 'idle',
    uploadStatus: 'queued',
    error: null,
    errorCode: null,
    retryable: null,
    attempts: 0,
    runtime: {
      originalFile: null,
      uploadFile: null,
      previewUrl: ''
    }
  });
});

test('ImportItem persistence excludes File, Blob, Object URL, and unknown input fields', () => {
  const api = loadImportJobCore();
  const originalFile = new Blob(['original'], { type: 'image/jpeg' });
  const uploadFile = new Blob(['upload'], { type: 'image/jpeg' });
  const item = api.createItem({
    id: 'item-1',
    sourceType: 'photo',
    sourceRef: 'IMG_0001.jpg',
    title: '旅行',
    tags: ['夏'],
    originalFile,
    uploadFile,
    previewUrl: 'blob:unexpected-top-level',
    runtime: {
      originalFile,
      uploadFile,
      previewUrl: 'blob:preview-1'
    }
  });

  assert.equal(item.runtime.originalFile, originalFile);
  assert.equal(item.runtime.uploadFile, uploadFile);
  assert.equal(item.runtime.previewUrl, 'blob:preview-1');

  const persistable = api.toPersistableItem(item);
  const json = JSON.stringify(persistable);
  assert.deepEqual(Object.keys(persistable), [
    'id', 'sourceType', 'sourceRef', 'title', 'description', 'lat', 'lng',
    'capturedAt', 'tags', 'links', 'color', 'icon', 'status', 'metadataStatus',
    'conversionStatus', 'uploadStatus', 'error', 'errorCode', 'retryable', 'attempts'
  ]);
  assert.equal(Object.prototype.hasOwnProperty.call(persistable, 'runtime'), false);
  assert.equal(json.includes('blob:'), false);
  assert.equal(json.includes('originalFile'), false);
  assert.equal(json.includes('uploadFile'), false);
  assert.deepEqual(plain(persistable.tags), ['夏']);
  assert.deepEqual(plain(persistable.links), []);
});

test('links are normalized, cloned, draft-editable, and persist without runtime data', () => {
  const api = loadImportJobCore();
  const inputLinks = ['https://example.com'];
  const job = api.createJob({
    id: 'links-job',
    items: [
      { id: 'item-1', links: inputLinks, runtime: { previewUrl: 'blob:one' } },
      { id: 'item-2' }
    ]
  });
  inputLinks.push('https://mutated.example');
  assert.deepEqual(plain(job.items[0].links), ['https://example.com']);
  assert.deepEqual(plain(job.items[1].links), []);

  const patchLinks = ['https://a.example', 'https://b.example'];
  const updated = api.updateDraftItem(job, 'item-1', { links: patchLinks });
  patchLinks.push('https://mutated.example');
  assert.deepEqual(plain(updated.items[0].links), ['https://a.example', 'https://b.example']);
  assert.deepEqual(plain(job.items[0].links), ['https://example.com']);

  const bulk = api.updateDraftItems(updated, [
    { itemId: 'item-1', patch: { links: [] } },
    { itemId: 'item-2', patch: { links: ['https://two.example'] } }
  ]);
  assert.deepEqual(plain(bulk.items.map((item) => item.links)), [[], ['https://two.example']]);
  assert.deepEqual(plain(api.toPersistableItem(bulk.items[1]).links), ['https://two.example']);
  assert.equal(Object.prototype.hasOwnProperty.call(api.toPersistableItem(bulk.items[1]), 'runtime'), false);
  assert.throws(
    () => api.updateDraftItem(job, 'item-1', { links: 'https://example.com' }),
    (error) => error.code === 'INVALID_IMPORT_DRAFT_FIELD'
  );
});

test('registration status is persisted, allowlisted, validated, and independent from upload status', () => {
  const api = loadImportJobCore();
  const job = api.createJob({
    id: 'status-job',
    items: [{ id: 'item-1', status: '', uploadStatus: 'queued' }]
  });

  const updated = api.updateDraftItem(job, 'item-1', { status: '対応中' });
  assert.equal(updated.items[0].status, '対応中');
  assert.equal(updated.items[0].uploadStatus, 'queued');
  assert.equal(job.items[0].status, '');
  assert.equal(api.toPersistableItem(updated.items[0]).status, '対応中');

  const cleared = api.updateDraftItem(updated, 'item-1', { status: '' });
  assert.equal(cleared.items[0].status, '');
  assert.throws(
    () => api.updateDraftItem(job, 'item-1', { status: '未知' }),
    (error) => error.code === 'INVALID_IMPORT_DRAFT_FIELD'
  );
});

test('bulk draft updates validate every patch before atomically creating one next job', () => {
  const api = loadImportJobCore();
  const job = api.createJob({
    id: 'bulk-job',
    items: [
      { id: 'item-1', tags: ['A'], status: '', runtime: { previewUrl: 'blob:1' } },
      { id: 'item-2', tags: ['B'], status: '未対応', runtime: { previewUrl: 'blob:2' } }
    ]
  });
  const before = plain(job);

  const updated = api.updateDraftItems(job, [
    { itemId: 'item-1', patch: { tags: ['共通'], status: '完了' } },
    { itemId: 'item-2', patch: { tags: [], status: '' } }
  ]);
  assert.notEqual(updated, job);
  assert.deepEqual(plain(updated.items.map((item) => item.tags)), [['共通'], []]);
  assert.deepEqual(plain(updated.items.map((item) => item.status)), ['完了', '']);
  assert.deepEqual(plain(job), before);

  assert.throws(() => api.updateDraftItems(job, [
    { itemId: 'item-1', patch: { status: '対応中' } },
    { itemId: 'item-2', patch: { status: '不正' } }
  ]), (error) => error.code === 'INVALID_IMPORT_DRAFT_FIELD');
  assert.deepEqual(plain(job), before);
});

test('ImportJob accepts up to 20 items and exposes initial counts', () => {
  const api = loadImportJobCore();
  const job = api.createJob({
    id: 'job-1',
    sourceType: 'photo',
    items: createItems(api, 20),
    createdAt: '2026-07-10T10:00:00.000Z'
  });

  assert.equal(job.items.length, 20);
  assert.equal(job.status, 'idle');
  assert.equal(job.cancelRequested, false);
  assert.equal(job.cancelRequestedAt, null);
  assert.equal(job.createdAt, '2026-07-10T10:00:00.000Z');
  assert.equal(job.startedAt, null);
  assert.equal(job.finishedAt, null);
  assert.deepEqual(plain(job.counts), {
    total: 20,
    succeeded: 0,
    failed: 0,
    processing: 0,
    waiting: 20
  });
  assert.deepEqual(plain(api.getCounts(job)), plain(job.counts));
});

test('ImportJob rejects a 21st item with an explicit limit error', () => {
  const api = loadImportJobCore();
  const twentyItems = createItems(api, 20);
  const job = api.createJob({ id: 'job-1', sourceType: 'photo', items: twentyItems });
  const extraItem = api.createItem({ id: 'item-21', sourceType: 'photo' });

  assert.throws(
    () => api.createJob({ id: 'too-large', sourceType: 'photo', items: createItems(api, 21) }),
    (error) => error.code === 'IMPORT_ITEM_LIMIT_EXCEEDED'
      && /20/.test(error.message)
  );
  assert.throws(
    () => api.addItem(job, extraItem),
    (error) => error.code === 'IMPORT_ITEM_LIMIT_EXCEEDED'
      && /20/.test(error.message)
  );
});

test('ImportJob rejects empty readiness and invalid or duplicate item ids', () => {
  const api = loadImportJobCore();
  const emptyJob = api.createJob({ id: 'empty-job', sourceType: 'photo' });

  assert.throws(
    () => api.readyJob(emptyJob),
    (error) => error.code === 'IMPORT_JOB_EMPTY'
  );
  assert.throws(
    () => api.createJob({
      id: 'empty-id-job',
      sourceType: 'photo',
      items: [api.createItem({ id: '', sourceType: 'photo' })]
    }),
    (error) => error.code === 'INVALID_IMPORT_ITEM_ID'
  );
  assert.throws(
    () => api.createJob({
      id: 'duplicate-id-job',
      sourceType: 'photo',
      items: [
        api.createItem({ id: 'duplicate', sourceType: 'photo' }),
        api.createItem({ id: 'duplicate', sourceType: 'photo' })
      ]
    }),
    (error) => error.code === 'DUPLICATE_IMPORT_ITEM_ID'
  );
});

test('items can be added only while a job is idle', () => {
  const api = loadImportJobCore();
  const idle = api.createJob({ id: 'job-1', sourceType: 'photo', items: createItems(api, 1) });
  const withSecondItem = api.addItem(
    idle,
    api.createItem({ id: 'item-2', sourceType: 'photo' })
  );
  assert.equal(withSecondItem.items.length, 2);
  assert.equal(idle.items.length, 1);

  const ready = api.readyJob(withSecondItem);
  assert.throws(
    () => api.addItem(ready, api.createItem({ id: 'item-3', sourceType: 'photo' })),
    (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
  );

  let completed = api.startJob(ready, '2026-07-10T10:00:00.000Z');
  completed = api.markItemProcessing(completed, 'item-1');
  completed = api.markItemSucceeded(completed, 'item-1', '2026-07-10T10:01:00.000Z');
  completed = api.markItemProcessing(completed, 'item-2');
  completed = api.markItemSucceeded(completed, 'item-2', '2026-07-10T10:02:00.000Z');
  assert.equal(completed.status, 'completed');
  assert.throws(
    () => api.addItem(completed, api.createItem({ id: 'item-3', sourceType: 'photo' })),
    (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
  );
});

test('ready jobs require queued work and reject invalid initial item states', () => {
  const api = loadImportJobCore();

  assert.throws(
    () => api.createItem({ id: 'item-1', sourceType: 'photo', uploadStatus: 'unknown' }),
    (error) => error.code === 'INVALID_IMPORT_ITEM_STATUS'
  );

  const succeededOnly = api.createJob({
    id: 'succeeded-job',
    sourceType: 'photo',
    items: [api.createItem({
      id: 'item-1', sourceType: 'photo', uploadStatus: 'succeeded'
    })]
  });
  assert.throws(
    () => api.readyJob(succeededOnly),
    (error) => error.code === 'IMPORT_JOB_NO_QUEUED_ITEMS'
  );

  const failedOnly = api.createJob({
    id: 'failed-job',
    sourceType: 'photo',
    items: [api.createItem({
      id: 'item-1', sourceType: 'photo', uploadStatus: 'failed'
    })]
  });
  assert.throws(
    () => api.readyJob(failedOnly),
    (error) => error.code === 'INVALID_IMPORT_JOB_ITEMS'
  );

  const retryShape = api.createJob({
    id: 'retry-job',
    sourceType: 'photo',
    items: [
      api.createItem({ id: 'item-1', sourceType: 'photo', uploadStatus: 'succeeded' }),
      api.createItem({ id: 'item-2', sourceType: 'photo', uploadStatus: 'queued' })
    ]
  });
  assert.equal(api.readyJob(retryShape).status, 'ready');
});

test('ImportJob follows idle to ready to running to completed item transitions', () => {
  const api = loadImportJobCore();
  const idle = api.createJob({
    id: 'job-1',
    sourceType: 'photo',
    items: createItems(api, 2),
    createdAt: '2026-07-10T10:00:00.000Z'
  });
  const ready = api.readyJob(idle);
  const running = api.startJob(ready, '2026-07-10T10:01:00.000Z');

  assert.equal(idle.status, 'idle');
  assert.equal(ready.status, 'ready');
  assert.equal(running.status, 'running');
  assert.equal(running.startedAt, '2026-07-10T10:01:00.000Z');
  assert.equal(api.getNextItem(running).id, 'item-1');

  const firstProcessing = api.markItemProcessing(running, 'item-1');
  assert.equal(running.items[0].uploadStatus, 'queued');
  assert.equal(firstProcessing.items[0].uploadStatus, 'processing');
  assert.equal(firstProcessing.items[0].attempts, 1);
  assert.equal(api.getNextItem(firstProcessing).id, 'item-2');

  const firstSucceeded = api.markItemSucceeded(
    firstProcessing,
    'item-1',
    '2026-07-10T10:02:00.000Z'
  );
  assert.equal(firstSucceeded.status, 'running');
  assert.deepEqual(plain(firstSucceeded.counts), {
    total: 2, succeeded: 1, failed: 0, processing: 0, waiting: 1
  });

  const secondProcessing = api.markItemProcessing(firstSucceeded, 'item-2');
  const completed = api.markItemSucceeded(
    secondProcessing,
    'item-2',
    '2026-07-10T10:03:00.000Z'
  );
  assert.equal(completed.status, 'completed');
  assert.equal(completed.finishedAt, '2026-07-10T10:03:00.000Z');
  assert.equal(api.getNextItem(completed), null);
  assert.deepEqual(plain(completed.counts), {
    total: 2, succeeded: 2, failed: 0, processing: 0, waiting: 0
  });
});

test('ImportJob aggregation reports succeeded, failed, processing, and waiting items', () => {
  const api = loadImportJobCore();
  let job = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 4)
  })), '2026-07-10T10:00:00.000Z');
  job = api.markItemProcessing(job, 'item-1');
  job = api.markItemSucceeded(job, 'item-1', '2026-07-10T10:01:00.000Z');
  job = api.markItemProcessing(job, 'item-2');
  job = api.markItemFailed(job, 'item-2', 'network error', '2026-07-10T10:02:00.000Z');
  job = api.markItemProcessing(job, 'item-3');

  assert.deepEqual(plain(api.getCounts(job)), {
    total: 4,
    succeeded: 1,
    failed: 1,
    processing: 1,
    waiting: 1
  });
  assert.deepEqual(plain(job.counts), plain(api.getCounts(job)));
});

test('retry queues only failed items and never resends succeeded items', () => {
  const api = loadImportJobCore();
  let job = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 2)
  })), '2026-07-10T10:00:00.000Z');
  job = api.markItemProcessing(job, 'item-1');
  job = api.markItemSucceeded(job, 'item-1', '2026-07-10T10:01:00.000Z');
  job = api.markItemProcessing(job, 'item-2');
  job = api.markItemFailed(job, 'item-2', new Error('upload failed'), '2026-07-10T10:02:00.000Z');

  assert.equal(job.status, 'completed');
  assert.equal(job.items[0].uploadStatus, 'succeeded');
  assert.equal(job.items[1].uploadStatus, 'failed');
  assert.match(job.items[1].error, /upload failed/);

  const retryReady = api.retryFailedItems(job);
  assert.notEqual(retryReady.items[0], job.items[0]);
  assert.notEqual(retryReady.items[0].links, job.items[0].links);
  assert.notEqual(retryReady.items[1].links, job.items[1].links);
  assert.equal(retryReady.status, 'ready');
  assert.equal(retryReady.cancelRequestedAt, null);
  assert.equal(retryReady.finishedAt, null);
  assert.equal(retryReady.items[0].uploadStatus, 'succeeded');
  assert.equal(retryReady.items[0].attempts, 1);
  assert.equal(retryReady.items[1].uploadStatus, 'queued');
  assert.equal(retryReady.items[1].attempts, 1);
  assert.equal(retryReady.items[1].error, null);

  const retryRunning = api.startJob(retryReady, '2026-07-10T10:03:00.000Z');
  assert.equal(api.getNextItem(retryRunning).id, 'item-2');
  assert.throws(
    () => api.markItemProcessing(retryRunning, 'item-1'),
    (error) => error.code === 'INVALID_IMPORT_ITEM_TRANSITION'
  );

  const retryProcessing = api.markItemProcessing(retryRunning, 'item-2');
  const succeeded = api.markItemSucceeded(
    retryProcessing,
    'item-2',
    '2026-07-10T10:04:00.000Z'
  );
  assert.equal(succeeded.status, 'completed');
  assert.equal(succeeded.items[0].attempts, 1);
  assert.equal(succeeded.items[1].attempts, 2);
  assert.deepEqual(plain(succeeded.counts), {
    total: 2, succeeded: 2, failed: 0, processing: 0, waiting: 0
  });
});

test('failed item metadata is persistable and non-retryable failures stay failed', () => {
  const api = loadImportJobCore();
  let job = api.startJob(api.readyJob(api.createJob({
    id: 'retry-policy',
    items: [{ id: 'permanent' }, { id: 'temporary' }]
  })), 'start');
  job = api.markItemProcessing(job, 'permanent');
  job = api.markItemFailed(job, 'permanent', Object.assign(
    new Error('入力が不正です。'),
    { code: 'INVALID_IMPORT_PAYLOAD', retryable: false }
  ), 'failed-1');
  job = api.markItemProcessing(job, 'temporary');
  job = api.markItemFailed(job, 'temporary', Object.assign(
    new Error('一時的な失敗です。'),
    { code: 'IMPORT_ITEM_SAVE_FAILED', retryable: true }
  ), 'failed-2');

  assert.equal(job.status, 'completed');
  assert.deepEqual(plain(api.toPersistableItem(job.items[0])), {
    id: 'permanent', sourceType: '', sourceRef: '', title: '', description: '',
    lat: null, lng: null, capturedAt: '', tags: [], links: [], color: '', icon: '', status: '',
    metadataStatus: 'idle', conversionStatus: 'idle', uploadStatus: 'failed',
    error: '入力が不正です。', errorCode: 'INVALID_IMPORT_PAYLOAD', retryable: false,
    attempts: 1
  });

  const retryReady = api.retryFailedItems(job);
  assert.equal(retryReady.status, 'ready');
  assert.equal(retryReady.items[0].uploadStatus, 'failed');
  assert.equal(retryReady.items[0].errorCode, 'INVALID_IMPORT_PAYLOAD');
  assert.equal(retryReady.items[0].retryable, false);
  assert.equal(retryReady.items[1].uploadStatus, 'queued');
  assert.equal(retryReady.items[1].error, null);
  assert.equal(retryReady.items[1].errorCode, null);
  assert.equal(retryReady.items[1].retryable, null);
});

test('cancellation prevents new items from starting but accepts in-flight results', () => {
  const api = loadImportJobCore();
  let running = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 3)
  })), '2026-07-10T10:00:00.000Z');
  running = api.markItemProcessing(running, 'item-1');

  const cancelled = api.requestCancel(running, '2026-07-10T10:01:00.000Z');
  assert.equal(cancelled.status, 'cancelled');
  assert.equal(cancelled.cancelRequested, true);
  assert.equal(cancelled.cancelRequestedAt, '2026-07-10T10:01:00.000Z');
  assert.equal(cancelled.finishedAt, null);
  assert.equal(api.getNextItem(cancelled), null);
  assert.throws(
    () => api.markItemProcessing(cancelled, 'item-2'),
    (error) => error.code === 'IMPORT_JOB_CANCELLED'
  );
  assert.throws(
    () => api.startJob(cancelled, '2026-07-10T10:02:00.000Z'),
    (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
  );
  assert.throws(
    () => api.retryFailedItems(cancelled),
    (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
  );
  assert.throws(
    () => api.resumeCancelledJob(cancelled),
    (error) => error.code === 'IMPORT_JOB_PROCESSING_ITEMS_REMAIN'
  );

  const settled = api.markItemSucceeded(cancelled, 'item-1', '2026-07-10T10:02:00.000Z');
  assert.equal(settled.status, 'cancelled');
  assert.equal(settled.cancelRequested, true);
  assert.equal(settled.cancelRequestedAt, '2026-07-10T10:01:00.000Z');
  assert.equal(settled.finishedAt, '2026-07-10T10:02:00.000Z');
  assert.deepEqual(plain(settled.counts), {
    total: 3, succeeded: 1, failed: 0, processing: 0, waiting: 2
  });
  assert.equal(api.getNextItem(settled), null);
});

test('cancelled jobs safely record success and failure for multiple in-flight items', () => {
  const api = loadImportJobCore();
  let job = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 3)
  })), '2026-07-10T10:00:00.000Z');
  job = api.markItemProcessing(job, 'item-1');
  job = api.markItemProcessing(job, 'item-2');
  job = api.requestCancel(job, '2026-07-10T10:01:00.000Z');
  assert.equal(job.finishedAt, null);
  job = api.markItemSucceeded(job, 'item-1', '2026-07-10T10:02:00.000Z');
  assert.equal(job.finishedAt, null);
  job = api.markItemFailed(job, 'item-2', 'cancelled in flight', '2026-07-10T10:03:00.000Z');

  assert.equal(job.status, 'cancelled');
  assert.equal(job.cancelRequested, true);
  assert.equal(job.cancelRequestedAt, '2026-07-10T10:01:00.000Z');
  assert.equal(job.finishedAt, '2026-07-10T10:03:00.000Z');
  assert.equal(api.getNextItem(job), null);
  assert.deepEqual(plain(job.counts), {
    total: 3, succeeded: 1, failed: 1, processing: 0, waiting: 1
  });
});

test('cancellation without processing items finishes immediately and resumes queued work', () => {
  const api = loadImportJobCore();
  const running = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 2)
  })), '2026-07-10T10:00:00.000Z');

  assert.throws(
    () => api.resumeCancelledJob(running),
    (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
  );

  const cancelled = api.requestCancel(running, '2026-07-10T10:01:00.000Z');
  assert.equal(cancelled.cancelRequestedAt, '2026-07-10T10:01:00.000Z');
  assert.equal(cancelled.finishedAt, '2026-07-10T10:01:00.000Z');

  const resumed = api.resumeCancelledJob(cancelled);
  assert.equal(resumed.status, 'ready');
  assert.equal(resumed.cancelRequested, false);
  assert.equal(resumed.cancelRequestedAt, null);
  assert.equal(resumed.finishedAt, null);
  assert.deepEqual(plain(resumed.items.map((item) => item.uploadStatus)), ['queued', 'queued']);
  assert.deepEqual(plain(resumed.counts), {
    total: 2, succeeded: 0, failed: 0, processing: 0, waiting: 2
  });
});

test('resuming cancellation starts queued items but preserves failures for explicit retry', () => {
  const api = loadImportJobCore();
  let job = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 4)
  })), '2026-07-10T10:00:00.000Z');
  job = api.markItemProcessing(job, 'item-1');
  job = api.markItemSucceeded(job, 'item-1', '2026-07-10T10:01:00.000Z');
  job = api.markItemProcessing(job, 'item-2');
  job = api.requestCancel(job, '2026-07-10T10:02:00.000Z');
  job = api.markItemFailed(job, 'item-2', 'cancelled failure', '2026-07-10T10:03:00.000Z');

  const resumed = api.resumeCancelledJob(job);
  assert.equal(resumed.status, 'ready');
  assert.equal(resumed.cancelRequested, false);
  assert.equal(resumed.cancelRequestedAt, null);
  assert.equal(resumed.finishedAt, null);
  assert.deepEqual(plain(resumed.items.map((item) => ({
    id: item.id,
    status: item.uploadStatus,
    error: item.error,
    attempts: item.attempts
  }))), [
    { id: 'item-1', status: 'succeeded', error: null, attempts: 1 },
    { id: 'item-2', status: 'failed', error: 'cancelled failure', attempts: 1 },
    { id: 'item-3', status: 'queued', error: null, attempts: 0 },
    { id: 'item-4', status: 'queued', error: null, attempts: 0 }
  ]);

  const retryRunning = api.startJob(resumed, '2026-07-10T10:04:00.000Z');
  assert.equal(api.getNextItem(retryRunning).id, 'item-3');
  assert.throws(
    () => api.markItemProcessing(retryRunning, 'item-1'),
    (error) => error.code === 'INVALID_IMPORT_ITEM_TRANSITION'
  );
  const retryProcessing = api.markItemProcessing(retryRunning, 'item-3');
  assert.equal(retryProcessing.items[1].attempts, 1);
  assert.equal(retryProcessing.items[2].attempts, 1);
  assert.equal(retryProcessing.items[0].uploadStatus, 'succeeded');
});

test('cancelled jobs with only retryable failures require explicit retry instead of resume', () => {
  const api = loadImportJobCore();
  let job = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 1)
  })), '2026-07-10T10:00:00.000Z');
  job = api.markItemProcessing(job, 'item-1');
  job = api.requestCancel(job, '2026-07-10T10:01:00.000Z');
  const failure = new Error('response lost');
  failure.code = 'IMPORT_ITEM_SAVE_FAILED';
  failure.retryable = true;
  job = api.markItemFailed(job, 'item-1', failure, '2026-07-10T10:02:00.000Z');

  assert.throws(
    () => api.resumeCancelledJob(job),
    (error) => error.code === 'IMPORT_JOB_NO_RESUMABLE_ITEMS'
  );
  const retry = api.retryFailedItems(job);
  assert.equal(retry.status, 'ready');
  assert.equal(retry.items[0].uploadStatus, 'queued');
  assert.equal(retry.items[0].attempts, 1);
});

test('cancelled jobs with only succeeded items have no resumable work', () => {
  const api = loadImportJobCore();
  let job = api.startJob(api.readyJob(api.createJob({
    id: 'job-1', sourceType: 'photo', items: createItems(api, 1)
  })), '2026-07-10T10:00:00.000Z');
  job = api.markItemProcessing(job, 'item-1');
  job = api.requestCancel(job, '2026-07-10T10:01:00.000Z');
  job = api.markItemSucceeded(job, 'item-1', '2026-07-10T10:02:00.000Z');

  assert.equal(job.finishedAt, '2026-07-10T10:02:00.000Z');
  assert.throws(
    () => api.resumeCancelledJob(job),
    (error) => error.code === 'IMPORT_JOB_NO_RESUMABLE_ITEMS'
  );
});

test('queue mutations reject missing items and invalid source states', () => {
  const api = loadImportJobCore();
  const idle = api.createJob({ id: 'job-1', sourceType: 'photo', items: createItems(api, 1) });
  const running = api.startJob(api.readyJob(idle), '2026-07-10T10:00:00.000Z');

  assert.throws(
    () => api.markItemSucceeded(running, 'item-1', '2026-07-10T10:01:00.000Z'),
    (error) => error.code === 'INVALID_IMPORT_ITEM_TRANSITION'
  );
  assert.throws(
    () => api.markItemProcessing(running, 'missing'),
    (error) => error.code === 'IMPORT_ITEM_NOT_FOUND'
  );
  assert.throws(
    () => api.readyJob(running),
    (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
  );
});

test('item resource release revokes its Object URL and clears runtime data', () => {
  const api = loadImportJobCore();
  const originalFile = new Blob(['original'], { type: 'image/jpeg' });
  const uploadFile = new Blob(['upload'], { type: 'image/jpeg' });
  const item = api.createItem({
    id: 'item-1',
    sourceType: 'photo',
    error: 'temporary failure',
    errorCode: 'IMPORT_ITEM_SAVE_FAILED',
    retryable: true,
    runtime: {
      originalFile,
      uploadFile,
      previewUrl: 'blob:preview-1'
    }
  });
  const revoked = [];

  const released = api.releaseItemResources(item, {
    revokeObjectURL(url) {
      revoked.push(url);
    }
  });

  assert.equal(released, item);
  assert.deepEqual(revoked, ['blob:preview-1']);
  assert.equal(item.runtime.previewUrl, '');
  assert.equal(item.runtime.originalFile, null);
  assert.equal(item.runtime.uploadFile, null);
  assert.equal(item.error, null);
  assert.equal(item.errorCode, null);
  assert.equal(item.retryable, null);
});

test('item resource release is safe and idempotent on the same item', () => {
  const api = loadImportJobCore();
  const item = api.createItem({
    id: 'item-1',
    sourceType: 'photo',
    runtime: {
      originalFile: new Blob(['original']),
      uploadFile: new Blob(['upload']),
      previewUrl: 'blob:preview-1'
    }
  });
  let revokeCount = 0;
  const urlApi = {
    revokeObjectURL() {
      revokeCount += 1;
      throw new Error('browser cleanup failure');
    }
  };

  assert.doesNotThrow(() => api.releaseItemResources(item, urlApi));
  assert.doesNotThrow(() => api.releaseItemResources(item, urlApi));
  assert.equal(revokeCount, 1);
  assert.doesNotThrow(() => api.releaseItemResources(null, null));
});

test('job resource release clears every item and remains repeat-safe', () => {
  const api = loadImportJobCore();
  const job = api.createJob({
    id: 'job-1',
    sourceType: 'photo',
    items: [
      api.createItem({
        id: 'item-1', sourceType: 'photo', error: 'first error',
        runtime: {
          originalFile: new Blob(['one']),
          uploadFile: new Blob(['one-upload']),
          previewUrl: 'blob:preview-1'
        }
      }),
      api.createItem({
        id: 'item-2', sourceType: 'photo', error: 'second error',
        runtime: {
          originalFile: new Blob(['two']),
          uploadFile: new Blob(['two-upload']),
          previewUrl: 'blob:preview-2'
        }
      })
    ]
  });
  const revoked = [];
  const urlApi = { revokeObjectURL: (url) => revoked.push(url) };

  assert.equal(api.releaseJobResources(job, urlApi), job);
  assert.equal(api.releaseJobResources(job, urlApi), job);
  assert.deepEqual(revoked, ['blob:preview-1', 'blob:preview-2']);
  job.items.forEach((item) => {
    assert.deepEqual(plain(item.runtime), {
      originalFile: null,
      uploadFile: null,
      previewUrl: ''
    });
    assert.equal(item.error, null);
  });
});

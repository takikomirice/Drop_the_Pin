const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function loadProcessor() {
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
      + 'globalThis.__processor = typeof ImportPinItemProcessor === "undefined" ? null : ImportPinItemProcessor;',
    context
  );
  return context.__processor;
}

function item(overrides = {}) {
  return {
    id: 'item-1', sourceType: 'csv', sourceRef: 'CSV 2行目 / source-99', sourceId: 'source-99',
    title: 'CSVピン', description: '説明', lat: 35.5, lng: 139.5,
    capturedAt: '2026-07-11T10:30', color: '#e53935', icon: 'default', status: '',
    tags: ['観察'], links: ['https://example.com'], uploadStatus: 'processing',
    runtime: { originalFile: { secret: true }, uploadFile: { secret: true }, previewUrl: 'blob:secret' },
    ...overrides
  };
}

function harness(overrides = {}) {
  const api = loadProcessor();
  assert.ok(api, 'Expected ImportPinItemProcessor');
  const audit = { gas: [], tokens: [], saved: [] };
  const processor = api.create({
    callGAS(method, payload) {
      audit.gas.push([method, payload]);
      return Promise.resolve({ ok: true, deduplicated: false, pin: { id: 'pin-1' } });
    },
    withEditToken(payload) {
      audit.tokens.push(payload);
      return { ...payload, __editToken: 'secret-token' };
    },
    onSaved(pin, context) { audit.saved.push([pin, context]); },
    ...overrides
  });
  return { processor, audit };
}

test('pin processor sends only the photo-less payload whitelist to saveImportPinItem', async () => {
  const { processor, audit } = harness();
  const result = await processor.processItem(item(), {
    jobId: ' csv-job ', itemId: ' csv-item ', attempt: 3, runtime: { secret: true }
  });
  assert.deepEqual(result, { id: 'pin-1' });
  assert.equal(audit.gas.length, 1);
  assert.equal(audit.gas[0][0], 'saveImportPinItem');
  const payload = audit.gas[0][1];
  assert.deepEqual(Object.keys(payload).sort(), [
    '__editToken', 'color', 'description', 'eventAt', 'icon', 'idempotencyKey',
    'itemId', 'jobId', 'lat', 'links', 'lng', 'status', 'tags', 'title'
  ].sort());
  assert.equal(payload.idempotencyKey, 'csv-job:csv-item');
  assert.equal(payload.eventAt, '2026-07-11T10:30');
  const serialized = JSON.stringify(payload);
  ['source-99', 'sourceRef', 'sourceId', 'runtime', 'blob:', 'attempt', 'base64', 'targetFolderId'].forEach((value) => {
    assert.equal(serialized.includes(value), false, value);
  });
});

test('pin processor keeps the idempotency key stable across attempts', async () => {
  const keys = [];
  const { processor } = harness({
    callGAS(_method, payload) {
      keys.push(payload.idempotencyKey);
      return Promise.resolve({ ok: true, pin: { id: 'pin' } });
    }
  });
  await processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 });
  await processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 9 });
  assert.deepEqual(keys, ['job:item', 'job:item']);
});

test('pin processor converts server errors and accepts deduplicated success with bounded onSaved context', async () => {
  const failed = harness({
    callGAS: () => Promise.resolve({
      ok: false, error: '保存中です。', errorCode: 'IMPORT_ITEM_IN_PROGRESS', retryable: true
    })
  });
  await assert.rejects(
    failed.processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
    (error) => error.code === 'IMPORT_ITEM_IN_PROGRESS' && error.retryable === true
  );

  const saved = [];
  const successful = harness({
    callGAS: () => Promise.resolve({ ok: true, deduplicated: true, pin: { id: 'same-pin' }, secret: 'raw' }),
    onSaved(pin, context) { saved.push([pin, context]); throw new Error('observer'); }
  });
  assert.deepEqual(
    await successful.processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 4 }),
    { id: 'same-pin' }
  );
  assert.deepEqual(JSON.parse(JSON.stringify(saved)), [[{ id: 'same-pin' }, {
    jobId: 'job', itemId: 'item', attempt: 4, deduplicated: true
  }]]);
});

test('pin processor sanitizes transport exceptions without exposing server internals', async () => {
  const { processor } = harness({
    callGAS: () => Promise.reject(new Error('secret server stack and payload'))
  });
  await assert.rejects(
    processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
    (error) => error.code === 'IMPORT_ITEM_SAVE_FAILED'
      && error.retryable === true
      && error.message === 'ピンを保存できませんでした。再試行してください。'
      && !error.message.includes('secret')
  );
});

test('pin processor rejects successful responses without a usable pin id as retryable failures', async () => {
  for (const pin of [null, {}, { id: '' }, { id: '   ' }, []]) {
    const { processor, audit } = harness({
      callGAS: () => Promise.resolve({ ok: true, pin })
    });
    await assert.rejects(
      processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
      (error) => error.code === 'IMPORT_ITEM_SAVE_FAILED'
        && error.retryable === true
        && !error.message.includes('undefined')
    );
    assert.equal(audit.saved.length, 0);
  }
});

test('pin processor rejects missing context without requiring runtime uploadFile or browser APIs', async () => {
  const { processor, audit } = harness();
  assert.deepEqual(
    await processor.processItem(item({ runtime: null }), { jobId: 'job', itemId: 'item', attempt: 1 }),
    { id: 'pin-1' }
  );
  await assert.rejects(
    processor.processItem(item(), { jobId: '', itemId: 'item' }),
    (error) => error.code === 'INVALID_IMPORT_PROCESS_CONTEXT' && error.retryable === false
  );
  assert.equal(audit.gas.length, 1);
});

test('pin processor source is independent from photo, Drive, DOM, queue, and direct GAS globals', () => {
  const start = indexHtml.indexOf('const ImportPinItemProcessor = (function() {');
  const end = indexHtml.indexOf('    const MultiPhotoImportWorkflow = (function()', start);
  assert.notEqual(start, -1);
  const source = indexHtml.slice(start, end);
  assert.doesNotMatch(source, /resizeWithOrientation|uploadFile|targetFolderId|base64|File\b|Blob\b|document\.|google\.script|ImportQueueRunner|ImportFlowController|ImportPreviewUI/);
});

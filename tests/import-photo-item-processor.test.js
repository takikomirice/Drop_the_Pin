const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function extractImportModules() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  return indexHtml.slice(start, end);
}

function loadProcessor() {
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
    `${extractImportModules()}\n`
      + 'globalThis.__processor = typeof ImportPhotoItemProcessor === "undefined" ? null : ImportPhotoItemProcessor;',
    context
  );
  return context.__processor;
}

function item(overrides = {}) {
  const uploadFile = { name: 'source.PNG', marker: 'upload-file' };
  return {
    id: 'item-1',
    sourceRef: 'source.PNG',
    title: '写真',
    description: '説明',
    lat: 35.5,
    lng: 139.5,
    capturedAt: '2026-07-11T10:30',
    color: '#e53935',
    icon: 'photo',
    status: '',
    tags: ['観察'],
    links: ['https://example.com'],
    uploadStatus: 'processing',
    runtime: {
      originalFile: { name: 'original.heic' },
      uploadFile,
      previewUrl: 'blob:secret-preview'
    },
    ...overrides
  };
}

function createHarness(overrides = {}) {
  const processorApi = loadProcessor();
  assert.ok(processorApi, 'Expected ImportPhotoItemProcessor');
  const audit = { resize: [], gas: [], token: [], saved: [] };
  const processor = processorApi.create({
    resizePhoto(file, maxSize) {
      audit.resize.push([file, maxSize]);
      return Promise.resolve('data:image/jpeg;base64,YWJj');
    },
    callGAS(method, payload) {
      audit.gas.push([method, payload]);
      return Promise.resolve({ ok: true, deduplicated: false, pin: { id: 'pin-1' } });
    },
    withEditToken(payload) {
      audit.token.push(payload);
      return { ...payload, __editToken: 'secret-token' };
    },
    getTargetFolderId() { return 'folder-1'; },
    onSaved(pin, context) { audit.saved.push([pin, context]); },
    ...overrides
  });
  return { processor, audit };
}

test('processor is lazy and sends only the whitelisted JPEG payload at process time', async () => {
  const source = item();
  const { processor, audit } = createHarness();
  assert.equal(audit.resize.length, 0);

  const pin = await processor.processItem(source, { jobId: ' job-1 ', itemId: ' item-1 ', attempt: 2 });

  assert.deepEqual(pin, { id: 'pin-1' });
  assert.deepEqual(audit.resize, [[source.runtime.uploadFile, 1920]]);
  assert.equal(audit.gas.length, 1);
  assert.equal(audit.gas[0][0], 'saveImportPhotoItem');
  const payload = audit.gas[0][1];
  assert.deepEqual(Object.keys(payload).sort(), [
    '__editToken', 'base64', 'color', 'description', 'eventAt', 'filename',
    'icon', 'idempotencyKey', 'itemId', 'jobId', 'lat', 'links', 'lng',
    'status', 'tags', 'targetFolderId', 'title'
  ].sort());
  assert.equal(payload.jobId, 'job-1');
  assert.equal(payload.itemId, 'item-1');
  assert.equal(payload.idempotencyKey, 'job-1:item-1');
  assert.equal(payload.filename, 'source.jpg');
  assert.equal(payload.eventAt, source.capturedAt);
  assert.equal(payload.icon, source.icon);
  assert.equal(payload.status, '');
  assert.equal(payload.base64, 'data:image/jpeg;base64,YWJj');
  assert.equal(JSON.stringify(payload).includes('blob:secret-preview'), false);
  assert.equal(JSON.stringify(payload).includes('marker'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(payload, 'runtime'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(payload, 'attempt'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(payload, 'sourceDriveFileId'), false);
});

test('Drive items send exactly one bounded source id while keeping other source metadata out', async () => {
  const source = item({ runtime: {
    originalFile: { name: 'original.heic' },
    uploadFile: { name: 'source.PNG', marker: 'upload-file' },
    previewUrl: 'blob:secret-preview',
    sourceDriveFileId: 'photo_AAAAAAAAAAA',
    sourceFolderId: 'folder-secret',
    owner: 'owner-secret'
  } });
  const { processor, audit } = createHarness();
  await processor.processItem(source, { jobId: 'job-1', itemId: 'item-1', attempt: 1 });
  const payload = audit.gas[0][1];
  assert.equal(payload.sourceDriveFileId, 'photo_AAAAAAAAAAA');
  assert.equal(payload.targetFolderId, 'folder-1');
  assert.equal(Object.hasOwn(payload, 'sourceFolderId'), false);
  assert.equal(Object.hasOwn(payload, 'owner'), false);
});

test('attach operation reuses the photo processor and sends only target identity and expected version', async () => {
  const { processor, audit } = createHarness({
    getPhotoOperation() {
      return {
        operationMode: 'attach-existing-pin',
        targetPinId: 'pin-existing-0001',
        expectedUpdatedAt: '2026-07-11T11:59:00.000Z'
      };
    }
  });

  await processor.processItem(item(), { jobId: 'attach-job', itemId: 'attach-item', attempt: 1 });

  const payload = audit.gas[0][1];
  assert.equal(payload.operationMode, 'attach-existing-pin');
  assert.equal(payload.targetPinId, 'pin-existing-0001');
  assert.equal(payload.expectedUpdatedAt, '2026-07-11T11:59:00.000Z');
  assert.equal(Object.hasOwn(payload, 'targetPin'), false);
  assert.equal(Object.hasOwn(payload, 'operation'), false);
});

test('attach operation errors use operation-specific safe client messages', async () => {
  const { processor } = createHarness({
    getPhotoOperation: () => ({
      operationMode: 'attach-existing-pin',
      targetPinId: 'pin-existing-0001',
      expectedUpdatedAt: ''
    }),
    callGAS: () => Promise.resolve({
      ok: false,
      errorCode: 'PIN_PHOTO_ATTACH_CONFLICT',
      error: 'private target data',
      retryable: false
    })
  });

  await assert.rejects(
    processor.processItem(item(), { jobId: 'attach-job', itemId: 'attach-item', attempt: 1 }),
    (error) => error.code === 'PIN_PHOTO_ATTACH_CONFLICT'
      && error.message === 'ピンが別の操作で更新されました。画面を再読み込みしてから再試行してください。'
      && error.retryable === false
  );
});

test('attempt changes do not change the idempotency key and ids do', async () => {
  const keys = [];
  const { processor } = createHarness({
    callGAS(_method, payload) {
      keys.push(payload.idempotencyKey);
      return Promise.resolve({ ok: true, pin: { id: 'pin' } });
    }
  });
  await processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 });
  await processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 9 });
  await processor.processItem(item(), { jobId: 'job', itemId: 'other', attempt: 1 });
  assert.deepEqual(keys, ['job:item', 'job:item', 'job:other']);
});

test('JPEG output names replace the final source extension exactly once', async () => {
  const names = [];
  const { processor } = createHarness({
    callGAS(_method, payload) {
      names.push(payload.filename);
      return Promise.resolve({ ok: true, pin: { id: `pin-${names.length}` } });
    }
  });
  for (const name of ['a.jpg', 'b.JPEG', 'c.png', 'd.webp', 'e.heic', '.heif', '']) {
    const source = item({
      sourceRef: name,
      runtime: { uploadFile: { name }, originalFile: null, previewUrl: '' }
    });
    await processor.processItem(source, { jobId: 'job', itemId: `item-${names.length}`, attempt: 1 });
  }
  assert.deepEqual(names, ['a.jpg', 'b.jpg', 'c.jpg', 'd.jpg', 'e.jpg', 'image.jpg', 'image.jpg']);
});

test('missing uploadFile rejects with a coded error before resize or GAS', async () => {
  const { processor, audit } = createHarness();
  await assert.rejects(
    processor.processItem(item({ runtime: { uploadFile: null } }), { jobId: 'job', itemId: 'item', attempt: 1 }),
    (error) => error.code === 'IMPORT_UPLOAD_FILE_REQUIRED'
  );
  assert.equal(audit.resize.length, 0);
  assert.equal(audit.gas.length, 0);
});

test('server failure rejects with errorCode and retryable metadata', async () => {
  const { processor } = createHarness({
    callGAS: () => Promise.resolve({
      ok: false, error: '保存処理中です。', errorCode: 'IMPORT_ITEM_IN_PROGRESS', retryable: true
    })
  });
  await assert.rejects(
    processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
    (error) => error.message === '保存処理中です。'
      && error.code === 'IMPORT_ITEM_IN_PROGRESS'
      && error.retryable === true
  );
});

test('transport rejection becomes a safe retryable response-loss failure', async () => {
  const { processor } = createHarness({
    callGAS: () => Promise.reject(new Error('private Apps Script stack and source metadata'))
  });
  await assert.rejects(
    processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
    (error) => error.code === 'IMPORT_ITEM_SAVE_FAILED'
      && error.message === '写真を保存できませんでした。再試行してください。'
      && error.retryable === true
      && !/private|stack|metadata/.test(error.message)
  );
});

test('malformed successful save responses stay retryable and never notify observers', async () => {
  for (const pin of [null, [], {}, { id: '' }, { id: 1 }, { id: {} }]) {
    let saved = 0;
    const { processor } = createHarness({
      callGAS: () => Promise.resolve({ ok: true, pin }),
      onSaved() { saved += 1; }
    });
    await assert.rejects(
      processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
      (error) => error.code === 'IMPORT_ITEM_SAVE_FAILED'
        && error.message === '写真を保存できませんでした。再試行してください。'
        && error.retryable === true
    );
    assert.equal(saved, 0);
  }
});

test('hostile successful response getters cannot expose provider details', async () => {
  const hostilePin = {};
  Object.defineProperty(hostilePin, 'id', {
    enumerable: true,
    get() { throw new Error('private pin getter and source id'); }
  });
  const hostileResponse = { ok: true };
  Object.defineProperty(hostileResponse, 'pin', {
    enumerable: true,
    get() { throw new Error('private response getter and stack'); }
  });
  const hostileOk = {};
  Object.defineProperty(hostileOk, 'ok', {
    enumerable: true,
    get() { throw new Error('private ok getter and stack'); }
  });
  const hostileErrorCode = { ok: false };
  Object.defineProperty(hostileErrorCode, 'errorCode', {
    enumerable: true,
    get() { throw new Error('private error code getter and stack'); }
  });
  for (const result of [
    { ok: true, pin: hostilePin }, hostileResponse, hostileOk, hostileErrorCode
  ]) {
    const { processor } = createHarness({ callGAS: () => Promise.resolve(result) });
    await assert.rejects(
      processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
      (error) => error.code === 'IMPORT_ITEM_SAVE_FAILED'
        && error.message === '写真を保存できませんでした。再試行してください。'
        && error.retryable === true
        && !/private|source id|stack/.test(error.message)
    );
  }
});

test('missing receipt setup uses safe actionable guidance without exposing sheet details', async () => {
  const { processor } = createHarness({
    callGAS: () => Promise.resolve({
      ok: false,
      error: 'import_receipts internal row details',
      errorCode: 'IMPORT_RECEIPT_SHEET_MISSING',
      retryable: false
    })
  });
  await assert.rejects(
    processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
    (error) => error.code === 'IMPORT_RECEIPT_SHEET_MISSING'
      && error.message === '初期設定が必要です。スプレッドシートの設定メニューから初期設定を実行してください。'
      && error.message.includes('internal') === false
  );
});

test('Drive management and map row failures use distinct safe actionable messages', async () => {
  const cases = [
    ['IMPORT_DRIVE_FILE_FAILED', '管理対象の写真を保存できませんでした。再試行してください。', true],
    ['DRIVE_SOURCE_NOT_EDITABLE', '選択したDrive写真を表示用ファイルとして利用できません。写真を選び直してください。', false],
    ['DRIVE_SOURCE_CHECK_FAILED', '選択したDrive写真を確認できませんでした。再試行してください。', true],
    ['DRIVE_LINK_SHARING_DENIED', '組織のGoogle Drive共有ポリシーにより写真を公開できません。公開可能な保存先を設定するか、Google Workspace管理者へリンク共有設定を確認してください。', false],
    ['DRIVE_LINK_SHARING_FAILED', '管理用写真のリンク共有を確認できませんでした。再試行してください。', true],
    ['DRIVE_MANAGED_COPY_CREATE_FAILED', '管理用の写真コピーを作成できませんでした。保存先Driveの作成権限を確認してください。', true],
    ['DRIVE_MANAGED_COPY_FINALIZE_FAILED', '管理用の写真コピーを確定できませんでした。再試行してください。', true],
    ['DRIVE_ORIGINAL_FOLDER_CREATE_FAILED', 'originalフォルダを作成できませんでした。Driveの作成権限を確認してください。', true],
    ['DRIVE_ORIGINAL_FOLDER_AMBIGUOUS', 'originalフォルダが複数あります。1つに整理してから再試行してください。', false],
    ['DRIVE_SOURCE_MOVE_FAILED', '元写真をoriginalフォルダへ移動できませんでした。ピンは登録していません。Driveの移動権限を確認してください。', true],
    ['DRIVE_SOURCE_MOVE_VERIFY_FAILED', '元写真の移動結果を確認できませんでした。ピンは登録していません。Driveを確認して再試行してください。', true],
    ['IMPORT_MAP_ROW_FAILED_AFTER_SOURCE_MOVE', '元写真の整理は完了しましたが、ピン情報を保存できませんでした。再試行してください。', true],
    ['IMPORT_MAP_ROW_FAILED', 'ピン情報を登録できませんでした。再試行してください。', true]
  ];
  for (const [errorCode, message, retryable] of cases) {
    const { processor } = createHarness({
      callGAS: () => Promise.resolve({
        ok: false, error: 'private Drive id and sheet row', errorCode, retryable
      })
    });
    await assert.rejects(
      processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
      (error) => error.code === errorCode
        && error.message === message
        && error.retryable === retryable
        && !/private|Drive id|sheet row/.test(error.message)
    );
  }
});

test('already-linked Drive sources are safe non-retryable failures', async () => {
  const { processor } = createHarness({
    callGAS: () => Promise.resolve({
      ok: false,
      error: 'private Drive source id',
      errorCode: 'DRIVE_SOURCE_ALREADY_LINKED',
      retryable: false
    })
  });
  await assert.rejects(
    processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 1 }),
    (error) => error.code === 'DRIVE_SOURCE_ALREADY_LINKED'
      && error.message === 'このDrive写真は既に別のピンへ紐づいています。'
      && error.retryable === false
      && !error.message.includes('source id')
  );
});

test('deduplicated responses succeed and onSaved receives only the bounded context', async () => {
  const calls = [];
  const { processor } = createHarness({
    callGAS: () => Promise.resolve({ ok: true, deduplicated: true, pin: { id: 'same-pin' }, secret: 'raw' }),
    onSaved(pin, context) { calls.push([pin, context]); }
  });
  const pin = await processor.processItem(item(), { jobId: 'job', itemId: 'item', attempt: 3 });
  assert.deepEqual(pin, { id: 'same-pin' });
  assert.deepEqual(JSON.parse(JSON.stringify(calls)), [[{ id: 'same-pin' }, {
    jobId: 'job', itemId: 'item', attempt: 3, deduplicated: true
  }]]);
  assert.equal(JSON.stringify(calls).includes('secret-token'), false);
  assert.equal(JSON.stringify(calls).includes('data:image'), false);
});

test('cleanup warning still succeeds and immediately notifies onSaved with the saved pin', async () => {
  const calls = [];
  const savedPin = { id: 'pin-warning', lat: null, lng: null };
  const { processor } = createHarness({
    callGAS: () => Promise.resolve({
      ok: true,
      deduplicated: false,
      warningCode: 'DRIVE_SOURCE_CLEANUP_PENDING',
      pin: savedPin
    }),
    onSaved(pin) { calls.push(pin); }
  });

  const result = await processor.processItem(item({ lat: null, lng: null }), {
    jobId: 'job', itemId: 'warning', attempt: 1
  });
  assert.deepEqual(result, savedPin);
  assert.deepEqual(calls, [savedPin]);
});

test('onSaved synchronous and async failures never turn a server success into failure', async () => {
  const first = createHarness({ onSaved() { throw new Error('observer'); } });
  assert.deepEqual(
    await first.processor.processItem(item(), { jobId: 'job', itemId: 'one', attempt: 1 }),
    { id: 'pin-1' }
  );
  const second = createHarness({ onSaved() { return Promise.reject(new Error('observer')); } });
  assert.deepEqual(
    await second.processor.processItem(item(), { jobId: 'job', itemId: 'two', attempt: 1 }),
    { id: 'pin-1' }
  );
  await new Promise((resolve) => setImmediate(resolve));
});

test('processor source stays independent from queue, flow, preview, DOM, and direct GAS globals', () => {
  const start = indexHtml.indexOf('const ImportPhotoItemProcessor = (function() {');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1);
  const source = indexHtml.slice(start, end);
  assert.doesNotMatch(source, /ImportQueueRunner|ImportFlowController|ImportPreviewUI|document\.|google\.script/);
});

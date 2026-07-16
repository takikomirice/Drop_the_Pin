const assert = require('node:assert/strict');
const test = require('node:test');
const {
  loadDriveClientModules, descriptor, fileResponse, imageFixtures
} = require('./drive-photo-import-client-harness');

class BrowserFile extends Blob {
  constructor(parts, name, options = {}) {
    super(parts, options);
    this.name = name;
    this.lastModified = options.lastModified;
  }
}

function environment(overrides = {}) {
  return {
    atob(value) { return Buffer.from(value, 'base64').toString('binary'); },
    Uint8Array,
    Blob,
    File: BrowserFile,
    ...overrides
  };
}

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function settle() {
  return new Promise((resolve) => setImmediate(resolve));
}

function values(count) {
  return Array.from({ length: count }, (_, index) => descriptor({
    id: `photo_${String(index).padStart(12, '0')}`,
    name: `${index}.jpg`
  }));
}

test('loader supports one and multiple files sequentially in selection order with send-time tokens', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  assert.ok(loaderApi);
  const calls = [];
  const progress = [];
  let active = 0;
  let maxActive = 0;
  let tokenNumber = 0;
  const instance = loaderApi.create({
    callGAS: async (method, payload) => {
      calls.push([method, { ...payload }]);
      active += 1;
      maxActive = Math.max(maxActive, active);
      await settle();
      active -= 1;
      const expected = values(3).find((item) => item.id === payload.fileId);
      return fileResponse(expected);
    },
    withEditToken(payload) {
      tokenNumber += 1;
      return { ...payload, __editToken: `token-${tokenNumber}` };
    },
    sourceCore,
    environment: environment(),
    onProgress(snapshot) {
      assert.equal(Object.isFrozen(snapshot), true);
      progress.push(snapshot);
    }
  });

  const one = await instance.start(values(1));
  assert.equal(one.length, 1);
  const files = await instance.start(values(3));
  assert.deepEqual(Array.from(files, (file) => file.name), ['0.jpg', '1.jpg', '2.jpg']);
  assert.equal(maxActive, 1);
  assert.deepEqual(calls.map((call) => call[0]), Array(4).fill('readDrivePhotoImportFile'));
  assert.deepEqual(calls.map((call) => Object.keys(call[1]).sort()),
    Array(4).fill(['__editToken', 'fileId']));
  assert.deepEqual(calls.map((call) => call[1].__editToken),
    ['token-1', 'token-2', 'token-3', 'token-4']);
  assert.equal(instance.isRunning(), false);
  progress.forEach((snapshot) => {
    assert.deepEqual(Object.keys(snapshot).sort(),
      ['completed', 'currentIndex', 'currentName', 'eventType', 'total']);
    assert.doesNotMatch(JSON.stringify(snapshot), /photo_\d|token-|base64|AQID|__editToken|fileId/);
  });
});

test('loader materializes Drive JPG PNG WebP HEIC and HEIF as BrowserFiles', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  const samples = [
    descriptor({ id: 'photo_JPGAAAAAAAA', name: 'photo.jpg', mimeType: 'image/jpeg', kind: 'jpeg', sizeBytes: imageFixtures.jpeg.length }),
    descriptor({ id: 'photo_PNGAAAAAAAA', name: 'photo.png', mimeType: 'image/png', kind: 'png', sizeBytes: imageFixtures.png.length }),
    descriptor({ id: 'photo_WEBPAAAAAAA', name: 'photo.webp', mimeType: 'image/webp', kind: 'webp', sizeBytes: imageFixtures.webp.length }),
    descriptor({ id: 'photo_HEICAAAAAAAA', name: 'photo.heic', mimeType: 'image/heic', kind: 'heic', sizeBytes: imageFixtures.heic.length }),
    descriptor({ id: 'photo_HEIFAAAAAAAA', name: 'photo.heif', mimeType: 'image/heif', kind: 'heic', sizeBytes: imageFixtures.heif.length })
  ];
  const fixtures = [
    imageFixtures.jpeg, imageFixtures.png, imageFixtures.webp, imageFixtures.heic, imageFixtures.heif
  ];
  const instance = loaderApi.create({
    callGAS: async (_method, payload) => {
      const index = samples.findIndex((sample) => sample.id === payload.fileId);
      return fileResponse(samples[index], fixtures[index].toString('base64'));
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'send-only-token' }),
    sourceCore,
    environment: environment()
  });

  const files = await instance.start(samples);

  assert.equal(files.every((file) => file instanceof BrowserFile), true);
  assert.deepEqual(Array.from(files, (file) => [file.name, file.type]), [
    ['photo.jpg', 'image/jpeg'],
    ['photo.png', 'image/png'],
    ['photo.webp', 'image/webp'],
    ['photo.heic', 'image/heic'],
    ['photo.heif', 'image/heif']
  ]);
  const materialized = await Promise.all(files.map(async (file) => Buffer.from(await file.arrayBuffer())));
  assert.deepEqual(materialized, fixtures);
});

test('loader materializes actual read bytes when listing metadata changes without retaining secrets', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  const listed = descriptor({
    name: 'listed.heic', mimeType: 'image/heic', sizeBytes: 3, kind: 'heic',
    modifiedAt: '2026-07-12T01:02:03.000Z'
  });
  const bytes = Buffer.alloc(4097, 0x4b);
  const events = [];
  const instance = loaderApi.create({
    callGAS: async (_method, payload) => fileResponse(descriptor({
      id: payload.fileId,
      name: 'current.jpeg',
      mimeType: 'IMAGE/JPG',
      sizeBytes: 1,
      modifiedAt: '2026-07-15T10:20:30.000Z',
      kind: 'ignored'
    }), bytes.toString('base64')),
    withEditToken: (payload) => ({ ...payload, __editToken: 'send-only-token' }),
    sourceCore,
    environment: environment(),
    onProgress: (snapshot) => events.push(snapshot)
  });

  const [file] = await instance.start([listed]);
  assert.equal(file.name, 'current.jpeg');
  assert.equal(file.type, 'image/jpeg');
  assert.equal(file.size, bytes.length);
  assert.equal(file.lastModified, Date.parse('2026-07-15T10:20:30.000Z'));
  assert.equal(instance.isRunning(), false);
  assert.doesNotMatch(JSON.stringify(events), /photo_A|base64|S0tL|send-only-token|__editToken|fileId/);
  assert.doesNotMatch(JSON.stringify(instance), /photo_A|base64|S0tL|send-only-token|__editToken|fileId/);
});

test('loader rejects start while active and fails atomically without starting later requests', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  const firstResponse = deferred();
  const inputs = values(3);
  let calls = 0;
  const instance = loaderApi.create({
    callGAS(method, payload) {
      calls += 1;
      if (calls === 1) return firstResponse.promise;
      if (calls === 2) return Promise.resolve(fileResponse(inputs[1], 'invalid***'));
      return Promise.resolve(fileResponse(inputs[2]));
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'secret' }),
    sourceCore,
    environment: environment()
  });

  const running = instance.start(inputs);
  await settle();
  await assert.rejects(instance.start(values(1)),
    (error) => error.code === 'DRIVE_IMPORT_ALREADY_RUNNING');
  firstResponse.resolve(fileResponse(inputs[0]));
  await assert.rejects(running,
    (error) => error.code === 'DRIVE_IMPORT_RESPONSE_INVALID' && !/invalid|AQID|photo_/.test(error.message));
  assert.equal(calls, 2);
  assert.equal(instance.isRunning(), false);
});

test('loader preserves only a whitelisted response diagnostic stage', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  const instance = loaderApi.create({
    callGAS: async () => fileResponse(descriptor()),
    withEditToken: (payload) => ({ ...payload, __editToken: 'secret' }),
    sourceCore,
    environment: environment({ atob() { throw new Error('private AQID photo_AAAAAAAAAAA'); } })
  });

  await assert.rejects(instance.start([descriptor()]), (error) => (
    error.code === 'DRIVE_IMPORT_RESPONSE_INVALID'
      && error.diagnosticStage === 'base64_decode_failed'
      && Object.keys(error).sort().join(',') === 'code,diagnosticStage'
      && !/private|AQID|photo_/.test(error.message)
  ));
  assert.equal(JSON.stringify(instance).includes('AQID'), false);
});

test('cancel is idempotent, ignores the active stale response, stops later calls, and permits reuse', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  const pending = deferred();
  const inputs = values(2);
  let calls = 0;
  const events = [];
  const instance = loaderApi.create({
    callGAS() {
      calls += 1;
      return calls === 1 ? pending.promise : Promise.resolve(fileResponse(values(1)[0]));
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'secret' }),
    sourceCore,
    environment: environment(),
    onProgress: (snapshot) => events.push(snapshot.eventType)
  });

  const running = instance.start(inputs);
  await settle();
  instance.cancel();
  instance.cancel();
  pending.resolve(fileResponse(inputs[0]));
  await assert.rejects(running, (error) => error.code === 'DRIVE_IMPORT_CANCELLED');
  assert.equal(calls, 1);
  assert.equal(events.filter((event) => event === 'cancelled').length, 1);
  assert.equal(instance.isRunning(), false);

  const reused = await instance.start(values(1));
  assert.equal(reused.length, 1);
  assert.equal(calls, 2);
});

test('release is idempotent and ignores a delayed response without retaining partial files', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  const pending = deferred();
  const input = values(1);
  const instance = loaderApi.create({
    callGAS: () => pending.promise,
    withEditToken: (payload) => ({ ...payload, __editToken: 'secret' }),
    sourceCore,
    environment: environment()
  });
  const running = instance.start(input);
  await settle();
  instance.release();
  instance.release();
  pending.resolve(fileResponse(input[0]));
  await assert.rejects(running, (error) => error.code === 'DRIVE_IMPORT_CANCELLED');
  assert.equal(instance.isRunning(), false);
  await assert.rejects(instance.start(input), (error) => error.code === 'DRIVE_IMPORT_RELEASED');
});

test('progress observer failures never stop loading and server failures use safe mapped errors', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  let notifications = 0;
  const successful = loaderApi.create({
    callGAS: (method, payload) => Promise.resolve(fileResponse(descriptor({ id: payload.fileId }))),
    withEditToken: (payload) => ({ ...payload, __editToken: 'secret' }),
    sourceCore,
    environment: environment(),
    onProgress() {
      notifications += 1;
      if (notifications % 2) throw new Error('observer private');
      return Promise.reject(new Error('async observer private'));
    }
  });
  assert.equal((await successful.start([descriptor()])).length, 1);

  const events = [];
  const failed = loaderApi.create({
    callGAS: () => Promise.resolve({
      ok: false,
      errorCode: 'DRIVE_IMPORT_FILE_TOO_LARGE',
      error: 'private server detail photo_AAAAAAAAAAA'
    }),
    withEditToken: (payload) => ({ ...payload, __editToken: 'secret' }),
    sourceCore,
    environment: environment(),
    onProgress: (snapshot) => events.push(snapshot)
  });
  await assert.rejects(failed.start([descriptor()]), (error) => (
    error.code === 'DRIVE_IMPORT_FILE_TOO_LARGE'
      && error.message === '1枚の写真は15MB以内にしてください。'
  ));
  assert.equal(events.at(-1).eventType, 'failed');
  assert.deepEqual(Object.keys(events.at(-1)).sort(),
    ['completed', 'currentIndex', 'currentName', 'eventType', 'total']);
  assert.doesNotMatch(JSON.stringify(events), /private|photo_A|error|stack|fileId/);
});

test('loader sanitizes a rejected value whose code getter throws', async () => {
  const { loader: loaderApi, sourceCore } = loadDriveClientModules();
  const hostile = {};
  Object.defineProperty(hostile, 'code', {
    get() { throw new Error('private getter detail photo_AAAAAAAAAAA'); }
  });
  const instance = loaderApi.create({
    callGAS: () => Promise.reject(hostile),
    withEditToken: (payload) => ({ ...payload, __editToken: 'secret' }),
    sourceCore,
    environment: environment()
  });
  await assert.rejects(instance.start([descriptor()]), (error) => (
    error.code === 'DRIVE_IMPORT_FILE_READ_FAILED'
      && !/private|photo_A/.test(error.message)
  ));
});

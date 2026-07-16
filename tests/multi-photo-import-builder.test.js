const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const readme = fs.readFileSync(path.join(root, 'README.md'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function extractImportModules() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1, 'Expected import modules');
  assert.notEqual(end, -1, 'Expected application state after import modules');
  return indexHtml.slice(start, end);
}

function loadDefaultPhotoPreparer(stubs = {}) {
  const start = indexHtml.indexOf('async function prepareMultiPhotoFile(');
  const end = indexHtml.indexOf('    function locationMessageForMetadataStatus(', start);
  assert.notEqual(start, -1, 'Expected default multi-photo preparer');
  assert.notEqual(end, -1, 'Expected metadata message helper after preparer');
  const context = { console, ...stubs };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\nglobalThis.__prepareMulti = prepareMultiPhotoFile;`,
    context
  );
  return context.__prepareMulti;
}

function loadBuilder(extra = {}) {
  const context = {
    console,
    Number,
    Object,
    Array,
    String,
    Error,
    Set,
    Promise,
    Date,
    Math,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935' }, { hex: '#2196f3' }],
    PIN_ICONS: [{ id: 'default' }, { id: 'photo' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    URL: { createObjectURL() { throw new Error('not configured'); }, revokeObjectURL() {} },
    crypto: { randomUUID: () => 'random-id' },
    ...extra
  };
  vm.createContext(context);
  vm.runInContext(
    `${extractImportModules()}\n`
      + 'globalThis.__core = ImportJobCore;\n'
      + 'globalThis.__builder = typeof MultiPhotoImportBuilder === "undefined" ? null : MultiPhotoImportBuilder;\n'
      + 'globalThis.__openMulti = typeof openMultiPhotoImportPreview === "undefined" ? null : openMultiPhotoImportPreview;',
    context
  );
  return { core: context.__core, builder: context.__builder, openMulti: context.__openMulti };
}

function defaults(overrides = {}) {
  return {
    tags: ['旅行'], color: '#e53935', icon: 'default', status: '未対応', ...overrides
  };
}

function photo(name, type = '') {
  return { name, type, lastModified: 123 };
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

function makeSession(builder, overrides = {}) {
  let nextId = 0;
  return builder.create({
    preparePhoto: async (file, details) => ({
      originalFile: file,
      uploadFile: file,
      metadataStatus: 'no-gps',
      conversionStatus: details.kind === 'heic' ? 'succeeded' : 'not-needed',
      lat: null,
      lng: null,
      capturedAt: ''
    }),
    createObjectURL: (file) => `blob:${file.name}`,
    revokeObjectURL() {},
    createId: () => `id-${++nextId}`,
    ...overrides
  });
}

test('multi-photo builder exposes the reusable session API and can be reused after completion', async () => {
  const { builder } = loadBuilder();
  assert.ok(builder);
  const session = makeSession(builder);
  assert.equal(typeof session.start, 'function');
  assert.equal(typeof session.cancel, 'function');
  assert.equal(typeof session.isRunning, 'function');
  assert.equal(typeof session.release, 'function');
  const first = await session.start([photo('first.jpg')], defaults());
  const second = await session.start([photo('second.jpg')], defaults());
  assert.deepEqual(first.items.map((item) => item.sourceRef), ['first.jpg']);
  assert.deepEqual(second.items.map((item) => item.sourceRef), ['second.jpg']);
});

test('Drive source ids stay ordered in runtime only and are released with file resources', async () => {
  const { builder, core } = loadBuilder();
  const session = makeSession(builder);
  const sourceIds = ['photo_AAAAAAAAAAA', 'photo_BBBBBBBBBBB'];
  const job = await session.start(
    [photo('first.jpg'), photo('second.jpg')],
    defaults(),
    { sourceDriveFileIds: sourceIds }
  );

  assert.deepEqual(job.items.map((item) => item.runtime.sourceDriveFileId), sourceIds);
  assert.equal(JSON.stringify(job.items.map(core.toPersistableItem)).includes('photo_AAAAAAAAAAA'), false);

  core.releaseJobResources(job, { revoke() {} });
  assert.equal(job.items.every((item) => !Object.hasOwn(item.runtime, 'sourceDriveFileId')), true);
});

test('local photo jobs never gain Drive source ids', async () => {
  const { builder } = loadBuilder();
  const job = await makeSession(builder).start([photo('local.jpg')], defaults());
  assert.equal(Object.hasOwn(job.items[0].runtime, 'sourceDriveFileId'), false);
});

test('file validation accepts the five families and keeps unsupported entries as failed items', async () => {
  const { builder } = loadBuilder();
  const prepared = [];
  const session = makeSession(builder, {
    preparePhoto: async (file, details) => {
      prepared.push([file.name, details.kind]);
      return {
        originalFile: file, uploadFile: file, metadataStatus: 'no-gps',
        conversionStatus: details.kind === 'heic' ? 'succeeded' : 'not-needed'
      };
    }
  });
  const files = [
    photo('a.jpg', 'image/jpeg'), photo('alias.jpg', 'image/jpg'), photo('b.JPEG'), photo('c.png', 'image/png'),
    photo('d.WEBP', ''), photo('e.heic', 'image/heic'), photo('f.HEIF', ''),
    photo('generic.heic', 'application/octet-stream'),
    photo('fake.jpg', 'image/svg+xml'), photo('fake.svg', 'image/jpeg'),
    photo('anim.gif', 'image/gif'), photo('movie.jpg', 'video/mp4'), photo('unknown.bin', ''), null
  ];

  const job = await session.start(files, defaults());

  assert.deepEqual(prepared, [
    ['a.jpg', 'jpeg'], ['alias.jpg', 'jpeg'], ['b.JPEG', 'jpeg'], ['c.png', 'png'],
    ['d.WEBP', 'webp'], ['e.heic', 'heic'], ['f.HEIF', 'heic'], ['generic.heic', 'heic']
  ]);
  assert.deepEqual(job.items.map((item) => item.uploadStatus), [
    'queued', 'queued', 'queued', 'queued', 'queued', 'queued', 'queued', 'queued',
    'failed', 'failed', 'failed', 'failed', 'failed', 'failed'
  ]);
  assert.equal(job.items[8].conversionStatus, 'unsupported');
  assert.match(job.items[8].error, /対応していない/);
});

test('public classifyFile keeps its legacy three-field shape and ignores an unused size getter', () => {
  const { builder } = loadBuilder();
  const file = photo('photo.jpg', 'image/jpeg');
  Object.defineProperty(file, 'size', {
    get() { throw new Error('unused size getter'); }
  });
  assert.deepEqual(plain(builder.classifyFile(file)), {
    supported: true,
    kind: 'jpeg',
    reason: ''
  });
  assert.deepEqual(plain(builder.classifyFile(photo('photo.svg', 'image/jpeg'))), {
    supported: false,
    kind: '',
    reason: 'unsupported'
  });
});

test('browser File availability rejects lookalike objects with a distinct safe item error', async () => {
  class BrowserFile {
    constructor(name, type) {
      this.name = name;
      this.type = type;
    }
  }
  const { builder } = loadBuilder({ File: BrowserFile });
  const real = new BrowserFile('real.jpg', 'image/jpeg');
  const lookalike = photo('lookalike.jpg', 'image/jpeg');
  const session = makeSession(builder);

  const job = await session.start([real, lookalike], defaults());

  assert.deepEqual(job.items.map((item) => item.uploadStatus), ['queued', 'failed']);
  assert.match(job.items[1].error, /ファイル参照/);
  assert.equal(job.items[1].runtime.originalFile, lookalike);
  assert.equal(job.items[1].runtime.uploadFile, null);
});

test('empty and oversized selections reject before preparation without partial adoption', async () => {
  const { core, builder } = loadBuilder();
  let prepareCalls = 0;
  let urlCalls = 0;
  const session = makeSession(builder, {
    preparePhoto: async () => { prepareCalls += 1; },
    createObjectURL: () => { urlCalls += 1; return 'blob:x'; }
  });

  await assert.rejects(session.start([], defaults()), (error) => error.code === 'MULTI_PHOTO_EMPTY');
  await assert.rejects(
    session.start(Array.from({ length: core.MAX_ITEMS + 1 }, (_, index) => photo(`${index}.jpg`)), defaults()),
    (error) => error.code === 'IMPORT_ITEM_LIMIT_EXCEEDED'
  );
  assert.equal(prepareCalls, 0);
  assert.equal(urlCalls, 0);
  assert.equal(session.isRunning(), false);
});

test('configuration and shared defaults are validated before workers start', async () => {
  const { builder } = loadBuilder();
  for (const concurrency of [0, 3, -1, '2']) {
    assert.throws(
      () => builder.create({ concurrency, preparePhoto() {} }),
      (error) => error.code === 'INVALID_MULTI_PHOTO_CONCURRENCY'
    );
  }
  const session = makeSession(builder);
  const invalidDefaults = [
    defaults({ tags: '旅行' }), defaults({ tags: ['1', '2', '3', '4', '5', '6'] }),
    defaults({ tags: [''] }), defaults({ color: '#ffffff' }), defaults({ color: '<script>' }),
    defaults({ icon: 'unknown' }), defaults({ status: 'unknown' })
  ];
  for (const value of invalidDefaults) {
    await assert.rejects(
      session.start([photo('a.jpg')], value),
      (error) => String(error.code || '').startsWith('INVALID_MULTI_PHOTO_DEFAULT')
    );
  }
  const emptyStatusJob = await session.start([photo('a.jpg')], defaults({ status: '' }));
  assert.equal(emptyStatusJob.items[0].status, '');
});

test('defaults are snapshotted, tags are copied per item, and running start is rejected', async () => {
  const { builder } = loadBuilder();
  const pending = deferred();
  const shared = defaults({ tags: ['開始時'] });
  const session = makeSession(builder, { preparePhoto: () => pending.promise });
  const first = session.start([photo('one.jpg'), photo('two.jpg')], shared);
  assert.equal(session.isRunning(), true);
  await assert.rejects(
    session.start([photo('other.jpg')], defaults()),
    (error) => error.code === 'MULTI_PHOTO_ALREADY_RUNNING'
  );
  shared.tags[0] = '変更後';
  shared.color = '#2196f3';
  pending.resolve({
    originalFile: photo('prepared.jpg'), uploadFile: photo('prepared.jpg'),
    metadataStatus: 'success', conversionStatus: 'not-needed'
  });
  const job = await first;
  assert.deepEqual(plain(job.items.map((item) => item.tags)), [['開始時'], ['開始時']]);
  assert.equal(job.items[0].color, '#e53935');
  assert.notEqual(job.items[0].tags, job.items[1].tags);
  assert.equal(session.isRunning(), false);
});

test('standard results preserve metadata, files, object URLs, safe titles, order, and persistable boundaries', async () => {
  const { core, builder } = loadBuilder();
  const first = photo('  <b>Tokyo</b>.JPG', 'image/jpeg');
  const second = photo(`${'長'.repeat(90)}.png`, 'image/png');
  const third = photo('.jpg', 'image/jpeg');
  const seenUrlFiles = [];
  let id = 0;
  const session = makeSession(builder, {
    createId: () => `same-${id++ % 2}`,
    preparePhoto: async (file) => ({
      originalFile: file,
      uploadFile: file,
      metadataStatus: file === second ? 'read-failed' : 'success',
      conversionStatus: 'not-needed',
      lat: file === first ? 35.6 : null,
      lng: file === first ? 139.7 : null,
      capturedAt: file === first ? '2026-07-11T09:30' : ''
    }),
    createObjectURL(file) { seenUrlFiles.push(file); return `blob:${seenUrlFiles.length}`; }
  });

  const job = await session.start([first, second, third], defaults());

  assert.equal(job.sourceType, 'multi-photo');
  assert.equal(new Set([job.id, ...job.items.map((item) => item.id)]).size, 4);
  assert.deepEqual(job.items.map((item) => item.sourceRef), [first.name, second.name, third.name]);
  assert.equal(job.items[0].title, '<b>Tokyo</b>');
  assert.equal(job.items[1].title.length, core.TITLE_MAX_LENGTH);
  assert.equal(job.items[2].title, '写真');
  assert.equal(job.items[0].runtime.originalFile, first);
  assert.equal(job.items[0].runtime.uploadFile, first);
  assert.equal(job.items[0].lat, 35.6);
  assert.equal(job.items[0].capturedAt, '2026-07-11T09:30');
  assert.equal(job.items[1].uploadStatus, 'queued');
  assert.equal(job.items[1].metadataStatus, 'read-failed');
  assert.deepEqual(seenUrlFiles, [first, second, third]);
  assert.equal(Object.prototype.hasOwnProperty.call(core.toPersistableItem(job.items[0]), 'runtime'), false);
  assert.doesNotMatch(JSON.stringify(core.toPersistableItem(job.items[0])), /blob:/);
});

test('all non-fatal metadata outcomes remain queued and duplicate filenames get unique item ids', async () => {
  const { builder } = loadBuilder();
  const statuses = ['success', 'no-gps', 'invalid-gps', 'unsupported', 'read-failed'];
  const session = makeSession(builder, {
    createId: () => 'duplicate-id',
    preparePhoto: async (file, details) => ({
      originalFile: file,
      uploadFile: file,
      metadataStatus: statuses[details.index],
      conversionStatus: 'not-needed'
    })
  });
  const files = statuses.map(() => photo('same-name.jpg', 'image/jpeg'));

  const job = await session.start(files, defaults());

  assert.deepEqual(job.items.map((item) => item.metadataStatus), statuses);
  assert.deepEqual(job.items.map((item) => item.uploadStatus), statuses.map(() => 'queued'));
  assert.deepEqual(job.items.map((item) => item.sourceRef), statuses.map(() => 'same-name.jpg'));
  assert.equal(new Set(job.items.map((item) => item.id)).size, statuses.length);
});

test('bounded workers refill free slots, continue after failures, and preserve selection order', async () => {
  const { builder } = loadBuilder();
  const gates = Array.from({ length: 5 }, deferred);
  const starts = [];
  let active = 0;
  let maxActive = 0;
  const session = makeSession(builder, {
    concurrency: 2,
    preparePhoto: async (file, details) => {
      starts.push(details.index);
      active += 1;
      maxActive = Math.max(maxActive, active);
      await gates[details.index].promise;
      active -= 1;
      if (details.index === 1) throw new Error('private stack detail');
      return {
        originalFile: file, uploadFile: file, metadataStatus: 'no-gps', conversionStatus: 'not-needed'
      };
    }
  });
  const run = session.start(Array.from({ length: 5 }, (_, i) => photo(`${i}.jpg`)), defaults());
  await settle();
  assert.deepEqual(starts, [0, 1]);
  gates[1].resolve();
  await settle();
  assert.deepEqual(starts, [0, 1, 2]);
  gates[2].resolve();
  await settle();
  assert.deepEqual(starts, [0, 1, 2, 3]);
  gates[3].resolve();
  await settle();
  assert.deepEqual(starts, [0, 1, 2, 3, 4]);
  gates[4].resolve();
  gates[0].resolve();
  const job = await run;
  assert.equal(maxActive, 2);
  assert.equal(new Set(starts).size, 5);
  assert.deepEqual(job.items.map((item) => item.sourceRef), ['0.jpg', '1.jpg', '2.jpg', '3.jpg', '4.jpg']);
  assert.equal(job.items[1].uploadStatus, 'failed');
  assert.doesNotMatch(job.items[1].error, /private stack detail/);
  assert.equal(job.items[4].uploadStatus, 'queued');
});

test('unexpected per-item getters are isolated as failures and remaining indexes continue', async () => {
  const { builder } = loadBuilder();
  const badName = {
    get name() { throw new Error('private filename getter'); },
    type: 'image/jpeg'
  };
  const badPrepared = photo('bad-prepared.jpg', 'image/jpeg');
  const good = photo('good.jpg', 'image/jpeg');
  const starts = [];
  const session = makeSession(builder, {
    onProgress() {},
    preparePhoto: async (file, details) => {
      starts.push(details.index);
      if (file === badPrepared) {
        return {
          originalFile: file,
          get uploadFile() { throw new Error('private upload getter'); }
        };
      }
      return {
        originalFile: file, uploadFile: file, metadataStatus: 'no-gps', conversionStatus: 'not-needed'
      };
    }
  });

  const job = await session.start([badName, badPrepared, good], defaults());

  assert.deepEqual(job.items.map((item) => item.uploadStatus), ['failed', 'failed', 'queued']);
  assert.deepEqual(starts, [1, 2]);
  assert.doesNotMatch(job.items[0].error, /private/);
  assert.doesNotMatch(job.items[1].error, /private/);
});

test('concurrency one never overlaps preparation', async () => {
  const { builder } = loadBuilder();
  let active = 0;
  let maxActive = 0;
  const session = makeSession(builder, {
    concurrency: 1,
    preparePhoto: async (file) => {
      active += 1;
      maxActive = Math.max(maxActive, active);
      await settle();
      active -= 1;
      return { originalFile: file, uploadFile: file, metadataStatus: 'no-gps', conversionStatus: 'not-needed' };
    }
  });
  await session.start([photo('1.jpg'), photo('2.jpg'), photo('3.jpg')], defaults());
  assert.equal(maxActive, 1);
});

test('progress snapshots are isolated from observer errors and never expose File objects', async () => {
  const { builder } = loadBuilder();
  const events = [];
  const session = makeSession(builder, {
    onProgress(snapshot) {
      events.push(plain(snapshot));
      if (snapshot.eventType === 'item-start') throw new Error('observer failure');
    }
  });
  const job = await session.start([photo('a.jpg'), photo('bad.svg', 'image/svg+xml')], defaults());
  assert.deepEqual(job.items.map((item) => item.uploadStatus), ['queued', 'failed']);
  assert.equal(events[0].eventType, 'start');
  assert.equal(events.at(-1).eventType, 'complete');
  assert.deepEqual(
    Object.keys(events.at(-1)).filter((key) => ['total', 'pending', 'processing', 'ready', 'failed', 'cancelled'].includes(key)).sort(),
    ['cancelled', 'failed', 'pending', 'processing', 'ready', 'total']
  );
  assert.equal(events.some((event) => Object.values(event).some((value) => value && value.name === 'a.jpg')), false);
});

test('cancel stops new claims, rejects predictably, revokes once, and ignores late results', async () => {
  const { builder } = loadBuilder();
  const gates = [deferred(), deferred(), deferred()];
  const starts = [];
  const revoked = [];
  const preparedObjects = [];
  const session = makeSession(builder, {
    concurrency: 1,
    preparePhoto: async (file, details) => {
      starts.push(details.index);
      await gates[details.index].promise;
      const result = {
        originalFile: file, uploadFile: file, metadataStatus: 'no-gps', conversionStatus: 'not-needed'
      };
      preparedObjects.push(result);
      return result;
    },
    createObjectURL: (file) => `blob:${file.name}`,
    revokeObjectURL: (url) => revoked.push(url)
  });
  const run = session.start([photo('0.jpg'), photo('1.jpg'), photo('2.jpg')], defaults());
  gates[0].resolve();
  await settle();
  assert.deepEqual(starts, [0, 1]);
  session.cancel();
  session.cancel();
  gates[1].resolve();
  await assert.rejects(run, (error) => error.code === 'MULTI_PHOTO_CANCELLED');
  assert.deepEqual(starts, [0, 1]);
  assert.deepEqual(revoked, ['blob:0.jpg']);
  assert.equal(session.isRunning(), false);
  session.release();
  session.release();
  assert.deepEqual(revoked, ['blob:0.jpg']);

  gates[2].resolve();
  const next = await session.start([photo('next.jpg')], defaults());
  assert.deepEqual(next.items.map((item) => item.sourceRef), ['next.jpg']);
});

test('conversion and object URL failures remain distinct failed items while other photos succeed', async () => {
  const { core, builder } = loadBuilder();
  const originalHeic = photo('broken.heic', 'image/heic');
  const urlFailure = photo('url.png', 'image/png');
  const okay = photo('okay.webp', 'image/webp');
  const session = makeSession(builder, {
    preparePhoto: async (file, details) => {
      if (file === originalHeic) {
        return {
          originalFile: file, uploadFile: null, metadataStatus: 'no-gps',
          conversionStatus: 'failed', errorCode: 'HEIC_CONVERSION_FAILED'
        };
      }
      return {
        originalFile: file, uploadFile: file, metadataStatus: 'no-gps',
        conversionStatus: details.kind === 'heic' ? 'succeeded' : 'not-needed'
      };
    },
    createObjectURL(file) {
      if (file === urlFailure) throw new Error('secret URL internals');
      return `blob:${file.name}`;
    }
  });
  const job = await session.start([originalHeic, urlFailure, okay], defaults());

  assert.deepEqual(job.items.map((item) => item.uploadStatus), ['failed', 'failed', 'queued']);
  assert.equal(job.items[0].conversionStatus, 'failed');
  assert.match(job.items[0].error, /変換/);
  assert.match(job.items[1].error, /プレビュー/);
  assert.doesNotMatch(job.items[1].error, /secret/);
  assert.equal(job.items[0].runtime.uploadFile, null);
  assert.equal(job.items[1].runtime.uploadFile, null);
  assert.throws(() => core.readyJob(job), (error) => error.code === 'INVALID_IMPORT_JOB_ITEMS');
});

test('the default browser preparation separates HEIC metadata and conversion statuses', async () => {
  const original = photo('iphone.HEIC', 'image/heic');
  const converted = photo('iphone.jpg', 'image/jpeg');
  const calls = [];
  const { builder } = loadBuilder({
    prepareMultiPhotoFile: async (file, details) => {
      calls.push([file, details.kind]);
      return {
        originalFile: file, uploadFile: converted, metadataStatus: 'no-gps',
        conversionStatus: 'succeeded', capturedAt: '2026-01-02T03:04'
      };
    }
  });
  const seenUrls = [];
  const session = builder.create({
    createObjectURL(file) { seenUrls.push(file); return 'blob:jpeg'; },
    revokeObjectURL() {},
    createId: (() => { let id = 0; return () => `id-${++id}`; })()
  });
  const job = await session.start([original], defaults());
  assert.deepEqual(calls, [[original, 'heic']]);
  assert.equal(job.items[0].runtime.originalFile, original);
  assert.equal(job.items[0].runtime.uploadFile, converted);
  assert.equal(job.items[0].metadataStatus, 'no-gps');
  assert.equal(job.items[0].conversionStatus, 'succeeded');
  assert.deepEqual(seenUrls, [converted]);
});

test('browser HEIC preparation reads original metadata before converting exactly once', async () => {
  const original = photo('iphone.heic', 'image/heic');
  const jpeg = photo('iphone.jpg', 'image/jpeg');
  const reader = { load() {} };
  function converter() {}
  function FileCtor() {}
  const calls = [];
  const prepare = loadDefaultPhotoPreparer({
    ExifReader: reader,
    HeicTo: converter,
    File: FileCtor,
    async readHeicMetadata(file, receivedReader) {
      calls.push('metadata');
      assert.equal(file, original);
      assert.equal(receivedReader, reader);
      return { status: 'no-gps', lat: null, lng: null, capturedAt: '2026-07-11T08:00' };
    },
    async convertHeicToJpeg(file, receivedConverter, receivedFileCtor) {
      calls.push('convert');
      assert.equal(file, original);
      assert.equal(receivedConverter, converter);
      assert.equal(receivedFileCtor, FileCtor);
      return jpeg;
    }
  });

  const result = await prepare(original, { kind: 'heic' });

  assert.deepEqual(calls, ['metadata', 'convert']);
  assert.equal(result.originalFile, original);
  assert.equal(result.uploadFile, jpeg);
  assert.equal(result.metadataStatus, 'no-gps');
  assert.equal(result.conversionStatus, 'succeeded');
  assert.equal(result.capturedAt, '2026-07-11T08:00');
});

test('browser standard preparation keeps the original queued when metadata reading fails', async () => {
  const original = photo('standard.webp', 'image/webp');
  const prepare = loadDefaultPhotoPreparer({
    async readExifGps(file) {
      assert.equal(file, original);
      throw new Error('private EXIF detail');
    },
    async readExifDateTimeOriginal(file) {
      assert.equal(file, original);
      return '2026-07-11T09:15';
    }
  });

  const result = await prepare(original, { kind: 'webp' });

  assert.equal(result.originalFile, original);
  assert.equal(result.uploadFile, original);
  assert.equal(result.metadataStatus, 'read-failed');
  assert.equal(result.conversionStatus, 'not-needed');
  assert.equal(result.capturedAt, '2026-07-11T09:15');
});

test('successful Object URL ownership transfers to the job and Core discard releases it once', async () => {
  const { core, builder } = loadBuilder();
  const revoked = [];
  const session = makeSession(builder, {
    createObjectURL: () => 'blob:owned-by-job',
    revokeObjectURL: (url) => revoked.push(url)
  });
  const job = await session.start([photo('owned.jpg')], defaults());

  session.release();
  assert.deepEqual(revoked, []);
  const urlApi = { revokeObjectURL: (url) => revoked.push(url) };
  core.releaseJobResources(job, urlApi);
  core.releaseJobResources(job, urlApi);
  assert.deepEqual(revoked, ['blob:owned-by-job']);
  assert.equal(job.items[0].runtime.originalFile, null);
  assert.equal(job.items[0].runtime.uploadFile, null);
  assert.equal(job.items[0].runtime.previewUrl, '');
});

test('session-level ids remain unique across completed starts even with a constant generator', async () => {
  const { builder } = loadBuilder();
  const session = makeSession(builder, { createId: () => 'constant-id' });
  const first = await session.start([photo('same.jpg')], defaults());
  const second = await session.start([photo('same.jpg')], defaults());
  const ids = [first.id, first.items[0].id, second.id, second.items[0].id];
  assert.equal(new Set(ids).size, ids.length);
});

test('README summarizes production multi-photo limits and recovery controls', () => {
  assert.match(readme, /端末またはGoogle Driveから最大20枚/);
  assert.match(readme, /複数写真[^\n]*1回1〜20枚[^\n]*合計100MB/);
  assert.match(readme, /キャンセル[^\n]*再開[^\n]*失敗項目の再試行[^\n]*応答喪失時の重複防止/);
  assert.match(readme, /端末とDriveから1枚／20枚[^\n]*キャンセル[^\n]*再開[^\n]*失敗再試行/);
});

test('multi-photo preview opener stays draft-only after production launch controls are added', () => {
  const { openMulti } = loadBuilder();
  assert.equal(typeof openMulti, 'function');
  assert.equal(/multiple(?:=|\s)/i.test(indexHtml.match(/<input[^>]*id="file-input"[^>]*>/)[0]), false);
  assert.equal(indexHtml.includes('id="multi-photo-file-input"'), true);
  assert.equal(indexHtml.includes('id="multi-photo-button"'), true);
  const functionSource = indexHtml.slice(
    indexHtml.indexOf('function openMultiPhotoImportPreview'),
    indexHtml.indexOf('    const state = {', indexHtml.indexOf('function openMultiPhotoImportPreview'))
  );
  assert.doesNotMatch(functionSource, /ImportQueueRunner|ImportFlowController|saveMapData|google\.script\.run/);
});

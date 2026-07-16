const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const {
  loadDriveClientModules, descriptor, fileResponse, imageFixtures
} = require('./drive-photo-import-client-harness');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

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

class FakeElement {
  constructor(tagName, id, owner) {
    this.tagName = String(tagName || 'div').toUpperCase();
    this.id = id || '';
    this.owner = owner;
    this.children = [];
    this.attributes = Object.create(null);
    this.listeners = Object.create(null);
    this.style = {};
    this.classList = { add() {}, remove() {} };
    this.textContent = '';
    this.value = '';
    this.max = 0;
    this.disabled = false;
    this.checked = false;
    this.hidden = false;
    this.type = '';
    this.className = '';
  }

  appendChild(child) { this.children.push(child); return child; }
  replaceChildren(...children) { this.children = children; }
  addEventListener(name, listener) { this.listeners[name] = listener; }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  focus() { this.owner.activeElement = this; }
}

function createDocument() {
  return {
    activeElement: null,
    elements: Object.create(null),
    createElement(tagName) { return new FakeElement(tagName, '', this); },
    getElementById(id) {
      if (!this.elements[id]) this.elements[id] = new FakeElement('div', id, this);
      return this.elements[id];
    }
  };
}

class BrowserFile extends Blob {
  constructor(parts, name, options = {}) {
    super(parts, options);
    this.name = name;
    this.lastModified = options.lastModified;
  }
}

function loadBuilder(preparePhoto) {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  const context = {
    console, Number, Object, Array, String, Error, Set, Promise, Date, Math,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'photo' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    File: BrowserFile,
    Blob,
    prepareMultiPhotoFile: preparePhoto || (async (file) => ({
      originalFile: file, uploadFile: file, lat: null, lng: null, capturedAt: '',
      metadataStatus: 'no-gps', conversionStatus: 'not-needed'
    })),
    URL: { createObjectURL(file) { return `blob:${file.name}`; }, revokeObjectURL() {} },
    crypto: { randomUUID: () => 'uuid' }
  };
  vm.createContext(context);
  vm.runInContext(`${indexHtml.slice(start, end)}\n`
    + 'globalThis.__builder = MultiPhotoImportBuilder; globalThis.__workflow = MultiPhotoImportWorkflow;'
    + 'globalThis.__processor = ImportPhotoItemProcessor; globalThis.__core = ImportJobCore;', context);
  return {
    builder: context.__builder, workflow: context.__workflow,
    processor: context.__processor, core: context.__core
  };
}

function loadDefaultPhotoPreparer(stubs) {
  const start = indexHtml.indexOf('async function prepareMultiPhotoFile(');
  const end = indexHtml.indexOf('    function locationMessageForMetadataStatus(', start);
  const context = { console, File: BrowserFile, ExifReader: {}, HeicTo() {}, ...stubs };
  vm.createContext(context);
  vm.runInContext(`${indexHtml.slice(start, end)}\nglobalThis.__prepare = prepareMultiPhotoFile;`, context);
  return context.__prepare;
}

function environment() {
  return {
    atob(value) { return Buffer.from(value, 'base64').toString('binary'); },
    Uint8Array, Blob, File: BrowserFile
  };
}

function inputs(count) {
  return Array.from({ length: count }, (_, index) => descriptor({
    id: `photo_${String(index).padStart(12, '0')}`,
    name: `${index}.jpg`
  }));
}

test('real Drive loader feeds one and multiple BrowserFiles into the existing ordered multi-photo builder', async () => {
  const { loader, sourceCore } = loadDriveClientModules();
  const { builder } = loadBuilder();
  const calls = [];
  const allDescriptors = inputs(20);
  const instance = loader.create({
    callGAS: async (method, payload) => {
      calls.push([method, { ...payload }]);
      const expected = allDescriptors.find((value) => value.id === payload.fileId);
      return fileResponse(expected);
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'send-only' }),
    sourceCore,
    environment: environment()
  });

  for (const count of [1, 20]) {
    const descriptors = allDescriptors.slice(0, count);
    const files = await instance.start(descriptors);
    const session = builder.create({
      concurrency: 2,
      preparePhoto: async (file) => ({
        originalFile: file, uploadFile: file, lat: null, lng: null,
        capturedAt: '', metadataStatus: 'no-gps', conversionStatus: 'not-needed'
      }),
      createObjectURL: (file) => `blob:${file.name}`,
      revokeObjectURL() {},
      createId: (_prefix, index) => `id-${count}-${index}`
    });
    const job = await session.start(files, {
      tags: [], color: '#e53935', icon: 'photo', status: ''
    }, { sourceDriveFileIds: descriptors.map((value) => value.id) });
    assert.equal(job.sourceType, 'multi-photo');
    assert.equal(job.items.length, count);
    assert.deepEqual(Array.from(job.items, (item) => item.runtime.originalFile.name),
      descriptors.map((value) => value.name));
    assert.equal(job.items.every((item) => item.runtime.originalFile instanceof BrowserFile), true);
    assert.equal(job.items.every((item) => !Object.hasOwn(item.runtime.originalFile, 'id')), true);
    assert.deepEqual(Array.from(job.items, (item) => item.runtime.sourceDriveFileId),
      descriptors.map((value) => value.id));
    assert.equal(job.items.every((item) => item.capturedAt === ''), true);
  }
  assert.equal(calls.every(([method]) => method === 'readDrivePhotoImportFile'), true);
  assert.equal(calls.every(([, payload]) => Object.keys(payload).sort().join(',') === '__editToken,fileId'), true);
});

test('Drive HEIC and HEIF both use the common conversion path and keep EXIF time distinct from modifiedAt', async () => {
  const { loader, sourceCore } = loadDriveClientModules();
  const modules = loadBuilder();
  const sources = [
    descriptor({
      id: 'heic_AAAAAAAAAAAA', name: 'drive.heic', mimeType: 'image/heic', kind: 'heic',
      modifiedAt: '2026-07-12T01:02:03.000Z'
    }),
    descriptor({
      id: 'heif_AAAAAAAAAAAA', name: 'drive.heif', mimeType: 'image/heif', kind: 'heic',
      modifiedAt: '2026-07-13T01:02:03.000Z'
    })
  ];
  const loaderInstance = loader.create({
    callGAS: async (_method, payload) => fileResponse(
      sources.find((source) => source.id === payload.fileId)
    ),
    withEditToken: (payload) => ({ ...payload, __editToken: 'send-only' }),
    sourceCore,
    environment: environment()
  });
  const driveFiles = await loaderInstance.start(sources);
  assert.deepEqual(Array.from(driveFiles, (file) => file.lastModified), sources.map(
    (source) => Date.parse(source.modifiedAt)
  ));
  const order = [];
  const prepare = loadDefaultPhotoPreparer({
    readHeicMetadata: async (file) => {
      order.push(['metadata', file.name]);
      return { status: 'ok', lat: 35, lng: 139, capturedAt: '2025-05-06T07:08' };
    },
    convertHeicToJpeg: async (file) => {
      order.push(['convert', file.name]);
      return new BrowserFile([new Uint8Array([1, 2, 3])], file.name.replace(/\.(?:heic|heif)$/i, '.jpg'), {
        type: 'image/jpeg', lastModified: file.lastModified
      });
    }
  });
  const created = [];
  const revoked = [];
  let nextId = 0;
  const session = modules.builder.create({
    preparePhoto: prepare,
    createObjectURL(file) { created.push(file); return `blob:${file.name}`; },
    revokeObjectURL(url) { revoked.push(url); },
    createId: () => `id-${++nextId}`
  });
  const job = await session.start(driveFiles, {
    tags: [], color: '#e53935', icon: 'photo', status: ''
  });
  assert.deepEqual(order, [
    ['metadata', 'drive.heic'], ['metadata', 'drive.heif'],
    ['convert', 'drive.heic'], ['convert', 'drive.heif']
  ]);
  assert.equal(job.items.every((item) => item.capturedAt === '2025-05-06T07:08'), true);
  assert.equal(job.items.every((item, index) => item.capturedAt !== sources[index].modifiedAt), true);
  assert.equal(job.items.every((item) => item.conversionStatus === 'succeeded'), true);
  assert.equal(job.items.every((item) => item.uploadStatus === 'queued'), true);
  assert.deepEqual(created.map((file) => file.name), ['drive.jpg', 'drive.jpg']);
  const resourceUrlApi = modules.builder.getResourceUrlApi(job);
  modules.core.releaseJobResources(job, resourceUrlApi);
  modules.core.releaseJobResources(job, resourceUrlApi);
  assert.deepEqual(revoked, ['blob:drive.jpg']);
});

test('production wiring reuses existing Preview track match processor and save retry boundaries', () => {
  const workflow = indexHtml.slice(
    indexHtml.indexOf('const MultiPhotoImportWorkflow ='),
    indexHtml.indexOf('const CsvInterchangeUI =')
  );
  assert.match(workflow, /builderApi\.create/);
  assert.match(workflow, /createTrackMatch/);
  assert.match(workflow, /processorApi\.create/);
  assert.match(workflow, /flowApi\.create/);
  assert.match(workflow, /title: '複数写真を確認'/);
  assert.match(indexHtml, /onFilesReady:[\s\S]*startPhotoImportFromFiles/);
  assert.match(indexHtml, /saveImportPhotoItem/);
  assert.equal(indexHtml.includes('saveDrivePhotoImportItem'), false);
  assert.equal(indexHtml.includes('DrivePhotoImportPreview'), false);
});

test('real Drive UI handoff resolves the root before Workflow and Builder completes every selection as ready', async () => {
  const { ui, loader, sourceCore } = loadDriveClientModules();
  const modules = loadBuilder(async (file) => ({
    originalFile: file,
    uploadFile: file,
    lat: null,
    lng: null,
    capturedAt: '',
    metadataStatus: 'no-gps',
    conversionStatus: 'not-needed'
  }));
  const descriptors = [
    descriptor({
      id: 'photo_JPGAAAAAAAA', name: 'drive.jpg', sizeBytes: imageFixtures.jpeg.length
    }),
    descriptor({
      id: 'photo_PNGAAAAAAAA', name: 'drive.png', mimeType: 'image/png', kind: 'png',
      sizeBytes: imageFixtures.png.length
    }),
    descriptor({
      id: 'photo_WEBPAAAAAAA', name: 'drive.webp', mimeType: 'image/webp', kind: 'webp',
      sizeBytes: imageFixtures.webp.length
    })
  ];
  const fixtures = [imageFixtures.jpeg, imageFixtures.png, imageFixtures.webp];
  const driveSourceState = {};
  const workflowState = {
    builder: null, controller: null, job: null, resourceUrlApi: null,
    targetFolderId: '', preparing: false, registering: false,
    cancellingPreparation: false, requestToken: 0
  };
  let completeProgress = null;
  let processorTargetFolderId = '';
  let previewJob = null;
  const workflow = modules.workflow.create({
    state: workflowState,
    builderApi: modules.builder,
    processorApi: {
      create(config) {
        processorTargetFolderId = config.getTargetFolderId();
        return modules.processor.create(config);
      }
    },
    flowApi: {
      create() {
        return {
          open(options) { previewJob = options.job; },
          isRunning() { return false; }
        };
      }
    },
    validator: { validateJob() { return true; } },
    resizePhoto: async () => 'data:image/jpeg;base64,/9j/',
    callGAS: async () => ({ ok: true }),
    withEditToken: (payload) => payload,
    resetTrackMatch() {},
    onProgress(progress) {
      if (progress.eventType === 'complete') completeProgress = { ...progress };
    }
  });
  const documentApi = createDocument();
  const startContext = {
    ImportJobCore: { MAX_ITEMS: 20 },
    driveSourceState,
    document: documentApi,
    resetMultiPhotoPreparationView() {},
    closeOverlay() {},
    openOverlay() {},
    multiPhotoWorkflow: workflow,
    refreshMultiPhotoButtonState() {}
  };
  vm.runInNewContext(
    `${sourceFunction(indexHtml, 'startPhotoImportFromFiles')}\n`
      + 'globalThis.__start = startPhotoImportFromFiles;',
    startContext
  );
  const returnMessages = [];
  const picker = ui.create({
    state: driveSourceState,
    documentApi,
    sourceCore,
    loaderApi: loader,
    environment: environment(),
    callGAS: async (method, payload) => {
      if (method === 'listDrivePhotoImportFolder') {
        return {
          ok: true,
          folder: { id: 'root_AAAAAAAAAAA', name: 'Root', isRoot: true },
          parent: null,
          folders: [],
          photos: descriptors,
          ignoredUnsupportedFileCount: 0,
          counts: { folders: 0, photos: descriptors.length }
        };
      }
      assert.equal(method, 'readDrivePhotoImportFile');
      const index = descriptors.findIndex((value) => value.id === payload.fileId);
      const source = descriptors[index];
      const bytes = fixtures[index];
      return fileResponse({
        ...source,
        mimeType: source.mimeType.toUpperCase(),
        sizeBytes: 1,
        modifiedAt: `2026-07-${15 + index}T10:20:30.000Z`,
        kind: 'ignored'
      }, bytes.toString('base64'));
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'send-only-token' }),
    canStart: () => true,
    getDefaults: () => ({
      tags: ['drive'], color: '#e53935', icon: 'photo', status: '', targetFolderId: ''
    }),
    onFilesReady: (files, snapshot, options) => startContext.__start(files, snapshot, {
      sourceDriveFileIds: options.sourceDriveFileIds
    }),
    openPicker() {},
    closePicker() {},
    returnToPhotoSource(message) { returnMessages.push(message); },
    onBusy() {}
  });

  assert.equal(await picker.open(), true);
  descriptors.forEach((value) => picker.toggle(value.id, true));
  assert.equal(await picker.confirm(), true);

  assert.equal(returnMessages.length, 0);
  assert.equal(completeProgress.eventType, 'complete');
  assert.equal(completeProgress.ready, descriptors.length);
  assert.equal(completeProgress.total, descriptors.length);
  assert.equal(processorTargetFolderId, 'root_AAAAAAAAAAA');
  assert.equal(workflowState.targetFolderId, 'root_AAAAAAAAAAA');
  assert.equal(previewJob.items.length, descriptors.length);
  assert.equal(previewJob.items.every((item) => item.uploadStatus === 'queued'), true);
  assert.deepEqual(Array.from(previewJob.items, (item) => item.runtime.originalFile.type),
    ['image/jpeg', 'image/png', 'image/webp']);
  assert.doesNotMatch(JSON.stringify(driveSourceState),
    /photo_(?:JPG|PNG|WEBP)|base64|AQID|send-only-token|__editToken/);
});

async function runDriveSaveScenario(count, sameSourceAndTarget, locations) {
  const { loader, sourceCore } = loadDriveClientModules();
  const modules = loadBuilder(async (file) => {
    const index = Number(String(file.name).replace(/\..*$/, ''));
    const location = Array.isArray(locations) ? locations[index] : null;
    const hasGps = !!location && Number.isFinite(location.lat) && Number.isFinite(location.lng);
    return {
      originalFile: file,
      uploadFile: file,
      lat: hasGps ? location.lat : null,
      lng: hasGps ? location.lng : null,
      capturedAt: '',
      metadataStatus: hasGps ? 'success' : 'no-gps',
      conversionStatus: 'not-needed'
    };
  });
  const descriptors = inputs(count);
  const sourceFolderId = 'root_AAAAAAAAAAA';
  const targetFolderId = sameSourceAndTarget ? sourceFolderId : 'target_AAAAAAAAA';
  const sourceAudit = { rename: 0, move: 0, trash: 0, share: 0, byteWrites: 0, reads: 0 };
  const sourceFiles = new Map(descriptors.map((value) => [value.id, value]));
  const loaderInstance = loader.create({
    callGAS: async (method, payload) => {
      assert.equal(method, 'readDrivePhotoImportFile');
      assert.equal(payload.__editToken, 'classroom-token');
      sourceAudit.reads += 1;
      return fileResponse(sourceFiles.get(payload.fileId));
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'classroom-token' }),
    sourceCore,
    environment: environment()
  });
  const files = await loaderInstance.start(descriptors);
  const state = {
    builder: null, controller: null, job: null, resourceUrlApi: null,
    targetFolderId: '', preparing: false, registering: false,
    cancellingPreparation: false, requestToken: 0
  };
  const managedLinks = [];
  const mapRows = [];
  const receipts = new Map();
  const savePayloads = [];
  const savedPins = [];
  const placedPins = [];
  const unplacedPins = [];
  let flowConfig;
  let previewOptions;
  let matcherInitialized = 0;
  let loseFirstResponse = true;
  const workflow = modules.workflow.create({
    state,
    builderApi: modules.builder,
    processorApi: modules.processor,
    flowApi: {
      create(config) {
        flowConfig = config;
        return { open(options) { previewOptions = options; } };
      }
    },
    validator: { validateJob() { return true; } },
    resizePhoto: async (file) => {
      assert.equal(file instanceof BrowserFile, true);
      return 'data:image/jpeg;base64,/9j/';
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'classroom-token' }),
    callGAS: async (method, payload) => {
      assert.equal(method, 'saveImportPhotoItem');
      savePayloads.push({ ...payload });
      const key = payload.idempotencyKey;
      let receipt = receipts.get(key);
      if (!receipt) {
        const pin = {
          id: `pin-${receipts.size + 1}`,
          fileId: `managed-${receipts.size + 1}`,
          lat: payload.lat,
          lng: payload.lng
        };
        managedLinks.push({
          sourceDriveFileId: payload.sourceDriveFileId,
          fileId: pin.fileId,
          filename: payload.filename
        });
        mapRows.push({ pinId: pin.id, fileId: pin.fileId, lat: pin.lat, lng: pin.lng });
        receipt = { pin, sourceDriveFileId: payload.sourceDriveFileId };
        receipts.set(key, receipt);
        if (loseFirstResponse) {
          loseFirstResponse = false;
          const error = new Error('response lost after commit');
          error.code = 'RESPONSE_LOST';
          throw error;
        }
      }
      return { ok: true, deduplicated: true, pin: receipt.pin };
    },
    createTrackMatch: function() {
      return {
        initialize() { matcherInitialized += 1; },
        getViewModel() { return {}; },
        cleanup() {}
      };
    },
    resetTrackMatch() {},
    onSaved(pin) {
      savedPins.push(pin.id);
      if (Number.isFinite(pin.lat) && Number.isFinite(pin.lng)) placedPins.push(pin.id);
      if (pin.lat === null && pin.lng === null) unplacedPins.push(pin.id);
    }
  });
  const job = await workflow.start(files, {
    tags: ['drive'], color: '#e53935', icon: 'photo', status: '', targetFolderId
  }, { sourceDriveFileIds: descriptors.map((value) => value.id) });
  assert.equal(previewOptions.job, job);
  assert.equal(previewOptions.trackMatch != null, true);
  assert.equal(matcherInitialized, 1);
  assert.equal(job.sourceType, 'multi-photo');
  assert.deepEqual(Array.from(job.items, (item) => item.runtime.originalFile.name),
    descriptors.map((value) => value.name));

  for (let index = 0; index < job.items.length; index += 1) {
    const item = job.items[index];
    const context = { jobId: job.id, itemId: item.id, attempt: 1 };
    if (index === 0) {
      await assert.rejects(flowConfig.processItem(item, context));
      context.attempt = 2;
    }
    await flowConfig.processItem(item, context);
  }

  assert.equal(managedLinks.length, count);
  assert.equal(mapRows.length, count);
  assert.equal(receipts.size, count);
  assert.equal(savedPins.length, count);
  assert.deepEqual(
    managedLinks.map((link) => link.sourceDriveFileId),
    descriptors.map((value) => value.id)
  );
  assert.equal(managedLinks.every((link) => link.fileId !== link.sourceDriveFileId), true);
  assert.equal(new Set(managedLinks.map((link) => link.fileId)).size, count);
  assert.equal(savePayloads[0].idempotencyKey, savePayloads[1].idempotencyKey);
  assert.equal(savePayloads.every((payload) => payload.targetFolderId === targetFolderId), true);
  assert.deepEqual(savePayloads.map((payload) => payload.sourceDriveFileId), [
    descriptors[0].id,
    ...descriptors.map((value) => value.id)
  ]);
  const forbidden = [
    'sourceFolderId', 'sourceModifiedAt', 'sourceDescriptor', 'importOrigin',
    'driveFileId', 'owner', 'permission', 'trackMatchMetadata'
  ];
  savePayloads.forEach((payload) => {
    assert.equal(payload.__editToken, 'classroom-token');
    forbidden.forEach((key) => assert.equal(Object.hasOwn(payload, key), false, key));
  });
  assert.deepEqual(sourceAudit, {
    rename: 0, move: 0, trash: 0, share: 0, byteWrites: 0, reads: count
  });
  return {
    managedLinks, mapRows, savePayloads, sourceFolderId, targetFolderId,
    savedPins, placedPins, unplacedPins
  };
}

test('Drive one-photo save uses the existing Preview matcher processor and passes the source for storage planning', async () => {
  const result = await runDriveSaveScenario(1, true);
  assert.equal(result.sourceFolderId, result.targetFolderId);
  assert.equal(result.managedLinks.length, 1);
  assert.equal(result.mapRows.length, 1);
});

test('Drive twenty-photo save preserves source order and response-loss retry deduplicates', async () => {
  const result = await runDriveSaveScenario(20, false);
  assert.notEqual(result.sourceFolderId, result.targetFolderId);
  assert.equal(result.managedLinks.length, 20);
  assert.equal(result.mapRows.length, 20);
  assert.equal(result.savePayloads.length, 21);
});

test('Drive GPS and no-GPS saves notify onSaved once into placed/map and unplaced buckets', async () => {
  const result = await runDriveSaveScenario(2, true, [
    { lat: 35.681236, lng: 139.767125 },
    { lat: null, lng: null }
  ]);

  assert.deepEqual(result.savedPins, ['pin-1', 'pin-2']);
  assert.deepEqual(result.placedPins, ['pin-1']);
  assert.deepEqual(result.unplacedPins, ['pin-2']);
  assert.deepEqual(result.mapRows.map((row) => [row.lat, row.lng]), [
    [35.681236, 139.767125],
    [null, null]
  ]);
  assert.equal(result.managedLinks[0].sourceDriveFileId !== result.managedLinks[0].fileId, true);
  assert.equal(result.managedLinks[1].sourceDriveFileId !== result.managedLinks[1].fileId, true);
});

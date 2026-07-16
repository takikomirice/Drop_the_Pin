const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const { loadDriveClientModules, descriptor } = require('./drive-photo-import-client-harness');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

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

function createStartHarness(driveSourceState = {}) {
  const helper = sourceFunction(indexHtml, 'startPhotoImportFromFiles');
  const calls = [];
  const elements = new Proxy({}, {
    get(target, key) {
      if (!target[key]) {
        target[key] = { classList: { add() {}, remove() {} }, value: '', textContent: '', style: {} };
      }
      return target[key];
    }
  });
  const context = {
    ImportJobCore: { MAX_ITEMS: 20 },
    driveSourceState,
    document: { getElementById(id) { return elements[id]; } },
    resetMultiPhotoPreparationView(count) { calls.push(['reset', count]); },
    closeOverlay(id) { calls.push(['close', id]); },
    openOverlay(id) { calls.push(['open', id]); },
    multiPhotoWorkflow: {
      start(input, snapshot, options) {
        calls.push(['start', input.slice(), snapshot, options]);
        return Promise.resolve('job');
      }
    },
    refreshMultiPhotoButtonState() {}
  };
  vm.runInNewContext(`${helper}\nthis.start = startPhotoImportFromFiles;`, context);
  return { start: context.start, calls };
}

test('local and Drive files use one common handoff into the existing multi-photo workflow', async () => {
  const helper = sourceFunction(indexHtml, 'startPhotoImportFromFiles');
  assert.match(helper, /MultiPhotoImport|multiPhotoWorkflow\.start/);
  assert.match(sourceFunction(indexHtml, 'handleMultiPhotoFilesSelected'), /startPhotoImportFromFiles/);
  assert.match(indexHtml, /onFilesReady:[\s\S]*startPhotoImportFromFiles/);
  assert.equal((indexHtml.match(/DriveMultiPhoto|DriveImportJob|DriveImportPreview/g) || []).length, 0);
  assert.doesNotMatch(helper, /upload-overlay|drive-photo-import-overlay/);

  const files = [{ name: 'one.jpg' }];
  const harness = createStartHarness({ rootFolderId: 'root_AAAAAAAAAAA' });
  const snapshot = { targetFolderId: 'target_AAAAAAAAA', tags: [], color: '#e53935', icon: 'photo', status: '' };
  const sourceDriveFileIds = ['photo_AAAAAAAAAAA'];
  const result = await harness.start(files, snapshot, { sourceDriveFileIds });
  assert.equal(result, 'job');
  assert.deepEqual(harness.calls.slice(0, 2), [
    ['reset', 1], ['open', 'multi-photo-preparation-overlay']
  ]);
  assert.equal(harness.calls[2][0], 'start');
  assert.deepEqual(JSON.parse(JSON.stringify(harness.calls[2][1])), files);
  assert.notEqual(harness.calls[2][2], snapshot);
  assert.equal(harness.calls[2][2].targetFolderId, 'target_AAAAAAAAA');
  assert.deepEqual(JSON.parse(JSON.stringify(harness.calls[2][3])), { sourceDriveFileIds });
  assert.match(
    indexHtml,
    /onFilesReady:\s*function\(files, snapshot, options\)\s*\{\s*return startPhotoImportFromFiles\(/
  );
});

test('Drive handoff fills an empty explicit destination from the validated root id', async () => {
  const harness = createStartHarness({ rootFolderId: 'root_AAAAAAAAAAA' });
  const snapshot = {
    targetFolderId: '', tags: [], color: '#e53935', icon: 'photo', status: ''
  };

  await harness.start([{ name: 'one.jpg' }], snapshot, {
    sourceDriveFileIds: ['photo_AAAAAAAAAAA']
  });

  const startCall = harness.calls.find(([name]) => name === 'start');
  assert.notEqual(startCall[2], snapshot);
  assert.equal(snapshot.targetFolderId, '');
  assert.equal(startCall[2].targetFolderId, 'root_AAAAAAAAAAA');
});

test('Drive handoff never treats the currently viewed child as the fallback destination', async () => {
  const harness = createStartHarness({
    rootFolderId: 'root_AAAAAAAAAAA',
    sourceFolderId: 'child_AAAAAAAAAA'
  });

  await harness.start([{ name: 'one.jpg' }], {
    targetFolderId: '', tags: [], color: '#e53935', icon: 'photo', status: ''
  }, {
    sourceDriveFileIds: ['photo_AAAAAAAAAAA']
  });

  const startCall = harness.calls.find(([name]) => name === 'start');
  assert.equal(startCall[2].targetFolderId, 'root_AAAAAAAAAAA');
  assert.notEqual(startCall[2].targetFolderId, 'child_AAAAAAAAAA');
});

test('Drive handoff prefers an explicit destination over the validated root id', async () => {
  const harness = createStartHarness({ rootFolderId: 'root_AAAAAAAAAAA' });
  const snapshot = {
    targetFolderId: ' target_AAAAAAAAA ', tags: [], color: '#e53935', icon: 'photo', status: ''
  };

  await harness.start([{ name: 'one.jpg' }], snapshot, {
    sourceDriveFileIds: ['photo_AAAAAAAAAAA']
  });

  const startCall = harness.calls.find(([name]) => name === 'start');
  assert.equal(startCall[2].targetFolderId, 'target_AAAAAAAAA');
});

test('Drive handoff safely aborts before Workflow when explicit and root destinations are missing', async () => {
  const harness = createStartHarness({ rootFolderId: '', sourceFolderId: 'child_AAAAAAAAAA' });

  await assert.rejects(
    harness.start([{ name: 'one.jpg' }], {
      targetFolderId: '', tags: [], color: '#e53935', icon: 'photo', status: ''
    }, {
      sourceDriveFileIds: ['photo_AAAAAAAAAAA']
    }),
    (error) => error.code === 'DRIVE_IMPORT_TARGET_FOLDER_REQUIRED'
      && error.message === 'Driveの保存先フォルダを確認できませんでした。もう一度開いてください。'
  );

  assert.equal(harness.calls.some(([name]) => name === 'start'), false);
});

test('local handoff keeps its explicit destination snapshot and ignores the Drive root', async () => {
  const harness = createStartHarness({ rootFolderId: 'root_AAAAAAAAAAA' });
  const snapshot = {
    targetFolderId: 'local_AAAAAAAAAA', tags: [], color: '#e53935', icon: 'photo', status: ''
  };

  await harness.start([{ name: 'local.jpg' }], snapshot);

  const startCall = harness.calls.find(([name]) => name === 'start');
  assert.equal(startCall[2], snapshot);
  assert.equal(startCall[2].targetFolderId, 'local_AAAAAAAAAA');
  assert.deepEqual(JSON.parse(JSON.stringify(startCall[3])), {});
});

test('photo import defaults and start preparation do not read or reset the single-pin form', async () => {
  const defaultsSource = sourceFunction(indexHtml, 'snapshotPhotoImportDefaults');
  assert.doesNotMatch(defaultsSource, /upload-(?:tags|status)|state\.upload|currentFolderIdOrRoot/);
  const defaultsContext = {
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'photo' }],
    state: { appSettings: { rootFolderId: 'root_AAAAAAAAAAA' } }
  };
  vm.runInNewContext(`${defaultsSource}\nthis.snapshot = snapshotPhotoImportDefaults;`, defaultsContext);
  assert.deepEqual(JSON.parse(JSON.stringify(defaultsContext.snapshot())), {
    tags: [], color: '#e53935', icon: 'photo', status: '未対応',
    targetFolderId: 'root_AAAAAAAAAAA'
  });

  const form = {
    title: '残すタイトル', description: '残す説明', tags: '#keep', status: '対応中'
  };
  let starts = 0;
  let resetCalls = 0;
  const prepareSource = sourceFunction(indexHtml, 'prepareAddMenuPhotoImport');
  const prepareContext = {
    state: { addMenuPreparing: false },
    canStartAddAction: () => true,
    canEdit: () => true,
    isProductionImportBusy: () => false,
    canStartMultiPhotoImport: () => true,
    refreshAddMenuState() {},
    closeOverlay() {},
    document: { getElementById() { return { textContent: '' }; } },
    resetUploadState: async () => {
      resetCalls += 1;
      form.title = '';
      form.description = '';
      form.tags = '';
      form.status = '未対応';
    }
  };
  vm.runInNewContext(`${prepareSource}\nthis.prepare = prepareAddMenuPhotoImport;`, prepareContext);

  await prepareContext.prepare(() => { starts += 1; });

  assert.equal(starts, 1);
  assert.equal(resetCalls, 0);
  assert.deepEqual(form, {
    title: '残すタイトル', description: '残す説明', tags: '#keep', status: '対応中'
  });
});

test('photo preparation returns only to the source selector and never opens single-pin upload', () => {
  const returnSource = sourceFunction(indexHtml, 'returnToPhotoSource');
  const calls = [];
  const elements = { 'multi-photo-conflict-message': { textContent: '' } };
  const context = {
    document: { getElementById(id) { return elements[id]; } },
    closeOverlay(id) { calls.push(['close', id]); },
    openOverlay(id) { calls.push(['open', id]); },
    refreshMultiPhotoButtonState() {
      calls.push(['refresh']);
      elements['multi-photo-conflict-message'].textContent = '';
    }
  };
  vm.runInNewContext(`${returnSource}\nthis.returnToSource = returnToPhotoSource;`, context);

  assert.equal(context.returnToSource('安全なメッセージ'), true);

  assert.deepEqual(calls, [
    ['close', 'multi-photo-preparation-overlay'],
    ['open', 'add-menu-overlay'],
    ['refresh']
  ]);
  assert.equal(elements['multi-photo-conflict-message'].textContent, '安全なメッセージ');
  assert.equal(calls.some(([, id]) => id === 'upload-overlay'), false);
});

test('Drive loader progress is safe and one or multiple files hand off in descriptor order without state retention', async () => {
  const { ui, sourceCore } = loadDriveClientModules();
  assert.ok(ui);
  const elements = Object.create(null);
  const documentApi = {
    createElement() { return { children: [], style: {}, appendChild(child) { this.children.push(child); },
      replaceChildren(...children) { this.children = children; }, addEventListener() {}, setAttribute() {}, focus() {} }; },
    getElementById(id) {
      if (!elements[id]) elements[id] = { children: [], style: {}, textContent: '', value: 0, max: 0,
        disabled: false, replaceChildren(...children) { this.children = children; }, setAttribute() {},
        addEventListener() {}, focus() {} };
      return elements[id];
    }
  };
  const photos = [descriptor({ id: 'photo_AAAAAAAAAAA', name: 'a.jpg' }),
    descriptor({ id: 'photo_BBBBBBBBBBB', name: 'b.jpg' })];
  const handedOff = [];
  let progressObserver;
  const state = {};
  const controller = ui.create({
    state, documentApi, sourceCore,
    callGAS: async () => ({
      ok: true, folder: { id: 'root_AAAAAAAAAAA', name: 'Root', isRoot: true }, parent: null,
      folders: [], photos, ignoredUnsupportedFileCount: 0, counts: { folders: 0, photos: 2 }
    }),
    withEditToken: (payload) => payload,
    loaderApi: { create(options) {
      progressObserver = options.onProgress;
      return { async start(values) {
        progressObserver({ eventType: 'item-start', total: values.length, completed: 0,
          currentIndex: 0, currentName: values[0].name, id: values[0].id, base64: 'secret' });
        return values.map((value) => ({ name: value.name, type: value.mimeType, size: value.sizeBytes }));
      }, cancel() {}, release() {}, isRunning() { return false; } };
    } },
    environment: {}, canStart: () => true,
    getDefaults: () => ({ targetFolderId: 'target_AAAAAAAAA', tags: [], color: '#e53935', icon: 'photo', status: '' }),
    onFilesReady: async (files, _defaults, options) => handedOff.push({
      names: files.map((file) => file.name),
      ids: options.sourceDriveFileIds.slice()
    }),
    openPicker() {}, closePicker() {}, returnToPhotoSource() {}, onBusy() {}
  });
  await controller.open();
  controller.toggle(photos[1].id, true);
  controller.toggle(photos[0].id, true);
  await controller.confirm();
  assert.deepEqual(handedOff, [{
    names: ['a.jpg', 'b.jpg'], ids: ['photo_AAAAAAAAAAA', 'photo_BBBBBBBBBBB']
  }]);
  assert.equal(Object.prototype.hasOwnProperty.call(state, 'files'), false);
  const serialized = JSON.stringify(state);
  assert.equal(serialized.includes('base64'), false);
  assert.equal(serialized.includes('photo_AAAAAAAAAAA'), false);
  assert.equal(state.progress, null);
});

test('Drive picker loading and handoff states feed production busy and beforeunload guards until cleanup', () => {
  const source = sourceFunction(indexHtml, 'isMultiPhotoImportBusyState');
  assert.match(source, /driveSource/);
  assert.match(source, /open|folderLoading/);
  assert.match(source, /fileLoading/);
  assert.match(source, /handingOff/);
  assert.match(sourceFunction(indexHtml, 'isShortcutOverlayOpen'), /drive-photo-import-overlay/);
  assert.match(sourceFunction(indexHtml, 'hasPendingMutationWork'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'initializeApp'), /beforeunload[\s\S]*hasPendingMutationWork\(\)/);
  const state = { multiPhotoImport: {
    preparing: false, registering: false, cancellingPreparation: false,
    builder: null, job: null, controller: null, driveSource: {}
  } };
  const context = { state };
  vm.runInNewContext(`${source}\nthis.busy = isMultiPhotoImportBusyState;`, context);
  for (const key of ['open', 'folderLoading', 'fileLoading', 'handingOff']) {
    state.multiPhotoImport.driveSource = { [key]: true };
    assert.equal(context.busy(), true, key);
  }
  state.multiPhotoImport.driveSource = {
    open: false, folderLoading: false, fileLoading: false, handingOff: false
  };
  assert.equal(context.busy(), false);
});

test('existing processor remains the only Drive save path and excludes unrelated source metadata', () => {
  const processor = indexHtml.slice(
    indexHtml.indexOf('const ImportPhotoItemProcessor ='),
    indexHtml.indexOf('const MultiPhotoImportWorkflow =')
  );
  assert.match(processor, /saveImportPhotoItem/);
  assert.match(processor, /sourceDriveFileId/);
  for (const forbidden of [
    'sourceFolderId', 'sourceModifiedAt', 'sourceDescriptor', 'importOrigin',
    'driveFileId', 'owner', 'permission', 'trackMatchMetadata'
  ]) assert.equal(processor.includes(forbidden), false, forbidden);
  assert.equal(indexHtml.includes("withGAS('saveDrivePhoto"), false);
});

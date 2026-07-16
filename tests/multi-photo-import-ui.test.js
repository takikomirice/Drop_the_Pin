const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

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

test('production UI adds a separate edit-only multi-photo input without changing the single input', () => {
  const single = indexHtml.match(/<input[^>]*id="file-input"[^>]*>/)[0];
  assert.equal(/\bmultiple\b/i.test(single), false);
  assert.match(single, /image\/gif/);
  const multi = indexHtml.match(/<input[^>]*id="multi-photo-file-input"[^>]*>/)[0];
  assert.match(multi, /\bmultiple\b/i);
  assert.match(multi, /hidden/);
  for (const value of ['.jpg', '.jpeg', '.png', '.webp', '.heic', '.heif', 'image/jpeg', 'image/png', 'image/webp', 'image/heic', 'image/heif']) {
    assert.ok(multi.includes(value), value);
  }
  assert.equal(multi.includes('image/gif'), false);
  assert.match(indexHtml, /id="multi-photo-button"[\s\S]*?add-menu-choice-title">端末から写真を取込</);
  assert.match(indexHtml, /1〜20枚を共通フローで登録/);
  assert.match(indexHtml, /body:not\(\.has-edit-token\) #multi-photo-button/);
  assert.match(indexHtml, /body:not\(\.edit-mode\) #multi-photo-button/);
});

test('preparation overlay exposes safe progress, cancellation, and error controls', () => {
  for (const id of [
    'multi-photo-preparation-overlay', 'multi-photo-preparation-status',
    'multi-photo-preparation-total', 'multi-photo-preparation-pending',
    'multi-photo-preparation-processing', 'multi-photo-preparation-ready',
    'multi-photo-preparation-failed', 'multi-photo-preparation-cancelled',
    'multi-photo-preparation-filename', 'multi-photo-preparation-progress',
    'multi-photo-preparation-error', 'multi-photo-preparation-cancel',
    'multi-photo-preparation-back'
  ]) assert.match(indexHtml, new RegExp(`id=["']${id}["']`), id);
  const renderSource = sourceFunction(indexHtml, 'renderMultiPhotoPreparationProgress');
  assert.doesNotMatch(renderSource, /innerHTML/);
  assert.match(renderSource, /textContent/);
  assert.match(sourceFunction(indexHtml, 'isShortcutOverlayOpen'), /multi-photo-preparation-overlay/);
});

test('selection guard protects single-photo resources and snapshots only approved defaults', () => {
  const guard = sourceFunction(indexHtml, 'canStartMultiPhotoImport');
  for (const needle of [
    'canEdit()', 'state.upload.converting', 'state.upload.originalFile',
    'state.upload.uploadFile', 'state.upload.previewUrl', 'state.upload.presetApplying',
    'multiPhotoWorkflow.isBusy()', 'isCsvImportBusy()'
  ]) assert.ok(guard.includes(needle), needle);
  assert.equal(guard.includes('clearUploadPhotoState'), false);

  const guardContext = {
    state: { upload: {
      converting: false, originalFile: null, uploadFile: null, previewUrl: '', presetApplying: true
    } },
    uploadFolderState: { loading: false },
    canEdit: () => true,
    multiPhotoWorkflow: { isBusy: () => false },
    isCsvImportBusy: () => false
  };
  vm.runInNewContext(`${guard}\nthis.canStart = canStartMultiPhotoImport;`, guardContext);
  assert.equal(guardContext.canStart(), false);
  const singleFile = { name: 'single.jpg' };
  guardContext.state.upload = {
    converting: false, originalFile: singleFile, uploadFile: singleFile,
    previewUrl: 'blob:single-photo', presetApplying: false
  };
  assert.equal(guardContext.canStart(), false);
  assert.equal(guardContext.state.upload.originalFile, singleFile);
  assert.equal(guardContext.state.upload.uploadFile, singleFile);
  assert.equal(guardContext.state.upload.previewUrl, 'blob:single-photo');
  guardContext.state.upload = {
    converting: false, originalFile: null, uploadFile: null, previewUrl: '', presetApplying: false
  };
  guardContext.uploadFolderState.loading = true;
  assert.equal(guardContext.canStart(), false);
  guardContext.uploadFolderState.loading = false;
  guardContext.isCsvImportBusy = () => true;
  assert.equal(guardContext.canStart(), false);

  const context = {
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'photo' }],
    state: { appSettings: { rootFolderId: 'folder-snapshot' } }
  };
  vm.runInNewContext(
    `${sourceFunction(indexHtml, 'snapshotPhotoImportDefaults')}\nthis.snapshot = snapshotPhotoImportDefaults;`,
    context
  );
  const snapshot = context.snapshot();
  assert.deepEqual(JSON.parse(JSON.stringify(snapshot)), {
    tags: [], color: '#e53935', icon: 'photo', status: '未対応',
    targetFolderId: 'folder-snapshot'
  });
  assert.equal(Object.prototype.hasOwnProperty.call(snapshot, 'title'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(snapshot, 'description'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(snapshot, 'eventAt'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(snapshot, 'links'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(snapshot, 'lat'), false);
});

test('target folder navigation commits the ID and label only after a validated request succeeds', async () => {
  let resolveFolder;
  const folderPromise = new Promise((resolve) => { resolveFolder = resolve; });
  const elements = {
    'folder-list': { innerHTML: '', appendChild() {} },
    'folder-breadcrumb': { textContent: 'Root' },
    'folder-back-btn': { style: {} }
  };
  const uploadFolderState = {
    stack: [], currentFolderId: 'root-folder', currentFolderUrl: 'root-url',
    loading: false, requestToken: 0
  };
  const context = {
    uploadFolderState,
    state: { appSettings: { rootFolderId: 'root-folder', rootFolderUrl: 'root-url' } },
    document: {
      getElementById(id) { return elements[id]; },
      createElement() {
        return { addEventListener() {}, appendChild() {}, style: {}, className: '', innerHTML: '' };
      }
    },
    withGAS: () => folderPromise,
    withEditToken: (value) => value,
    refreshMultiPhotoButtonState() {},
    updateCurrentDriveLink() {},
    escHtml: (value) => value
  };
  vm.runInNewContext(
    `${sourceFunction(indexHtml, 'currentFolderIdOrRoot')}\n`
      + `async ${sourceFunction(indexHtml, 'loadUploadFolder')}\n`
      + 'this.current = currentFolderIdOrRoot; this.load = loadUploadFolder;',
    context
  );
  uploadFolderState.stack = [{ id: 'optimistic-child', name: 'Child' }];
  assert.equal(context.current(), 'root-folder');
  uploadFolderState.stack = [];
  const pending = context.load('child-folder', false, [{ id: 'child-folder', name: 'Child' }]);
  assert.equal(uploadFolderState.loading, true);
  assert.equal(uploadFolderState.currentFolderId, 'root-folder');
  assert.deepEqual(uploadFolderState.stack, []);
  resolveFolder({
    folderId: 'child-folder', folderUrl: 'child-url', folderName: 'Child', items: []
  });
  assert.equal(await pending, true);
  assert.equal(uploadFolderState.currentFolderId, 'child-folder');
  assert.deepEqual(uploadFolderState.stack, [{ id: 'child-folder', name: 'Child' }]);
  assert.equal(elements['folder-breadcrumb'].textContent, 'Child');
  assert.equal(uploadFolderState.loading, false);
});

test('production handlers recheck authorization, reset the multi input, and keep shared view untouched', () => {
  const handler = sourceFunction(indexHtml, 'handleMultiPhotoFilesSelected');
  assert.match(handler, /canStartMultiPhotoImport\(\)/);
  assert.match(handler, /input\.value = ''/);
  assert.match(handler, /snapshotPhotoImportDefaults\(\)/);
  assert.match(handler, /files\.length > ImportJobCore\.MAX_ITEMS/);
  assert.match(handler, /startPhotoImportFromFiles/);
  assert.equal(handler.includes('clearUploadPhotoState'), false);
  assert.equal(sharedHtml.includes('multi-photo-file-input'), false);
  assert.equal(sharedHtml.includes('MultiPhotoImportWorkflow'), false);
  assert.equal(sharedHtml.includes('saveImportPhotoItem'), false);
  assert.equal(sharedHtml.includes('saveImportPinItem'), false);
  assert.equal(sharedHtml.includes('ImportPinItemProcessor'), false);
});

test('browser navigation warns while route saving or either production import workflow is active', () => {
  assert.match(sourceFunction(indexHtml, 'hasPendingRouteMutationWork'), /hasRouteSaveInFlight\(\)/);
  assert.match(sourceFunction(indexHtml, 'hasPendingMutationWork'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'initializeApp'),
    /window\.addEventListener\('beforeunload',[\s\S]*?hasPendingMutationWork\(\)[\s\S]*?event\.returnValue = ''/);
});

test('single add and CSV launch are mutually exclusive with active production imports', () => {
  assert.match(sourceFunction(indexHtml, 'openUploadModal'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'saveNewPin'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'openCsvImportPreview'), /isMultiPhotoImportBusyState\(\)/);
  assert.match(indexHtml, /csv-import-button[\s\S]*?isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'openSettingsModal'), /isProductionImportBusy\(\)/);
  assert.match(indexHtml, /canExport:\s*function\(\)\s*\{\s*return !isProductionImportBusy\(\);\s*\}/);
});

test('upsertImportedPin replaces by id and refreshes map, panels, filters, and folder cache', () => {
  const calls = [];
  const state = { pins: [{ id: 'pin-1', title: 'old' }] };
  const context = {
    state,
    clonePin(pin) { return { ...pin, cloned: true }; },
    cachePinFolderUrl_(id, url) { calls.push(['cache', id, url]); },
    renderPins() { calls.push(['pins']); },
    renderSidePanel() { calls.push(['side']); },
    renderColorFilterUI() { calls.push(['color']); },
    renderIconFilterUI() { calls.push(['icon']); },
    renderTagFilterUI() { calls.push(['tag']); }
  };
  vm.runInNewContext(
    `${sourceFunction(indexHtml, 'upsertImportedPin')}\nthis.upsert = upsertImportedPin;`,
    context
  );
  context.upsert({ id: 'pin-1', title: 'new', folderUrl: 'folder-url' });
  assert.equal(state.pins.length, 1);
  assert.deepEqual(state.pins[0], { id: 'pin-1', title: 'new', folderUrl: 'folder-url', cloned: true });
  context.upsert({ id: 'pin-2', title: 'second', folderUrl: '' });
  assert.equal(state.pins.length, 2);
  assert.deepEqual(calls.slice(0, 6), [
    ['cache', 'pin-1', 'folder-url'], ['pins'], ['side'], ['color'], ['icon'], ['tag']
  ]);
});

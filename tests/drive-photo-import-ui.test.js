const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const {
  loadDriveClientModules, plain, descriptor, fileResponse
} = require('./drive-photo-import-client-harness');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

const DRIVE_IDS = [
  'drive-photo-import-overlay', 'drive-photo-import-title',
  'drive-photo-import-folder-name',
  'drive-photo-import-parent', 'drive-photo-import-folder-list',
  'drive-photo-import-photo-list', 'drive-photo-import-unsupported-count',
  'drive-photo-import-photo-count', 'drive-photo-import-photo-empty',
  'drive-photo-import-hide-imported',
  'drive-photo-import-selection-count', 'drive-photo-import-selection-size',
  'drive-photo-import-select-all', 'drive-photo-import-clear-selection',
  'drive-photo-import-confirm', 'drive-photo-import-cancel',
  'drive-photo-import-storage-note',
  'drive-photo-import-status', 'drive-photo-import-error',
  'drive-photo-import-progress', 'drive-photo-import-progress-current'
];

class FakeElement {
  constructor(tagName, id, owner) {
    this.tagName = String(tagName || 'div').toUpperCase();
    this.id = id || '';
    this.owner = owner;
    this.children = [];
    this.attributes = Object.create(null);
    this.listeners = Object.create(null);
    this.style = {};
    this.textContent = '';
    this.value = '';
    this.max = 0;
    this.disabled = false;
    this.checked = false;
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
  const documentApi = {
    activeElement: null,
    elements: Object.create(null),
    createElement(tagName) { return new FakeElement(tagName, '', documentApi); },
    getElementById(id) {
      if (!this.elements[id]) this.elements[id] = new FakeElement('div', id, documentApi);
      return this.elements[id];
    }
  };
  DRIVE_IDS.forEach((id) => documentApi.getElementById(id));
  return documentApi;
}

function folderResponse(overrides = {}) {
  return {
    ok: true,
    folder: { id: 'root_AAAAAAAAAAA', name: 'Root', isRoot: true },
    parent: null,
    folders: [{ id: 'child_AAAAAAAAAA', name: 'Child' }],
    photos: [descriptor()],
    ignoredUnsupportedFileCount: 0,
    counts: { folders: 1, photos: 1 },
    ...overrides
  };
}

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

class BrowserFile extends Blob {
  constructor(parts, name, options = {}) {
    super(parts, options);
    this.name = name;
    this.lastModified = options.lastModified;
  }
}

function createPicker(overrides = {}) {
  const { sourceCore, ui } = loadDriveClientModules();
  assert.ok(ui, 'Expected DrivePhotoImportUI');
  const documentApi = createDocument();
  const state = {};
  const calls = [];
  const controller = ui.create({
    state,
    documentApi,
    callGAS: async (method, payload) => {
      calls.push([method, payload]);
      return folderResponse();
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'send-only-token' }),
    sourceCore,
    loaderApi: { create() { throw new Error('loader not configured'); } },
    environment: {},
    canStart: () => true,
    getDefaults: () => ({
      tags: [], color: '#e53935', icon: 'photo', status: '', targetFolderId: 'target_AAAAAAAAA'
    }),
    onFilesReady: async () => {},
    openPicker: () => calls.push(['open']),
    closePicker: () => calls.push(['close']),
    returnToPhotoSource: (message) => calls.push(['return', message]),
    onBusy: () => {},
    ...overrides
  });
  return { controller, state, documentApi, calls, sourceCore };
}

test('production exposes separate local and Drive entries with an accessible source picker only in edit UI', () => {
  assert.match(indexHtml, /id="multi-photo-button"[\s\S]*?add-menu-choice-title">端末から写真を取込</);
  assert.match(indexHtml, /id="drive-photo-import-button"[\s\S]*?add-menu-choice-title">Driveから写真を取込</);
  assert.match(indexHtml, /body:not\(\.has-edit-token\)[^}]*#drive-photo-import-button/);
  assert.match(indexHtml, /body:not\(\.edit-mode\)[^}]*#drive-photo-import-button/);
  DRIVE_IDS.forEach((id) => assert.match(indexHtml, new RegExp(`id=["']${id}["']`), id));
  assert.match(indexHtml, /登録成功後、表示用の管理JPEGを作成し、元写真はoriginalフォルダへ移動します。元写真自体は削除・変換しません。/);
  assert.match(indexHtml, /取込待ちのフォルダ/);
  assert.doesNotMatch(indexHtml, /管理用コピーの保存先/);
  assert.match(indexHtml, /選択した写真を読み込む/);
  assert.doesNotMatch(indexHtml, /drive-photo-import-consent/);
  assert.doesNotMatch(indexHtml, /同意して確認へ進む/);
  assert.match(indexHtml, /id="drive-photo-import-title"[^>]*tabindex="-1"/);
  assert.match(indexHtml, /id="import-preview-title"[^>]*tabindex="-1"/);
  const controllerStart = indexHtml.indexOf('const drivePhotoImportController = DrivePhotoImportUI.create');
  const productionWiring = indexHtml.slice(
    controllerStart,
    indexHtml.indexOf('function handleMultiPhotoButtonClick', controllerStart)
  );
  assert.doesNotMatch(productionWiring, /drive-photo-import-(?:button|title)['"]\)\.focus\(\)/);
  assert.match(productionWiring, /returnToPhotoSource/);
  assert.doesNotMatch(productionWiring, /upload-overlay|upload-form-status|loadUploadInputPresets|returnToUpload/);
  const driveUiSource = indexHtml.slice(
    indexHtml.indexOf('const DrivePhotoImportUI ='),
    indexHtml.indexOf('const drivePhotoImportController =')
  );
  assert.match(driveUiSource, /resetFields\(true\);\s*render\(\);\s*openPicker\(\)/);
  assert.doesNotMatch(driveUiSource, /consent|agreeSelection|backFromConsent/);
  assert.doesNotMatch(driveUiSource, /returnToUpload/);
  assert.equal(sharedHtml.includes('drive-photo-import-button'), false);
  assert.equal(sharedHtml.includes('DrivePhotoImportUI'), false);
  assert.equal(sharedHtml.includes('listDrivePhotoImportFolder'), false);
  assert.equal(sharedHtml.includes('readDrivePhotoImportFile'), false);
});

test('production Drive environment preserves the browser receiver for atob', () => {
  const controllerStart = indexHtml.indexOf('const drivePhotoImportController = DrivePhotoImportUI.create');
  const productionWiring = indexHtml.slice(
    controllerStart,
    indexHtml.indexOf('function handleMultiPhotoButtonClick', controllerStart)
  );
  const atobInitializer = productionWiring.match(
    /environment:\s*\{\s*atob:\s*([\s\S]*?),\s*Uint8Array:/
  );
  assert.ok(atobInitializer, 'production Drive environment must define atob');
  const browserContext = {
    browserGlobal: true,
    decoded: '',
    atob(value) {
      if (!this.browserGlobal) throw new TypeError('Illegal invocation');
      return value === 'AQID' ? '\x01\x02\x03' : '';
    }
  };
  browserContext.window = browserContext;
  vm.runInNewContext(`
    const environment = { atob: ${atobInitializer[1]} };
    decoded = environment.atob('AQID');
  `, browserContext);
  assert.equal(browserContext.decoded, '\x01\x02\x03');
});

test('folder response is adopted only after validation and names render as text nodes without IDs', async () => {
  const hostile = folderResponse({
    folder: { id: 'root_AAAAAAAAAAA', name: '<img src=x onerror=1>', isRoot: true },
    folders: [{ id: 'child_AAAAAAAAAA', name: '<b>folder</b>' }],
    photos: [descriptor({ name: '<svg onload=1>.jpg' })]
  });
  const picker = createPicker({ callGAS: async () => hostile });
  await picker.controller.open();
  assert.equal(picker.state.folder.name, '<img src=x onerror=1>');
  assert.equal(picker.documentApi.getElementById('drive-photo-import-folder-name').textContent,
    '<img src=x onerror=1>');
  const folderButton = picker.documentApi.getElementById('drive-photo-import-folder-list').children[0];
  assert.equal(folderButton.children.some((child) => child.textContent === '<b>folder</b>'), true);
  assert.equal(folderButton.children.some((child) => child.textContent === 'child_AAAAAAAAAA'), false);
  const photoRow = picker.documentApi.getElementById('drive-photo-import-photo-list').children[0];
  assert.equal(photoRow.children.some((child) => child.textContent === '<svg onload=1>.jpg'), true);
});

test('imported filter defaults to all photos, badges imported photos, and selects only visible photos', async () => {
  const photos = [
    descriptor({ id: 'photo_new_AAAAAAAA', name: 'new.jpg', imported: false }),
    descriptor({ id: 'photo_done_AAAAAAA', name: 'done.jpg', imported: true }),
    descriptor({ id: 'photo_done_BBBBBBB', name: 'done-2.jpg', imported: true })
  ];
  const picker = createPicker({
    callGAS: async () => folderResponse({ photos, counts: { folders: 1, photos: 3 } })
  });
  await picker.controller.open();
  const list = picker.documentApi.getElementById('drive-photo-import-photo-list');
  const filter = picker.documentApi.getElementById('drive-photo-import-hide-imported');
  assert.equal(filter.checked, false);
  assert.equal(list.children.length, 3);
  assert.equal(list.children.flatMap((row) => row.children).filter((child) => child.textContent === '取込済み').length, 2);
  assert.equal(picker.documentApi.getElementById('drive-photo-import-photo-count').textContent,
    '写真3件・取込済み2件');

  picker.controller.toggle(photos[0].id, true);
  assert.equal(picker.controller.toggle(photos[1].id, true), false);
  assert.equal(list.children[1].children[0].disabled, true);
  filter.checked = true;
  filter.listeners.change();
  assert.equal(list.children.length, 1);
  assert.deepEqual(plain(picker.state.selectedIds), [photos[0].id]);
  assert.equal(picker.state.statusMessage, '');
  picker.controller.clearSelection();
  assert.equal(picker.controller.selectAll(), true);
  assert.deepEqual(plain(picker.state.selectedIds), [photos[0].id]);
  const renderedText = list.children.flatMap((row) => row.children)
    .map((child) => `${child.textContent} ${Object.values(child.attributes).join(' ')}`).join(' ');
  assert.equal(renderedText.includes(photos[0].id), false);
  assert.equal(renderedText.includes(photos[1].id), false);
});

test('imported filter survives folder navigation and resets after closing the picker', async () => {
  const picker = createPicker({
    callGAS: async (_method, payload) => payload.folderId
      ? folderResponse({
        folder: { id: 'child_AAAAAAAAAA', name: 'Child', isRoot: false },
        parent: { id: 'root_AAAAAAAAAAA', name: 'Root' },
        folders: [], photos: [descriptor({ id: 'child_photo_AAAAA', imported: true })],
        counts: { folders: 0, photos: 1 }
      })
      : folderResponse()
  });
  await picker.controller.open();
  const filter = picker.documentApi.getElementById('drive-photo-import-hide-imported');
  filter.checked = true;
  filter.listeners.change();
  await picker.controller.navigate('child_AAAAAAAAAA');
  assert.equal(picker.state.hideImported, true);
  assert.equal(filter.checked, true);
  picker.controller.cancel();
  await picker.controller.open();
  assert.equal(picker.state.hideImported, false);
  assert.equal(filter.checked, false);
});

test('validated Drive root id survives navigation while the current source folder changes', async () => {
  const picker = createPicker({
    callGAS: async (_method, payload) => payload.folderId
      ? folderResponse({
        folder: { id: 'child_AAAAAAAAAA', name: 'Child', isRoot: false },
        parent: { id: 'root_AAAAAAAAAAA', name: 'Root' },
        folders: [], photos: [descriptor({ id: 'child_photo_AAAAA' })],
        counts: { folders: 0, photos: 1 }
      })
      : folderResponse()
  });

  await picker.controller.open();
  assert.equal(picker.state.rootFolderId, 'root_AAAAAAAAAAA');
  assert.equal(picker.state.sourceFolderId, 'root_AAAAAAAAAAA');

  await picker.controller.navigate('child_AAAAAAAAAA');
  assert.equal(picker.state.sourceFolderId, 'child_AAAAAAAAAA');
  assert.equal(picker.state.rootFolderId, 'root_AAAAAAAAAAA');
});

test('selection is list ordered, permits one and twenty, and rejects the 21st and 100MB overflow atomically', async () => {
  const photos = Array.from({ length: 21 }, (_, index) => descriptor({
    id: `photo_${String(index).padStart(12, '0')}`,
    name: `${index}.jpg`,
    sizeBytes: index < 20 ? 1024 : 1
  }));
  const picker = createPicker({
    callGAS: async () => folderResponse({ photos, counts: { folders: 1, photos: 21 } })
  });
  await picker.controller.open();
  picker.controller.toggle(photos[2].id, true);
  picker.controller.toggle(photos[0].id, true);
  assert.deepEqual(plain(picker.state.selectedIds), [photos[0].id, photos[2].id]);
  photos.slice(0, 20).forEach((photo) => picker.controller.toggle(photo.id, true));
  assert.equal(picker.state.selectedIds.length, 20);
  picker.controller.toggle(photos[20].id, true);
  assert.equal(picker.state.selectedIds.length, 20);
  assert.equal(picker.state.errorCode, 'DRIVE_IMPORT_SELECTION_LIMIT_EXCEEDED');

  const large = Array.from({ length: 7 }, (_, index) => descriptor({
    id: `large_${String(index).padStart(12, '0')}`,
    name: `large-${index}.jpg`, sizeBytes: 15 * 1024 * 1024
  }));
  const largePicker = createPicker({
    callGAS: async () => folderResponse({ photos: large, counts: { folders: 1, photos: 7 } })
  });
  await largePicker.controller.open();
  large.slice(0, 6).forEach((photo) => largePicker.controller.toggle(photo.id, true));
  largePicker.controller.toggle(large[6].id, true);
  assert.equal(largePicker.state.selectedIds.length, 6);
  assert.equal(largePicker.state.errorCode, 'DRIVE_IMPORT_SELECTION_TOO_LARGE');
});

test('existing-pin attach mode accepts exactly one Drive photo without partial multi-selection', async () => {
  const first = descriptor({ id: 'photo_ATTACHFIRSTA', name: 'first.jpg' });
  const second = descriptor({ id: 'photo_ATTACHSECONDA', name: 'second.jpg' });
  const picker = createPicker({
    getSelectionLimit: () => 1,
    callGAS: async () => folderResponse({
      folders: [], photos: [first, second], counts: { folders: 0, photos: 2 }
    })
  });
  await picker.controller.open();

  assert.equal(picker.controller.toggle(first.id, true), true);
  assert.equal(picker.controller.toggle(second.id, true), false);
  assert.deepEqual(plain(picker.state.selectedIds), [first.id]);
  assert.equal(picker.state.errorCode, 'DRIVE_IMPORT_SELECTION_INVALID');
  assert.equal(picker.controller.selectAll(), false);
  assert.deepEqual(plain(picker.state.selectedIds), [first.id]);
});

test('select all is atomic, clear uses a null-prototype map, and navigation clears visible selection', async () => {
  const photos = Array.from({ length: 20 }, (_, index) => descriptor({
    id: `photo_${String(index).padStart(12, '0')}`,
    name: `${index}.jpg`, sizeBytes: 5 * 1024 * 1024
  }));
  const picker = createPicker({
    callGAS: async (_method, payload) => payload.folderId
      ? folderResponse({
        folder: { id: 'child_AAAAAAAAAA', name: 'Child', isRoot: false },
        parent: { id: 'root_AAAAAAAAAAA', name: 'Root' }, folders: [], photos: [],
        counts: { folders: 0, photos: 0 }
      })
      : folderResponse({ photos, counts: { folders: 1, photos: 20 } })
  });
  await picker.controller.open();
  assert.equal(picker.controller.selectAll(), true);
  assert.equal(picker.state.selectedIds.length, 20);
  picker.controller.clearSelection();
  assert.equal(Object.getPrototypeOf(picker.state.selectedById), null);
  picker.controller.toggle(photos[0].id, true);
  await picker.controller.navigate('child_AAAAAAAAAA');
  assert.equal(picker.state.selectedIds.length, 0);
  assert.match(picker.state.statusMessage, /選択を解除/);
  assert.equal(picker.state.sourceFolderId, 'child_AAAAAAAAAA');
});

test('failed all-select preserves the prior selection and current folder failure preserves the validated listing', async () => {
  const photos = Array.from({ length: 21 }, (_, index) => descriptor({
    id: `photo_${String(index).padStart(12, '0')}`,
    name: `${index}.jpg`, sizeBytes: 1024
  }));
  let calls = 0;
  const picker = createPicker({
    callGAS: async () => {
      calls += 1;
      if (calls === 1) return folderResponse({ photos, counts: { folders: 1, photos: 21 } });
      throw Object.assign(new Error('private Drive detail'), { code: 'DRIVE_IMPORT_FOLDER_NOT_FOUND' });
    }
  });
  await picker.controller.open();
  picker.controller.toggle(photos[3].id, true);
  assert.equal(picker.controller.selectAll(), false);
  assert.deepEqual(plain(picker.state.selectedIds), [photos[3].id]);
  const previousFolder = picker.state.folder;
  const previousFolders = picker.state.folders;
  const previousPhotos = picker.state.photos;
  assert.equal(await picker.controller.navigate('child_AAAAAAAAAA'), false);
  assert.equal(picker.state.folder, previousFolder);
  assert.equal(picker.state.folders, previousFolders);
  assert.equal(picker.state.photos, previousPhotos);
  assert.deepEqual(plain(picker.state.selectedIds), [photos[3].id]);
  assert.equal(picker.state.errorCode, 'DRIVE_IMPORT_FOLDER_NOT_FOUND');
  assert.equal(picker.documentApi.activeElement,
    picker.documentApi.getElementById('drive-photo-import-error'));
});

test('newer navigation supersedes an in-flight request and only the adopted folder clears selection', async () => {
  const rootRequest = deferred();
  const childRequest = deferred();
  const picker = createPicker({
    callGAS: (_method, payload) => payload.folderId ? childRequest.promise : rootRequest.promise
  });
  const opening = picker.controller.open();
  const navigating = picker.controller.navigate('child_AAAAAAAAAA');
  rootRequest.resolve(folderResponse());
  await opening;
  childRequest.resolve(folderResponse({
    folder: { id: 'child_AAAAAAAAAA', name: 'Child', isRoot: false },
    parent: { id: 'root_AAAAAAAAAAA', name: 'Root' },
    folders: [], photos: [], counts: { folders: 0, photos: 0 }
  }));
  assert.equal(await navigating, true);
  assert.equal(picker.state.folder.id, 'child_AAAAAAAAAA');
  assert.equal(picker.state.errorCode, '');
});

test('checkbox rerender restores keyboard focus to the same photo control', async () => {
  const picker = createPicker();
  await picker.controller.open();
  const oldCheckbox = picker.documentApi.getElementById('drive-photo-import-photo-list').children[0].children[0];
  oldCheckbox.focus();
  oldCheckbox.checked = true;
  oldCheckbox.listeners.change();
  const newCheckbox = picker.documentApi.getElementById('drive-photo-import-photo-list').children[0].children[0];
  assert.notEqual(newCheckbox, oldCheckbox);
  assert.equal(picker.documentApi.activeElement, newCheckbox);
  assert.equal(newCheckbox.checked, true);
});

test('close and reopen ignore stale folder success and stale errors without replacing the new session', async () => {
  const oldRequest = deferred();
  const newRequest = deferred();
  let requestCount = 0;
  const picker = createPicker({
    callGAS: () => (++requestCount === 1 ? oldRequest.promise : newRequest.promise)
  });
  const oldOpen = picker.controller.open();
  picker.controller.cancel();
  const newOpen = picker.controller.open();
  newRequest.resolve(folderResponse({
    folder: { id: 'newroot_AAAAAAAAAA', name: 'New root', isRoot: true }
  }));
  await newOpen;
  oldRequest.resolve(folderResponse({
    folder: { id: 'oldroot_AAAAAAAAAA', name: 'Old root', isRoot: true }
  }));
  await oldOpen;
  assert.equal(picker.state.folder.name, 'New root');
  assert.equal(picker.state.errorCode, '');

  const staleFailure = deferred();
  const latest = deferred();
  requestCount = 0;
  const other = createPicker({
    callGAS: () => (++requestCount === 1 ? staleFailure.promise : latest.promise)
  });
  const staleOpen = other.controller.open();
  other.controller.cancel();
  const latestOpen = other.controller.open();
  latest.resolve(folderResponse());
  await latestOpen;
  staleFailure.reject(Object.assign(new Error('private server detail'), { code: 'DRIVE_IMPORT_ACCESS_DENIED' }));
  await staleOpen;
  assert.equal(other.state.errorCode, '');
  assert.equal(other.state.folder.name, 'Root');
});

test('selection confirmation starts the loader directly and blocks a double confirmation', async () => {
  const loading = deferred();
  let starts = 0;
  const picker = createPicker({
    loaderApi: { create() {
      return {
        start() { starts += 1; return loading.promise; },
        cancel() {}, release() {}, isRunning() { return starts > 0; }
      };
    } },
    onFilesReady: async () => ({ id: 'job' })
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  const firstConfirmation = picker.controller.confirm();
  await Promise.resolve();
  assert.equal(starts, 1);
  assert.equal(await picker.controller.confirm(), false);
  loading.resolve([{ name: 'photo.jpg' }]);
  assert.equal(await firstConfirmation, true);
  assert.equal(starts, 1);
});

test('Drive cancel closes the picker before returning to photo source selection', async () => {
  const picker = createPicker();
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);

  assert.equal(picker.controller.cancel(), true);

  assert.equal(picker.state.open, false);
  assert.deepEqual(picker.calls.slice(-2), [
    ['close'],
    ['return', undefined]
  ]);
});

test('Drive read failure keeps picker and selection open without returning to another screen', async () => {
  const readError = Object.assign(new Error('private read detail'), {
    code: 'DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED'
  });
  const picker = createPicker({
    loaderApi: { create() {
      return {
        start: async () => { throw readError; },
        cancel() {}, release() {}, isRunning() { return false; }
      };
    } }
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await picker.controller.confirm(), false);

  assert.equal(picker.state.open, true);
  assert.equal(picker.state.fileLoading, false);
  assert.equal(picker.state.errorCode, 'DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED');
  assert.deepEqual(plain(picker.state.selectedIds), ['photo_AAAAAAAAAAA']);
  assert.equal(picker.calls.filter(([name]) => name === 'close').length, 0);
  assert.equal(picker.calls.filter(([name]) => name === 'return').length, 0);
});

test('Drive response failure keeps selection and renders only its safe diagnostic stage', async () => {
  const readError = Object.assign(new Error('private AQID photo_AAAAAAAAAAA'), {
    code: 'DRIVE_IMPORT_RESPONSE_INVALID',
    diagnosticStage: 'base64_decode_failed',
    base64: 'AQID'
  });
  const picker = createPicker({
    loaderApi: { create() {
      return {
        start: async () => { throw readError; },
        cancel() {}, release() {}, isRunning() { return false; }
      };
    } }
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await picker.controller.confirm(), false);

  assert.equal(picker.state.open, true);
  assert.equal(picker.state.errorCode, 'DRIVE_IMPORT_RESPONSE_INVALID');
  assert.equal(picker.state.errorDiagnosticStage, 'base64_decode_failed');
  assert.deepEqual(plain(picker.state.selectedIds), ['photo_AAAAAAAAAAA']);
  assert.equal(
    picker.documentApi.getElementById('drive-photo-import-error').textContent,
    'Driveの内容を確認できませんでした。もう一度開いてください。\n診断コード: base64_decode_failed'
  );
  assert.doesNotMatch(JSON.stringify(picker.state), /AQID|private/);
  assert.equal(picker.calls.filter(([name]) => name === 'close').length, 0);
  assert.equal(picker.calls.filter(([name]) => name === 'return').length, 0);
});

test('Drive response failure discards an unrecognized diagnostic stage', async () => {
  const picker = createPicker({
    loaderApi: { create() {
      return {
        start: async () => {
          throw Object.assign(new Error('private'), {
            code: 'DRIVE_IMPORT_RESPONSE_INVALID',
            diagnosticStage: 'base64_decode_failed\nAQID'
          });
        },
        cancel() {}, release() {}, isRunning() { return false; }
      };
    } }
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await picker.controller.confirm(), false);

  assert.equal(picker.state.errorDiagnosticStage, '');
  assert.equal(
    picker.documentApi.getElementById('drive-photo-import-error').textContent,
    'Driveの内容を確認できませんでした。もう一度開いてください。'
  );
});

test('cancelled loader settlement cannot clean or reopen over a newer picker session', async () => {
  const oldLoad = deferred();
  const handedOff = [];
  let loaderNumber = 0;
  const picker = createPicker({
    loaderApi: { create() {
      loaderNumber += 1;
      if (loaderNumber === 1) {
        return {
          start: () => oldLoad.promise,
          cancel() {}, release() {}, isRunning() { return true; }
        };
      }
      return {
        start: async (values) => values.map((value) => ({ name: value.name })),
        cancel() {}, release() {}, isRunning() { return false; }
      };
    } },
    onFilesReady: async (files) => {
      handedOff.push(files.map((file) => file.name));
      return { id: 'new-job' };
    }
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  const oldConfirm = picker.controller.confirm();
  await Promise.resolve();
  picker.controller.cancel();
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await picker.controller.confirm(), true);
  assert.deepEqual(handedOff, [['photo.jpg']]);
  const settledState = JSON.stringify(picker.state);
  oldLoad.resolve([{ name: 'stale.jpg' }]);
  assert.equal(await oldConfirm, false);
  assert.equal(JSON.stringify(picker.state), settledState);
  assert.deepEqual(handedOff, [['photo.jpg']]);
  assert.equal(picker.state.open, false);
  assert.equal(picker.state.fileLoading, false);
});

test('handoff rejection releases Drive ownership and returns with a safe actionable message', async () => {
  const picker = createPicker({
    loaderApi: { create() {
      return {
        start: async (values) => values.map((value) => ({ name: value.name })),
        cancel() {}, release() {}, isRunning() { return false; }
      };
    } },
    onFilesReady: async () => { throw new Error('private handoff failure'); }
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await picker.controller.confirm(), false);
  assert.equal(picker.state.open, false);
  assert.equal(picker.state.handingOff, false);
  assert.equal(picker.state.loader, null);
  assert.equal(picker.calls.filter(([name]) => name === 'return').length, 1);
  assert.deepEqual(picker.calls.find(([name]) => name === 'return'), [
    'return', 'Drive写真を複数写真の確認画面へ渡せませんでした。もう一度選択してください。'
  ]);
});

test('missing Drive destination returns to photo source selection with the dedicated safe message', async () => {
  const picker = createPicker({
    loaderApi: { create() {
      return {
        start: async (values) => values.map((value) => ({ name: value.name })),
        cancel() {}, release() {}, isRunning() { return false; }
      };
    } },
    onFilesReady: async () => {
      const error = new Error('private missing destination detail');
      error.code = 'DRIVE_IMPORT_TARGET_FOLDER_REQUIRED';
      throw error;
    }
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await picker.controller.confirm(), false);
  assert.equal(picker.calls.filter(([name]) => name === 'return').length, 1);
  assert.deepEqual(picker.calls.find(([name]) => name === 'return'), [
    'return', 'Driveの保存先フォルダを確認できませんでした。もう一度開いてください。'
  ]);
});

test('every falsy handoff settlement is a failure with one safe rollback', async () => {
  for (const settlement of [null, false, undefined, 0, '', NaN]) {
    const picker = createPicker({
      loaderApi: { create() {
        return {
          start: async (values) => values.map((value) => ({ name: value.name })),
          cancel() {}, release() {}, isRunning() { return false; }
        };
      } },
      onFilesReady: async () => settlement
    });
    await picker.controller.open();
    picker.controller.toggle('photo_AAAAAAAAAAA', true);
    assert.equal(await picker.controller.confirm(), false);
    assert.equal(picker.state.open, false);
    assert.equal(picker.state.handingOff, false);
    assert.equal(picker.state.loader, null);
    assert.equal(picker.calls.filter(([name]) => name === 'return').length, 1);
    assert.equal(picker.calls.find(([name]) => name === 'return')[1],
      'Drive写真を複数写真の確認画面へ渡せませんでした。もう一度選択してください。');
  }
});

test('UI accepts only bounded loader progress that names the selected descriptor', async () => {
  let picker;
  const seen = [];
  picker = createPicker({
    loaderApi: { create(options) {
      return {
        async start(values) {
          options.onProgress({ eventType: 'private', total: -1, completed: 99,
            currentIndex: 0, currentName: 'secret', base64: 'AQID' });
          options.onProgress({ eventType: 'item-start', total: values.length, completed: 0,
            currentIndex: 0, currentName: 'wrong.jpg' });
          options.onProgress({ eventType: 'item-start', total: values.length, completed: 0,
            currentIndex: 0, currentName: values[0].name, arbitrary: 'secret' });
          return values.map((value) => ({ name: value.name }));
        },
        cancel() {}, release() {}, isRunning() { return false; }
      };
    } },
    onFilesReady: async () => ({ id: 'job-1' }),
    onBusy() {
      if (picker && picker.state.progress) seen.push({ ...picker.state.progress });
    }
  });
  await picker.controller.open();
  picker.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await picker.controller.confirm(), true);
  assert.equal(seen.some((progress) => progress.eventType === 'private'), false);
  assert.equal(seen.some((progress) => progress.currentName === 'wrong.jpg'), false);
  const valid = seen.find((progress) => progress.eventType === 'item-start');
  assert.deepEqual(Object.keys(valid).sort(),
    ['completed', 'currentIndex', 'currentName', 'eventType', 'total']);
  assert.equal(JSON.stringify(seen).includes('secret'), false);
});

test('Classroom edit access sends a fresh token for list read and save while view and share access call no Drive API', async () => {
  for (const access of [
    { hasEditToken: false, editMode: true, shareMode: false },
    { hasEditToken: true, editMode: false, shareMode: false },
    { hasEditToken: true, editMode: true, shareMode: true }
  ]) {
    let calls = 0;
    const denied = createPicker({
      canStart: () => access.hasEditToken && access.editMode && !access.shareMode,
      callGAS: async () => { calls += 1; return folderResponse(); }
    });
    assert.equal(await denied.controller.open(), false);
    assert.equal(calls, 0);
  }

  const access = { hasEditToken: true, editMode: true, shareMode: false };
  const calls = [];
  let tokenNumber = 0;
  const { loader: loaderApi } = loadDriveClientModules();
  const withFreshToken = (payload) => ({ ...payload, __editToken: `token-${++tokenNumber}` });
  async function callGAS(method, payload) {
    calls.push([method, { ...payload }]);
    if (method === 'listDrivePhotoImportFolder') return folderResponse();
    if (method === 'readDrivePhotoImportFile') return fileResponse(descriptor());
    if (method === 'saveImportPhotoItem') return { ok: true, pin: { id: 'pin-1' } };
    throw new Error('unexpected method');
  }
  const allowed = createPicker({
    canStart: () => access.hasEditToken && access.editMode && !access.shareMode,
    callGAS,
    withEditToken: withFreshToken,
    loaderApi,
    environment: {
      atob(value) { return Buffer.from(value, 'base64').toString('binary'); },
      Uint8Array, Blob, File: BrowserFile
    },
    onFilesReady: async (files) => {
      assert.equal(files.length, 1);
      assert.equal(files[0] instanceof BrowserFile, true);
      assert.equal(Object.hasOwn(files[0], '__editToken'), false);
      return callGAS('saveImportPhotoItem', withFreshToken({ filename: files[0].name }));
    }
  });
  await allowed.controller.open();
  allowed.controller.toggle('photo_AAAAAAAAAAA', true);
  assert.equal(await allowed.controller.confirm(), true);
  assert.deepEqual(calls.map(([method]) => method), [
    'listDrivePhotoImportFolder', 'readDrivePhotoImportFile', 'saveImportPhotoItem'
  ]);
  assert.deepEqual(calls.map(([, payload]) => payload.__editToken), ['token-1', 'token-2', 'token-3']);
  assert.equal(JSON.stringify(allowed.state).includes('token-'), false);
  assert.equal(Object.hasOwn(descriptor(), '__editToken'), false);
});

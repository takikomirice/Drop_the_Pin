const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function extractFunction(source, name) {
  const functionStart = source.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `Expected function ${name} to exist`);
  const start = source.slice(Math.max(0, functionStart - 6), functionStart) === 'async '
    ? functionStart - 6
    : functionStart;
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, index + 1);
  }
  assert.fail(`Could not parse function ${name}`);
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

function createElement() {
  let html = '';
  const element = {
    style: {},
    children: [],
    appendChild(child) {
      this.children.push(child);
    }
  };
  Object.defineProperty(element, 'innerHTML', {
    get: () => html,
    set(value) {
      html = String(value);
      element.children.length = 0;
    }
  });
  return element;
}

function createHarness(options = {}) {
  const detailDrive = createElement();
  const editorDrive = createElement();
  const uploadSubmit = { disabled: false };
  const calls = [];
  const hints = [];
  const pins = options.pins || [
    { id: 'pin-a', fileId: 'file-a', folderUrl: '' },
    { id: 'pin-b', fileId: 'file-b', folderUrl: '' },
    { id: 'pin-empty', fileId: '', folderUrl: '' }
  ];
  const state = {
    editMode: options.editMode !== false,
    shareMode: options.shareMode === true,
    pins,
    shownPinId: options.shownPinId || null,
    editingPinId: options.editingPinId || null,
    pinDriveMetaCache: Object.create(null),
    pinDriveMetaRequests: Object.create(null)
  };
  const context = {
    Promise,
    Date,
    console: { error() {} },
    state,
    canEdit: () => options.hasEditToken !== false && state.editMode && !state.shareMode,
    getPinById: (pinId) => state.pins.find((pin) => pin.id === pinId) || null,
    withEditToken: (payload) => ({ ...payload, __editToken: 'test-token' }),
    withGAS: (methodName, payload) => {
      calls.push({ methodName, payload });
      return options.withGAS ? options.withGAS(methodName, payload) : Promise.resolve({ ok: true, folderUrl: '' });
    },
    escHtml: (value) => String(value),
    safeColor: (value) => String(value || '#e53935'),
    normalizeIcon: (value) => String(value || 'default'),
    showHint: (message) => hints.push(message),
    resizeWithOrientation: () => Promise.reject(new Error('unexpected image resize')),
    currentFolderIdOrRoot: () => 'target-folder',
    closeOverlay() {},
    renderPins() {},
    renderSidePanel() {},
    updateUnplacedBadge() {},
    renderColorFilterUI() {},
    renderIconFilterUI() {},
    renderTagFilterUI() {},
    clearUploadPhotoState() {},
    document: {
      getElementById: (id) => {
        if (id === 'pin-detail-drive') return detailDrive;
        if (id === 'edit-drive-area') return editorDrive;
        if (id === 'upload-submit') return uploadSubmit;
        return null;
      },
      createElement: () => ({})
    }
  };
  const functionNames = [
    'cachePinFolderUrl_',
    'clonePin',
    'renderPinDetailDrive_',
    'renderPinEditorDrive_',
    'requestPinDriveMeta_',
    'loadPinDriveMetaForCurrentViews_',
    'saveNewPin'
  ];
  const functionSource = functionNames.map((name) => extractFunction(indexHtml, name)).join('\n');
  vm.runInNewContext(`${functionSource}
globalThis.__lazyApi = {
  cachePinFolderUrl_,
  renderPinDetailDrive_,
  renderPinEditorDrive_,
  requestPinDriveMeta_,
  loadPinDriveMetaForCurrentViews_,
  saveNewPin
};`, context);
  return { api: context.__lazyApi, calls, detailDrive, editorDrive, hints, pins, state };
}

function renderedHref(element) {
  if (element.children.length > 0) return element.children[0].href || '';
  const match = element.innerHTML.match(/href="([^"]+)"/);
  return match ? match[1] : '';
}

test('view-only and shared modes never request pin Drive metadata', async () => {
  let harness = createHarness({ hasEditToken: false, shownPinId: 'pin-a' });
  await harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  assert.equal(harness.calls.length, 0);

  harness = createHarness({ shareMode: true, shownPinId: 'pin-a' });
  await harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  assert.equal(harness.calls.length, 0);
});

test('simultaneous displays of one pin share a single GAS request', async () => {
  const pending = deferred();
  const harness = createHarness({
    shownPinId: 'pin-a',
    editingPinId: 'pin-a',
    withGAS: () => pending.promise
  });

  const detailRequest = harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  const editorRequest = harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  assert.equal(harness.calls.length, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(harness.calls[0])), {
    methodName: 'getPinDriveMeta',
    payload: { pinId: 'pin-a', __editToken: 'test-token' }
  });

  pending.resolve({ ok: true, folderUrl: 'https://drive.google.com/drive/folders/a' });
  await Promise.all([detailRequest, editorRequest]);
  assert.equal(renderedHref(harness.detailDrive), 'https://drive.google.com/drive/folders/a');
  assert.equal(renderedHref(harness.editorDrive), 'https://drive.google.com/drive/folders/a');
});

test('normal empty results are cached without a GAS request for photo-less pins', async () => {
  const harness = createHarness({ shownPinId: 'pin-empty' });
  const emptyPin = harness.pins[2];

  await harness.api.loadPinDriveMetaForCurrentViews_(emptyPin);
  await harness.api.loadPinDriveMetaForCurrentViews_(emptyPin);

  assert.equal(harness.calls.length, 0);
  assert.equal(Object.prototype.hasOwnProperty.call(harness.state.pinDriveMetaCache, 'pin-empty'), true);
  assert.equal(harness.state.pinDriveMetaCache['pin-empty'], '');
});

test('Drive and transport failures are not cached and retry on the next display', async () => {
  const responses = [
    () => Promise.resolve({ ok: false, folderUrl: '', error: 'drive_meta_unavailable' }),
    () => Promise.reject(new Error('temporary transport failure')),
    () => Promise.resolve({ ok: true, folderUrl: 'https://drive.google.com/drive/folders/recovered' })
  ];
  const harness = createHarness({
    shownPinId: 'pin-a',
    withGAS: () => responses.shift()()
  });

  await harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  assert.equal(Object.prototype.hasOwnProperty.call(harness.state.pinDriveMetaCache, 'pin-a'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(harness.state.pinDriveMetaRequests, 'pin-a'), false);

  await harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  assert.equal(Object.prototype.hasOwnProperty.call(harness.state.pinDriveMetaCache, 'pin-a'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(harness.state.pinDriveMetaRequests, 'pin-a'), false);

  await harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  assert.equal(harness.calls.length, 3);
  assert.equal(harness.state.pinDriveMetaCache['pin-a'], 'https://drive.google.com/drive/folders/recovered');
});

test('a stale pin A response never overwrites pin B detail DOM', async () => {
  const pending = deferred();
  const harness = createHarness({ shownPinId: 'pin-a', withGAS: () => pending.promise });

  const request = harness.api.loadPinDriveMetaForCurrentViews_(harness.pins[0]);
  harness.state.shownPinId = 'pin-b';
  harness.api.cachePinFolderUrl_('pin-b', 'https://drive.google.com/drive/folders/b');
  harness.api.renderPinDetailDrive_(harness.pins[1]);
  pending.resolve({ ok: true, folderUrl: 'https://drive.google.com/drive/folders/a' });
  await request;

  assert.equal(renderedHref(harness.detailDrive), 'https://drive.google.com/drive/folders/b');
  assert.equal(harness.pins[0].folderUrl, 'https://drive.google.com/drive/folders/a');
});

test('the map goto action clears the shown pin before closing its detail overlay', () => {
  const openPinDetailSource = extractFunction(indexHtml, 'openPinDetail');
  const gotoHandlerStart = openPinDetailSource.indexOf('gotoBtn.onclick');
  const gotoHandlerEnd = openPinDetailSource.indexOf('};', gotoHandlerStart);
  const gotoHandler = openPinDetailSource.slice(gotoHandlerStart, gotoHandlerEnd);

  assert.match(gotoHandler, /state\.shownPinId = null;[\s\S]*closeOverlay\('pin-detail-overlay'\)/);
});

test('a new pin cache seed displays its known folder without another Drive request', async () => {
  const newPin = { id: 'pin-new', fileId: 'file-new', folderUrl: '' };
  const harness = createHarness({ pins: [newPin], shownPinId: 'pin-new' });

  harness.api.cachePinFolderUrl_('pin-new', 'https://drive.google.com/drive/folders/new');
  await harness.api.loadPinDriveMetaForCurrentViews_(newPin);

  assert.equal(harness.calls.length, 0);
  assert.equal(renderedHref(harness.detailDrive), 'https://drive.google.com/drive/folders/new');
});

test('saveNewPin seeds its returned folder URL before the new pin is displayed', async () => {
  const harness = createHarness({
    pins: [],
    withGAS: (methodName) => {
      assert.equal(methodName, 'saveMapData');
      return Promise.resolve({
        ok: true,
        id: 'pin-saved',
        fileId: 'file-saved',
        imageUrl: 'https://example.com/saved.jpg',
        folderUrl: 'https://drive.google.com/drive/folders/saved'
      });
    }
  });
  const draft = {
    file: null,
    title: 'Saved pin',
    description: '',
    eventAt: '',
    color: '#e53935',
    icon: 'default',
    links: [],
    status: '未対応',
    tags: []
  };

  await harness.api.saveNewPin(draft, { lat: 35, lng: 139 });
  harness.state.shownPinId = 'pin-saved';
  await harness.api.loadPinDriveMetaForCurrentViews_(harness.state.pins[0]);

  assert.equal(harness.calls.length, 1);
  assert.equal(harness.state.pinDriveMetaCache['pin-saved'], 'https://drive.google.com/drive/folders/saved');
  assert.equal(renderedHref(harness.detailDrive), 'https://drive.google.com/drive/folders/saved');
});

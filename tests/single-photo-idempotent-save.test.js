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

function createSaveHarness(overrides = {}) {
  const calls = [];
  const clears = [];
  let uuidCalls = 0;
  let resizeCalls = 0;
  const button = { disabled: false };
  const state = {
    pins: [],
    upload: {
      saving: false,
      saveError: '',
      photoSaveIdentity: null,
      submittedPhotoPayload: null,
      savePromise: null
    }
  };
  const context = {
    Promise,
    Date,
    Math,
    Object,
    Array,
    String,
    state,
    window: {
      crypto: {
        randomUUID() {
          uuidCalls += 1;
          return 'stable-selection-uuid';
        }
      }
    },
    document: { getElementById: () => button },
    isProductionImportBusy: () => false,
    resizeWithOrientation: async () => {
      resizeCalls += 1;
      return 'data:image/jpeg;base64,/9j/AA==';
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'stable-edit-token' }),
    withGAS: async (method, payload) => {
      calls.push({ method, payload });
      return {
        ok: true,
        deduplicated: false,
        pin: {
          id: 'pin-1', title: payload.title, description: payload.description,
          lat: payload.lat, lng: payload.lng, color: payload.color, icon: payload.icon,
          imageUrl: 'image', fileId: 'file-1', timestamp: 'saved', eventAt: payload.eventAt,
          updatedAt: '', links: payload.links, folderUrl: 'folder', status: payload.status,
          tags: payload.tags
        }
      };
    },
    currentFolderIdOrRoot: () => 'folder-1',
    clonePin: (pin) => ({ ...pin }),
    cachePinFolderUrl_() {},
    clearUploadPhotoState(options) {
      clears.push(options || null);
      if (options && options.afterSave) {
        state.upload.photoSaveIdentity = null;
        state.upload.submittedPhotoPayload = null;
      }
    },
    closeOverlay() {},
    renderPins() {},
    renderSidePanel() {},
    renderColorFilterUI() {},
    renderIconFilterUI() {},
    renderTagFilterUI() {},
    refreshUploadFormState() {},
    upsertImportedPin(pin) {
      const index = state.pins.findIndex((existing) => existing.id === pin.id);
      if (index === -1) state.pins.push({ ...pin });
      else state.pins[index] = { ...pin };
      return pin;
    },
    ...overrides
  };
  const names = [
    'createSinglePhotoSaveIdentity',
    'ensureSinglePhotoSaveIdentity',
    'createSinglePhotoSavePayload',
    'saveNewPin'
  ];
  vm.runInNewContext(
    `${names.map((name) => extractFunction(indexHtml, name)).join('\n')}
globalThis.__save = saveNewPin;`,
    context
  );
  return {
    context,
    state,
    calls,
    clears,
    button,
    getUuidCalls: () => uuidCalls,
    getResizeCalls: () => resizeCalls
  };
}

function photoDraft() {
  return {
    file: { name: 'photo.jpg', type: 'image/jpeg' },
    filename: 'photo.jpg',
    title: '固定タイトル',
    description: '固定説明',
    eventAt: '2026-07-13T12:00',
    color: '#e53935',
    icon: 'photo',
    links: ['https://example.com/original'],
    status: '未対応',
    tags: ['固定']
  };
}

test('single-photo explicit retry keeps identifiers, edit token, and the first payload unchanged', async () => {
  let attempt = 0;
  const harness = createSaveHarness({
    withGAS: async (method, payload) => {
      harness.calls.push({ method, payload });
      attempt += 1;
      if (attempt === 1) throw new Error('response lost');
      return {
        ok: true,
        deduplicated: true,
        pin: {
          id: 'pin-1', title: payload.title, description: payload.description,
          lat: payload.lat, lng: payload.lng, color: payload.color, icon: payload.icon,
          imageUrl: 'image', fileId: 'file-1', timestamp: 'saved', eventAt: payload.eventAt,
          updatedAt: '', links: payload.links, folderUrl: 'folder', status: payload.status,
          tags: payload.tags
        }
      };
    }
  });
  const draft = photoDraft();
  const firstCoords = { lat: 35.25, lng: 139.5 };
  await assert.rejects(harness.context.__save(draft, firstCoords), /response lost/);
  const retainedPayload = harness.state.upload.submittedPhotoPayload;
  assert.ok(retainedPayload);
  assert.equal(harness.clears.length, 0);

  draft.title = '変更後タイトル';
  draft.links.push('https://example.com/changed');
  await harness.context.__save(draft, { lat: 1, lng: 2 });

  assert.equal(harness.calls.length, 2);
  assert.equal(harness.calls[0].method, 'saveImportPhotoItem');
  assert.equal(harness.calls[1].method, 'saveImportPhotoItem');
  assert.equal(harness.calls[0].payload, harness.calls[1].payload);
  assert.equal(harness.calls[0].payload, retainedPayload);
  assert.equal(harness.calls[1].payload.title, '固定タイトル');
  assert.deepEqual(Array.from(harness.calls[1].payload.links), ['https://example.com/original']);
  assert.equal(harness.calls[1].payload.lat, 35.25);
  assert.equal(harness.calls[1].payload.lng, 139.5);
  assert.equal(harness.calls[1].payload.__editToken, 'stable-edit-token');
  assert.equal(harness.calls[0].payload.idempotencyKey,
    `${harness.calls[0].payload.jobId}:${harness.calls[0].payload.itemId}`);
  assert.equal(harness.getUuidCalls(), 1);
  assert.equal(harness.getResizeCalls(), 1);
  assert.equal(harness.state.pins.length, 1);
  assert.equal(harness.state.pins[0].id, 'pin-1');
  assert.equal(harness.clears.length, 1);
  assert.equal(harness.clears[0].afterSave, true);
});

test('parallel double save shares one photo request and applies the pin once', async () => {
  let release;
  const pending = new Promise((resolve) => { release = resolve; });
  const harness = createSaveHarness({
    withGAS: async (method, payload) => {
      harness.calls.push({ method, payload });
      await pending;
      return {
        ok: true,
        deduplicated: false,
        pin: {
          id: 'pin-1', title: payload.title, description: payload.description,
          lat: payload.lat, lng: payload.lng, color: payload.color, icon: payload.icon,
          imageUrl: 'image', fileId: 'file-1', timestamp: 'saved', eventAt: payload.eventAt,
          updatedAt: '', links: payload.links, folderUrl: 'folder', status: payload.status,
          tags: payload.tags
        }
      };
    }
  });
  const draft = photoDraft();
  const first = harness.context.__save(draft, { lat: 35, lng: 139 });
  const second = harness.context.__save(draft, { lat: 35, lng: 139 });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(harness.calls.length, 1);
  assert.equal(harness.state.upload.saving, true);
  release();
  await Promise.all([first, second]);

  assert.equal(harness.calls.length, 1);
  assert.equal(harness.getResizeCalls(), 1);
  assert.equal(harness.state.pins.length, 1);
  assert.equal(harness.clears.length, 1);
  assert.equal(harness.state.upload.saving, false);
  assert.equal(harness.state.upload.savePromise, null);
});

test('photo selection state owns stable save identity and clears it only with the photo session', () => {
  assert.match(indexHtml, /photoSaveIdentity:\s*null/);
  assert.match(indexHtml, /submittedPhotoPayload:\s*null/);
  assert.match(indexHtml, /savePromise:\s*null/);
  const selection = extractFunction(indexHtml, 'handleUploadPhotoSelected');
  assert.match(selection, /state\.upload\.photoSaveIdentity\s*=\s*createSinglePhotoSaveIdentity\(\)/);
  const clear = extractFunction(indexHtml, 'clearUploadPhotoState');
  assert.match(clear, /state\.upload\.photoSaveIdentity\s*=\s*null/);
  assert.match(clear, /state\.upload\.submittedPhotoPayload\s*=\s*null/);
  assert.match(clear, /state\.upload\.saving.*afterSave/);
});

test('photo save uses the receipt API while photo-less save keeps saveMapData', () => {
  const source = extractFunction(indexHtml, 'saveNewPin');
  assert.match(source, /draft\.file[\s\S]*withGAS\('saveImportPhotoItem'/);
  assert.match(source, /withGAS\('saveMapData'/);
});

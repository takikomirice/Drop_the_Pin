const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function functionSource(name) {
  const functionStart = indexHtml.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `Expected function ${name}`);
  const start = indexHtml.slice(Math.max(0, functionStart - 6), functionStart) === 'async '
    ? functionStart - 6
    : functionStart;
  const bodyStart = indexHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < indexHtml.length; index += 1) {
    const character = indexHtml[index];
    if (quote) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'" || character === '`') {
      quote = character;
      continue;
    }
    if (character === '{') depth += 1;
    if (character === '}') {
      depth -= 1;
      if (depth === 0) return indexHtml.slice(start, index + 1);
    }
  }
  throw new Error(`Unterminated function ${name}`);
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

function preset(id, orderIndex = 0) {
  return {
    presetId: id,
    name: `Preset ${id}`,
    enabled: true,
    orderIndex,
    tagsMode: 'set',
    tags: ['植物'],
    colorMode: 'keep',
    color: null,
    iconMode: 'keep',
    icon: null,
    statusMode: 'keep',
    status: null,
    createdAt: '2026-01-01T00:00:00.000Z',
    updatedAt: '2026-01-01T00:00:00.000Z'
  };
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function createStateHarness(initial = {}, options = {}) {
  const responses = [];
  const gasCalls = [];
  const overlays = [];
  const renderCounts = { list: 0, editor: 0, controls: 0, sortable: 0 };
  const state = {
    inputPresets: Object.assign({
      items: [], loading: false, loaded: false, loadPromise: null,
      saving: false, editing: null, sortable: null
    }, initial)
  };
  const elements = {
    'input-preset-name': { focus() {} }
  };
  const context = {
    state,
    settingsSavePending: false,
    inputPresetEditorOpener: null,
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'default' }],
    PIN_STATUSES: ['未対応'],
    canEdit: () => options.canEdit !== false,
    withEditToken: (payload) => Object.assign({ __editToken: 'token' }, payload),
    withGAS(method, payload) {
      gasCalls.push({ method, payload: clone(payload) });
      const response = responses.shift();
      if (!response) return Promise.reject(new Error(`Unexpected GAS call: ${method}`));
      return response;
    },
    startupDiagnostics: {
      measurePromise(stage, operation) {
        try {
          return Promise.resolve(operation());
        } catch (error) {
          return Promise.reject(error);
        }
      }
    },
    closeOverlay: (id) => overlays.push(`close:${id}`),
    closeSettingsModal: () => { overlays.push('close:settings-overlay'); return true; },
    openOverlay: (id) => overlays.push(`open:${id}`),
    renderInputPresetList: () => { renderCounts.list += 1; },
    renderInputPresetEditor: () => { renderCounts.editor += 1; },
    updateInputPresetEditorControls: () => { renderCounts.controls += 1; },
    destroyInputPresetSortable: () => { state.inputPresets.sortable = null; },
    clearInputPresetError: () => { state.inputPresets.error = ''; },
    setInputPresetError: (error) => { state.inputPresets.error = error && error.message ? error.message : String(error); },
    showHint: () => {},
    buildInputPresetPayload: () => clone(state.inputPresets.editing || {}),
    confirm: () => true,
    document: { getElementById: (id) => elements[id] || { focus() {} } },
    window: { Sortable: { create: () => { renderCounts.sortable += 1; return { destroy() {} }; } } },
    Sortable: { create: () => { renderCounts.sortable += 1; return { destroy() {} }; } },
    console
  };
  const names = [
    'cloneInputPreset', 'compactInputPresetPayload', 'createBlankInputPreset', 'createInputPresetDuplicate',
    'sortInputPresetItems', 'isInputPresetManagerBusy', 'hasLoadedInputPresets', 'editInputPreset',
    'getEnabledInputPresets', 'loadInputPresetCatalog', 'ensureInputPresetsLoaded',
    'reloadInputPresets',
    'rememberInputPresetEditorOpener', 'duplicateInputPreset', 'addInputPreset', 'saveInputPresetEditor',
    'deleteInputPresetFromUi', 'toggleInputPresetEnabled', 'persistInputPresetOrder',
    'moveInputPreset', 'attachInputPresetSortable', 'openInputPresetManager',
    'closeInputPresetManager'
  ];
  vm.runInNewContext(`${names.map(functionSource).join('\n')}
    this.api = {
      addInputPreset, editInputPreset, duplicateInputPreset, saveInputPresetEditor,
      getEnabledInputPresets, ensureInputPresetsLoaded, reloadInputPresets,
      deleteInputPresetFromUi, toggleInputPresetEnabled, persistInputPresetOrder,
      moveInputPreset, attachInputPresetSortable, openInputPresetManager,
      closeInputPresetManager
    };`, context);
  return {
    api: context.api,
    state,
    gasCalls,
    overlays,
    renderCounts,
    enqueue(response) { responses.push(response); }
  };
}

function createElement() {
  const element = {
    children: [],
    disabled: false,
    className: '',
    textContent: '',
    dataset: {},
    style: {},
    classList: { toggle() {}, add() {}, remove() {} },
    appendChild(child) { this.children.push(child); return child; },
    removeChild(child) { this.children.splice(this.children.indexOf(child), 1); },
    addEventListener() {},
    setAttribute() {},
    querySelectorAll() { return []; }
  };
  Object.defineProperty(element, 'firstChild', { get() { return this.children[0] || null; } });
  return element;
}

test('render disables add until the preset list has loaded but always permits close after loading ends', () => {
  const elements = {
    'input-presets-list': createElement(),
    'input-preset-add': createElement(),
    'input-presets-close': createElement()
  };
  const context = {
    state: { inputPresets: { items: [], loading: false, loaded: false, saving: false, editing: null, sortable: null } },
    document: {
      getElementById: (id) => elements[id],
      createElement
    },
    canEdit: () => true,
    destroyInputPresetSortable() {},
    isInputPresetManagerBusy: () => false,
    attachInputPresetSortable() {}
  };
  vm.runInNewContext(`${functionSource('hasLoadedInputPresets')}\n${functionSource('renderInputPresetList')}\nthis.render = renderInputPresetList;`, context);

  context.render();
  assert.equal(elements['input-preset-add'].disabled, true);
  assert.equal(elements['input-presets-close'].disabled, false);

  context.state.inputPresets.loaded = true;
  context.render();
  assert.equal(elements['input-preset-add'].disabled, false);
});

test('shared catalog deduplicates in-flight loads, reuses success, and returns ordered enabled clones', async () => {
  const harness = createStateHarness();
  const pending = deferred();
  harness.enqueue(pending.promise);

  const first = harness.api.ensureInputPresetsLoaded();
  const second = harness.api.ensureInputPresetsLoaded();
  assert.equal(first, second);
  assert.equal(harness.state.inputPresets.loading, true);
  await Promise.resolve();
  assert.equal(harness.gasCalls.length, 1);

  pending.resolve({
    ok: true,
    presets: [preset('later', 20), Object.assign(preset('disabled', 0), { enabled: false }), preset('first', 10)]
  });
  await first;
  assert.equal(harness.state.inputPresets.loaded, true);
  assert.equal(harness.state.inputPresets.loading, false);
  assert.equal(harness.state.inputPresets.loadPromise, null);

  const candidates = harness.api.getEnabledInputPresets();
  assert.deepEqual(clone(candidates).map((item) => item.presetId), ['first', 'later']);
  candidates[0].name = 'mutated outside';
  assert.equal(harness.state.inputPresets.items.find((item) => item.presetId === 'first').name, 'Preset first');

  await harness.api.ensureInputPresetsLoaded();
  assert.equal(harness.gasCalls.length, 1);
});

test('forced reload clears old candidates and a failure leaves one shared unloaded catalog', async () => {
  const harness = createStateHarness({ items: [preset('old')], loaded: true });
  const pending = deferred();
  harness.enqueue(pending.promise);
  const loading = harness.api.reloadInputPresets();

  assert.equal(harness.state.inputPresets.loaded, false);
  assert.deepEqual(clone(harness.state.inputPresets.items), []);
  assert.deepEqual(clone(harness.api.getEnabledInputPresets()), []);
  pending.reject(new Error('<b>private GAS detail</b>'));
  await assert.rejects(loading, /プリセットの読み込みに失敗/);
  assert.equal(harness.state.inputPresets.loaded, false);
  assert.deepEqual(clone(harness.state.inputPresets.items), []);
  assert.equal(harness.state.inputPresets.loading, false);
});

test('catalog refuses to call GAS outside editable authenticated mode', async () => {
  const harness = createStateHarness({}, { canEdit: false });
  await assert.rejects(harness.api.ensureInputPresetsLoaded(), /編集モード/);
  assert.deepEqual(harness.gasCalls, []);
  assert.equal(harness.state.inputPresets.loaded, false);
});

test('management mutations are immediately reflected by the shared enabled candidates', async () => {
  const editing = Object.assign(preset('new', 1), { name: 'New preset' });
  delete editing.presetId;
  const harness = createStateHarness({ items: [preset('existing', 0)], loaded: true, editing });
  harness.enqueue(Promise.resolve({ ok: true, preset: preset('new', 1) }));
  await harness.api.saveInputPresetEditor();
  assert.deepEqual(clone(harness.api.getEnabledInputPresets()).map((item) => item.presetId), ['existing', 'new']);

  const disabled = Object.assign(preset('existing', 0), { enabled: false });
  harness.enqueue(Promise.resolve({ ok: true, preset: disabled }));
  await harness.api.toggleInputPresetEnabled('existing');
  assert.deepEqual(clone(harness.api.getEnabledInputPresets()).map((item) => item.presetId), ['new']);
});

test('reopening clears stale items before fetch and a failed fetch leaves the manager unloaded and closable', async () => {
  const harness = createStateHarness({ items: [preset('stale')], loaded: true });
  const first = deferred();
  harness.enqueue(first.promise);
  const firstOpen = harness.api.openInputPresetManager();
  assert.equal(harness.state.inputPresets.loading, true);
  assert.equal(harness.state.inputPresets.loaded, false);
  assert.deepEqual(clone(harness.state.inputPresets.items), []);
  first.resolve({ ok: true, presets: [preset('fresh')] });
  await firstOpen;
  assert.equal(harness.state.inputPresets.loaded, true);
  assert.deepEqual(clone(harness.state.inputPresets.items).map((item) => item.presetId), ['fresh']);

  harness.api.closeInputPresetManager();
  const second = deferred();
  harness.enqueue(second.promise);
  const secondOpen = harness.api.openInputPresetManager();
  assert.equal(harness.state.inputPresets.loaded, false);
  assert.deepEqual(clone(harness.state.inputPresets.items), []);
  second.reject(new Error('list failed'));
  await secondOpen;
  assert.equal(harness.state.inputPresets.loading, false);
  assert.equal(harness.state.inputPresets.loaded, false);
  assert.deepEqual(clone(harness.state.inputPresets.items), []);
  assert.match(harness.state.inputPresets.error, /再度開いてください/);
  harness.api.closeInputPresetManager();
  assert.equal(harness.overlays.at(-1), 'close:input-presets-overlay');

  harness.enqueue(Promise.resolve({ ok: true, presets: [preset('recovered')] }));
  await harness.api.openInputPresetManager();
  assert.equal(harness.state.inputPresets.loaded, true);
  harness.api.addInputPreset();
  assert.notEqual(harness.state.inputPresets.editing, null);
});

test('unloaded preset state blocks every mutation path and Sortable without relying on disabled DOM', async () => {
  const originalItems = [preset('a', 0), preset('b', 1)];
  const sentinelEditing = { name: 'sentinel' };
  const harness = createStateHarness({ items: clone(originalItems), loaded: false, editing: sentinelEditing });

  harness.api.addInputPreset();
  harness.api.editInputPreset('a');
  harness.api.duplicateInputPreset('a');
  await harness.api.deleteInputPresetFromUi('a');
  await harness.api.toggleInputPresetEnabled('a');
  harness.api.moveInputPreset('a', 1);
  await Promise.resolve();
  await harness.api.persistInputPresetOrder(['b', 'a']);
  harness.api.attachInputPresetSortable({ querySelectorAll: () => [] });
  await harness.api.saveInputPresetEditor();

  assert.equal(harness.state.inputPresets.editing, sentinelEditing);
  assert.deepEqual(clone(harness.state.inputPresets.items), originalItems);
  assert.deepEqual(harness.gasCalls, []);
  assert.equal(harness.renderCounts.sortable, 0);
});

test('every mutation entry point explicitly requires a loaded preset list', () => {
  [
    'addInputPreset', 'editInputPreset', 'duplicateInputPreset', 'deleteInputPresetFromUi',
    'toggleInputPresetEnabled', 'moveInputPreset', 'persistInputPresetOrder',
    'attachInputPresetSortable', 'saveInputPresetEditor'
  ].forEach((name) => {
    assert.match(functionSource(name), /hasLoadedInputPresets\(\)/, name);
  });
  assert.match(functionSource('updateInputPresetEditorControls'), /!hasLoadedInputPresets\(\)/);
});

test('save failure keeps a loaded list and unsaved editor while order failure still rolls back', async () => {
  const originalItems = [preset('a', 0), preset('b', 1)];
  const editing = Object.assign(preset('a', 0), { name: 'Unsaved name' });
  const saveHarness = createStateHarness({ items: clone(originalItems), loaded: true, editing });
  saveHarness.enqueue(Promise.reject(new Error('save failed')));
  await saveHarness.api.saveInputPresetEditor();
  assert.equal(saveHarness.state.inputPresets.loaded, true);
  assert.deepEqual(clone(saveHarness.state.inputPresets.items), originalItems);
  assert.equal(saveHarness.state.inputPresets.editing.name, 'Unsaved name');

  const orderHarness = createStateHarness({ items: clone(originalItems), loaded: true });
  orderHarness.enqueue(Promise.reject(new Error('order failed')));
  await orderHarness.api.persistInputPresetOrder(['b', 'a']);
  assert.equal(orderHarness.state.inputPresets.loaded, true);
  assert.deepEqual(clone(orderHarness.state.inputPresets.items), originalItems);
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionSource(name) {
  const start = indexHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let cursor = open; cursor < indexHtml.length; cursor += 1) {
    const character = indexHtml[cursor];
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
    if (character === '}' && --depth === 0) return indexHtml.slice(start, cursor + 1);
  }
  assert.fail(`Could not parse function ${name}`);
}

function functionBody(name) {
  const source = functionSource(name);
  return source.slice(source.indexOf('{') + 1, -1);
}

function createReadinessHarness(overrides = {}) {
  const button = { disabled: true };
  const state = {
    initializing: false,
    editMode: true,
    previewMode: false,
    narrowView: false,
    shareMode: false,
    upload: { saving: false },
    trackDeleteConfirming: '',
    addMenuPreparing: false,
    placement: null,
    duplicatePlacement: null
  };
  Object.assign(state, overrides.state || {});
  if (overrides.state && overrides.state.upload) {
    state.upload = Object.assign({ saving: false }, overrides.state.upload);
  }
  const context = {
    state,
    hasEditToken: overrides.hasEditToken !== false,
    document: { getElementById: (id) => (id === 'pin-add-btn' ? button : null) },
    isProductionImportBusy: () => overrides.busy === true,
    startupDiagnostics: { recordPinAddReady() {} }
  };
  vm.runInNewContext(`
    ${functionSource('canEdit')}
    ${functionSource('canStartAddAction')}
    ${functionSource('refreshPinAddButtonState')}
    this.api = { canEdit, canStartAddAction, refreshPinAddButtonState };
  `, context);
  return { api: context.api, button, state };
}

test('startup busy release refreshes pin add without opening Data', () => {
  const initialize = functionBody('initializeApp');
  assert.match(initialize, /state\.initializing = false;[\s\S]*?refreshPinAddButtonState\(\)/);

  const dataWorkbench = functionBody('openDataWorkbench');
  assert.doesNotMatch(dataWorkbench, /state\.initializing\s*=/);
  assert.doesNotMatch(dataWorkbench, /state\.(?:editMode|previewMode|narrowView|shareMode)\s*=/);

  const harness = createReadinessHarness({ state: { initializing: true } });
  assert.equal(harness.api.refreshPinAddButtonState(), false);
  assert.equal(harness.button.disabled, true);
  harness.state.initializing = false;
  assert.equal(harness.api.refreshPinAddButtonState(), true);
  assert.equal(harness.button.disabled, false);
});

test('pin add disabled state has one shared writer backed by add-action policy', () => {
  const refresh = functionBody('refreshPinAddButtonState');
  assert.equal((refresh.match(/\.disabled\s*=/g) || []).length, 1);
  assert.doesNotMatch(indexHtml, /singleAddButton\.disabled\s*=/);
  assert.match(refresh, /canStartAddAction\(\)/);
  assert.match(functionBody('renderTrackImportBusy'), /refreshPinAddButtonState\(\)/);

  const addPolicy = functionBody('canStartAddAction');
  for (const guard of [
    /canEdit\(\)/,
    /isProductionImportBusy\(\)/,
    /state\.upload\.saving/,
    /state\.trackDeleteConfirming/,
    /state\.addMenuPreparing/,
    /state\.placement/,
    /state\.duplicatePlacement/
  ]) assert.match(addPolicy, guard);
});

test('read-only narrow preview share and production busy modes never enable pin add', () => {
  const scenarios = [
    { hasEditToken: false },
    { state: { editMode: false } },
    { state: { previewMode: true } },
    { state: { narrowView: true } },
    { state: { shareMode: true } },
    { busy: true }
  ];
  scenarios.forEach((scenario) => {
    const harness = createReadinessHarness(scenario);
    assert.equal(harness.api.refreshPinAddButtonState(), false);
    assert.equal(harness.button.disabled, true);
  });
});

test('local add preparation save and placement states keep pin add disabled', () => {
  const scenarios = [
    { upload: { saving: true } },
    { trackDeleteConfirming: 'track-1' },
    { addMenuPreparing: true },
    { placement: { mode: 'new' } },
    { duplicatePlacement: { sourcePinId: 'pin-1' } }
  ];
  scenarios.forEach((state) => {
    const harness = createReadinessHarness({ state });
    assert.equal(harness.api.refreshPinAddButtonState(), false);
    assert.equal(harness.button.disabled, true);
  });
});

test('pin load success and failure plus access-mode settlement reevaluate readiness', () => {
  assert.match(functionBody('loadMapData'), /finally\s*\{[\s\S]*?refreshPinAddButtonState\(\)/);
  assert.match(functionBody('setEditMode'), /refreshPinAddButtonState\(\)|renderTrackImportBusy\(\)/);
  assert.match(functionBody('setPreviewMode'), /refreshPinAddButtonState\(\)|renderTrackImportBusy\(\)/);
  assert.match(functionBody('setNarrowView'), /refreshPinAddButtonState\(\)/);
});

test('opening Data does not change the formal pin-add-ready condition', () => {
  const readiness = functionBody('canStartAddAction');
  assert.doesNotMatch(readiness, /data-overlay|openDataWorkbench|classList\.contains\('open'\)/);
  const before = createReadinessHarness();
  const after = createReadinessHarness();
  assert.equal(before.api.canStartAddAction(), after.api.canStartAddAction());
});

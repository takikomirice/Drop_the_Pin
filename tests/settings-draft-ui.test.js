const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionSource(name) {
  const functionStart = indexHtml.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `Expected function ${name}`);
  const start = indexHtml.slice(Math.max(0, functionStart - 6), functionStart) === 'async '
    ? functionStart - 6
    : functionStart;
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
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

function createClassList() {
  const names = new Set();
  return {
    add(name) { names.add(name); },
    remove(...values) { values.forEach((value) => names.delete(value)); },
    contains(name) { return names.has(name); },
    toggle(name, force) {
      const enabled = force === undefined ? !names.has(name) : !!force;
      if (enabled) names.add(name);
      else names.delete(name);
      return enabled;
    }
  };
}

function createSettingsHarness(options = {}) {
  const elements = new Proxy({}, {
    get(target, key) {
      if (!target[key]) {
        target[key] = {
          classList: createClassList(),
          style: {},
          textContent: '',
          disabled: false,
          attributes: {},
          setAttribute(name, value) { this.attributes[name] = String(value); }
        };
      }
      return target[key];
    }
  });
  const state = {
    appSettings: {
      rootFolderId: 'root-id',
      rootFolderUrl: 'https://example.test/root',
      renameFileWithTitle: !!options.renameFileWithTitle
    },
    settingsDraft: null,
    upload: { uploadFile: { name: 'photo.jpg' } },
    editingPinId: 'pin-1'
  };
  const gasCalls = [];
  const gasResults = [];
  const notifications = [];
  const hints = [];
  const context = {
    state,
    settingsSavePending: false,
    document: {
      getElementById(id) { return elements[id]; }
    },
    canEdit() { return true; },
    isProductionImportBusy() { return false; },
    closeMoreMenu() {},
    closeDataWorkbench() { elements['data-overlay'].classList.remove('open'); return true; },
    openOverlay(id) { elements[id].classList.add('open'); },
    closeOverlay(id) { elements[id].classList.remove('open'); },
    showHint(message) { hints.push(message); },
    getPinById(id) { return id === 'pin-1' ? { id, fileId: 'file-1' } : null; },
    updateInputPresetManagerButton() {},
    withEditToken(payload) { return Object.assign({ editToken: 'edit-token' }, payload); },
    withGAS(method, payload) {
      gasCalls.push({ method, payload });
      if (!gasResults.length) throw new Error('missing queued GAS result');
      return gasResults.shift();
    },
    showAppNotification(options) { notifications.push(options); }
  };
  context.setActionButtonLabel = vm.runInNewContext(`(${functionSource('setActionButtonLabel')})`, context);
  [
    'createSettingsDraft', 'discardSettingsDraft', 'updateRenameToggle',
    'setRenameNote', 'refreshRenameNotes', 'closeSettingsModal',
    'openSettingsModal', 'toggleRenameSyncDraft', 'saveAppSettings'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });
  return { context, state, elements, gasCalls, gasResults, notifications, hints };
}

test('settings open creates a fresh draft and toggle/cancel never mutate canonical settings', () => {
  const harness = createSettingsHarness({ renameFileWithTitle: false });

  harness.context.openSettingsModal();
  assert.deepEqual({ ...harness.state.settingsDraft }, { renameFileWithTitle: false });
  assert.equal(harness.elements['rename-sync-toggle'].attributes['aria-pressed'], 'false');

  harness.context.toggleRenameSyncDraft();
  assert.equal(harness.state.settingsDraft.renameFileWithTitle, true);
  assert.equal(harness.state.appSettings.renameFileWithTitle, false);
  harness.context.refreshRenameNotes();
  assert.match(harness.elements['upload-rename-note'].textContent, /変更しません/);
  assert.match(harness.elements['edit-rename-note'].textContent, /変更しません/);

  assert.equal(harness.context.closeSettingsModal(), true);
  assert.equal(harness.state.settingsDraft, null);
  assert.equal(harness.state.appSettings.renameFileWithTitle, false);
  assert.equal(harness.state.appSettings.rootFolderId, 'root-id');
  assert.equal(harness.state.appSettings.rootFolderUrl, 'https://example.test/root');

  harness.context.openSettingsModal();
  assert.deepEqual({ ...harness.state.settingsDraft }, { renameFileWithTitle: false });
  assert.equal(harness.elements['rename-sync-toggle'].attributes['aria-pressed'], 'false');
});

test('every settings close route delegates to closeSettingsModal so drafts are discarded', () => {
  assert.match(functionSource('closeOverlayFromBackdrop'), /dismissOverlayById\(record\.id\)/);
  assert.match(functionSource('dismissOverlayById'), /settings-overlay[\s\S]*closeSettingsModal\(\)/);
  assert.match(functionSource('openDataWorkbench'), /settings-overlay[\s\S]*closeSettingsModal\(\{\s*restoreFocus:\s*false\s*\}\)/);
  assert.match(functionSource('dismissEditOnlySurfacesForPreview'), /settings-overlay[\s\S]*closeSettingsModal\(\)/);
  assert.match(functionSource('openInputPresetManager'), /closeSettingsModal\(\{\s*restoreFocus:\s*false\s*\}\)/);
});

test('input preset manager renders its loading state before common initial focus runs', () => {
  const source = functionSource('openInputPresetManager');
  const reloadIndex = source.indexOf('const loading = reloadInputPresets()');
  const renderIndex = source.indexOf('renderInputPresetList()');
  const openIndex = source.indexOf("openOverlay('input-presets-overlay')");

  assert.ok(reloadIndex >= 0);
  assert.ok(renderIndex > reloadIndex);
  assert.ok(openIndex > renderIndex);
});

test('successful save commits only the server response and snapshots the draft payload', async () => {
  const harness = createSettingsHarness({ renameFileWithTitle: false });
  const response = deferred();
  harness.gasResults.push(response.promise);
  harness.context.openSettingsModal();
  harness.context.toggleRenameSyncDraft();

  const saving = harness.context.saveAppSettings();
  assert.equal(harness.gasCalls.length, 1);
  assert.equal(harness.gasCalls[0].method, 'updateAppSettings');
  assert.deepEqual({ ...harness.gasCalls[0].payload }, {
    editToken: 'edit-token',
    renameFileWithTitle: true
  });

  harness.state.settingsDraft.renameFileWithTitle = false;
  assert.equal(harness.state.appSettings.renameFileWithTitle, false);
  assert.equal(harness.gasCalls[0].payload.renameFileWithTitle, true);

  response.resolve({ ok: true, renameFileWithTitle: true });
  await saving;

  assert.equal(harness.state.appSettings.renameFileWithTitle, true);
  assert.equal(harness.state.settingsDraft, null);
  assert.equal(harness.elements['settings-overlay'].classList.contains('open'), false);
  assert.match(harness.elements['upload-rename-note'].textContent, /同期します/);
  assert.match(harness.elements['edit-rename-note'].textContent, /同期します/);
  assert.equal(harness.context.settingsSavePending, false);
});

test('failed save preserves the draft for retry and leaves canonical settings unchanged', async () => {
  const harness = createSettingsHarness({ renameFileWithTitle: false });
  harness.gasResults.push(Promise.reject(new Error('temporary failure')));
  harness.context.openSettingsModal();
  harness.context.toggleRenameSyncDraft();

  await harness.context.saveAppSettings();

  assert.equal(harness.state.appSettings.renameFileWithTitle, false);
  assert.equal(harness.state.settingsDraft.renameFileWithTitle, true);
  assert.equal(harness.elements['settings-overlay'].classList.contains('open'), true);
  assert.equal(harness.context.settingsSavePending, false);
  assert.equal(harness.notifications.length, 1);
  assert.match(harness.notifications[0].message, /temporary failure/);
});

test('rename notes and photo save paths never read the unsaved settings draft', () => {
  assert.match(functionSource('setRenameNote'), /state\.appSettings\.renameFileWithTitle/);
  assert.match(functionSource('setRenameNote'), /元のDrive写真は変更しません/);
  assert.doesNotMatch(functionSource('setRenameNote'), /settingsDraft/);
  assert.doesNotMatch(functionSource('createSinglePhotoSavePayload'), /settingsDraft/);
  assert.doesNotMatch(functionSource('saveNewPin'), /settingsDraft/);
});

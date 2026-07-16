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
  assert.fail(`Could not parse ${name}`);
}

function createClassList() {
  const values = new Set();
  return {
    add(value) { values.add(value); },
    remove(...items) { items.forEach((item) => values.delete(item)); },
    contains(value) { return values.has(value); },
    toggle(value, force) {
      const enabled = force === undefined ? !values.has(value) : !!force;
      if (enabled) values.add(value);
      else values.delete(value);
      return enabled;
    }
  };
}

function createElement(id) {
  const classList = createClassList();
  if (String(id).endsWith('-overlay')) classList.add('sheet-overlay');
  const attributes = {};
  return {
    id,
    classList,
    style: { zIndex: '' },
    value: '',
    src: '',
    textContent: '',
    innerHTML: '',
    disabled: false,
    inert: false,
    onclick: null,
    getAttribute(name) {
      return Object.prototype.hasOwnProperty.call(attributes, name) ? attributes[name] : null;
    },
    setAttribute(name, value) {
      attributes[name] = String(value);
      if (name === 'inert') this.inert = true;
    },
    removeAttribute(name) {
      delete attributes[name];
      if (name === 'inert') this.inert = false;
    },
    hasAttribute(name) { return Object.prototype.hasOwnProperty.call(attributes, name); },
    closest() { return null; },
    focus() {}
  };
}

function createOpener(label) {
  return {
    label,
    connected: true,
    disabled: false,
    focusCount: 0,
    getAttribute() { return null; },
    setAttribute() {},
    closest() { return null; },
    focus() { this.focusCount += 1; }
  };
}

function escapeEvent() {
  return {
    key: 'Escape',
    preventCount: 0,
    stopCount: 0,
    preventDefault() { this.preventCount += 1; },
    stopPropagation() { this.stopCount += 1; }
  };
}

function createHarness() {
  const elements = new Proxy({}, {
    get(target, key) {
      if (!target[key]) target[key] = createElement(String(key));
      return target[key];
    }
  });
  const revokedUrls = [];
  const hints = [];
  const renderCounts = { pins: 0, tracks: 0, panel: 0, presetEditor: 0 };
  const controllerCalls = { drive: 0, importPreview: 0, trackDiscard: 0, trackClose: 0, preparation: 0 };
  const state = {
    duplicatePlacement: null,
    placement: null,
    addMenuPreparing: false,
    settingsDraft: null,
    appSettings: { renameFileWithTitle: false },
    upload: {
      saving: false, originalFile: null, uploadFile: null, previewUrl: '', lat: null, lng: null,
      capturedAt: '', metadataStatus: 'idle', converting: false, conversionError: '',
      photoSaveIdentity: null, submittedPhotoPayload: null, positionMode: 'map',
      positionModeManual: false, selectionToken: 0
    },
    inputPresets: { loading: false, saving: false, editing: null, sortable: null },
    bulkMetadata: { ids: [], saving: false },
    bulkRouteAdd: { pinIds: [] },
    bulkDeletePending: false,
    deleteMutationPending: false,
    shownPinId: null,
    contextPinId: null,
    editingPinId: null,
    routeEditOpenId: null,
    shareManager: {
      loading: false, mutationPending: false, previewing: false,
      sessionGeneration: 0, listRequestId: 0
    },
    trackImport: { owner: '', saving: false, saved: false },
    trackEdit: { trackId: '', saving: false, saved: false },
    multiPhotoImport: { preparing: false, registering: false, cancellingPreparation: false }
  };
  const context = {
    state,
    overlayOpenRecords: [],
    overlayBackgroundStateRecords: new Map(),
    OVERLAY_STACK_Z_INDEX_BASE: 1400,
    OVERLAY_STACK_Z_INDEX_STEP: 10,
    inputPresetEditorOpener: null,
    MAIN_DISMISSIBLE_OVERLAY_IDS: [
      'add-menu-overlay', 'upload-overlay', 'drive-photo-import-overlay', 'settings-overlay',
      'data-overlay', 'input-presets-overlay', 'location-choice-overlay', 'pin-detail-overlay',
      'photo-viewer-overlay',
      'pin-menu-overlay', 'edit-overlay', 'bulk-metadata-overlay',
      'bulk-route-add-overlay', 'delete-overlay', 'route-edit-overlay', 'share-overlay',
      'share-qr-overlay', 'multi-photo-preparation-overlay', 'import-preview-overlay',
      'track-import-preview-overlay', 'dup-warning-overlay', 'route-name-overlay',
      'route-track-import-overlay'
    ],
    BACKDROP_DISMISS_OVERLAY_IDS: [
      'help-overlay', 'add-menu-overlay', 'upload-overlay', 'drive-photo-import-overlay',
      'settings-overlay', 'data-overlay', 'location-choice-overlay', 'pin-detail-overlay',
      'photo-viewer-overlay',
      'pin-menu-overlay', 'delete-overlay', 'share-overlay',
      'share-qr-overlay', 'dup-warning-overlay', 'bulk-metadata-overlay',
      'bulk-route-add-overlay', 'input-presets-overlay'
    ],
    settingsSavePending: false,
    window: { matchMedia() { return { matches: false }; } },
    Map,
    document: {
      activeElement: null,
      body: { children: [], classList: createClassList() },
      getElementById(id) { return elements[id]; },
      contains(node) { return !!node && node.connected !== false; }
    },
    URL: { revokeObjectURL(url) { revokedUrls.push(url); } },
    csvInterchangeController: { invalidate() { return true; } },
    geoJsonInterchangeController: { invalidate() { return true; } },
    trackImportController: {
      busy: false,
      discard() {
        controllerCalls.trackDiscard += 1;
        if (this.busy) return false;
        context.closeOverlay('track-import-preview-overlay');
        state.trackImport.owner = '';
        return true;
      },
      close() {
        controllerCalls.trackClose += 1;
        if (this.busy) return false;
        context.closeOverlay('track-import-preview-overlay');
        state.trackImport.owner = '';
        return true;
      },
      invalidate() { return !this.busy; }
    },
    drivePhotoImportController: {
      busy: false,
      cancel() {
        controllerCalls.drive += 1;
        if (this.busy) return false;
        context.closeOverlay('drive-photo-import-overlay');
        return true;
      }
    },
    ImportPreviewUI: {
      busy: false,
      close(options) {
        controllerCalls.importPreview += 1;
        assert.equal(options.discard, true);
        if (this.busy) return false;
        context.closeOverlay('import-preview-overlay');
        return true;
      }
    },
    multiPhotoWorkflow: {
      cancelPreparation() { controllerCalls.preparation += 1; return true; }
    },
    isProductionImportBusy() { return false; },
    isInputPresetManagerBusy() { return state.inputPresets.loading || state.inputPresets.saving; },
    destroyInputPresetSortable() { state.inputPresets.sortable = null; },
    clearInputPresetError() {},
    renderInputPresetEditor() { renderCounts.presetEditor += 1; },
    refreshUploadPhotoStatus() {}, refreshUploadSubmitState() {}, refreshUploadPositionControls() {},
    refreshRenameNotes() {}, loadUploadInputPresets() {}, refreshMultiPhotoButtonState() {},
    refreshPinAddButtonState() {},
    hideHint() {}, renderPins() { renderCounts.pins += 1; },
    renderTrackLayers() { renderCounts.tracks += 1; }, renderSidePanel() { renderCounts.panel += 1; },
    showHint(message) { hints.push(message); },
    handleDuplicateShortcut() { context.shortcutCalls += 1; },
    focusOverlayInitial() {},
    clearOverlayFallbackFocus() {},
    trapOverlayFocus() { return false; },
    closePhotoViewerForTrigger() { return false; },
    updatePhotoViewerTrigger() { return false; },
    closePhotoViewer() { return context.closeOverlay('photo-viewer-overlay'); },
    closeRouteAddMenu() { return false; },
    closeRouteNameDialog() { context.closeOverlay('route-name-overlay'); return true; },
    closeRouteTrackImportDialog() { context.closeOverlay('route-track-import-overlay'); return true; },
    shortcutCalls: 0
  };

  [
    'setActionButtonLabel',
    'setOverlayInteractionAttribute', 'captureOverlayInteractionState',
    'restoreOverlayInteractionState', 'removeOverlayOpenRecord', 'getOverlayOpenRecord',
    'isDockedPinDetailOverlay', 'getOpenSheetOverlayRecords', 'getTopOpenSheetOverlayRecord',
    'getTopModalSheetOverlayRecord', 'restoreOverlayBackgroundState', 'syncOverlayInteractionState',
    'isOverlayElementVisible', 'openOverlay', 'closeOverlay',
    'canRestoreSurfaceFocus', 'restoreSurfaceFocus', 'getTopOpenOverlayRecord',
    'discardSettingsDraft', 'closeSettingsModal', 'closeDataWorkbench',
    'clearUploadPhotoState', 'cancelUpload', 'rememberInputPresetEditorOpener',
    'cancelInputPresetEditor', 'closeInputPresetManager', 'closeBulkMetadataOverlay',
    'closeBulkRouteAdd', 'closePinDetail', 'closePinMenu', 'cancelPinEditor',
    'cancelDeleteConfirmation', 'closeRouteEditOverlay', 'closeShareQr',
    'advanceShareManagerSession', 'closeShareDialog', 'closeDuplicateWarning',
    'cancelLocationChoice', 'clearDuplicatePlacement', 'clearPlacement', 'cancelPlacement',
    'returnToPhotoSource', 'returnFromMultiPhotoPreparation',
    'closeHelpPanel', 'setMoreMenuOpen', 'closeMoreMenu',
    'isTrackDisplaySettingsEditorOpen', 'dismissOverlayById', 'closeOverlayFromBackdrop', 'dispatchEscape',
    'handleGlobalKeydown'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });

  return {
    context, state, elements, revokedUrls, hints, renderCounts, controllerCalls,
    open(id, opener) {
      context.document.activeElement = opener || null;
      context.openOverlay(id);
    },
    pressEscape() {
      const event = escapeEvent();
      context.handleGlobalKeydown(event);
      return event;
    },
    dismissBackdrop(id) {
      return context.closeOverlayFromBackdrop(id);
    }
  };
}

test('dispatcher inventory routes every overlay through its existing cleanup contract', () => {
  const inventoryMatch = indexHtml.match(/const MAIN_DISMISSIBLE_OVERLAY_IDS = Object\.freeze\(\[([\s\S]*?)\]\);/);
  assert.ok(inventoryMatch, 'main overlay inventory');
  const inventoryIds = [...inventoryMatch[1].matchAll(/'([^']+-overlay)'/g)].map((match) => match[1]).sort();
  const domOverlayIds = [...indexHtml.matchAll(/<div id="([^"]+-overlay)" class="[^"]*sheet-overlay[^"]*"/g)]
    .map((match) => match[1])
    .filter((id) => id !== 'help-overlay')
    .sort();
  assert.deepEqual(inventoryIds, domOverlayIds);

  const source = functionSource('dismissOverlayById');
  const expected = {
    'help-overlay': 'closeHelpPanel',
    'settings-overlay': 'closeSettingsModal',
    'data-overlay': 'closeDataWorkbench',
    'upload-overlay': 'cancelUpload',
    'drive-photo-import-overlay': 'drivePhotoImportController.cancel',
    'input-presets-overlay': 'closeInputPresetManager',
    'share-overlay': 'closeShareDialog',
    'share-qr-overlay': 'closeShareQr',
    'bulk-route-add-overlay': 'closeBulkRouteAdd',
    'bulk-metadata-overlay': 'closeBulkMetadataOverlay',
    'photo-viewer-overlay': 'closePhotoViewer',
    'pin-detail-overlay': 'closePinDetail',
    'pin-menu-overlay': 'closePinMenu',
    'edit-overlay': 'cancelPinEditor',
    'delete-overlay': 'cancelDeleteConfirmation',
    'route-edit-overlay': 'closeRouteEditOverlay',
    'dup-warning-overlay': 'closeDuplicateWarning',
    'location-choice-overlay': 'cancelLocationChoice',
    'multi-photo-preparation-overlay': 'multiPhotoWorkflow.cancelPreparation',
    'import-preview-overlay': 'ImportPreviewUI.close',
    'track-import-preview-overlay': 'trackImportController',
    'add-menu-overlay': 'closeOverlay'
  };
  Object.entries(expected).forEach(([id, cleanup]) => {
    assert.ok(source.includes(id), id);
    assert.ok(source.includes(cleanup), `${id} -> ${cleanup}`);
  });
  const allDialogIds = [...indexHtml.matchAll(/<div id="([^"]+-overlay)" class="[^"]*sheet-overlay[^"]*"/g)]
    .map((match) => match[1]);
  allDialogIds.forEach((id) => assert.ok(source.includes(id), `${id} dispatcher contract`));
  assert.match(functionSource('dismissOverlayById'), /ImportPreviewUI\.close\(\{\s*discard:\s*true\s*\}\)/);
  assert.match(functionSource('openOverlay'), /document\.activeElement/);
  assert.match(functionSource('openOverlay'), /overlayOpenRecords\.push/);
  assert.match(functionSource('closeOverlay'), /removeOverlayOpenRecord/);
  assert.match(indexHtml, /openOverlay\('import-preview-overlay'\)/);
  assert.match(indexHtml, /openOverlay\('track-import-preview-overlay'\)/);
});

test('backdrop dismissal uses the common dispatcher and closes only the top nested dialog', () => {
  const harness = createHarness();
  const parentOpener = createOpener('share');
  const childOpener = createOpener('qr');
  harness.state.shareManager.previewing = true;
  harness.open('share-overlay', parentOpener);
  harness.open('share-qr-overlay', childOpener);

  const dismissed = harness.dismissBackdrop('share-overlay');

  assert.equal(dismissed, true);
  assert.equal(harness.elements['share-qr-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['share-overlay'].classList.contains('open'), true);
  assert.equal(harness.state.shareManager.previewing, true);
  assert.equal(childOpener.focusCount, 1);
  assert.equal(parentOpener.focusCount, 0);
  assert.deepEqual(harness.context.overlayOpenRecords.map((record) => record.id), ['share-overlay']);
  assert.equal(harness.elements['share-overlay'].inert, false);
});

test('backdrop busy refusal preserves focus, stack, and interaction isolation', () => {
  const harness = createHarness();
  const opener = createOpener('settings');
  const focusedControl = createOpener('settings-control');
  const background = harness.elements['app-shell'];
  harness.context.document.body.children.push(background);
  harness.state.settingsDraft = { renameFileWithTitle: true };
  harness.context.settingsSavePending = true;
  harness.open('settings-overlay', opener);
  harness.context.document.activeElement = focusedControl;

  const dismissed = harness.dismissBackdrop('settings-overlay');

  assert.equal(dismissed, false);
  assert.equal(harness.elements['settings-overlay'].classList.contains('open'), true);
  assert.equal(harness.context.document.activeElement, focusedControl);
  assert.equal(opener.focusCount, 0);
  assert.deepEqual(harness.context.overlayOpenRecords.map((record) => record.id), ['settings-overlay']);
  assert.equal(harness.elements['settings-overlay'].inert, false);
  assert.equal(background.inert, true);
  assert.equal(background.getAttribute('aria-hidden'), 'true');
  assert.deepEqual({ ...harness.state.settingsDraft }, { renameFileWithTitle: true });
});

test('desktop docked pin detail ignores backdrop dismissal', () => {
  const harness = createHarness();
  harness.context.window.matchMedia = function(query) {
    return { matches: query === '(min-width: 900px)' };
  };
  harness.context.document.body.classList.add('panel-visible');
  harness.state.shownPinId = 'pin-1';
  harness.open('pin-detail-overlay', createOpener('pin'));

  const dismissed = harness.dismissBackdrop('pin-detail-overlay');

  assert.equal(dismissed, false);
  assert.equal(harness.elements['pin-detail-overlay'].classList.contains('open'), true);
  assert.equal(harness.state.shownPinId, 'pin-1');
  assert.deepEqual(harness.context.overlayOpenRecords.map((record) => record.id), ['pin-detail-overlay']);
  assert.equal(harness.elements['pin-detail-overlay'].getAttribute('aria-modal'), 'false');
});

test('one Escape closes only the top nested overlay and restores each live opener', () => {
  const harness = createHarness();
  const shareOpener = createOpener('share');
  const qrOpener = createOpener('qr');
  harness.state.shareManager.previewing = true;
  harness.open('share-overlay', shareOpener);
  harness.open('share-qr-overlay', qrOpener);

  let event = harness.pressEscape();
  assert.equal(harness.elements['share-qr-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['share-overlay'].classList.contains('open'), true);
  assert.equal(harness.state.shareManager.previewing, true);
  assert.equal(qrOpener.focusCount, 1);
  assert.equal(shareOpener.focusCount, 0);
  assert.equal(event.preventCount, 1);
  assert.equal(event.stopCount, 1);

  event = harness.pressEscape();
  assert.equal(harness.elements['share-overlay'].classList.contains('open'), false);
  assert.equal(harness.state.shareManager.previewing, false);
  assert.equal(harness.state.shareManager.sessionGeneration, 1);
  assert.deepEqual(harness.renderCounts, { pins: 1, tracks: 1, panel: 1, presetEditor: 0 });
  assert.equal(shareOpener.focusCount, 1);
  assert.equal(event.preventCount, 1);
  assert.equal(event.stopCount, 1);
});

test('settings Escape discards only its draft while saving refuses close and focus change', () => {
  const harness = createHarness();
  const opener = createOpener('settings');
  harness.state.settingsDraft = { renameFileWithTitle: true };
  harness.open('settings-overlay', opener);
  harness.pressEscape();
  assert.equal(harness.state.settingsDraft, null);
  assert.equal(harness.state.appSettings.renameFileWithTitle, false);
  assert.equal(harness.elements['settings-overlay'].classList.contains('open'), false);
  assert.equal(opener.focusCount, 1);

  harness.state.settingsDraft = { renameFileWithTitle: true };
  harness.context.settingsSavePending = true;
  harness.open('settings-overlay', opener);
  const event = harness.pressEscape();
  assert.equal(harness.elements['settings-overlay'].classList.contains('open'), true);
  assert.deepEqual({ ...harness.state.settingsDraft }, { renameFileWithTitle: true });
  assert.equal(opener.focusCount, 1);
  assert.equal(event.preventCount, 1);
  assert.equal(event.stopCount, 1);
  assert.match(harness.hints.at(-1), /保存/);
});

test('upload Escape releases photo ownership and a saving upload remains untouched', () => {
  const harness = createHarness();
  harness.state.upload = Object.assign(harness.state.upload, {
    originalFile: { name: 'source.jpg' }, uploadFile: { name: 'upload.jpg' },
    previewUrl: 'blob:photo', lat: 35, lng: 139, capturedAt: '2026-01-01',
    metadataStatus: 'success', photoSaveIdentity: { jobId: 'job' },
    submittedPhotoPayload: { idempotencyKey: 'key' }, selectionToken: 4
  });
  harness.open('upload-overlay', createOpener('upload'));
  harness.pressEscape();
  assert.deepEqual(harness.revokedUrls, ['blob:photo']);
  assert.equal(harness.state.upload.originalFile, null);
  assert.equal(harness.state.upload.uploadFile, null);
  assert.equal(harness.state.upload.previewUrl, '');
  assert.equal(harness.state.upload.photoSaveIdentity, null);
  assert.equal(harness.state.upload.submittedPhotoPayload, null);
  assert.equal(harness.state.upload.selectionToken, 5);
  assert.equal(harness.elements['upload-overlay'].classList.contains('open'), false);

  harness.state.upload.saving = true;
  harness.state.upload.previewUrl = 'blob:busy';
  harness.open('upload-overlay', createOpener('busy-upload'));
  harness.pressEscape();
  assert.equal(harness.elements['upload-overlay'].classList.contains('open'), true);
  assert.equal(harness.state.upload.previewUrl, 'blob:busy');
  assert.deepEqual(harness.revokedUrls, ['blob:photo']);
});

test('preset editor closes before its manager and a busy editor refuses both closes', () => {
  const harness = createHarness();
  const managerOpener = createOpener('manager');
  const editorOpener = createOpener('editor');
  harness.open('input-presets-overlay', managerOpener);
  harness.context.document.activeElement = editorOpener;
  harness.context.rememberInputPresetEditorOpener();
  harness.state.inputPresets.editing = { name: 'draft' };

  harness.pressEscape();
  assert.equal(harness.state.inputPresets.editing, null);
  assert.equal(harness.elements['input-presets-overlay'].classList.contains('open'), true);
  assert.equal(editorOpener.focusCount, 1);
  assert.equal(managerOpener.focusCount, 0);

  harness.pressEscape();
  assert.equal(harness.elements['input-presets-overlay'].classList.contains('open'), false);

  harness.open('input-presets-overlay', managerOpener);
  harness.context.document.activeElement = editorOpener;
  harness.context.rememberInputPresetEditorOpener();
  harness.state.inputPresets.editing = { name: 'busy' };
  harness.state.inputPresets.saving = true;
  harness.pressEscape();
  assert.notEqual(harness.state.inputPresets.editing, null);
  assert.equal(harness.elements['input-presets-overlay'].classList.contains('open'), true);
});

test('duplicate and normal placement take priority over every overlay', () => {
  const harness = createHarness();
  harness.open('pin-detail-overlay', createOpener('pin'));
  harness.state.duplicatePlacement = { sourcePinId: 'pin-1' };
  harness.elements['duplicate-placement-bar'].classList.add('open');
  harness.pressEscape();
  assert.equal(harness.state.duplicatePlacement, null);
  assert.equal(harness.elements['pin-detail-overlay'].classList.contains('open'), true);

  let removed = 0;
  harness.state.placement = { mode: 'existing', marker: { remove() { removed += 1; } } };
  harness.pressEscape();
  assert.equal(harness.state.placement, null);
  assert.equal(removed, 1);
  assert.equal(harness.elements['pin-detail-overlay'].classList.contains('open'), true);

  harness.pressEscape();
  assert.equal(harness.elements['pin-detail-overlay'].classList.contains('open'), false);
});

test('ordinary overlay precedes help and more menu in the common Escape order', () => {
  const harness = createHarness();
  const overlayOpener = createOpener('overlay');
  const helpOpener = createOpener('help');
  const moreToggle = createOpener('more');
  harness.elements['more-menu-toggle'] = moreToggle;
  harness.open('pin-detail-overlay', overlayOpener);
  harness.open('help-overlay', helpOpener);
  harness.elements['more-menu'].classList.add('open');

  harness.pressEscape();
  assert.equal(harness.elements['pin-detail-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['help-overlay'].classList.contains('open'), true);
  assert.equal(harness.elements['more-menu'].classList.contains('open'), true);

  harness.pressEscape();
  assert.equal(harness.elements['help-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['more-menu'].classList.contains('open'), true);

  harness.pressEscape();
  assert.equal(harness.elements['more-menu'].classList.contains('open'), false);
  assert.equal(moreToggle.focusCount, 1);
});

test('busy import, track, bulk, and share surfaces refuse Escape without focus loss', () => {
  const cases = [
    ['import-preview-overlay', (h) => { h.context.ImportPreviewUI.busy = true; }],
    ['track-import-preview-overlay', (h) => { h.state.trackImport.saving = true; }],
    ['multi-photo-preparation-overlay', (h) => { h.state.multiPhotoImport.registering = true; }],
    ['bulk-metadata-overlay', (h) => { h.state.bulkMetadata.saving = true; }],
    ['share-overlay', (h) => { h.state.shareManager.mutationPending = true; }]
  ];
  cases.forEach(([id, makeBusy]) => {
    const harness = createHarness();
    const opener = createOpener(id);
    makeBusy(harness);
    harness.open(id, opener);
    const event = harness.pressEscape();
    assert.equal(harness.elements[id].classList.contains('open'), true, id);
    assert.equal(opener.focusCount, 0, id);
    assert.equal(event.preventCount, 1, id);
    assert.equal(event.stopCount, 1, id);
  });
});

test('asynchronous photo preparation cancellation transitions without restoring the old opener', () => {
  const harness = createHarness();
  const opener = createOpener('photo-import');
  harness.state.multiPhotoImport.preparing = true;
  harness.open('multi-photo-preparation-overlay', opener);

  const event = harness.pressEscape();
  assert.equal(event.preventCount, 1);
  assert.equal(harness.controllerCalls.preparation, 1);
  assert.equal(harness.elements['multi-photo-preparation-overlay'].classList.contains('open'), true);
  assert.equal(opener.focusCount, 0);

  harness.state.multiPhotoImport.preparing = false;
  harness.context.returnFromMultiPhotoPreparation();
  assert.equal(harness.elements['multi-photo-preparation-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['add-menu-overlay'].classList.contains('open'), true);
  assert.equal(harness.elements['upload-overlay'].classList.contains('open'), false);
  assert.equal(opener.focusCount, 0);
});

test('pin, bulk, and confirmation surfaces use cleanup state and close one at a time', () => {
  const cases = [
    ['pin-detail-overlay', (h) => { h.state.shownPinId = 'pin-1'; }, (h) => h.state.shownPinId === null],
    ['pin-menu-overlay', (h) => { h.state.contextPinId = 'pin-1'; }, (h) => h.state.contextPinId === null],
    ['edit-overlay', (h) => { h.state.editingPinId = 'pin-1'; }, (h) => h.state.editingPinId === null],
    ['bulk-route-add-overlay', (h) => { h.state.bulkRouteAdd.pinIds = ['pin-1']; }, (h) => h.state.bulkRouteAdd.pinIds.length === 0],
    ['bulk-metadata-overlay', (h) => { h.state.bulkMetadata.ids = ['a', 'b']; }, (h) => h.state.bulkMetadata.ids.length === 0],
    ['delete-overlay', (h) => { h.state.bulkDeletePending = true; }, (h) => h.state.bulkDeletePending === false]
  ];
  cases.forEach(([id, arrange, cleaned]) => {
    const harness = createHarness();
    arrange(harness);
    harness.open(id, createOpener(id));
    harness.pressEscape();
    assert.equal(harness.elements[id].classList.contains('open'), false, id);
    assert.equal(cleaned(harness), true, id);
  });
});

test('removed opener is ignored and non-Escape duplicate shortcuts remain wired', () => {
  const harness = createHarness();
  const opener = createOpener('removed');
  harness.open('pin-detail-overlay', opener);
  opener.connected = false;
  assert.doesNotThrow(() => harness.pressEscape());
  assert.equal(opener.focusCount, 0);

  const event = { key: 'c' };
  harness.context.handleGlobalKeydown(event);
  assert.equal(harness.context.shortcutCalls, 1);
  assert.equal(indexHtml.match(/document\.addEventListener\('keydown'/g).length, 1);
  assert.match(functionSource('handleGlobalKeydown'), /handleDuplicateShortcut\(event\)/);
  assert.match(functionSource('handleGlobalKeydown'), /preventDefault\(\)[\s\S]*stopPropagation\(\)/);
});

test('an open geocode result panel consumes Escape before the overlay dispatcher', () => {
  assert.match(
    indexHtml,
    /if \(panelEl\.style\.display !== 'none'\) \{\s*closeCurrentPanel\(\);\s*e\.preventDefault\(\);\s*e\.stopPropagation\(\);/,
  );
});

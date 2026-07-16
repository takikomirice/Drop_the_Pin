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
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

function createPendingHarness() {
  const state = {
    routeSaving: {},
    routePendingSave: {},
    routeOrderSaving: false,
    trackOrderSaving: false,
    routeMapAdd: null,
    upload: { saving: false },
    bulkUnplaceSaving: false,
    bulkMetadata: { saving: false },
    inputPresets: { loading: false, saving: false },
    shareManager: { loading: false },
    bulkDeletePending: false
  };
  const context = {
    state,
    settingsSavePending: false,
    productionImportBusy: false,
    isProductionImportBusy() { return context.productionImportBusy; },
    isInputPresetManagerBusy() {
      return state.inputPresets.loading || state.inputPresets.saving;
    },
    sameRoutePinOrder(left, right) {
      return JSON.stringify(left || []) === JSON.stringify(right || []);
    }
  };
  context.hasRouteSaveInFlight = vm.runInNewContext(`(${functionSource('hasRouteSaveInFlight')})`, context);
  context.hasPendingRouteMutationWork = vm.runInNewContext(`(${functionSource('hasPendingRouteMutationWork')})`, context);
  context.hasPendingMutationWork = vm.runInNewContext(`(${functionSource('hasPendingMutationWork')})`, context);
  return { state, context };
}

function createTransitionHarness() {
  const { state, context } = createPendingHarness();
  const group = { id: 'route-a', routeId: 'route-a', pinIds: ['original', 'draft'] };
  const hints = [];
  const cancelStates = [];
  let cancelCalls = 0;
  let destroySortableCalls = 0;
  Object.assign(state, {
    routeGroups: [group],
    routeMapAdd: {
      routeId: 'route-a',
      originalPinIds: ['original'],
      draftPinIds: ['original', 'draft']
    },
    shareMode: false,
    narrowView: false,
    initializing: false,
    editMode: true,
    previewMode: false,
    routeVisibilityOverrides: {},
    routeEditOpenId: null,
    editingPinId: null,
    selectedPinIds: new Set(),
    placement: null,
    duplicatePlacement: null,
    draggedPinId: null,
    routeDraggedGroupId: null,
    routeDraggedPinId: null,
    listStatusFilter: 'all'
  });
  const elements = new Proxy({}, {
    get(target, key) {
      if (!target[key]) {
        target[key] = {
          classList: { contains() { return false; }, toggle() {}, remove() {} },
          style: {}, value: '', textContent: '', setAttribute() {}
        };
      }
      return target[key];
    }
  });
  Object.assign(context, {
    hasEditToken: true,
    Set,
    document: {
      body: { classList: { toggle() {}, remove() {} } },
      getElementById(id) { return elements[id]; }
    },
    getRouteGroupById(routeId) {
      return routeId === 'route-a' ? state.routeGroups[0] || null : null;
    },
    setRouteGroupPinIds(routeId, pinIds) {
      if (routeId !== 'route-a' || !state.routeGroups.length) return false;
      state.routeGroups[0] = Object.assign({}, state.routeGroups[0], { pinIds: pinIds.slice() });
      return true;
    },
    renderPins() {}, renderSidePanel() {}, renderRouteMapAddMode() {}, renderAccessMode() {},
    updateInputPresetManagerButton() {}, renderCsvInterchangeBusy() {}, renderGeoJsonInterchangeBusy() {},
    renderTrackImportBusy() {}, closeInputPresetManager() {},
    closeRouteEditOverlay() {}, closeBulkRouteAdd() {}, closePinMenu() {}, closeShareQr() {},
    closeShareDialog() {}, closeOverlay() {}, clearPlacement() {}, clearDuplicatePlacement() {},
    cancelRouteDockResize() {}, closeRouteAddMenu() {}, closeRouteTrackImportDialog() {},
    hideDeleteZone() {}, updateBulkBar() {}, showEditAccessNotice() {},
    destroyRouteSortables() { destroySortableCalls += 1; },
    showHint(message) { hints.push(message); }
  });
  const cancel = vm.runInNewContext(`(${functionSource('cancelRouteMapAdd')})`, context);
  context.cancelRouteMapAdd = function() {
    cancelCalls += 1;
    cancelStates.push({ editMode: state.editMode, previewMode: state.previewMode });
    return cancel();
  };
  context.hasPreviewModeBlockingWork = vm.runInNewContext(`(${functionSource('hasPreviewModeBlockingWork')})`, context);
  context.dismissEditOnlySurfacesForPreview = vm.runInNewContext(
    `(${functionSource('dismissEditOnlySurfacesForPreview')})`, context
  );
  context.setPreviewMode = vm.runInNewContext(`(${functionSource('setPreviewMode')})`, context);
  context.setEditMode = vm.runInNewContext(`(${functionSource('setEditMode')})`, context);
  return {
    state, context, hints, cancelStates,
    getCancelCalls: () => cancelCalls,
    getDestroySortableCalls: () => destroySortableCalls
  };
}

test('route pending predicate includes debounce timers, in-flight saves, ordering, and dirty drafts only', () => {
  const { state, context } = createPendingHarness();
  assert.equal(context.hasPendingRouteMutationWork(), false);

  state.routePendingSave['route-a'] = { kind: 'pins', timer: 123 };
  assert.equal(context.hasPendingRouteMutationWork(), true, 'debounce timer');
  delete state.routePendingSave['route-a'];

  state.routeSaving['route-a'] = true;
  assert.equal(context.hasPendingRouteMutationWork(), true, 'in-flight route save');
  delete state.routeSaving['route-a'];

  state.routeOrderSaving = true;
  assert.equal(context.hasPendingRouteMutationWork(), true, 'route order save');
  state.routeOrderSaving = false;

  state.trackOrderSaving = true;
  assert.equal(context.hasPendingRouteMutationWork(), true, 'track order save');
  state.trackOrderSaving = false;

  state.routeMapAdd = { routeId: 'route-a', originalPinIds: ['a'], draftPinIds: ['a', 'b'] };
  assert.equal(context.hasPendingRouteMutationWork(), true, 'dirty route-map draft');
  assert.equal(context.hasPendingRouteMutationWork({ includeRouteMapAddDraft: false }), false,
    'transition check can exclude only the route-map draft');
  assert.equal(context.hasPendingMutationWork({ includeRouteMapAddDraft: false }), false,
    'common transition check can exclude only the route-map draft');
  state.routePendingSave['route-a'] = { kind: 'pins', timer: 123 };
  assert.equal(context.hasPendingMutationWork({ includeRouteMapAddDraft: false }), true,
    'other pending work remains blocking when the draft is excluded');
  delete state.routePendingSave['route-a'];
  state.routeMapAdd.draftPinIds = ['a'];
  assert.equal(context.hasPendingRouteMutationWork(), false, 'unchanged route-map draft');
});

test('dirty route-map draft alone is cancelled before entering preview', () => {
  const harness = createTransitionHarness();

  harness.context.setPreviewMode(true);

  assert.equal(harness.state.previewMode, true);
  assert.equal(harness.state.routeMapAdd, null);
  assert.deepEqual(harness.state.routeGroups[0].pinIds, ['original']);
  assert.equal(harness.getCancelCalls(), 1);
  assert.equal(harness.getDestroySortableCalls(), 1);
  assert.deepEqual(harness.cancelStates, [{ editMode: true, previewMode: false }]);
  assert.deepEqual(harness.hints, []);
});

test('dirty route-map draft alone is cancelled before leaving edit mode', () => {
  const harness = createTransitionHarness();

  harness.context.setEditMode(false);

  assert.equal(harness.state.editMode, false);
  assert.equal(harness.state.routeMapAdd, null);
  assert.deepEqual(harness.state.routeGroups[0].pinIds, ['original']);
  assert.equal(harness.getCancelCalls(), 1);
  assert.equal(harness.getDestroySortableCalls(), 1);
  assert.deepEqual(harness.cancelStates, [{ editMode: true, previewMode: false }]);
  assert.deepEqual(harness.hints, []);
});

test('other pending route save blocks transitions without cancelling the dirty draft', () => {
  ['preview', 'edit'].forEach((transition) => {
    const harness = createTransitionHarness();
    harness.state.routePendingSave['route-a'] = { kind: 'pins', timer: 123 };

    if (transition === 'preview') harness.context.setPreviewMode(true);
    else harness.context.setEditMode(false);

    assert.equal(harness.state.previewMode, false, transition);
    assert.equal(harness.state.editMode, true, transition);
    assert.notEqual(harness.state.routeMapAdd, null, transition);
    assert.deepEqual(harness.state.routeGroups[0].pinIds, ['original', 'draft'], transition);
    assert.equal(harness.getCancelCalls(), 0, transition);
    assert.equal(harness.getDestroySortableCalls(), 0, transition);
    assert.match(harness.hints.at(-1), /完了|保存/, transition);
  });
});

test('common pending predicate aggregates imports, settings, photo, bulk, preset, share, and route work', () => {
  const { state, context } = createPendingHarness();
  const cases = [
    ['production import', () => { context.productionImportBusy = true; }, () => { context.productionImportBusy = false; }],
    ['settings', () => { context.settingsSavePending = true; }, () => { context.settingsSavePending = false; }],
    ['photo', () => { state.upload.saving = true; }, () => { state.upload.saving = false; }],
    ['bulk unplace', () => { state.bulkUnplaceSaving = true; }, () => { state.bulkUnplaceSaving = false; }],
    ['bulk metadata', () => { state.bulkMetadata.saving = true; }, () => { state.bulkMetadata.saving = false; }],
    ['bulk delete', () => { state.bulkDeletePending = true; }, () => { state.bulkDeletePending = false; }],
    ['preset', () => { state.inputPresets.saving = true; }, () => { state.inputPresets.saving = false; }],
    ['share', () => { state.shareManager.loading = true; }, () => { state.shareManager.loading = false; }],
    ['route debounce', () => { state.routePendingSave['route-a'] = { timer: 1 }; }, () => { delete state.routePendingSave['route-a']; }]
  ];

  assert.equal(context.hasPendingMutationWork(), false);
  cases.forEach(([label, enable, disable]) => {
    enable();
    assert.equal(context.hasPendingMutationWork(), true, label);
    disable();
    assert.equal(context.hasPendingMutationWork(), false, `${label} cleanup`);
  });
});

test('beforeunload and preview/edit transitions delegate to the common pending predicate', () => {
  const initialize = functionSource('initializeApp');
  assert.match(initialize, /beforeunload[\s\S]*?hasPendingMutationWork\(\)[\s\S]*?event\.returnValue = ''/);
  assert.match(functionSource('hasPreviewModeBlockingWork'), /hasPendingMutationWork/);
  assert.match(functionSource('setPreviewMode'), /includeRouteMapAddDraft:\s*false/);
  assert.match(functionSource('setEditMode'), /!next[\s\S]*includeRouteMapAddDraft:\s*false/);
});

test('orphan routePinSaving state and references are fully removed', () => {
  assert.equal(indexHtml.includes('routePinSaving'), false);
});

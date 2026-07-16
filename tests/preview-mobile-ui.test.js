const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionBody(name) {
  const start = indexHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${name}`);
}

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

function createNarrowHarness(options = {}) {
  const state = {
    narrowView: false,
    editMode: options.editMode === true,
    previewMode: options.previewMode === true,
    initializing: false,
    routeMapAdd: options.routeMapAdd || null
  };
  let pending = options.pending === true;
  let cancelCalls = 0;
  let dismissCalls = 0;
  let readinessRefreshCalls = 0;
  const context = {
    state,
    document: {
      body: { classList: { toggle() {}, remove() {} } }
    },
    window: { matchMedia() { return { matches: true }; } },
    cancelRouteMapAdd() { cancelCalls += 1; state.routeMapAdd = null; },
    cancelRouteDockResize() {},
    closeRouteAddMenu() {},
    hasPreviewModeBlockingWork() { return pending; },
    dismissEditOnlySurfacesForPreview() { dismissCalls += 1; },
    refreshPinAddButtonState() { readinessRefreshCalls += 1; },
    renderAccessMode() {}, renderPanelToggle() {}, renderMobileSheetState() {},
    renderPins() {}, renderSidePanel() {}, invalidateMapAfterPanelTransition() {}
  };
  const setNarrowView = vm.runInNewContext(
    `(function setNarrowView(next) {${functionBody('setNarrowView')}})`, context
  );
  return {
    state,
    setNarrowView,
    setPending(value) { pending = value; },
    getCancelCalls() { return cancelCalls; },
    getDismissCalls() { return dismissCalls; },
    getReadinessRefreshCalls() { return readinessRefreshCalls; }
  };
}

test('edit preview keeps the edit session while disabling every mutation path', () => {
  assert.match(indexHtml, /previewMode:\s*false/);
  const canEditBody = functionBody('canEdit');
  assert.match(canEditBody, /state\.previewMode\s*!==\s*true/);
  assert.match(canEditBody, /state\.narrowView\s*!==\s*true/);

  const canEdit = vm.runInNewContext(`(function() {${canEditBody}})`, {
    hasEditToken: true,
    state: { initializing: false, editMode: true, previewMode: true, narrowView: false, shareMode: false }
  });
  assert.equal(canEdit(), false, 'preview must expose browse behavior even with a valid edit token');

  const preview = functionBody('setPreviewMode');
  assert.match(preview, /state\.previewMode\s*=\s*next/);
  assert.match(preview, /classList\.toggle\('preview-mode'/);
  assert.match(preview, /renderPins\(\)/);
  assert.match(preview, /renderSidePanel\(\)/);
  assert.match(indexHtml, /プレビュー中/);
  assert.match(indexHtml, /編集に戻る/);
  assert.match(indexHtml, />プレビュー</);
});

test('preview switching is rejected while imports or saves are unsettled', () => {
  const blocking = functionBody('hasPreviewModeBlockingWork');
  assert.match(blocking, /hasPendingMutationWork\(options\)/);
  const common = functionBody('hasPendingMutationWork');
  assert.match(common, /isProductionImportBusy\(\)/);
  assert.match(common, /settingsSavePending/);
  assert.match(common, /state\.upload\.saving/);
  const route = functionBody('hasPendingRouteMutationWork');
  assert.match(route, /hasRouteSaveInFlight\(\)/);
  assert.match(route, /state\.routePendingSave/);

  const preview = functionBody('setPreviewMode');
  assert.match(preview, /hasPreviewModeBlockingWork\(pendingOptions\)/);
  assert.match(preview, /includeRouteMapAddDraft:\s*false/);
  assert.match(preview, /完了までお待ちください/);
  assert.match(preview, /return;/);
});

test('preview and normal URLs do not expose edit-only actions or a persistent readonly label', () => {
  [
    'body.preview-mode #pin-add-btn',
    'body.preview-mode #data-toggle',
    'body.preview-mode #share-open-btn',
    'body.preview-mode #bulk-action-bar',
    'body.preview-mode #route-add-btn',
    'body:not(.has-edit-token) #pin-add-btn',
    'body:not(.has-edit-token) #data-toggle',
    'body:not(.has-edit-token) #share-open-btn',
    'body:not(.has-edit-token) #pin-detail-share'
  ].forEach((selector) => assert.ok(indexHtml.includes(selector), `missing access gate ${selector}`));

  assert.doesNotMatch(indexHtml, /body\.view-only:not\(\.share-mode\) #readonly-banner\s*\{\s*display:\s*block/);
  assert.match(functionBody('updateReadonlyBanner'), /deniedEditUrl[\s\S]*banner\.style\.display\s*=\s*'none'/);
  assert.match(indexHtml, /body\.preview-mode #mode-badge/);
  assert.match(functionBody('openPinDetail'), /openOverlay\('pin-detail-overlay'\)/);
  assert.match(indexHtml, /pin-detail-share[\s\S]*?if \(!canEdit\(\)\) return/);
});

test('narrow view provides the Pencil mobile map, search and bottom sheet structure', () => {
  [
    'mobile-sheet-header',
    'mobile-sheet-handle',
    'mobile-sheet-tabs',
    'mobile-sheet-pin-tab',
    'mobile-sheet-route-tab',
    'mobile-sheet-status',
    'mobile-sheet-retry'
  ].forEach((id) => assert.equal(countId(id), 1, `${id} must exist once`));

  assert.match(indexHtml, /@media \(max-width:\s*640px\)[\s\S]*?#map-search-bar[\s\S]*?top:\s*calc\(var\(--app-header-height\) \+ 12px\)/);
  assert.match(indexHtml, /body\.narrow-view #side-panel\s*\{[\s\S]*?bottom:\s*0[\s\S]*?border-radius:\s*20px 20px 0 0/);
  assert.match(indexHtml, /body\.mobile-sheet-routes #dock-pin-region/);
  assert.match(indexHtml, /body:not\(\.mobile-sheet-routes\) #dock-route-region/);
  assert.match(indexHtml, /body\.narrow-view #pin-detail-overlay\.open[\s\S]*?align-items:\s*flex-end/);
  assert.match(indexHtml, /body\.narrow-view #pin-detail-overlay \.sheet-body[\s\S]*?border-radius:\s*20px 20px 0 0/);
});

test('narrow view hides edit-only controls, keeps core topbar actions, and restores desktop rendering cleanly', () => {
  const canEditBody = functionBody('canEdit');
  const canEdit = vm.runInNewContext(`(function() {${canEditBody}})`, {
    hasEditToken: true,
    state: { initializing: false, editMode: true, previewMode: false, narrowView: true, shareMode: false }
  });
  assert.equal(canEdit(), false);

  const canEditRouteControlsBody = functionBody('canEditRouteControls');
  const canEditRouteControls = vm.runInNewContext(`(function() {${canEditRouteControlsBody}})`, {
    hasEditToken: true,
    state: { initializing: false, editMode: true, previewMode: false, narrowView: true, shareMode: false }
  });
  assert.equal(canEditRouteControls(), true, 'route card controls remain editable on narrow screens');
  assert.doesNotMatch(canEditRouteControlsBody, /narrowView/);

  [
    'body.narrow-view #edit-toggle',
    'body.narrow-view #bulk-action-bar',
    'body.narrow-view #route-add-btn'
  ].forEach((selector) => assert.ok(indexHtml.includes(selector), `missing narrow gate ${selector}`));
  assert.doesNotMatch(indexHtml, /body\.narrow-view #route-map-add-bar,/);
  assert.match(indexHtml, /body\.narrow-view \.mobile-icon-action[\s\S]*?width:\s*44px/);
  assert.match(indexHtml, /body\.narrow-view \.mobile-icon-action \.topbar-action-label\s*\{\s*display:\s*none/);
  assert.doesNotMatch(indexHtml, /body\.narrow-view #(?:data-toggle|share-open-btn),/);

  const viewport = functionBody('setNarrowView');
  assert.match(viewport, /state\.narrowView\s*=\s*narrow/);
  assert.match(viewport, /classList\.toggle\('narrow-view'/);
  assert.match(viewport, /renderPins\(\)/);
  assert.match(viewport, /renderSidePanel\(\)/);
  assert.match(viewport, /invalidateMapAfterPanelTransition\(\)/);
  assert.doesNotMatch(viewport, /state\.previewMode\s*=/);
  assert.doesNotMatch(viewport, /state\.editMode\s*=/);
});

test('narrow is derived presentation state and preserves preview, edit, and browse access modes', () => {
  [
    { label: 'edit preview', editMode: true, previewMode: true },
    { label: 'edit', editMode: true, previewMode: false },
    { label: 'browse', editMode: false, previewMode: false }
  ].forEach((entry) => {
    const harness = createNarrowHarness(entry);
    harness.setNarrowView(true);
    assert.equal(harness.state.narrowView, true, entry.label);
    assert.equal(harness.state.editMode, entry.editMode, entry.label);
    assert.equal(harness.state.previewMode, entry.previewMode, entry.label);

    harness.setNarrowView(false);
    assert.equal(harness.state.narrowView, false, entry.label);
    assert.equal(harness.state.editMode, entry.editMode, entry.label);
    assert.equal(harness.state.previewMode, entry.previewMode, entry.label);
    assert.equal(harness.getReadinessRefreshCalls(), 2, entry.label);
  });
});

test('pending resize preserves access state while narrow entry still cancels route map add', () => {
  const pendingHarness = createNarrowHarness({ editMode: true, previewMode: true, pending: true });
  pendingHarness.setNarrowView(true);
  pendingHarness.setNarrowView(false);
  assert.equal(pendingHarness.state.editMode, true);
  assert.equal(pendingHarness.state.previewMode, true);
  assert.equal(pendingHarness.getDismissCalls(), 0);

  const routeHarness = createNarrowHarness({
    editMode: true,
    previewMode: false,
    routeMapAdd: { routeId: 'route-a', originalPinIds: ['a'], draftPinIds: ['a', 'b'] }
  });
  routeHarness.setNarrowView(true);
  assert.equal(routeHarness.getCancelCalls(), 1);
  assert.equal(routeHarness.state.routeMapAdd, null);
  assert.equal(routeHarness.state.editMode, true);
  assert.equal(routeHarness.state.previewMode, false);
});

test('mobile bottom sheet switches between pin list and all unified routes', () => {
  const setTab = functionBody('setMobileSheetTab');
  assert.match(setTab, /tab === 'routes'/);
  assert.match(setTab, /state\.mobileSheetTab/);
  assert.match(setTab, /setRouteDockExpanded\(true\)/);
  assert.match(setTab, /renderMobileSheetState\(\)/);

  const renderState = functionBody('renderMobileSheetState');
  assert.match(renderState, /mobile-sheet-routes/);
  assert.match(renderState, /aria-selected/);
  assert.match(renderState, /読み込み中/);
  assert.match(renderState, /読み込めませんでした/);
  assert.match(renderState, /state\.viewLoadError/);

  const routePanel = functionBody('renderRoutePanel');
  assert.match(routePanel, /getUnifiedRouteEntriesForPanel\(\)/);
  assert.match(routePanel, /buildRouteItem/);
  assert.match(routePanel, /buildUnifiedTrackRouteItem/);
  assert.match(functionBody('openPinDetail'), /openOverlay\('pin-detail-overlay'\)/);
});

test('viewing cards are keyboard-sized detail links and only editable rows receive drag handles', () => {
  const listItem = functionBody('buildListItem');
  assert.match(listItem, /view-card-chevron/);
  assert.match(listItem, /if \(!canEdit\(\)\)/);
  assert.doesNotMatch(listItem, /dragHandle\.draggable\s*=\s*true/);
  assert.doesNotMatch(listItem, /button\.draggable\s*=\s*true/);
  assert.match(listItem, /aria-label', 'ピンを並べ替え'/);
  assert.match(functionBody('attachPinOrderSortable'), /draggable:\s*'\.pin-list-row'[\s\S]*handle:\s*'\.pin-drag-handle'/);
  assert.match(indexHtml, /\.view-card-chevron/);
  assert.match(indexHtml, /body\.narrow-view #side-panel \.pin-list-row\s*\{[^}]*min-height:\s*66px[^}]*margin-bottom:\s*5px/);
  assert.match(indexHtml, /#mobile-sheet-tabs button\s*\{[\s\S]*?min-height:\s*44px/);
  assert.match(indexHtml, /:where\(a, button, input, textarea, select, \[tabindex\]\):focus-visible/);
  assert.match(indexHtml, /@media \(prefers-reduced-motion:\s*reduce\)/);
});

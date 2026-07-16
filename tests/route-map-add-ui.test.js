const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

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

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function installPlacedRoutePinHelper(context) {
  context.isPlacedRoutePin = vm.runInNewContext(`(${functionSource('isPlacedRoutePin')})`, context);
}

function createAtomicAddHarness(pins, existingPinIds = ['existing']) {
  const group = { id: 'route-a', routeId: 'route-a', pinIds: existingPinIds.slice() };
  const optimisticCalls = [];
  const hints = [];
  const context = {
    canEdit: () => true,
    getRouteGroupById: (routeId) => routeId === 'route-a' ? group : null,
    getRouteId: (routeGroup) => routeGroup && (routeGroup.routeId || routeGroup.id),
    canQueueRoutePinChange: () => true,
    getPinById: (pinId) => pins[String(pinId || '').trim()] || null,
    getRoutePinIdsForState: (routeGroup) => routeGroup.pinIds.slice(),
    updateRoutePinsOptimistic: (routeId, pinIds) => optimisticCalls.push({ routeId, pinIds: pinIds.slice() }),
    showTransientHint: (message) => hints.push(message)
  };
  context.mergeRoutePinIds = vm.runInNewContext(`(${functionSource('mergeRoutePinIds')})`, { Set });
  installPlacedRoutePinHelper(context);
  context.addPinsToRouteAtomically = vm.runInNewContext(`(${functionSource('addPinsToRouteAtomically')})`, context);
  return { context, group, optimisticCalls, hints };
}

function createRouteMapAddSessionHarness() {
  const group = { id: 'route-a', routeId: 'route-a', name: 'Route A', pinIds: ['original'] };
  const pins = {
    original: { id: 'original', lat: 35, lng: 139 },
    first: { id: 'first', lat: 36, lng: 140 },
    second: { id: 'second', lat: 37, lng: 141 },
    outside: { id: 'outside', lat: 38, lng: 142 }
  };
  const undoCalls = [];
  const queueCalls = [];
  const optimisticCalls = [];
  const hints = [];
  const state = {
    routeGroups: [group],
    routeMapAdd: null,
    routeSaving: {},
    routePendingSave: {},
    routeOrderSaving: false,
    routeUndoStack: [],
    placement: null,
    duplicatePlacement: null
  };
  const context = {
    state,
    Set,
    Number,
    isPanelVisible: true,
    canEdit: () => true,
    canEditRouteControls: () => true,
    getRouteGroupById: (routeId) => routeId === 'route-a' ? state.routeGroups[0] || null : null,
    getRouteGroupIndex: (routeId) => routeId === 'route-a' && state.routeGroups.length ? 0 : -1,
    getRouteId: (routeGroup) => routeGroup && (routeGroup.routeId || routeGroup.id),
    getRoutePinIdsForState: (routeGroup) => routeGroup ? routeGroup.pinIds.slice() : [],
    getPinById: (pinId) => pins[String(pinId || '').trim()] || null,
    normalizeRoutePinIds: (pinIds) => Array.from(new Set((pinIds || []).map(String).filter(Boolean))),
    cloneRouteGroupForState: (routeGroup) => Object.assign({}, routeGroup, { pinIds: routeGroup.pinIds.slice() }),
    setRouteGroupPinIds(routeId, pinIds) {
      if (routeId !== 'route-a' || !state.routeGroups.length) return false;
      state.routeGroups[0] = Object.assign({}, state.routeGroups[0], { pinIds: pinIds.slice() });
      return true;
    },
    pushRoutePinUndo: (routeId, pinIds) => undoCalls.push({ routeId, pinIds: pinIds.slice() }),
    queueRoutePinsSave: (routeId, pinIds, rollbackSnapshot) => queueCalls.push({
      routeId,
      pinIds: pinIds.slice(),
      rollbackSnapshot
    }),
    renderPins() {},
    renderSidePanel() {},
    renderRoutePanel() {},
    refreshOpenPinDetailRouteUi() {},
    renderRouteMapAddMode() {},
    setRouteDockExpanded() {},
    renderPanelToggle() {},
    invalidateMapAfterPanelTransition() {},
    showTransientHint: (message) => hints.push(message),
    document: { body: { classList: { remove() {}, toggle() {} } } }
  };
  context.sameRoutePinOrder = vm.runInNewContext(`(${functionSource('sameRoutePinOrder')})`, context);
  context.mergeRoutePinIds = vm.runInNewContext(`(${functionSource('mergeRoutePinIds')})`, context);
  installPlacedRoutePinHelper(context);
  context.isRouteMapAddSessionLocked = function(routeId) {
    return !!(state.routeMapAdd && state.routeMapAdd.routeId === String(routeId || '').trim());
  };
  context.isRouteBusy = vm.runInNewContext(`(${functionSource('isRouteBusy')})`, context);
  context.canQueueRoutePinChange = vm.runInNewContext(`(${functionSource('canQueueRoutePinChange')})`, context);
  const update = vm.runInNewContext(`(${functionSource('updateRoutePinsOptimistic')})`, context);
  context.updateRoutePinsOptimistic = function(routeId, pinIds, options) {
    optimisticCalls.push({ routeId, pinIds: pinIds.slice(), options: options || null });
    return update(routeId, pinIds, options);
  };
  context.addPinsToRouteAtomically = vm.runInNewContext(`(${functionSource('addPinsToRouteAtomically')})`, context);
  context.beginRouteMapAdd = vm.runInNewContext(`(${functionSource('beginRouteMapAdd')})`, context);
  context.cancelRouteMapAdd = vm.runInNewContext(`(${functionSource('cancelRouteMapAdd')})`, context);
  context.completeRouteMapAdd = vm.runInNewContext(`(${functionSource('completeRouteMapAdd')})`, context);
  context.handleRouteMapAddPinClick = vm.runInNewContext(`(${functionSource('handleRouteMapAddPinClick')})`, context);
  return { context, state, pins, undoCalls, queueCalls, optimisticCalls, hints };
}

test('map-add keeps explicit paths while the bulk toolbar route entry is removed', () => {
  ['route-map-add-bar', 'route-map-add-message', 'route-map-add-cancel', 'route-map-add-done',
    'bulk-route-add-overlay', 'bulk-route-add-select',
    'bulk-route-add-cancel', 'bulk-route-add-apply']
    .forEach((id) => assert.equal(countId(id), 1, `${id} must exist exactly once`));
  assert.equal(countId('bulk-route-add-btn'), 0);
  assert.match(indexHtml, /id="route-map-add-bar"[^>]*role="status"/);
  assert.match(indexHtml, /id="bulk-route-add-overlay"[^>]*role="dialog"/);
  assert.doesNotMatch(indexHtml, /bulk-route-add-btn/);
});

test('route pin merge preserves order and prevents duplicates', () => {
  const merge = vm.runInNewContext(`(${functionSource('mergeRoutePinIds')})`);
  const result = merge(['a', 'b', 'a', ''], ['b', 'c', 'c', 'd', null]);
  assert.deepEqual(JSON.parse(JSON.stringify(result)), {
    pinIds: ['a', 'b', 'c', 'd'],
    addedPinIds: ['c', 'd']
  });
});

test('map marker click is intercepted only while route-add mode is active', () => {
  const begin = functionSource('beginRouteMapAdd');
  assert.match(begin, /if \(!canEditRouteControls\(\) \|\| state\.routeMapAdd\) return false/);
  assert.match(begin, /getRouteGroupById/);
  assert.match(begin, /state\.routeMapAdd\s*=/);
  assert.match(begin, /setRouteDockExpanded\(true\)/);

  const marker = functionSource('bindMarkerInteractions');
  assert.match(marker, /if \(state\.routeMapAdd\)/);
  assert.match(marker, /handleRouteMapAddPinClick\(pin\.id\)/);
  assert.ok(marker.indexOf('handleRouteMapAddPinClick(pin.id)') < marker.indexOf('focusPin(pin)'));

  const handle = functionSource('handleRouteMapAddPinClick');
  assert.match(handle, /isPlacedRoutePin\(pin\)/);
  assert.match(handle, /draftPinIds/);
  assert.doesNotMatch(handle, /addPinsToRouteAtomically|updateRoutePinsOptimistic|queueRoutePinsSave/);
  assert.match(handle, /すでにルートに追加済み/);
});

test('map-add begin snapshots original and draft memberships', () => {
  const harness = createRouteMapAddSessionHarness();

  assert.equal(harness.context.beginRouteMapAdd('route-a'), true);

  assert.deepEqual(plain(harness.state.routeMapAdd), {
    routeId: 'route-a',
    originalPinIds: ['original'],
    draftPinIds: ['original']
  });
});

test('map-add begin refuses existing debounce and in-flight route work', () => {
  const harness = createRouteMapAddSessionHarness();
  harness.state.routePendingSave['route-a'] = { kind: 'pins', timer: 1 };
  assert.equal(harness.context.beginRouteMapAdd('route-a'), false);

  delete harness.state.routePendingSave['route-a'];
  harness.state.routeSaving['route-a'] = true;
  assert.equal(harness.context.beginRouteMapAdd('route-a'), false);

  delete harness.state.routeSaving['route-a'];
  assert.equal(harness.context.beginRouteMapAdd('route-a'), true);
});

test('multiple map clicks update only the draft display without Undo or save queue work', () => {
  const harness = createRouteMapAddSessionHarness();
  harness.context.beginRouteMapAdd('route-a');

  assert.equal(harness.context.handleRouteMapAddPinClick('first'), true);
  assert.equal(harness.context.handleRouteMapAddPinClick('second'), true);
  assert.equal(harness.context.handleRouteMapAddPinClick('first'), false);

  assert.deepEqual(Array.from(harness.state.routeMapAdd.draftPinIds || []), ['original', 'first', 'second']);
  assert.deepEqual(plain(harness.state.routeGroups[0].pinIds), ['original', 'first', 'second']);
  assert.deepEqual(harness.undoCalls, []);
  assert.deepEqual(harness.queueCalls, []);
  assert.deepEqual(harness.optimisticCalls, []);
});

test('map click rejects every invalid coordinate without changing draft, Undo, or save queue', () => {
  const invalidPins = [
    { id: 'null-lat', lat: null, lng: 139 },
    { id: 'nan-lng', lat: 35, lng: Number.NaN },
    { id: 'infinite-lat', lat: Number.POSITIVE_INFINITY, lng: 139 },
    { id: 'infinite-lng', lat: 35, lng: Number.NEGATIVE_INFINITY },
    { id: 'lat-high', lat: 90.0001, lng: 139 },
    { id: 'lat-low', lat: -90.0001, lng: 139 },
    { id: 'lng-high', lat: 35, lng: 180.0001 },
    { id: 'lng-low', lat: 35, lng: -180.0001 }
  ];

  invalidPins.forEach((pin) => {
    const harness = createRouteMapAddSessionHarness();
    harness.pins[pin.id] = pin;
    harness.context.beginRouteMapAdd('route-a');

    assert.equal(harness.context.handleRouteMapAddPinClick(pin.id), false, pin.id);
    assert.deepEqual(Array.from(harness.state.routeMapAdd.draftPinIds), ['original'], pin.id);
    assert.deepEqual(plain(harness.state.routeGroups[0].pinIds), ['original'], pin.id);
    assert.deepEqual(harness.undoCalls, [], pin.id);
    assert.deepEqual(harness.queueCalls, [], pin.id);
    assert.deepEqual(harness.optimisticCalls, [], pin.id);
  });
});

test('route add, map click, and D&D use one complete placed-pin predicate', () => {
  const predicate = vm.runInNewContext(`(${functionSource('isPlacedRoutePin')})`, { Number });
  [
    null,
    { id: 'null', lat: null, lng: 139 },
    { id: 'nan', lat: Number.NaN, lng: 139 },
    { id: 'infinity', lat: 35, lng: Number.POSITIVE_INFINITY },
    { id: 'lat-range', lat: -91, lng: 139 },
    { id: 'lng-range', lat: 35, lng: 181 }
  ].forEach((pin) => assert.equal(predicate(pin), false));
  assert.equal(predicate({ id: 'min', lat: -90, lng: -180 }), true);
  assert.equal(predicate({ id: 'max', lat: 90, lng: 180 }), true);

  ['addPinsToRouteAtomically', 'handleRouteMapAddPinClick', 'getRoutePinDropInfo'].forEach((name) => {
    assert.match(functionSource(name), /isPlacedRoutePin\(/, name);
  });

  const pins = {
    invalid: { id: 'invalid', lat: 35, lng: Number.POSITIVE_INFINITY }
  };
  const context = {
    state: { shareMode: false, placement: null, duplicatePlacement: null },
    canEdit: () => true,
    getRouteGroupById: () => ({ id: 'route-a', routeId: 'route-a', pinIds: [] }),
    getRouteId: (group) => group && group.routeId,
    getPinById: (id) => pins[id] || null,
    isRouteBusy: () => false,
    canQueueRoutePinChange: () => true,
    isPinInRoute: () => false,
    Number
  };
  installPlacedRoutePinHelper(context);
  const getDropInfo = vm.runInNewContext(`(${functionSource('getRoutePinDropInfo')})`, context);
  assert.deepEqual(plain(getDropInfo('route-a', 'invalid')), {
    acceptsDrop: false,
    reason: 'unplaced'
  });
});

test('map-add Cancel restores the exact original membership without Undo or save work', () => {
  const harness = createRouteMapAddSessionHarness();
  harness.context.beginRouteMapAdd('route-a');
  harness.context.handleRouteMapAddPinClick('first');
  harness.context.handleRouteMapAddPinClick('second');

  harness.context.cancelRouteMapAdd();

  assert.equal(harness.state.routeMapAdd, null);
  assert.deepEqual(plain(harness.state.routeGroups[0].pinIds), ['original']);
  assert.deepEqual(harness.undoCalls, []);
  assert.deepEqual(harness.queueCalls, []);
  assert.deepEqual(harness.optimisticCalls, []);
});

test('map-add Done commits all draft clicks through one existing optimistic update', () => {
  const harness = createRouteMapAddSessionHarness();
  harness.context.beginRouteMapAdd('route-a');
  harness.context.handleRouteMapAddPinClick('first');
  harness.context.handleRouteMapAddPinClick('second');

  harness.context.completeRouteMapAdd();

  assert.equal(harness.state.routeMapAdd, null);
  assert.deepEqual(plain(harness.state.routeGroups[0].pinIds), ['original', 'first', 'second']);
  assert.equal(harness.optimisticCalls.length, 1);
  assert.equal(harness.undoCalls.length, 1);
  assert.equal(harness.queueCalls.length, 1);
  assert.deepEqual(plain(harness.undoCalls[0]), { routeId: 'route-a', pinIds: ['original'] });
  assert.deepEqual(plain(harness.queueCalls[0].pinIds), ['original', 'first', 'second']);
});

test('map-add Done without changes exits without Undo or save work', () => {
  const harness = createRouteMapAddSessionHarness();
  harness.context.beginRouteMapAdd('route-a');

  harness.context.completeRouteMapAdd();

  assert.equal(harness.state.routeMapAdd, null);
  assert.deepEqual(plain(harness.state.routeGroups[0].pinIds), ['original']);
  assert.deepEqual(harness.undoCalls, []);
  assert.deepEqual(harness.queueCalls, []);
  assert.deepEqual(harness.optimisticCalls, []);
});

test('map-add session lock rejects external membership changes for only the target route', () => {
  const harness = createRouteMapAddSessionHarness();
  harness.context.beginRouteMapAdd('route-a');

  assert.equal(harness.context.canQueueRoutePinChange('route-a'), false);
  assert.deepEqual(plain(harness.context.addPinsToRouteAtomically('route-a', ['outside'])), {
    ok: false,
    reason: 'busy',
    addedPinIds: []
  });
  harness.context.updateRoutePinsOptimistic('route-a', ['original', 'outside']);

  assert.deepEqual(plain(harness.state.routeGroups[0].pinIds), ['original']);
  assert.deepEqual(harness.queueCalls, []);
  assert.match(functionSource('isRouteBusy'), /isRouteMapAddSessionLocked/);
  assert.match(functionSource('deleteRouteGroupFromUi'), /isRouteMapAddSessionLocked/);
});

test('preview, narrow, edit exit, and missing-route cleanup all use Cancel restoration', () => {
  assert.match(functionSource('dismissEditOnlySurfacesForPreview'), /cancelRouteMapAdd/);
  assert.match(functionSource('setEditMode'), /cancelRouteMapAdd/);
  const narrow = functionSource('setNarrowView');
  assert.match(narrow, /cancelRouteMapAdd/);
  assert.ok(narrow.indexOf('cancelRouteMapAdd') < narrow.indexOf('state.narrowView = narrow'));
  const render = functionSource('renderRouteMapAddMode');
  assert.match(render, /cancelRouteMapAdd/);
  assert.doesNotMatch(render, /state\.routeMapAdd\s*=\s*null/);
});

test('multiple pins are submitted as one route update and existing rollback remains active', () => {
  const add = functionSource('addPinsToRouteAtomically');
  assert.match(add, /mergeRoutePinIds/);
  assert.equal((add.match(/updateRoutePinsOptimistic/g) || []).length, 1);
  assert.doesNotMatch(add, /forEach[\s\S]*updateRoutePinsOptimistic/);

  const save = functionSource('saveRoutePinsQueued');
  assert.match(save, /withGAS\('setRoutePins',[\s\S]*routeId:[\s\S]*pinIds:/);
  assert.match(save, /catch \(error\)[\s\S]*rollbackRouteGroup\(routeId, rollbackSnapshot\)/);
});

test('atomic route addition rejects unknown, unplaced, invalid, and mixed candidates before optimistic state', () => {
  const invalidPins = [
    { id: 'null-lat', lat: null, lng: 139 },
    { id: 'null-lng', lat: 35, lng: null },
    { id: 'nan-lat', lat: Number.NaN, lng: 139 },
    { id: 'lat-range', lat: 91, lng: 139 },
    { id: 'lng-range', lat: 35, lng: -181 }
  ];
  invalidPins.forEach((pin) => {
    const harness = createAtomicAddHarness({ [pin.id]: pin });
    assert.deepEqual(plain(harness.context.addPinsToRouteAtomically('route-a', [pin.id])), {
      ok: false,
      reason: 'unplaced',
      addedPinIds: []
    }, pin.id);
    assert.deepEqual(harness.optimisticCalls, [], `${pin.id}: optimistic state must not change`);
  });

  let harness = createAtomicAddHarness({ placed: { id: 'placed', lat: 35, lng: 139 } });
  assert.deepEqual(plain(harness.context.addPinsToRouteAtomically('route-a', ['missing'])), {
    ok: false,
    reason: 'no-pin',
    addedPinIds: []
  });
  assert.deepEqual(harness.optimisticCalls, []);

  harness = createAtomicAddHarness({
    placed: { id: 'placed', lat: 35, lng: 139 },
    unplaced: { id: 'unplaced', lat: null, lng: null }
  });
  assert.deepEqual(plain(harness.context.addPinsToRouteAtomically('route-a', ['placed', 'unplaced'])), {
    ok: false,
    reason: 'unplaced',
    addedPinIds: []
  });
  assert.deepEqual(harness.optimisticCalls, []);

  const busyHarness = createAtomicAddHarness({
    unplaced: { id: 'unplaced', lat: null, lng: null }
  });
  busyHarness.context.canQueueRoutePinChange = () => false;
  assert.deepEqual(plain(busyHarness.context.addPinsToRouteAtomically('route-a', ['unplaced'])), {
    ok: false,
    reason: 'unplaced',
    addedPinIds: []
  });
  assert.deepEqual(busyHarness.optimisticCalls, []);
});

test('placed atomic route addition keeps one update and duplicate behavior', () => {
  const harness = createAtomicAddHarness({
    first: { id: 'first', lat: -90, lng: -180 },
    second: { id: 'second', lat: 90, lng: 180 },
    existing: { id: 'existing', lat: 35, lng: 139 }
  });

  assert.deepEqual(plain(harness.context.addPinsToRouteAtomically('route-a', ['first', 'second'])), {
    ok: true,
    reason: '',
    addedPinIds: ['first', 'second'],
    pinIds: ['existing', 'first', 'second']
  });
  assert.deepEqual(plain(harness.optimisticCalls), [{
    routeId: 'route-a',
    pinIds: ['existing', 'first', 'second']
  }]);

  const duplicateHarness = createAtomicAddHarness({ existing: { id: 'existing', lat: 35, lng: 139 } });
  assert.deepEqual(plain(duplicateHarness.context.addPinsToRouteAtomically('route-a', ['existing'])), {
    ok: false,
    reason: 'duplicate',
    addedPinIds: [],
    pinIds: ['existing']
  });
  assert.deepEqual(duplicateHarness.optimisticCalls, []);
});

test('bulk route addition keeps a mixed selection and reports the unplaced atomic rejection', () => {
  const harness = createAtomicAddHarness({
    placed: { id: 'placed', lat: 35, lng: 139 },
    unplaced: { id: 'unplaced', lat: null, lng: null }
  });
  const selectedPinIds = new Set(['placed', 'unplaced']);
  let closeCalls = 0;
  let bulkBarCalls = 0;
  let renderCalls = 0;
  Object.assign(harness.context, {
    state: { bulkRouteAdd: { pinIds: ['placed', 'unplaced'] }, selectedPinIds },
    document: { getElementById: () => ({ value: 'route-a' }) },
    closeBulkRouteAdd: () => { closeCalls += 1; },
    updateBulkBar: () => { bulkBarCalls += 1; },
    renderSidePanel: () => { renderCalls += 1; },
    Set
  });
  const apply = vm.runInNewContext(`(${functionSource('applyBulkRouteAdd')})`, harness.context);

  apply();

  assert.deepEqual(harness.optimisticCalls, []);
  assert.deepEqual(harness.hints, ['未配置ピンを含むため、ルートへ追加できません。']);
  assert.equal(harness.context.state.selectedPinIds, selectedPinIds);
  assert.equal(closeCalls, 0);
  assert.equal(bulkBarCalls, 0);
  assert.equal(renderCalls, 0);
});

test('bulk route addition snapshots every selected ID instead of silently filtering unknown pins', () => {
  const count = { textContent: '' };
  const select = { options: [{}] };
  const state = {
    selectedPinIds: new Set(['placed', 'missing']),
    bulkRouteAdd: { pinIds: [] }
  };
  const context = {
    state,
    canEdit: () => true,
    getPinById: (pinId) => pinId === 'placed' ? { id: 'placed', lat: 35, lng: 139 } : null,
    renderBulkRouteAddTargets() {},
    document: {
      getElementById(id) {
        if (id === 'bulk-route-add-select') return select;
        if (id === 'bulk-route-add-count') return count;
        return null;
      }
    },
    showTransientHint() {},
    setRouteDockExpanded() {},
    openOverlay() {}
  };
  const open = vm.runInNewContext(`(${functionSource('openBulkRouteAdd')})`, context);

  open();

  assert.deepEqual(plain(state.bulkRouteAdd.pinIds), ['placed', 'missing']);
  assert.equal(count.textContent, '2件を追加します');
});

test('pin detail route addition reports unplaced rejection and keeps successful behavior', () => {
  let harness = createAtomicAddHarness({ unplaced: { id: 'unplaced', lat: null, lng: null } });
  let addOne = vm.runInNewContext(`(${functionSource('addPinToRoute')})`, harness.context);
  assert.deepEqual(plain(addOne('route-a', 'unplaced')), {
    ok: false,
    reason: 'unplaced',
    addedPinIds: []
  });
  assert.deepEqual(harness.optimisticCalls, []);
  assert.deepEqual(harness.hints, ['未配置ピンを含むため、ルートへ追加できません。']);

  harness = createAtomicAddHarness({ placed: { id: 'placed', lat: 35, lng: 139 } });
  addOne = vm.runInNewContext(`(${functionSource('addPinToRoute')})`, harness.context);
  assert.equal(addOne('route-a', 'placed').ok, true);
  assert.equal(harness.optimisticCalls.length, 1);
  assert.deepEqual(harness.hints, []);
});

test('failed atomic save restores the original route snapshot', async () => {
  const rollbackCalls = [];
  const context = {
    state: { routeSaving: {}, routePendingSave: {} },
    normalizeRoutePinIds: (values) => values,
    withGAS: async () => { throw new Error('network'); },
    withEditToken: (payload) => payload,
    renderRoutePanel() {},
    refreshOpenPinDetailRouteUi() {},
    renderPins() {},
    renderSidePanel() {},
    rollbackRouteGroup(routeId, snapshot) { rollbackCalls.push([routeId, snapshot]); },
    clearTimeout() {},
    showAppNotification() {}
  };
  const save = vm.runInNewContext(`(async ${functionSource('saveRoutePinsQueued')})`, context);
  const snapshot = { id: 'route-a', pinIds: ['before'] };

  await save('route-a', { pinIds: ['before', 'new'] }, snapshot);

  assert.deepEqual(JSON.parse(JSON.stringify(rollbackCalls)), [['route-a', snapshot]]);
  assert.equal(context.state.routeSaving['route-a'], undefined);
});

test('server pin_unplaced race rejection rolls optimistic membership back without a partial pending save', async () => {
  const rollbackCalls = [];
  const gasCalls = [];
  const context = {
    state: { routeSaving: {}, routePendingSave: {} },
    normalizeRoutePinIds: (values) => values,
    withGAS: async (method, payload) => {
      gasCalls.push({ method, payload });
      return { ok: false, error: 'pin_unplaced', pinId: 'raced-pin' };
    },
    withEditToken: (payload) => Object.assign({ __editToken: 'edit-token' }, payload),
    renderRoutePanel() {},
    refreshOpenPinDetailRouteUi() {},
    renderPins() {},
    renderSidePanel() {},
    rollbackRouteGroup(routeId, snapshot) { rollbackCalls.push([routeId, snapshot]); },
    clearTimeout() {},
    showAppNotification() {}
  };
  const save = vm.runInNewContext(`(async ${functionSource('saveRoutePinsQueued')})`, context);
  const snapshot = { id: 'route-a', pinIds: ['before'] };

  await save('route-a', { pinIds: ['before', 'raced-pin'] }, snapshot);

  assert.deepEqual(plain(gasCalls), [{
    method: 'setRoutePins',
    payload: { __editToken: 'edit-token', routeId: 'route-a', pinIds: ['before', 'raced-pin'] }
  }]);
  assert.deepEqual(plain(rollbackCalls), [['route-a', snapshot]]);
  assert.equal(context.state.routePendingSave['route-a'], undefined);
  assert.equal(context.state.routeSaving['route-a'], undefined);
});

test('route panel refreshes map-add status after optimistic save or rollback renders', () => {
  assert.match(functionSource('renderRoutePanel'), /renderRouteMapAddMode\(\)/);
});

test('bulk target chooser contains pin-created routes only and applies one atomic addition', () => {
  const render = functionSource('renderBulkRouteAddTargets');
  assert.match(render, /state\.routeGroups/);
  assert.doesNotMatch(render, /state\.tracks/);
  assert.match(render, /getRouteId/);

  const apply = functionSource('applyBulkRouteAdd');
  assert.match(apply, /addPinsToRouteAtomically/);
  assert.equal((apply.match(/addPinsToRouteAtomically/g) || []).length, 1);
  assert.match(apply, /state\.bulkRouteAdd\.pinIds/);
});

test('pin cards are drop targets while imported route cards stay non-targets', () => {
  assert.match(functionSource('buildRouteItem'), /attachRoutePinDropTarget/);
  assert.doesNotMatch(functionSource('buildUnifiedTrackRouteItem'), /attachRoutePinDropTarget/);
  const refresh = functionSource('refreshRoutePinDropTargetsForDraggedPin');
  assert.match(refresh, /\.imported-route-item\[data-route-kind\]/);
  assert.match(refresh, /reason:\s*'imported-route'/);
  assert.match(refresh, /setRoutePinDropTargetState\(target, getRoutePinDropInfo\(routeId, id\), target\.classList\.contains\('route-item'\)\)/);
  assert.match(functionSource('attachPinOrderSortable'), /setRouteDockExpanded\(true\)[\s\S]*refreshRoutePinDropTargetsForDraggedPin/);
});

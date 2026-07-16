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
    if (indexHtml[index] === '}' && --depth === 0) return indexHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

function pinOrderCoreSource() {
  const start = indexHtml.indexOf('const PinOrderCore = (function() {');
  assert.notEqual(start, -1, 'Expected PinOrderCore');
  const end = indexHtml.indexOf('\n    })();', start);
  assert.notEqual(end, -1, 'Expected PinOrderCore closure');
  return indexHtml.slice(start, end + 10);
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function createClassList() {
  const names = new Set();
  return {
    add(...values) { values.forEach((value) => names.add(value)); },
    remove(...values) { values.forEach((value) => names.delete(value)); },
    contains(value) { return names.has(value); }
  };
}

function createRow(id, bucket) {
  const row = {
    dataset: { pinId: id, pinBucket: bucket },
    classList: createClassList(),
    parentNode: null
  };
  row.classList.add('pin-list-row');
  return row;
}

function createContainer(id, bucket, ids) {
  const container = {
    id,
    bucket,
    children: ids.map((pinId) => createRow(pinId, bucket)),
    querySelectorAll(selector) {
      return selector === '.pin-list-row' ? this.children.slice() : [];
    }
  };
  container.children.forEach((row) => { row.parentNode = container; });
  return container;
}

function moveRow(container, fromIndex, toIndex) {
  const [row] = container.children.splice(fromIndex, 1);
  container.children.splice(toIndex, 0, row);
  return row;
}

function createEventTarget(id) {
  const handlers = new Map();
  return {
    id,
    classList: createClassList(),
    handlers,
    addEventListener(type, handler) {
      if (!handlers.has(type)) handlers.set(type, []);
      handlers.get(type).push(handler);
    },
    contains() { return false; },
    async dispatch(type, event) {
      for (const handler of handlers.get(type) || []) await handler(event);
    }
  };
}

function createHarness(options = {}) {
  const pins = options.pins || [
    { id: 'A', lat: 35, lng: 139 },
    { id: 'B', lat: 36, lng: 140 },
    { id: 'C', lat: 37, lng: 141 },
    { id: 'U1', lat: null, lng: null },
    { id: 'U2', lat: null, lng: null }
  ];
  const values = new Map();
  const storage = {
    getItem(key) { return values.has(key) ? values.get(key) : null; },
    setItem(key, value) { values.set(key, value); }
  };
  const counters = { created: 0, destroyed: 0, live: 0, sideRenders: 0, mapRenders: 0 };
  const instances = [];
  const Sortable = {
    create(container, sortableOptions) {
      const instance = {
        container,
        options: sortableOptions,
        destroyed: false,
        destroy() {
          assert.equal(this.destroyed, false, 'Sortable must only be destroyed once');
          this.destroyed = true;
          counters.destroyed += 1;
          counters.live -= 1;
        }
      };
      instances.push(instance);
      counters.created += 1;
      counters.live += 1;
      return instance;
    }
  };
  const targets = Object.fromEntries([
    'side-placed', 'side-unplaced', 'pin-tab-placed', 'pin-tab-unplaced', 'dnd-delete-zone', 'map'
  ].map((id) => [id, createEventTarget(id)]));
  const placed = createContainer('side-placed', 'placed', pins.filter((pin) => pin.lat != null && pin.lng != null).map((pin) => pin.id));
  const unplaced = createContainer('side-unplaced', 'unplaced', pins.filter((pin) => pin.lat == null || pin.lng == null).map((pin) => pin.id));
  targets['side-placed'] = Object.assign(placed, targets['side-placed']);
  targets['side-unplaced'] = Object.assign(unplaced, targets['side-unplaced']);
  const gasCalls = [];
  const state = {
    pins,
    pinOrders: {
      placed: placed.children.map((row) => row.dataset.pinId),
      unplaced: unplaced.children.map((row) => row.dataset.pinId)
    },
    pinOrderSortables: { placed: null, unplaced: null },
    draggedPinId: null,
    draggedPinBucket: null,
    suppressClickUntil: 0,
    placement: null,
    duplicatePlacement: null,
    shareMode: false,
    listQuery: '',
    listStatusFilter: 'all',
    listTagFilter: [],
    listColorFilter: [],
    listIconFilter: [],
    listSort: 'manual'
  };
  const document = {
    getElementById(id) { return targets[id] || null; },
    querySelectorAll(selector) {
      if (selector.includes('.pin-list-row')) return placed.children.concat(unplaced.children);
      return [];
    },
    querySelector() {
      return { textContent: '' };
    }
  };
  const context = {
    state,
    window: { Sortable, localStorage: storage },
    Sortable,
    document,
    console,
    JSON,
    String,
    Array,
    Object,
    Set,
    Map,
    Number,
    Math,
    Date,
    canEdit: () => true,
    getBasePins: () => pins,
    getPinById: (id) => pins.find((pin) => pin.id === id) || null,
    getActivePinFilterState: () => ({
      query: state.listQuery,
      status: state.listStatusFilter,
      tags: state.listTagFilter,
      colors: state.listColorFilter,
      icons: state.listIconFilter
    }),
    showDeleteZone() {},
    hideDeleteZone() {},
    setRouteDockExpanded() {},
    refreshRoutePinDropTargetsForDraggedPin() {},
    clearRoutePinDropTargetStates() {},
    renderPinReorderState() {},
    renderSidePanel() { counters.sideRenders += 1; },
    renderPins() { counters.mapRenders += 1; },
    setPinDockTab() {},
    openOverlay() {},
    showAppNotification() {},
    map: {
      getCenter() { return { lat: 35.5, lng: 139.5 }; },
      getContainer() { return { getBoundingClientRect: () => ({ left: 10, top: 20 }) }; },
      containerPointToLatLng(point) { return { lat: point.y / 10, lng: point.x / 10 }; }
    },
    L: { point(x, y) { return { x, y }; } },
    withEditToken(payload) { return payload; },
    async withGAS(name, payload) {
      gasCalls.push({ name, payload });
      return { ok: true };
    }
  };
  vm.createContext(context);
  const names = [
    'destroySortableInstance',
    'destroyPinOrderSortable',
    'destroyPinOrderSortables',
    'canReorderPinList',
    'persistPinOrderFromContainer',
    'attachPinOrderSortable',
    'setupDndDropTargets',
    'setupMapDropTarget'
  ];
  vm.runInContext(`${pinOrderCoreSource()}\n${names.map(functionSource).join('\n')}\nthis.api = {
    core: PinOrderCore,
    attach: attachPinOrderSortable,
    destroyOne: destroyPinOrderSortable,
    destroyAll: destroyPinOrderSortables,
    setupDnd: setupDndDropTargets,
    setupMap: setupMapDropTarget
  };`, context);
  return { context, state, storage, values, counters, instances, targets, placed, unplaced, gasCalls, api: context.api };
}

function rowIds(container) {
  return container.children.map((row) => row.dataset.pinId);
}

function reorder(instance, container, fromIndex, toIndex) {
  const item = container.children[fromIndex];
  instance.options.onStart({ item, from: container, oldIndex: fromIndex });
  assert.equal(instance.options.onMove({ dragged: item, from: container, to: container }), true);
  moveRow(container, fromIndex, toIndex);
  instance.options.onEnd({ item, from: container, to: container, oldIndex: fromIndex, newIndex: toIndex });
}

test('Sortable callbacks move A, B, C to C, A, B and persist the final DOM order across reload', () => {
  const harness = createHarness();
  const instance = harness.api.attach(harness.placed, 'placed');
  const beforeEnd = Date.now();

  reorder(instance, harness.placed, 2, 0);

  assert.deepEqual(rowIds(harness.placed), ['C', 'A', 'B']);
  assert.deepEqual(plain(harness.state.pinOrders.placed), ['C', 'A', 'B']);
  assert.deepEqual(JSON.parse(harness.values.get(harness.api.core.STORAGE_KEYS.placed)), ['C', 'A', 'B']);
  assert.deepEqual(JSON.parse(harness.values.get(harness.api.core.STORAGE_KEYS.unplaced)), ['U1', 'U2']);
  assert.equal(harness.state.draggedPinId, null);
  assert.equal(harness.state.draggedPinBucket, null);
  assert.ok(harness.state.suppressClickUntil >= beforeEnd);
  assert.equal(harness.placed.children.some((row) => row.classList.contains('is-pin-order-dragging')), false);

  const reloaded = harness.api.core.load(harness.storage);
  assert.deepEqual(
    plain(harness.api.core.orderBucket(harness.state.pins, 'placed', reloaded).map((pin) => pin.id)),
    ['C', 'A', 'B']
  );
});

test('placed and unplaced Sortables persist independent DOM orders', () => {
  const harness = createHarness();
  const placedSortable = harness.api.attach(harness.placed, 'placed');
  const unplacedSortable = harness.api.attach(harness.unplaced, 'unplaced');

  reorder(placedSortable, harness.placed, 2, 0);
  reorder(unplacedSortable, harness.unplaced, 1, 0);

  assert.notEqual(placedSortable, unplacedSortable);
  assert.deepEqual(plain(harness.state.pinOrders), {
    placed: ['C', 'A', 'B'],
    unplaced: ['U2', 'U1']
  });
  assert.deepEqual(JSON.parse(harness.values.get(harness.api.core.STORAGE_KEYS.placed)), ['C', 'A', 'B']);
  assert.deepEqual(JSON.parse(harness.values.get(harness.api.core.STORAGE_KEYS.unplaced)), ['U2', 'U1']);
});

test('search disables same-list movement and clearing it re-enables movement without reload', () => {
  const harness = createHarness();
  harness.state.listQuery = '東京';
  const blocked = harness.api.attach(harness.placed, 'placed');
  const item = harness.placed.children[2];
  blocked.options.onStart({ item, from: harness.placed, oldIndex: 2 });
  assert.equal(blocked.options.sort, false);
  assert.equal(blocked.options.onMove({ dragged: item, from: harness.placed, to: harness.placed }), false);
  blocked.options.onEnd({ item, from: harness.placed, to: harness.placed, oldIndex: 2, newIndex: 2 });
  assert.deepEqual(rowIds(harness.placed), ['A', 'B', 'C']);

  harness.state.listQuery = '';
  const enabled = harness.api.attach(harness.placed, 'placed');
  assert.equal(enabled.options.sort, true);
  reorder(enabled, harness.placed, 2, 0);
  assert.deepEqual(rowIds(harness.placed), ['C', 'A', 'B']);
});

test('reattaching both lists destroys prior Sortables and keeps exactly two live instances', () => {
  const harness = createHarness();
  for (let cycle = 0; cycle < 3; cycle += 1) {
    harness.api.attach(harness.placed, 'placed');
    harness.api.attach(harness.unplaced, 'unplaced');
  }

  assert.equal(harness.counters.created, 6);
  assert.equal(harness.counters.destroyed, 4);
  assert.equal(harness.counters.live, 2);
  assert.equal(harness.state.pinOrderSortables.placed.destroyed, false);
  assert.equal(harness.state.pinOrderSortables.unplaced.destroyed, false);

  assert.equal(harness.api.destroyAll(), 2);
  assert.equal(harness.counters.destroyed, 6);
  assert.equal(harness.counters.live, 0);
  assert.deepEqual(plain(harness.state.pinOrderSortables), { placed: null, unplaced: null });
});

test('same-bucket drop is not consumed by placement-state D&D before Sortable onEnd', async () => {
  const harness = createHarness();
  harness.api.setupDnd();
  harness.state.draggedPinId = 'A';
  harness.state.draggedPinBucket = 'placed';
  let prevented = false;

  await harness.targets['side-placed'].dispatch('drop', {
    preventDefault() { prevented = true; },
    stopPropagation() {}
  });

  assert.equal(prevented, false);
  assert.equal(harness.state.draggedPinId, 'A');
});

test('Sortable onStart still supplies the existing unplaced-pin map drop path', async () => {
  const harness = createHarness();
  harness.state.listQuery = '検索中';
  const sortable = harness.api.attach(harness.unplaced, 'unplaced');
  assert.equal(sortable.options.sort, false);
  harness.api.setupMap();
  const item = harness.unplaced.children[0];
  sortable.options.onStart({ item, from: harness.unplaced, oldIndex: 0 });
  let dragoverPrevented = false;

  await harness.targets.map.dispatch('dragover', {
    preventDefault() { dragoverPrevented = true; },
    dataTransfer: {}
  });
  await harness.targets.map.dispatch('drop', {
    preventDefault() {},
    clientX: 50,
    clientY: 80
  });
  sortable.options.onEnd({ item, from: harness.unplaced, to: harness.unplaced, oldIndex: 0, newIndex: 0 });

  assert.equal(dragoverPrevented, true);
  assert.deepEqual(plain(harness.gasCalls), [{ name: 'movePin', payload: { id: 'U1', lat: 6, lng: 4 } }]);
  assert.equal(harness.state.pins.find((pin) => pin.id === 'U1').lat, 6);
  assert.equal(harness.state.pins.find((pin) => pin.id === 'U1').lng, 4);
  assert.equal(harness.counters.mapRenders, 1);
  assert.equal(harness.state.draggedPinId, null);
});

test('render owns two list Sortables and shared keeps no Sortable or drag handle', () => {
  const render = functionSource('renderSidePanel');
  assert.ok(render.indexOf('destroyPinOrderSortables()') < render.indexOf("unplacedContainer.innerHTML = ''"));
  assert.match(render, /attachPinOrderSortable\(placedContainer,\s*'placed'\)/);
  assert.match(render, /attachPinOrderSortable\(unplacedContainer,\s*'unplaced'\)/);

  const build = functionSource('buildListItem');
  assert.doesNotMatch(build, /\.draggable\s*=\s*true|addEventListener\('dragstart'|addEventListener\('dragend'/);
  const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');
  assert.doesNotMatch(sharedHtml, /Sortable|pin-drag-handle/);
});

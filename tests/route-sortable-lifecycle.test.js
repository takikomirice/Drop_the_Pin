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

function createHarness() {
  const counters = { created: 0, destroyed: 0, live: 0 };
  const instances = [];
  const orderCalls = { routes: [], tracks: [] };
  const Sortable = {
    create(list, options) {
      const instance = {
        kind: options.draggable === '.route-pin-row' ? 'pin' : 'group',
        list,
        options,
        destroyed: false,
        destroy() {
          assert.equal(this.destroyed, false, 'a Sortable instance must be destroyed only once');
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
  const presetSortable = { destroyCalls: 0, destroy() { this.destroyCalls += 1; } };
  const state = {
    routeGroupSortable: null,
    routePinSortables: new Map(),
    routeDraggedPinId: null,
    routeDraggedGroupId: null,
    suppressClickUntil: 0,
    routeGroups: [{ routeId: 'r1' }, { routeId: 'r2' }],
    tracks: [{ trackId: 't1' }, { trackId: 't2' }],
    inputPresets: { sortable: presetSortable }
  };
  const context = {
    state,
    window: { Sortable },
    Sortable,
    console,
    canEdit: () => true,
    canSortRoutePins: () => true,
    canSortRouteGroups: () => true,
    renderRoutePanel() {},
    getRouteGroupIdsForState: () => state.routeGroups.map((group) => group.routeId),
    getTrackIdsForState: () => state.tracks.map((track) => track.trackId),
    getRouteGroupOrderFromList: (list) => list.querySelectorAll('.route-item').map((item) => item.dataset.routeId),
    getTrackOrderFromList: (list) => list.querySelectorAll('.imported-route-item').map((item) => item.dataset.trackId),
    sameRouteGroupOrder: (left, right) => JSON.stringify(left) === JSON.stringify(right),
    sameTrackOrder: (left, right) => JSON.stringify(left) === JSON.stringify(right),
    updateRoutesOrderOptimistic: (ids) => orderCalls.routes.push(ids.slice()),
    updateTracksOrderOptimistic: (ids) => orderCalls.tracks.push(ids.slice())
  };
  const names = [
    'destroySortableInstance', 'destroyRouteGroupSortable', 'destroyRoutePinSortable',
    'destroyRoutePinSortables', 'destroyRouteSortables',
    'attachRoutePinSortable', 'attachRouteGroupSortable'
  ];
  vm.runInNewContext(`${names.map(functionSource).join('\n')}
    this.api = {
      destroyGroup: destroyRouteGroupSortable,
      destroyPin: destroyRoutePinSortable,
      destroyPins: destroyRoutePinSortables,
      destroyAll: destroyRouteSortables,
      attachPin: attachRoutePinSortable,
      attachGroup: attachRouteGroupSortable
    };`, context);
  return { context, state, counters, instances, presetSortable, orderCalls };
}

function renderCycle(harness) {
  harness.context.api.destroyAll();
  harness.context.api.attachGroup({ id: 'group-list' }, [{ id: 'r1' }, { id: 'r2' }]);
  harness.context.api.attachPin({ id: 'pin-list-r1' }, 'r1', ['p1', 'p2']);
  harness.context.api.attachPin({ id: 'pin-list-r2' }, 'r2', ['p3', 'p4']);
}

test('three route renders destroy old Sortables and keep only the required live instances', () => {
  const harness = createHarness();
  renderCycle(harness);
  assert.equal(harness.counters.created, 3);
  assert.equal(harness.counters.destroyed, 0);
  assert.equal(harness.counters.live, 3);

  renderCycle(harness);
  renderCycle(harness);
  assert.equal(harness.counters.created, 9);
  assert.equal(harness.counters.destroyed, 6);
  assert.equal(harness.counters.live, 3);
  assert.equal(harness.state.routeGroupSortable.kind, 'group');
  assert.equal(harness.state.routePinSortables.size, 2);
  assert.equal(harness.state.routePinSortables.get('r1').kind, 'pin');
  assert.equal(harness.state.routePinSortables.get('r2').kind, 'pin');
});

test('route deletion and edit exit destroy the right instances without touching preset Sortable', () => {
  const harness = createHarness();
  renderCycle(harness);
  const group = harness.state.routeGroupSortable;
  const route1 = harness.state.routePinSortables.get('r1');
  const route2 = harness.state.routePinSortables.get('r2');

  harness.context.api.destroyPin('r1');
  assert.equal(route1.destroyed, true);
  assert.equal(route2.destroyed, false);
  assert.equal(group.destroyed, false);
  assert.equal(harness.state.routePinSortables.has('r1'), false);
  assert.equal(harness.counters.live, 2);

  harness.context.api.destroyAll();
  assert.equal(group.destroyed, true);
  assert.equal(route2.destroyed, true);
  assert.equal(harness.state.routeGroupSortable, null);
  assert.equal(harness.state.routePinSortables.size, 0);
  assert.equal(harness.counters.live, 0);
  assert.equal(harness.presetSortable.destroyCalls, 0);
  assert.equal(harness.state.inputPresets.sortable, harness.presetSortable);
});

test('reattaching one target replaces its instance and Sortable absence is safe', () => {
  const harness = createHarness();
  const groupList = { id: 'group-list' };
  const pinList = { id: 'pin-list' };
  harness.context.api.attachGroup(groupList, [{ id: 'r1' }, { id: 'r2' }]);
  const oldGroup = harness.state.routeGroupSortable;
  harness.context.api.attachGroup(groupList, [{ id: 'r1' }, { id: 'r2' }]);
  assert.equal(oldGroup.destroyed, true);

  harness.context.api.attachPin(pinList, 'r1', ['p1', 'p2']);
  const oldPin = harness.state.routePinSortables.get('r1');
  harness.context.api.attachPin(pinList, 'r1', ['p1', 'p2']);
  assert.equal(oldPin.destroyed, true);
  assert.equal(harness.counters.live, 2);

  harness.context.window.Sortable = null;
  harness.context.Sortable = null;
  assert.doesNotThrow(() => harness.context.api.attachGroup(groupList, [{ id: 'r1' }, { id: 'r2' }]));
  assert.doesNotThrow(() => harness.context.api.attachPin(pinList, 'r1', ['p1', 'p2']));
  assert.equal(harness.state.routeGroupSortable, null);
  assert.equal(harness.state.routePinSortables.size, 0);
  assert.equal(harness.counters.live, 0);
});

test('unified route Sortable reorders within sections and blocks pin-track crossing', () => {
  const harness = createHarness();
  harness.context.api.attachGroup(
    { id: 'route-list' },
    [{ id: 'r1' }, { id: 'r2' }],
    [{ trackId: 't1' }, { trackId: 't2' }]
  );
  const options = harness.state.routeGroupSortable.options;

  assert.equal(options.draggable, '.route-item, .imported-route-item');
  assert.equal(options.onMove({
    dragged: { dataset: { routeKind: 'pin' } },
    related: { dataset: { routeKind: 'pin' } }
  }), true);
  assert.equal(options.onMove({
    dragged: { dataset: { routeKind: 'gpx' } },
    related: { dataset: { routeKind: 'geojson' } }
  }), true);
  assert.equal(options.onMove({
    dragged: { dataset: { routeKind: 'pin' } },
    related: { dataset: { routeKind: 'gpx' } }
  }), false);
});

test('unified route Sortable persists imported card order through the track API path', () => {
  const harness = createHarness();
  const list = {
    querySelectorAll(selector) {
      if (selector === '.imported-route-item') {
        return [
          { dataset: { trackId: 't2', routeKind: 'geojson' } },
          { dataset: { trackId: 't1', routeKind: 'gpx' } }
        ];
      }
      return [
        { dataset: { routeId: 'r1', routeKind: 'pin' } },
        { dataset: { routeId: 'r2', routeKind: 'pin' } }
      ];
    }
  };
  harness.context.api.attachGroup(list, harness.state.routeGroups, harness.state.tracks);

  harness.state.routeGroupSortable.options.onEnd({ item: { dataset: { routeKind: 'gpx' } } });

  assert.deepEqual(harness.orderCalls.tracks, [['t2', 't1']]);
  assert.deepEqual(harness.orderCalls.routes, []);
});

test('render, delete, mode transitions and page lifecycle use the route destroy contract', () => {
  assert.match(indexHtml, /routeGroupSortable:\s*null/);
  assert.match(indexHtml, /routePinSortables:\s*new Map\(\)/);

  const render = functionSource('renderRoutePanel');
  assert.ok(render.indexOf('destroyRouteSortables()') < render.indexOf("list.innerHTML = ''"));
  assert.match(functionSource('deleteRouteGroupFromUi'), /destroyRoutePinSortable\(routeId\)/);
  assert.match(functionSource('dismissEditOnlySurfacesForPreview'), /destroyRouteSortables\(\)/);
  assert.match(functionSource('setEditMode'), /if \(!next\)[\s\S]*destroyRouteSortables\(\)/);
  assert.match(functionSource('enterShareMode'), /destroyRouteSortables\(\)/);
  assert.match(functionSource('initializeApp'), /addEventListener\('pagehide',\s*destroyRouteSortables/);

  const preset = functionSource('attachInputPresetSortable');
  assert.match(preset, /state\.inputPresets\.sortable\s*=\s*Sortable\.create/);
  assert.doesNotMatch(preset, /route(?:Group|Pin)Sortable/);
});

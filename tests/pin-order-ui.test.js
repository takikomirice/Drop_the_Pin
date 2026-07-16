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

function loadCore() {
  const start = indexHtml.indexOf('const PinOrderCore = (function() {');
  assert.notEqual(start, -1, 'Expected PinOrderCore');
  const end = indexHtml.indexOf('\n    })();', start);
  assert.notEqual(end, -1, 'Expected PinOrderCore closure');
  const context = { JSON, String, Array, Object, Set };
  vm.createContext(context);
  vm.runInContext(`${indexHtml.slice(start, end + 10)}\nglobalThis.__core = PinOrderCore;`, context);
  return context.__core;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

test('placed and unplaced orders reconcile independently with new and deleted pins', () => {
  const core = loadCore();
  const pins = [
    { id: 'placed-a', lat: 35, lng: 139 },
    { id: 'unplaced-a', lat: null, lng: null },
    { id: 'placed-new', lat: 36, lng: 140 },
    { id: 'unplaced-new', lat: null, lng: null }
  ];
  const reconciled = core.reconcile(pins, {
    placed: ['deleted', 'placed-a', 'placed-a', 'unplaced-a'],
    unplaced: ['deleted', 'unplaced-a', 'placed-a']
  });

  assert.deepEqual(plain(reconciled), {
    placed: ['placed-a', 'placed-new'],
    unplaced: ['unplaced-a', 'unplaced-new']
  });
  assert.deepEqual(
    plain(core.orderBucket(pins, 'placed', reconciled).map((pin) => pin.id)),
    ['placed-a', 'placed-new']
  );
  assert.deepEqual(
    plain(core.orderBucket(pins, 'unplaced', reconciled).map((pin) => pin.id)),
    ['unplaced-a', 'unplaced-new']
  );
});

test('manual reorder persists as separate pin-id arrays and survives storage errors', () => {
  const core = loadCore();
  const values = new Map();
  const storage = {
    getItem(key) { return values.has(key) ? values.get(key) : null; },
    setItem(key, value) { values.set(key, value); }
  };
  let orders = { placed: ['p1', 'p2', 'p3'], unplaced: ['u1', 'u2'] };
  orders = core.move(orders, 'placed', 'p3', 0);
  assert.deepEqual(plain(orders), { placed: ['p3', 'p1', 'p2'], unplaced: ['u1', 'u2'] });
  assert.equal(core.save(storage, orders), true);
  assert.deepEqual(JSON.parse(values.get(core.STORAGE_KEYS.placed)), ['p3', 'p1', 'p2']);
  assert.deepEqual(JSON.parse(values.get(core.STORAGE_KEYS.unplaced)), ['u1', 'u2']);
  assert.deepEqual(plain(core.load(storage)), plain(orders));

  const blocked = { getItem() { throw new Error('blocked'); }, setItem() { throw new Error('blocked'); } };
  assert.deepEqual(plain(core.load(blocked)), { placed: [], unplaced: [] });
  assert.equal(core.save(blocked, orders), false);
});

test('a pin moved between tabs is removed from the source and appended to the destination', () => {
  const core = loadCore();
  const before = { placed: ['p1', 'moved', 'p2'], unplaced: ['u1'] };
  const pins = [
    { id: 'p1', lat: 35, lng: 139 },
    { id: 'p2', lat: 36, lng: 140 },
    { id: 'u1', lat: null, lng: null },
    { id: 'moved', lat: null, lng: null }
  ];
  assert.deepEqual(plain(core.reconcile(pins, before)), {
    placed: ['p1', 'p2'],
    unplaced: ['u1', 'moved']
  });
});

test('search and filters expose a visible reason and disable only same-tab reordering', () => {
  const core = loadCore();
  assert.equal(core.getBlockReason({ query: '', status: 'all', tags: [], colors: [], icons: [] }, 'manual'), '');
  assert.match(core.getBlockReason({ query: '東京', status: 'all', tags: [], colors: [], icons: [] }, 'manual'), /検索・絞り込み中/);
  assert.match(core.getBlockReason({ query: '', status: '完了', tags: [], colors: [], icons: [] }, 'manual'), /検索・絞り込み中/);
  assert.match(core.getBlockReason({ query: '', status: 'all', tags: [], colors: [], icons: [] }, 'title'), /手動順/);

  assert.match(indexHtml, /id="pin-reorder-note"[^>]*role="status"/);
  assert.match(functionSource('renderPinReorderState'), /getBlockReason/);
  assert.match(functionSource('attachPinOrderSortable'), /sort:\s*sortEnabled/);
  assert.match(functionSource('attachPinOrderSortable'), /onMove:[\s\S]*canReorderPinList/);
  assert.match(functionSource('attachPinOrderSortable'), /onEnd:[\s\S]*persistPinOrderFromContainer/);
  assert.doesNotMatch(functionSource('setupDndDropTargets'), /setupPinOrderDropTarget/);
});

test('manual order is the index default while shared remains untouched', () => {
  assert.match(indexHtml, /<option value="manual">手動順<\/option>/);
  assert.match(indexHtml, /listSort:\s*'manual'/);
  assert.match(functionSource('renderSidePanel'), /reconcilePinOrders\(getBasePins\(\)\)/);
  assert.match(functionSource('renderSidePanel'), /PinOrderCore\.orderBucket/);

  const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');
  assert.equal(sharedHtml.includes('PinOrderCore'), false);
  assert.equal(sharedHtml.includes('pin-drag-handle'), false);
  assert.equal(sharedHtml.includes('pin-add-btn'), false);
});

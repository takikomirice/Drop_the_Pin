const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');

function functionSource(name) {
  const start = sharedHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = sharedHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < sharedHtml.length; index += 1) {
    if (sharedHtml[index] === '{') depth += 1;
    if (sharedHtml[index] === '}') depth -= 1;
    if (depth === 0) return sharedHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse function ${name}`);
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

function sameRealm(value) {
  return JSON.parse(JSON.stringify(value));
}

function createHarness(responseHandler, roadRouteCache = {}) {
  const calls = [];
  let handler = responseHandler;
  const context = {
    state: { token: 'share-token', roadRouteCache },
    sharedRoadRoutePendingRequests: Object.create(null),
    withGAS(method, payload) {
      calls.push({ method, payload });
      return handler(method, payload, calls.length);
    }
  };
  vm.runInNewContext([
    functionSource('normalizeSharedRoadRouteCoords'),
    functionSource('getSharedCachedRoadRouteCoords'),
    functionSource('requestSharedRoadRouteCache'),
    'this.request = requestSharedRoadRouteCache;'
  ].join('\n'), context);
  return {
    request: context.request,
    pending: context.sharedRoadRoutePendingRequests,
    calls,
    setHandler(nextHandler) { handler = nextHandler; }
  };
}

test('same unresolved route ID shares one Promise and one GAS request', async () => {
  assert.match(sharedHtml, /let sharedRoadRoutePendingRequests = Object\.create\(null\);/);
  const gate = deferred();
  const harness = createHarness(() => gate.promise);

  const first = harness.request(' route-1 ');
  const second = harness.request('route-1');
  await Promise.resolve();

  assert.strictEqual(first, second);
  assert.equal(harness.calls.length, 1);
  assert.deepEqual(sameRealm(harness.calls[0]), {
    method: 'getSharedRoadRouteCache',
    payload: { token: 'share-token', routeId: 'route-1' }
  });
  assert.strictEqual(harness.pending['route-1'], first);

  gate.resolve({ ok: true, coords: [[35, 139], [35.1, 139.1]] });
  await first;
});

test('successful shared request resolves both callers with one coords object and clears pending state', async () => {
  const gate = deferred();
  const harness = createHarness(() => gate.promise);
  const first = harness.request('route-success');
  const second = harness.request('route-success');
  await Promise.resolve();

  gate.resolve({ ok: true, coords: [{ lat: '35', lng: '139' }, [35.1, 139.1]] });
  const [firstCoords, secondCoords] = await Promise.all([first, second]);

  assert.strictEqual(firstCoords, secondCoords);
  assert.deepEqual(sameRealm(firstCoords), [[35, 139], [35.1, 139.1]]);
  assert.equal(Object.hasOwn(harness.pending, 'route-success'), false);
  assert.equal(harness.calls.length, 1);
});

test('failed shared request rejects both callers and clears pending state', async () => {
  const gate = deferred();
  const harness = createHarness(() => gate.promise);
  const first = harness.request('route-failure');
  const second = harness.request('route-failure');
  await Promise.resolve();

  const failure = new Error('network failed');
  gate.reject(failure);
  const results = await Promise.allSettled([first, second]);

  assert.strictEqual(first, second);
  assert.deepEqual(results.map((result) => result.status), ['rejected', 'rejected']);
  assert.strictEqual(results[0].reason, failure);
  assert.strictEqual(results[1].reason, failure);
  assert.equal(Object.hasOwn(harness.pending, 'route-failure'), false);
  assert.equal(harness.calls.length, 1);
});

test('a failed route request can retry with a new GAS call and succeed', async () => {
  const firstGate = deferred();
  const harness = createHarness(() => firstGate.promise);
  const first = harness.request('route-retry');
  await Promise.resolve();
  firstGate.reject(new Error('temporary failure'));
  await assert.rejects(first, /temporary failure/);
  assert.equal(Object.hasOwn(harness.pending, 'route-retry'), false);

  harness.setHandler(() => Promise.resolve({ ok: true, coords: [[36, 140], [36.1, 140.1]] }));
  const retryCoords = await harness.request('route-retry');

  assert.deepEqual(sameRealm(retryCoords), [[36, 140], [36.1, 140.1]]);
  assert.equal(harness.calls.length, 2);
  assert.equal(Object.hasOwn(harness.pending, 'route-retry'), false);
});

test('completed cache returns normalized coords without a GAS request', async () => {
  const harness = createHarness(
    () => Promise.reject(new Error('must not request')),
    { cached: { ok: true, coords: [{ lat: '34', lng: '138' }, [34.1, 138.1], ['bad', 1]] } }
  );

  const coords = await harness.request(' cached ');

  assert.deepEqual(sameRealm(coords), [[34, 138], [34.1, 138.1]]);
  assert.equal(harness.calls.length, 0);
  assert.equal(Object.keys(harness.pending).length, 0);
});

test('empty route ID skips communication and resolves null', async () => {
  const harness = createHarness(() => Promise.reject(new Error('must not request')));

  const coords = await harness.request('   ');

  assert.equal(coords, null);
  assert.equal(harness.calls.length, 0);
  assert.equal(Object.keys(harness.pending).length, 0);
});

test('different route IDs keep independent pending requests', async () => {
  const gates = { first: deferred(), second: deferred() };
  const harness = createHarness((_method, payload) => gates[payload.routeId].promise);

  const first = harness.request('first');
  const second = harness.request('second');
  await Promise.resolve();

  assert.notStrictEqual(first, second);
  assert.equal(harness.calls.length, 2);
  assert.strictEqual(harness.pending.first, first);
  assert.strictEqual(harness.pending.second, second);

  gates.first.resolve({ ok: true, coords: [[35, 139], [35.1, 139.1]] });
  gates.second.resolve({ ok: true, coords: [[36, 140], [36.1, 140.1]] });
  await Promise.all([first, second]);
  assert.equal(Object.keys(harness.pending).length, 0);
});

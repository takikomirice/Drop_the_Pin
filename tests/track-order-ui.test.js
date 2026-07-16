const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionSource(name) {
  const asyncStart = indexHtml.indexOf(`async function ${name}(`);
  const start = asyncStart === -1 ? indexHtml.indexOf(`function ${name}(`) : asyncStart;
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

function createHarness(options = {}) {
  const state = {
    tracks: [
      { id: 'track-a', trackId: 'track-a', orderIndex: 0 },
      { id: 'track-b', trackId: 'track-b', orderIndex: 1 }
    ],
    trackOrderSaving: false
  };
  const calls = [];
  const notifications = [];
  const context = {
    state,
    canEdit: () => true,
    getTrackId: (track) => String(track && (track.trackId || track.id) || ''),
    withEditToken: (payload) => ({ ...payload, __editToken: 'fresh-token' }),
    withGAS(method, payload) {
      calls.push({ method, payload });
      if (options.reject) return Promise.reject(new Error('raw order failure'));
      return Promise.resolve({ ok: true, tracks: state.tracks.map((track) => ({ ...track })) });
    },
    renderTrackLayers() {},
    renderTrackPanel() {},
    renderTrackImportBusy() {},
    showAppNotification(value) { notifications.push(value); }
  };
  const names = [
    'getTrackIdsForState', 'sameTrackOrder', 'setTracksOrder', 'rollbackTracksOrder',
    'saveTracksOrder', 'updateTracksOrderOptimistic'
  ];
  vm.runInNewContext(`${names.map(functionSource).join('\n')}\nthis.api = { updateTracksOrderOptimistic };`, context);
  return { api: context.api, state, calls, notifications };
}

test('imported route ordering uses a guarded updateTracksOrder mutation and a dedicated pending state', () => {
  assert.match(indexHtml, /trackOrderSaving:\s*false/);
  assert.match(functionSource('saveTracksOrder'), /withGAS\('updateTracksOrder',\s*withEditToken\(/);
  assert.match(functionSource('attachRouteGroupSortable'), /\.route-item, \.imported-route-item/);
  assert.match(functionSource('attachRouteGroupSortable'), /updateTracksOrderOptimistic/);
});

test('optimistic imported route reorder persists and keeps the returned formal order', async () => {
  const harness = createHarness();

  await harness.api.updateTracksOrderOptimistic(['track-b', 'track-a']);

  assert.deepEqual(harness.state.tracks.map((track) => [track.trackId, track.orderIndex]), [
    ['track-b', 0], ['track-a', 1]
  ]);
  assert.deepEqual(harness.calls, [{
    method: 'updateTracksOrder',
    payload: { orderedIds: ['track-b', 'track-a'], __editToken: 'fresh-token' }
  }]);
  assert.equal(harness.state.trackOrderSaving, false);
});

test('failed imported route reorder restores the previous order without exposing raw errors', async () => {
  const harness = createHarness({ reject: true });

  await harness.api.updateTracksOrderOptimistic(['track-b', 'track-a']);

  assert.deepEqual(harness.state.tracks.map((track) => track.trackId), ['track-a', 'track-b']);
  assert.equal(JSON.stringify(harness.notifications).includes('raw order failure'), false);
  assert.match(harness.notifications.at(-1).message, /順序を保存できません/);
  assert.equal(harness.state.trackOrderSaving, false);
});

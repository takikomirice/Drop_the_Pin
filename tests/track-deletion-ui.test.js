const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

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

function deferred() {
  let resolve;
  const promise = new Promise((done) => { resolve = done; });
  return { promise, resolve };
}

function createHarness(options = {}) {
  const layer = {
    removed: false,
    remove() {
      if (options.layerRemoveFails) throw new Error('raw leaflet failure');
      this.removed = true;
    }
  };
  const track = {
    id: 'track-a', trackId: 'track-a', name: '<Track A>', sourceType: 'gpx',
    orderIndex: 0, visible: true, segments: []
  };
  const state = {
    tracks: [track],
    trackLayers: { 'track-a': layer },
    trackVisibilityOverrides: { 'track-a': false },
    trackDeleteConfirming: '',
    trackDeletePending: '',
    routeUiOpenKey: 'gpx:track-a',
    routeUiOpenGroupId: null,
    multiPhotoImport: {
      trackMatch: {
        enabled: !!options.matching,
        open: !!options.matching,
        running: !!options.matching,
        trackId: options.matching ? 'track-a' : '',
        trackRevisionId: options.matching ? 'rev-a' : '',
        result: options.matching ? { trackId: 'track-a' } : null
      }
    }
  };
  const confirmations = [];
  const gasCalls = [];
  const notifications = [];
  let renders = 0;
  const context = {
    state,
    canEdit: () => options.canEdit !== false,
    canEditRouteControls: () => options.canEdit !== false,
    getTrackId: (value) => value ? String(value.trackId || value.id || '') : '',
    getTrackById: (id) => state.tracks.find((value) => value.trackId === id) || null,
    isProductionImportBusy: () => !!options.importBusy,
    hasPendingRouteMutationWork: () => !!options.routeBusy,
    requestAppConfirmation(config) {
      confirmations.push(config);
      return options.confirmationPromise || Promise.resolve(options.confirm !== false);
    },
    withEditToken: (payload) => ({ ...payload, __editToken: 'fresh-token' }),
    withGAS(method, payload) {
      gasCalls.push({ method, payload });
      return Promise.resolve(options.result || { ok: true, deleted: true, removedShareReferences: 1 });
    },
    renderTrackPanel() { renders += 1; },
    renderTrackLayers() { renders += 1; },
    renderTrackImportBusy() { renders += 1; },
    showAppNotification(config) { notifications.push(config); return Promise.resolve(true); }
  };
  const names = [
    'isTrackSelectedForPhotoMatch', 'trackDeletionFailureMessage',
    'applyDeletedTrackToClient', 'deleteTrackFromUi'
  ];
  vm.runInNewContext(`${names.map(functionSource).join('\n')}\nthis.api = { deleteTrackFromUi };`, context);
  return {
    api: context.api, state, layer, confirmations, gasCalls, notifications,
    renderCount: () => renders
  };
}

test('GeoJSON and GPX unified cards expose icon-only edit and delete controls while Shared stays read-only', () => {
  const source = functionSource('buildUnifiedTrackRouteItem');
  assert.match(source, /canEditRouteControls\(\)/);
  assert.match(source, /openTrackDisplaySettingsEditor\(trackId\)/);
  assert.match(source, /deleteTrackFromUi\(trackId\)/);
  assert.match(source, /createActionIconElement\(document, 'settings'\)/);
  assert.match(source, /createActionIconElement\(document, 'delete'\)/);
  assert.doesNotMatch(source, /textContent = '削除'/);
  assert.match(source, /trackDeletePending/);
  assert.doesNotMatch(sharedHtml, /deleteTrackFromUi|deleteTrack\s*\(/);
});

test('cancellation makes no server call and confirmation contains the safe track name and scope', async () => {
  const harness = createHarness({ confirm: false });

  assert.equal(await harness.api.deleteTrackFromUi('track-a'), false);

  assert.equal(harness.gasCalls.length, 0);
  assert.equal(harness.state.tracks.length, 1);
  assert.equal(harness.confirmations.length, 1);
  assert.match(harness.confirmations[0].message, /<Track A>/);
  assert.match(harness.confirmations[0].message, /保存済みのGeoJSON／GPXルートデータ/);
});

test('double activation confirms and sends deleteTrack only once with a fresh token', async () => {
  const gate = deferred();
  const harness = createHarness({ confirmationPromise: gate.promise });

  const first = harness.api.deleteTrackFromUi('track-a');
  const second = harness.api.deleteTrackFromUi('track-a');
  gate.resolve(true);

  assert.equal(await second, false);
  assert.equal(await first, true);
  assert.deepEqual(harness.gasCalls, [{
    method: 'deleteTrack', payload: { trackId: 'track-a', __editToken: 'fresh-token' }
  }]);
});

test('success removes state layer and transient references immediately', async () => {
  const harness = createHarness();

  assert.equal(await harness.api.deleteTrackFromUi('track-a'), true);

  assert.deepEqual(harness.state.tracks, []);
  assert.equal(harness.layer.removed, true);
  assert.equal(Object.prototype.hasOwnProperty.call(harness.state.trackLayers, 'track-a'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(harness.state.trackVisibilityOverrides, 'track-a'), false);
  assert.equal(harness.state.routeUiOpenKey, null);
  assert.equal(harness.state.multiPhotoImport.trackMatch.trackId, '');
  assert.equal(harness.state.trackDeletePending, '');
  assert.ok(harness.renderCount() > 0);
});

test('server failure leaves the target visible and maps fixed safe messages', async () => {
  const harness = createHarness({
    result: { ok: false, errorCode: 'TRACK_SEGMENTS_DELETE_FAILED', error: 'raw server details' }
  });

  assert.equal(await harness.api.deleteTrackFromUi('track-a'), false);

  assert.equal(harness.state.tracks.length, 1);
  assert.equal(harness.layer.removed, false);
  assert.equal(JSON.stringify(harness.notifications).includes('raw server details'), false);
  assert.match(harness.notifications.at(-1).message, /ルート区間データ/);
});

test('server-deleted client reconciliation failure keeps state and asks for reload', async () => {
  const harness = createHarness({ layerRemoveFails: true });

  assert.equal(await harness.api.deleteTrackFromUi('track-a'), false);

  assert.equal(harness.state.tracks.length, 1);
  assert.match(harness.notifications.at(-1).message, /再読み込み/);
  assert.equal(JSON.stringify(harness.notifications).includes('raw leaflet failure'), false);
});

test('active import matching and route mutations reject deletion before confirmation', async () => {
  for (const options of [{ matching: true }, { importBusy: true }, { routeBusy: true }]) {
    const harness = createHarness(options);
    assert.equal(await harness.api.deleteTrackFromUi('track-a'), false);
    assert.equal(harness.confirmations.length, 0);
    assert.equal(harness.gasCalls.length, 0);
    assert.equal(harness.state.tracks.length, 1);
    assert.match(harness.notifications.at(-1).message, /完了|使用中/);
  }
});

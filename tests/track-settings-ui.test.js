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

test('saved track editing reuses the import preview fields with isolated state', () => {
  assert.match(indexHtml, /trackEdit:\s*\{[\s\S]*?trackId:\s*''[\s\S]*?submittedPayload:\s*null[\s\S]*?saving:\s*false/);
  const open = functionSource('openTrackDisplaySettingsEditor');
  for (const id of [
    'track-import-preview-overlay', 'track-import-preview-name', 'track-import-preview-description',
    'track-import-preview-color', 'track-import-preview-line-style', 'track-import-preview-visible'
  ]) assert.match(open, new RegExp(id));
  assert.match(open, /renderColorPaletteButtons/);
  assert.match(open, /ルート表示設定を編集/);
  assert.match(open, /openOverlay\('track-import-preview-overlay'\)/);
  assert.doesNotMatch(open, /track-import-preview-line-width/);
  assert.doesNotMatch(open, /normalizeTrackForState/);
  assert.match(open, /snapshot\s*=\s*\{[\s\S]*trackId:\s*id[\s\S]*lineStyle:/);
  assert.doesNotMatch(open, /snapshot\s*=\s*\{[\s\S]*segments:/);
});

test('saved track settings submit once with an edit token and retry the frozen payload', () => {
  const save = functionSource('saveTrackDisplaySettings');
  const submit = functionSource('submitTrackDisplaySettings');
  const retry = functionSource('retryTrackDisplaySettings');
  assert.match(save, /validateTrackImportPreviewForm\(\)/);
  assert.match(save, /submittedPayload/);
  assert.match(submit, /updateTrackDisplaySettings/);
  assert.match(submit, /withEditToken\(/);
  assert.match(submit, /lineWidth\s*!==\s*4/);
  assert.match(submit, /applyTrackDisplaySettingsToClient/);
  assert.match(submit, /TRACK_CLIENT_SYNC_FAILED/);
  assert.match(submit, /retryable\s*=\s*serverAcknowledged\s*\?\s*false/);
  assert.match(retry, /submittedPayload/);
  assert.match(retry, /submitTrackDisplaySettings/);
});

test('track preview actions and Escape dispatch by import or edit mode', () => {
  for (const name of [
    'handleTrackPreviewSave', 'handleTrackPreviewRetry',
    'handleTrackPreviewDiscard', 'handleTrackPreviewClose'
  ]) {
    const source = functionSource(name);
    assert.match(source, /isTrackDisplaySettingsEditorOpen/);
    assert.match(source, /trackImportController/);
  }
  const dismiss = functionSource('dismissOverlayById');
  assert.match(dismiss, /track-import-preview-overlay[\s\S]*isTrackDisplaySettingsEditorOpen/);
  assert.match(dismiss, /state\.trackEdit\.saving/);
});

test('client metadata merge preserves geometry and immutable track fields', () => {
  const original = {
    id: 'track-a', trackId: 'track-a', revisionId: 'rev-a', name: 'Old', description: '',
    color: '#2196f3', sourceType: 'gpx', sourceName: 'route.gpx', orderIndex: 7,
    visible: true, lineStyle: 'solid', lineWidth: 4, createdAt: 'created', updatedAt: 'old',
    segments: [{ index: 0, points: [{ lat: 35, lng: 139, elevation: 10, time: '2026-01-01T00:00:00Z' }] }]
  };
  const state = {
    tracks: [JSON.parse(JSON.stringify(original))],
    trackVisibilityOverrides: { 'track-a': false }
  };
  let layerRenders = 0;
  let panelRenders = 0;
  const context = {
    state,
    getTrackId(track) { return track && String(track.trackId || track.id || ''); },
    normalizeTrackForState(track) { return JSON.parse(JSON.stringify(track)); },
    renderTrackLayers() { layerRenders += 1; },
    renderTrackPanel() { panelRenders += 1; }
  };
  vm.runInNewContext(`${functionSource('applyTrackDisplaySettingsToClient')}\nthis.apply = applyTrackDisplaySettingsToClient;`, context);

  const applied = context.apply({
    trackId: 'track-a', name: 'New', description: 'changed', color: '#e53935',
    visible: false, lineStyle: 'dotted', lineWidth: 4, updatedAt: 'new-time'
  });

  assert.equal(applied, true);
  const updated = state.tracks[0];
  assert.equal(updated.name, 'New');
  assert.equal(updated.description, 'changed');
  assert.equal(updated.color, '#e53935');
  assert.equal(updated.visible, false);
  assert.equal(updated.lineStyle, 'dotted');
  assert.equal(updated.updatedAt, 'new-time');
  for (const key of ['revisionId', 'sourceType', 'sourceName', 'orderIndex', 'createdAt']) {
    assert.deepEqual(updated[key], original[key], key);
  }
  assert.deepEqual(updated.segments, original.segments);
  assert.equal(Object.prototype.hasOwnProperty.call(state.trackVisibilityOverrides, 'track-a'), false);
  assert.equal(layerRenders, 1);
  assert.equal(panelRenders, 1);
  assert.equal(context.apply({ trackId: 'missing', lineWidth: 4 }), false);
});

test('track edit participates in production busy and mutation policy', () => {
  assert.match(functionSource('isProductionImportBusy'), /isTrackDisplaySettingsEditBusy\(\)/);
  assert.match(functionSource('isTrackDisplaySettingsEditBusy'), /trackEdit/);
  assert.match(indexHtml, /withGAS\('updateTrackDisplaySettings',\s*withEditToken\(/);
});

test('track settings errors use fixed messages and never expose raw server details', () => {
  const context = {};
  vm.runInNewContext(`${functionSource('trackDisplaySettingsFailureMessage')}\nthis.message = trackDisplaySettingsFailureMessage;`, context);
  assert.match(context.message({ errorCode: 'TRACK_NOT_FOUND', error: 'raw secret' }), /見つかりません/);
  assert.match(context.message({ errorCode: 'TRACK_METADATA_UPDATE_FAILED', error: 'raw secret' }), /更新できません/);
  assert.match(context.message({ errorCode: 'TRACK_CLIENT_SYNC_FAILED', error: 'raw secret' }), /再読み込み/);
  assert.doesNotMatch(context.message({ errorCode: 'UNKNOWN', error: 'raw secret' }), /raw secret/);
});

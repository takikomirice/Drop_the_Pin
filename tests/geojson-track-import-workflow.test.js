const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const codeJs = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(__dirname, '..', 'shared.html'), 'utf8');
const readme = fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8');

function sourceFunction(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let cursor = open; cursor < source.length; cursor += 1) {
    if (source[cursor] === '{') depth += 1;
    if (source[cursor] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, cursor + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function track(overrides = {}) {
  return {
    id: 'track-a', trackId: 'track-a', revisionId: 'rev-a', name: 'A', description: '',
    color: '#e53935', sourceType: 'geojson', sourceName: 'a.geojson',
    segments: [{ index: 0, points: [{ lat: 35, lng: 139, elevation: null, time: '' }] }],
    orderIndex: 1, visible: true, lineStyle: 'solid', lineWidth: 4,
    ...overrides
  };
}

test('production state and controller keep common track ownership separate from Point import', () => {
  assert.match(indexHtml, /trackImport:\s*\{\s*owner:\s*'',\s*loading:\s*false,\s*requestToken:\s*0,/s);
  assert.match(indexHtml, /const trackImportController = TrackFileImportUI\.create\(\{/);
  assert.match(indexHtml, /adapters:\s*\[GeoJsonTrackImportAdapter, GpxTrackImportAdapter\]/);
  assert.match(indexHtml, /callGAS:\s*withGAS/);
  assert.match(indexHtml, /withEditToken:\s*withEditToken/);
  assert.match(indexHtml, /onSaved:\s*upsertTrack/);
  assert.match(indexHtml, /function isTrackImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'isProductionImportBusy'), /isTrackImportBusy\(\)/);

  const pointPreview = sourceFunction(indexHtml, 'openGeoJsonImportPreview');
  assert.match(pointPreview, /ImportPinItemProcessor\.create/);
  assert.match(pointPreview, /isTrackImportBusy\(\)/);
  assert.doesNotMatch(pointPreview, /saveTrackBundle|GeoJsonTrackInterchangeCore/);
  const processorStart = indexHtml.indexOf('const ImportPinItemProcessor =');
  const processorEnd = indexHtml.indexOf('const MultiPhotoImportWorkflow =', processorStart);
  assert.match(indexHtml.slice(processorStart, processorEnd), /saveImportPinItem/);
  assert.doesNotMatch(indexHtml.slice(processorStart, processorEnd), /saveTrackBundle/);
});

test('every production import entry point and edit exit guard includes common track ownership', () => {
  assert.match(sourceFunction(indexHtml, 'openCsvImportPreview'), /isTrackImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'openGeoJsonImportPreview'), /isTrackImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'canStartMultiPhotoImport'), /isTrackImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'openUploadModal'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'openSettingsModal'), /isProductionImportBusy\(\)/);
  assert.match(indexHtml, /canStartImport:\s*function\(\)\s*\{\s*return !isCsvImportBusy\(\)\s*&&\s*!isGeoJsonImportBusy\(\)\s*&&\s*!isMultiPhotoImportBusyState\(\);\s*\}/s);
  const setEditMode = sourceFunction(indexHtml, 'setEditMode');
  assert.match(setEditMode, /isProductionImportBusy\(\)/);
  assert.match(setEditMode, /閲覧モード/);
  assert.match(sourceFunction(indexHtml, 'hasPendingMutationWork'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'initializeApp'), /beforeunload[\s\S]*hasPendingMutationWork\(\)/);
});

test('settings close preserves common track import only for preview handoff', () => {
  const closeOverlay = sourceFunction(indexHtml, 'closeOverlay');
  assert.match(closeOverlay, /preserveTrackImport/);
  assert.match(closeOverlay, /trackImportController\.invalidate\(\)/);
  const importFileStart = indexHtml.indexOf('function importFile(kind, file)', indexHtml.indexOf('const TrackFileImportUI'));
  const importFileEnd = indexHtml.indexOf('function handleFileSelected', importFileStart);
  const importFileSource = indexHtml.slice(importFileStart, importFileEnd);
  assert.match(importFileSource, /preserveTrackImport:\s*true/);
  assert.equal((indexHtml.match(/preserveTrackImport:\s*true/g) || []).length, 1);
});

test('track preview blocks background shortcuts and has no implicit backdrop close path', () => {
  assert.match(sourceFunction(indexHtml, 'isShortcutOverlayOpen'), /track-import-preview-overlay/);
  const backdropIdsStart = indexHtml.indexOf('const BACKDROP_DISMISS_OVERLAY_IDS =');
  const backdropIdsEnd = indexHtml.indexOf('];', backdropIdsStart);
  assert.doesNotMatch(indexHtml.slice(backdropIdsStart, backdropIdsEnd), /track-import-preview-overlay/);
});

test('upsertTrack normalizes by server trackId, sorts, rerenders, and never duplicates replay', () => {
  const state = {
    tracks: [track({ trackId: 'track-b', id: 'track-b', orderIndex: 2, name: 'B' })],
    trackVisibilityOverrides: Object.create(null)
  };
  state.trackVisibilityOverrides['track-a'] = false;
  let layersRendered = 0;
  let panelRendered = 0;
  const context = {
    state,
    normalizeTrackForState(value) {
      assert.ok(value.trackId);
      return plain({ ...value, id: value.trackId, trackId: value.trackId });
    },
    renderTrackLayers() { layersRendered += 1; },
    renderTrackPanel() { panelRendered += 1; }
  };
  vm.runInNewContext(`${sourceFunction(indexHtml, 'upsertTrack')}\nglobalThis.upsertTrack = upsertTrack;`, context);

  const inserted = context.upsertTrack(track({ orderIndex: 1, name: 'first' }));
  assert.equal(inserted.trackId, 'track-a');
  assert.deepEqual(state.tracks.map((item) => item.trackId), ['track-a', 'track-b']);
  assert.equal(state.trackVisibilityOverrides['track-a'], false);
  context.upsertTrack(track({ orderIndex: 0, name: 'deduplicated replay' }));
  assert.equal(state.tracks.length, 2);
  assert.equal(state.tracks[0].name, 'deduplicated replay');
  assert.equal(layersRendered, 2);
  assert.equal(panelRendered, 2);
  assert.equal(context.upsertTrack({ id: 'feature-id' }), null);
});

test('track import remains isolated from routes Drive Point API and shared view', () => {
  const moduleStart = indexHtml.indexOf('const TrackFileImportUI =');
  const moduleEnd = indexHtml.indexOf('\n    const state = {', moduleStart);
  const moduleSource = indexHtml.slice(moduleStart, moduleEnd);
  assert.match(moduleSource, /saveTrackBundle/);
  assert.doesNotMatch(moduleSource, /saveImportPinItem|route_pins|route_cache|saveRoute|Drive|pointsJson|map_info/);
  assert.doesNotMatch(sharedHtml, /GeoJsonTrackImportUI|GeoJsonTrackInterchangeCore|saveTrackBundle/);
  assert.equal((codeJs.match(/function saveTrackBundle\(/g) || []).length, 1);
  assert.doesNotMatch(codeJs, /function (?:import|save)GeoJsonTrack/);
});

test('README summarizes GeoJSON track formats, limits, segment behavior, and verification', () => {
  assert.match(
    readme,
    /GeoJSONトラック[^\n]*LineString／MultiLineString[^\n]*5MB[^\n]*100,000 source point/
  );
  assert.match(
    readme,
    /GeoJSON[^\n]*完全重複[^\n]*最大5秒[^\n]*形状[^\n]*4時間[^\n]*20,000 point/
  );
  assert.match(readme, /segment順を維持し[^\n]*segment間を自動接続しません/);
  assert.match(readme, /Nodeテスト[^\n]*Chromiumテスト[^\n]*Apps Script[^\n]*実Drive[^\n]*未実施/);
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

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

test('production uses one owner-aware track import state for GeoJSON and GPX', () => {
  assert.match(indexHtml, /trackImport:\s*\{\s*owner:\s*'',\s*loading:\s*false,\s*requestToken:\s*0,/s);
  assert.match(indexHtml, /sourceKind:\s*'',\s*sourceName:\s*''/s);
  assert.doesNotMatch(indexHtml, /geoJsonTrackImport:\s*\{/);
  assert.equal((indexHtml.match(/state\.trackImport/g) || []).length > 0, true);
  assert.match(indexHtml, /const trackImportController = TrackFileImportUI\.create\(\{/);
  assert.match(indexHtml, /const geoJsonTrackImportController = trackImportController\.forKind\('geojson'\)/);
  assert.match(indexHtml, /const gpxTrackImportController = trackImportController\.forKind\('gpx'\)/);
});

test('settings handoff has one common preserve flag and normal close invalidates common ownership', () => {
  assert.equal((indexHtml.match(/preserveTrackImport:\s*true/g) || []).length, 1);
  assert.doesNotMatch(indexHtml, /preserveGeoJsonTrackImport|preserveGpxTrackImport/);
  const closeOverlay = sourceFunction(indexHtml, 'closeOverlay');
  assert.match(closeOverlay, /preserveTrackImport/);
  assert.match(closeOverlay, /trackImportController\.invalidate\(\)/);
});

test('production busy and beforeunload read the common track import state once', () => {
  const trackBusy = sourceFunction(indexHtml, 'isTrackImportBusy');
  assert.match(trackBusy, /state\.trackImport/);
  assert.match(trackBusy, /loading|draft|submittedPayload|saving|saved/);
  const productionBusy = sourceFunction(indexHtml, 'isProductionImportBusy');
  assert.match(productionBusy, /isTrackImportBusy\(\)/);
  assert.doesNotMatch(productionBusy, /isGeoJsonTrackImportBusy|isGpxTrackImportBusy/);
  assert.match(sourceFunction(indexHtml, 'hasPendingMutationWork'), /isProductionImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'initializeApp'), /beforeunload[\s\S]*hasPendingMutationWork\(\)/);
});

test('shared preview exposes summary counts without raw coordinate or source markup', () => {
  for (const id of [
    'track-import-preview-format', 'track-import-preview-stats', 'track-import-preview-warnings'
  ]) assert.match(indexHtml, new RegExp(`id="${id}"`), id);
  assert.match(indexHtml, /id="track-import-preview-points"/);
  assert.doesNotMatch(indexHtml, /track-import-preview-(?:coordinates|xml|geojson)/);
});

test('track import busy rendering synchronizes both interchange buttons multi-photo and single add', () => {
  const renderStart = indexHtml.indexOf('function renderTrackImportBusy()');
  const renderEnd = indexHtml.indexOf('\n    function isCsvImportBusy', renderStart);
  const renderSource = indexHtml.slice(renderStart, renderEnd);
  assert.match(renderSource, /geojson-track-import-button/);
  assert.match(renderSource, /gpx-track-import-button/);
  assert.match(renderSource, /multi-photo-button|refreshMultiPhotoButtonState/);
  assert.match(renderSource, /refreshPinAddButtonState\(\)/);

  const csvStart = indexHtml.indexOf('const csvInterchangeController =');
  const geoJsonStart = indexHtml.indexOf('const geoJsonInterchangeController =', csvStart);
  const trackStart = indexHtml.indexOf('const trackImportController =', geoJsonStart);
  assert.match(indexHtml.slice(csvStart, geoJsonStart), /onBusy:[\s\S]*renderTrackImportBusy\(\)/);
  assert.match(indexHtml.slice(geoJsonStart, trackStart), /onBusy:[\s\S]*renderTrackImportBusy\(\)/);
  assert.match(sourceFunction(indexHtml, 'setEditMode'), /renderTrackImportBusy\(\)/);
});

test('route preview delegates initial focus and production busy guidance stays format neutral', () => {
  const trackUiStart = indexHtml.indexOf('const TrackFileImportUI =');
  const trackUiEnd = indexHtml.indexOf('\n    const GeoJsonTrackImportUI =', trackUiStart);
  const trackUiSource = indexHtml.slice(trackUiStart, trackUiEnd);
  const renderPreview = sourceFunction(trackUiSource, 'renderPreview');
  assert.match(renderPreview, /openOverlay\('track-import-preview-overlay'\)/);
  assert.doesNotMatch(renderPreview, /\.focus\(/);

  for (const name of ['openSettingsModal', 'openUploadModal', 'saveNewPin']) {
    assert.match(sourceFunction(indexHtml, name), /進行中の取込/, name);
  }
});

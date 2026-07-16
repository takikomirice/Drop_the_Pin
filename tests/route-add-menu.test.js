const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

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

test('route add button exposes a two-choice menu with synchronized accessibility state', () => {
  assert.match(indexHtml, /id="route-add-btn"[^>]*aria-haspopup="menu"[^>]*aria-controls="route-add-menu"[^>]*aria-expanded="false"/);
  assert.match(indexHtml, /id="route-add-menu"[^>]*role="menu"/);
  assert.match(indexHtml, /id="route-add-pin-item"[^>]*role="menuitem"[^>]*>[^<]*ピンからルートを作成/);
  assert.match(indexHtml, /id="route-add-file-item"[^>]*role="menuitem"[^>]*>[^<]*GPX \/ GeoJSONからルートを作成/);
  const setter = functionSource('setRouteAddMenuOpen');
  assert.match(setter, /aria-expanded/);
  assert.match(setter, /menu\.hidden/);
  assert.match(setter, /focus/);
  assert.match(functionSource('dispatchEscape'), /closeRouteAddMenu/);
  assert.match(indexHtml, /document\.addEventListener\('click',[\s\S]*closeRouteAddMenu/);
});

test('pin route choice closes the menu and reuses the existing route-name overlay path', () => {
  const source = functionSource('startPinRouteFromRouteAddMenu');
  assert.match(source, /closeRouteAddMenu/);
  assert.match(source, /createRouteGroup\(\)/);
  assert.match(functionSource('openRouteNameDialog'), /openOverlay\('route-name-overlay'\)/);
  assert.match(indexHtml, /'ピンからルートを作成'/);
});

test('file route choice opens one dedicated source dialog wired into the common Escape dispatcher', () => {
  assert.match(indexHtml, /id="route-track-import-overlay"[^>]*role="dialog"/);
  assert.match(indexHtml, /id="route-track-import-gpx"[^>]*>[^<]*GPXを取込/);
  assert.match(indexHtml, /id="route-track-import-geojson"[^>]*>[^<]*GeoJSONを取込/);
  assert.match(indexHtml, /id="route-track-import-cancel"[^>]*>[^<]*キャンセル/);
  assert.match(functionSource('openRouteTrackImportDialog'), /openOverlay\('route-track-import-overlay'\)/);
  assert.match(functionSource('dismissOverlayById'), /route-track-import-overlay[\s\S]*closeRouteTrackImportDialog/);
  assert.match(indexHtml, /MAIN_DISMISSIBLE_OVERLAY_IDS[\s\S]*'route-track-import-overlay'/);
});

test('GPX and GeoJSON choices reuse existing track inputs and controllers only', () => {
  const start = functionSource('startRouteTrackImport');
  assert.match(start, /gpx-track-file-input/);
  assert.match(start, /geojson-track-file-input/);
  assert.doesNotMatch(start, /geojson-file-input/);
  assert.match(start, /closeRouteTrackImportDialog/);
  assert.ok(
    start.indexOf('closeRouteTrackImportDialog') < start.indexOf('input.click()'),
    'the modal must release inert background inputs before opening the native picker'
  );
  assert.match(start, /watchRouteTrackImportFilePickerCancel/);
  const change = functionSource('handleTrackFileInputChange');
  assert.match(change, /gpxTrackImportController/);
  assert.match(change, /geoJsonTrackImportController/);
  assert.match(change, /handleFileSelected/);
  assert.match(change, /closeRouteTrackImportDialog/);
  assert.match(change, /finally/);
  assert.match(change, /clearRouteTrackImportFilePickerWatcher/);

  assert.equal((indexHtml.match(/function createTrackFileImportAdapter\(/g) || []).length, 1);
  assert.equal((indexHtml.match(/callGAS\('saveTrackBundle'/g) || []).length, 1);
  assert.match(indexHtml, /geojson-track-file-input[\s\S]*geoJsonTrackImportController/);
  assert.match(indexHtml, /gpx-track-file-input[\s\S]*gpxTrackImportController/);
});

test('route import handoff keeps asynchronous parsing alive after its source dialog closes', () => {
  const openCheck = functionSource('isTrackImportEntrySurfaceOpen');
  assert.match(openCheck, /data-overlay/);
  assert.match(openCheck, /route-track-import-overlay/);
  assert.match(openCheck, /routeTrackImportHandoffKind/);
  const controller = indexHtml.slice(
    indexHtml.indexOf('const trackImportController ='),
    indexHtml.indexOf('const geoJsonTrackImportController =')
  );
  assert.match(controller, /isSettingsOpen:\s*isTrackImportEntrySurfaceOpen/);
  assert.match(controller, /closeSettings:\s*closeTrackImportEntrySurface/);
});

test('route import validation failures return to the source dialog with a visible safe error', () => {
  assert.match(indexHtml, /id="route-track-import-error"[^>]*role="alert"[^>]*hidden/);
  assert.match(functionSource('openRouteTrackImportDialog'), /setRouteTrackImportError/);
  const change = functionSource('handleTrackFileInputChange');
  assert.match(change, /value\s*==\s*null/);
  assert.match(change, /gpx-track-operation-error/);
  assert.match(change, /geojson-track-operation-error/);
  assert.match(change, /openRouteTrackImportDialog/);
});

test('native picker handoff preserves the route add button as the final focus return target', () => {
  assert.match(indexHtml, /let routeTrackImportReturnFocusTarget\s*=\s*null/);
  assert.match(functionSource('startRouteTrackImport'), /getOverlayOpenRecord\('route-track-import-overlay'\)/);
  assert.match(functionSource('openRouteTrackImportDialog'), /adoptRouteTrackImportReturnFocus\('route-track-import-overlay'\)/);
  assert.match(functionSource('handleTrackFileInputChange'), /adoptRouteTrackImportReturnFocus\('track-import-preview-overlay'\)/);
});

test('route creation entry stays hidden outside editable desktop mode', () => {
  assert.match(indexHtml, /body:not\(\.has-edit-token\)[\s\S]*#route-add-btn/);
  assert.match(indexHtml, /body:not\(\.edit-mode\):not\(\.share-mode\)[\s\S]*#route-add-btn/);
  assert.match(indexHtml, /body\.preview-mode[\s\S]*#route-add-btn/);
  assert.match(indexHtml, /body\.narrow-view[\s\S]*#route-add-btn/);
  assert.match(functionSource('renderAccessMode'), /closeRouteAddMenu/);
  assert.match(functionSource('setNarrowView'), /closeRouteAddMenu/);
  assert.match(functionSource('renderPanelToggle'), /closeRouteAddMenu/);
});

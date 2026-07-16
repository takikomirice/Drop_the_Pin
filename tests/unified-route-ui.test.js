const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

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

test('pin, GPX and GeoJSON routes are projected into one collision-safe list', () => {
  const context = {
    state: {
      routeGroups: [{ id: 'same', routeId: 'same', name: 'Pin route' }],
      tracks: [
        { id: 'same', trackId: 'same', sourceType: 'gpx', name: 'GPX route' },
        { id: 'geo', trackId: 'geo', sourceType: 'geojson', name: 'Geo route' },
        { id: 'legacy', trackId: 'legacy', sourceType: 'manual', name: 'Legacy track' }
      ]
    },
    getRouteId: (route) => route.routeId || route.id,
    getTrackId: (track) => track.trackId || track.id
  };
  const getEntries = vm.runInNewContext(`(${functionSource('getUnifiedRouteEntries')})`, context);
  const entries = getEntries();

  assert.deepEqual(
    JSON.parse(JSON.stringify(entries.map(({ key, kind }) => ({ key, kind })))),
    [
      { key: 'pin:same', kind: 'pin' },
      { key: 'gpx:same', kind: 'gpx' },
      { key: 'geojson:geo', kind: 'geojson' }
    ]
  );
});

test('right dock keeps legacy IDs but exposes one visible route region', () => {
  ['dock-route-region', 'side-routes', 'route-list', 'side-tracks', 'track-list',
    'route-panel-header', 'route-dock-toggle', 'route-count']
    .forEach((id) => assert.equal(countId(id), 1, `${id} must exist exactly once`));
  assert.match(indexHtml, /id="side-tracks"[^>]*hidden[^>]*aria-hidden="true"/);
  assert.match(indexHtml, /id="route-dock-toggle"[^>]*aria-controls="route-list"[^>]*aria-expanded="false"/);

  const render = functionSource('renderRoutePanel');
  assert.match(render, /getUnifiedRouteEntriesForPanel\(\)/);
  assert.match(render, /buildRouteItem/);
  assert.match(render, /buildUnifiedTrackRouteItem/);
  assert.match(render, /getElementById\('route-list'\)/);
});

test('route dock actions are borderless and use stable inline SVG icons', () => {
  assert.match(
    indexHtml,
    /#route-add-btn, #route-dock-toggle\s*\{[^}]*border:\s*0[^}]*background:\s*transparent/
  );
  assert.match(
    indexHtml,
    /id="route-add-btn"[^>]*aria-label="ルートを追加"[\s\S]*?<svg[^>]*class="route-panel-action-icon route-add-icon"[\s\S]*?<path d="M12 5v14M5 12h14"/
  );
  assert.match(
    indexHtml,
    /id="route-dock-toggle"[^>]*aria-label="ルート領域を展開"[\s\S]*?<svg[^>]*route-dock-icon-collapsed[\s\S]*?<svg[^>]*route-dock-icon-expanded/
  );
  assert.match(indexHtml, /#route-dock-toggle\[aria-expanded="true"\] \.route-dock-icon-collapsed\s*\{\s*display:\s*none/);
  assert.match(indexHtml, /#route-dock-toggle\[aria-expanded="true"\] \.route-dock-icon-expanded\s*\{\s*display:\s*block/);

  const render = functionSource('renderRouteDockState');
  assert.doesNotMatch(render, /toggle\.textContent/);
  assert.match(render, /toggle\.setAttribute\('aria-label', expanded \? 'ルート領域を最小化' : 'ルート領域を展開'\)/);
  assert.match(render, /toggle\.title = expanded \? 'ルート領域を最小化' : 'ルート領域を展開'/);
});

test('unified cards show the three approved type badges', () => {
  const badge = functionSource('createRouteTypeBadge');
  assert.match(badge, /pin:\s*'ピンで作成'/);
  assert.match(badge, /gpx:\s*'GPX'/);
  assert.match(badge, /geojson:\s*'GeoJSON'/);
  assert.match(badge, /textContent/);
  assert.match(indexHtml, /\.route-type-badge/);
  assert.match(indexHtml, /\.route-type-badge\.type-pin/);
  assert.match(indexHtml, /\.route-type-badge\.type-gpx/);
  assert.match(indexHtml, /\.route-type-badge\.type-geojson/);
});

test('pin-created and imported route cards expose their approved edit controls', () => {
  const pinCard = functionSource('buildRouteItem');
  assert.match(pinCard, /createRouteTypeBadge\('pin'\)/);
  assert.match(pinCard, /beginRouteMapAdd/);
  assert.match(pinCard, /buildRoutePinList/);
  assert.match(pinCard, /attachRoutePinDropTarget/);

  const importedCard = functionSource('buildUnifiedTrackRouteItem');
  assert.match(importedCard, /createRouteTypeBadge\(kind\)/);
  assert.match(importedCard, /toggleTrackVisibility/);
  assert.match(importedCard, /fitTrackBounds/);
  assert.match(importedCard, /canEditRouteControls\(\)[\s\S]*openTrackDisplaySettingsEditor[\s\S]*deleteTrackFromUi/);
  assert.doesNotMatch(importedCard, /unified-route-card imported-route-item track-item/);
  assert.match(importedCard, /createActionIconElement\(document, 'settings'\)/);
  assert.match(importedCard, /createActionIconElement\(document, 'delete'\)/);
  ['beginRouteMapAdd', 'buildRoutePinList', 'attachRoutePinDropTarget', 'openRouteEdit']
    .forEach((needle) => assert.equal(importedCard.includes(needle), false, `Imported route must not use ${needle}`));
});

test('pin route actions are five icon-only controls above the owned-pin heading', () => {
  const pinCard = functionSource('buildRouteItem');
  const addIndex = pinCard.indexOf("title = '追加'");
  const chronologicalIndex = pinCard.indexOf("title = '時系列'");
  const undoIndex = pinCard.indexOf("title = 'Undo'");
  const settingsIndex = pinCard.indexOf("title = '編集'");
  const deleteIndex = pinCard.indexOf("title = '削除'");
  const headingIndex = pinCard.indexOf("'所属ピン'");
  assert.ok(addIndex !== -1 && addIndex < chronologicalIndex);
  assert.ok(chronologicalIndex < undoIndex);
  assert.ok(undoIndex < settingsIndex);
  assert.ok(settingsIndex < deleteIndex);
  assert.ok(deleteIndex < headingIndex);

  const contracts = [
    ['ピンを追加', 'beginRouteMapAdd', '追加'],
    ['所属ピンを時系列に戻す', 'resetRoutePinsToChronological', '時系列'],
    ['直前の所属ピン変更を元に戻す', 'undoRoutePinChange', 'Undo'],
    ['ルートを編集', 'openRouteEditOverlay', '編集'],
    ['ルートを削除', 'deleteRouteGroupFromUi', '削除']
  ];
  contracts.forEach(([label, handler, title]) => {
    assert.match(pinCard, new RegExp(`aria-label', '${label.replace(/[.*+?^${}()|[\\]\\]/g, '\\$&')}`));
    assert.match(pinCard, new RegExp(`title = '${title}`));
    assert.match(pinCard, new RegExp(handler));
  });
  assert.match(pinCard, /createActionIconElement/);
  assert.match(pinCard, /const editable = canEditRouteControls\(\)/);
  assert.doesNotMatch(pinCard, /className = 'action-label'/);
  assert.match(indexHtml, /\.route-item-actions\s*\{[^}]*grid-template-columns:\s*repeat\(5, 44px\)[^}]*flex-wrap:\s*nowrap/);
  assert.match(indexHtml, /\.route-item-actions \.ghost-btn\s*\{[^}]*width:\s*44px[^}]*min-width:\s*44px[^}]*height:\s*44px[^}]*min-height:\s*44px/);
  assert.match(indexHtml, /body\.narrow-view \.route-detail\s*\{[^}]*padding-left:\s*10px/);
});

test('pin route deletion is no longer embedded in the edit overlay', () => {
  assert.doesNotMatch(indexHtml, /id="route-edit-delete"/);
  assert.doesNotMatch(functionSource('renderRouteEditOverlay'), /deleteRouteGroupFromUi/);
});

test('route dock defaults to minimized and expands for selection or drag', () => {
  assert.match(indexHtml, /routeDockExpanded:\s*false/);
  const render = functionSource('renderRouteDockState');
  assert.match(render, /route-dock-expanded/);
  assert.match(render, /aria-expanded/);
  assert.match(render, /routeDockExpanded/);
  assert.match(functionSource('setRouteDockExpanded'), /renderRouteDockState/);
  assert.match(functionSource('attachPinOrderSortable'), /setRouteDockExpanded\(true\)/);
  assert.match(indexHtml, /#side-panel\.route-dock-expanded/);
});

test('pin add stays in the pin header and the map FAB positioning is removed', () => {
  assert.match(indexHtml, /class="pin-panel-actions"[\s\S]*?id="pin-add-btn"/);
  assert.doesNotMatch(indexHtml, /\bid="fab"|#fab\s*\{/);
});

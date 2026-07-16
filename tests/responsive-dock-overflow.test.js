const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];

function blockAfter(source, marker) {
  const markerIndex = source.indexOf(marker);
  assert.notEqual(markerIndex, -1, `missing ${marker}`);
  const openIndex = source.indexOf('{', markerIndex + marker.length);
  assert.notEqual(openIndex, -1, `missing block for ${marker}`);
  let depth = 0;
  for (let index = openIndex; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(openIndex + 1, index);
  }
  assert.fail(`unterminated block for ${marker}`);
}

test('expanded desktop dock keeps headers outside its independent scroll areas', () => {
  const desktop = blockAfter(css, '@media (min-width: 641px)');
  const fixedHeaders = blockAfter(
    desktop,
    '#side-panel.route-dock-expanded :is(.dock-pin-header, .dock-pin-tabs, #route-panel-header)'
  );
  const scrollAreas = blockAfter(
    desktop,
    '#side-panel.route-dock-expanded :is(#side-placed, #side-unplaced, #route-list)'
  );

  assert.match(fixedHeaders, /flex:\s*0 0 auto/);
  assert.match(scrollAreas, /min-height:\s*0/);
  assert.match(scrollAreas, /overflow-x:\s*hidden/);
  assert.match(scrollAreas, /overflow-y:\s*auto/);
  assert.match(scrollAreas, /overscroll-behavior:\s*contain/);
  assert.match(scrollAreas, /scrollbar-gutter:\s*stable/);
});

test('the complete pin and route flex chain permits child scrolling', () => {
  const dockRegions = blockAfter(css, '#dock-pin-region, #dock-route-region');
  const pinRegionSource = css.slice(css.indexOf('#dock-pin-region {'));
  const groupedRouteIndex = css.indexOf('#dock-route-region {');
  const routeRegionSource = css.slice(css.indexOf('#dock-route-region {', groupedRouteIndex + 1));
  const pinRegion = blockAfter(pinRegionSource, '#dock-pin-region');
  const routeRegion = blockAfter(routeRegionSource, '#dock-route-region');
  const sideRoutes = blockAfter(css, '#side-routes');
  const routeList = blockAfter(css, '#route-list');
  const routeCard = blockAfter(css, '.unified-route-card');
  const routePinList = blockAfter(css, '.route-pin-list');
  const expanded = blockAfter(css, '#side-panel.route-dock-expanded');

  assert.match(dockRegions, /min-height:\s*0/);
  assert.match(pinRegion, /overflow:\s*hidden/);
  assert.match(routeRegion, /overflow:\s*hidden/);
  assert.match(sideRoutes, /min-height:\s*0/);
  assert.match(sideRoutes, /flex:\s*1 1 auto/);
  assert.match(sideRoutes, /overflow:\s*hidden/);
  assert.match(routeList, /overflow-x:\s*hidden/);
  assert.match(routeList, /overflow-y:\s*auto/);
  assert.match(routeList, /overscroll-behavior:\s*contain/);
  assert.match(routeCard, /flex:\s*0 0 auto/);
  assert.match(routePinList, /max-height:/);
  assert.match(routePinList, /overflow-x:\s*hidden/);
  assert.match(routePinList, /overflow-y:\s*auto/);
  assert.doesNotMatch(routePinList, /overscroll-behavior/, 'owned pins must chain at their edge to the outer route list');
  assert.match(expanded, /minmax\(\s*var\(--dock-route-region-effective-min-height,\s*var\(--dock-route-region-min-height\)\)\s*,\s*1fr\s*\)/);
});

test('route and owned-pin scroll regions are keyboard reachable without changing Sortable ownership', () => {
  assert.match(indexHtml, /id="route-list"[^>]*tabindex="0"[^>]*aria-label="ルート一覧"/);
  const pinList = indexHtml.slice(
    indexHtml.indexOf('function buildRoutePinList'),
    indexHtml.indexOf('function buildRouteItem')
  );
  assert.match(pinList, /list\.tabIndex\s*=\s*0/);
  assert.match(pinList, /setAttribute\('aria-label', '所属ピン一覧'\)/);
  assert.match(pinList, /attachRoutePinSortable\(list, routeId, pinIds\)/);
  assert.match(indexHtml, /#route-list[\s\S]*addEventListener\('pointerdown',[\s\S]*stopPropagation/);
});

test('overflow isolation does not alter mobile sheet geometry or route state wiring', () => {
  const mobileSource = css.slice(css.indexOf('/* Phase 9.5-F UI-7'));
  const mobile = blockAfter(mobileSource, '@media (max-width: 640px)');
  const mobilePanel = blockAfter(mobile, 'body.narrow-view #side-panel');

  assert.match(mobilePanel, /height:\s*220px/);
  assert.match(mobilePanel, /min-height:\s*220px/);
  assert.match(
    mobile,
    /body\.narrow-view\.mobile-sheet-expanded #side-panel\s*\{\s*height:\s*min\(70dvh, 586px\)/
  );
  assert.match(
    indexHtml,
    /function renderRouteDockState\(\)[\s\S]*?classList\.toggle\('route-dock-expanded', expanded\)/
  );
  assert.match(indexHtml, /const narrowViewMedia = window\.matchMedia\('\(max-width: 640px\)'\);/);
});

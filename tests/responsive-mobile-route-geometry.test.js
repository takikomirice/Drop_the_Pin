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

function pixelToken(source, name) {
  const match = source.match(new RegExp(`${name}:\\s*(\\d+)px`));
  assert.ok(match, `missing ${name} pixel token`);
  return Number(match[1]);
}

test('route tab uses its own bounded mobile sheet geometry', () => {
  const rootBlock = blockAfter(css, ':root');
  const mobileSource = css.slice(css.indexOf('/* Phase 9.5-F UI-7'));
  const mobile = blockAfter(mobileSource, '@media (max-width: 640px)');
  const routeSheet = blockAfter(
    mobile,
    'body.narrow-view.mobile-sheet-expanded.mobile-sheet-routes #side-panel'
  );

  assert.match(rootBlock, /--mobile-route-map-min-height:\s*154px/);
  assert.match(rootBlock, /--mobile-route-sheet-max-height:\s*626px/);
  assert.match(routeSheet, /height:\s*min\(/);
  assert.match(routeSheet, /100dvh/);
  assert.match(routeSheet, /var\(--app-header-height\)/);
  assert.match(routeSheet, /var\(--mobile-route-map-min-height\)/);
  assert.match(routeSheet, /var\(--mobile-route-sheet-max-height\)/);

  const pinSheet = blockAfter(
    mobile,
    'body.narrow-view.mobile-sheet-expanded #side-panel'
  );
  assert.match(pinSheet, /height:\s*min\(70dvh,\s*586px\)/);
});

test('route sheet preserves map space and matches the 390 by 844 Pencil frame', () => {
  const rootBlock = blockAfter(css, ':root');
  const headerHeight = pixelToken(rootBlock, '--app-header-height');
  const mapMinimum = pixelToken(rootBlock, '--mobile-route-map-min-height');
  const sheetMaximum = pixelToken(rootBlock, '--mobile-route-sheet-max-height');

  [800, 812, 844, 932].forEach((viewportHeight) => {
    const appHeight = viewportHeight - headerHeight;
    const sheetHeight = Math.min(appHeight - mapMinimum, sheetMaximum);
    const visibleMapHeight = appHeight - sheetHeight;

    assert.ok(sheetHeight > 0, `${viewportHeight}px route sheet must have usable height`);
    assert.ok(sheetHeight <= appHeight, `${viewportHeight}px route sheet must fit the app shell`);
    assert.ok(
      visibleMapHeight >= mapMinimum,
      `${viewportHeight}px route sheet must preserve the minimum map height`
    );
  });

  const pencilViewportHeight = 844;
  const pencilSheetHeight = Math.min(
    pencilViewportHeight - headerHeight - mapMinimum,
    sheetMaximum
  );
  assert.equal(pencilSheetHeight, 626);
  assert.equal(pencilViewportHeight - pencilSheetHeight, 218);
});

test('mobile state thresholds and sheet transition wiring remain unchanged', () => {
  assert.match(indexHtml, /const narrowViewMedia = window\.matchMedia\('\(max-width: 640px\)'\);/);
  assert.match(
    indexHtml,
    /let isPanelVisible = narrowViewMedia\.matches \|\| window\.matchMedia\('\(min-width: 900px\)'\)\.matches;/
  );
  assert.match(
    indexHtml,
    /function setMobileSheetTab\(tab\)[\s\S]*?classList\.add\('mobile-sheet-expanded'\)/
  );
  assert.match(
    indexHtml,
    /function setMobileSheetTab\(tab\)[\s\S]*?if \(showingRoutes\) setRouteDockExpanded\(true\)/
  );
});

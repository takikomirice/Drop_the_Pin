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

function parseDockClamp(source) {
  const match = source.match(
    /--app-dock-width:\s*clamp\(\s*(\d+)px,\s*([\d.]+)vw,\s*(\d+)px\s*\)/
  );
  assert.ok(match, 'desktop dock width must use a pixel/vw/pixel clamp');
  return {
    minimum: Number(match[1]),
    fluid: Number(match[2]),
    maximum: Number(match[3])
  };
}

const clamp = (low, preferred, high) => Math.min(high, Math.max(low, preferred));

test('desktop dock width is fluid while leaving more room for the map', () => {
  const rootBlock = blockAfter(css, ':root');
  const dock = parseDockClamp(rootBlock);

  assert.equal(dock.minimum, 340);
  assert.ok(dock.fluid >= 26 && dock.fluid <= 27);
  assert.equal(dock.maximum, 360);

  assert.match(blockAfter(css, '#side-panel'), /width:\s*min\(var\(--app-dock-width\),\s*100%\)/);
  assert.match(blockAfter(css, '#panel-toggle'), /right:\s*var\(--app-dock-width\)/);

  const desktop = blockAfter(css, '@media (min-width: 900px)');
  assert.match(desktop, /#map[\s\S]*?right:\s*var\(--app-dock-width\)/);
});

test('1024 through 1280 preserve map, search, toggle, and dock operating space', () => {
  const rootBlock = blockAfter(css, ':root');
  const dock = parseDockClamp(rootBlock);
  const searchLeft = Number(rootBlock.match(/--sp-6:\s*(\d+)px/)[1]);
  const searchMaximum = Number(
    blockAfter(css, '#map-search-bar').match(/width:\s*min\((\d+)px/)[1]
  );
  const toggleWidth = 36;
  const minimumGap = Number(rootBlock.match(/--sp-3:\s*(\d+)px/)[1]);

  [1024, 1100, 1200, 1280].forEach((viewportWidth) => {
    const dockWidth = clamp(
      dock.minimum,
      viewportWidth * dock.fluid / 100,
      dock.maximum
    );
    const mapWidth = viewportWidth - dockWidth;
    const searchWidth = Math.min(searchMaximum, viewportWidth - (searchLeft * 2));
    const searchRight = searchLeft + searchWidth;
    const toggleLeft = viewportWidth - dockWidth - toggleWidth;

    assert.ok(dockWidth >= 340 && dockWidth <= 360, `${viewportWidth}px dock must stay bounded`);
    assert.ok(mapWidth >= 684, `${viewportWidth}px map must retain its effective width`);
    assert.ok(
      searchRight <= toggleLeft - minimumGap,
      `${viewportWidth}px search and toggle must keep an operating gap`
    );
    assert.ok(toggleLeft + toggleWidth <= viewportWidth - dockWidth, 'toggle must meet the dock edge');
  });
});

test('compact and desktop dock widths stay continuous at 1023 and 1024 pixels', () => {
  const rootBlock = blockAfter(css, ':root');
  const dock = parseDockClamp(rootBlock);
  const compact = blockAfter(css, '@media (min-width: 641px) and (max-width: 1023px)');
  const compactRoot = blockAfter(compact, ':root');
  const compactMatch = compactRoot.match(
    /--app-effective-dock-width:\s*clamp\(\s*(\d+)px,\s*([\d.]+)vw,\s*var\(--app-dock-width\)\s*\)/
  );
  assert.ok(compactMatch, 'compact dock maximum must follow the desktop dock token');

  const desktopAt = (width) => clamp(dock.minimum, width * dock.fluid / 100, dock.maximum);
  const compactAt1023 = clamp(
    Number(compactMatch[1]),
    1023 * Number(compactMatch[2]) / 100,
    desktopAt(1023)
  );
  const desktopAt1024 = desktopAt(1024);

  assert.ok(Math.abs(compactAt1023 - desktopAt1024) <= 2);
  assert.ok(Math.abs(desktopAt(1366) - 360) <= 1, '1366px must use the slimmer dock maximum');
  assert.ok(Math.abs(desktopAt(1440) - 360) <= 1, 'wide desktop must retain the slimmer maximum');
});

test('responsive width work leaves mobile and JavaScript thresholds intact', () => {
  assert.match(indexHtml, /const narrowViewMedia = window\.matchMedia\('\(max-width: 640px\)'\);/);
  assert.match(
    indexHtml,
    /let isPanelVisible = narrowViewMedia\.matches \|\| window\.matchMedia\('\(min-width: 900px\)'\)\.matches;/
  );
  const mobileSource = css.slice(css.indexOf('/* Phase 9.5-F UI-7'));
  const mobile = blockAfter(mobileSource, '@media (max-width: 640px)');
  assert.match(blockAfter(mobile, 'body.narrow-view #side-panel'), /width:\s*100%/);
  assert.match(
    blockAfter(mobile, 'body.narrow-view #map-search-bar'),
    /width:\s*calc\(100vw - 32px\)/
  );
});

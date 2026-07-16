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

function functionBody(name) {
  return blockAfter(indexHtml, `function ${name}`);
}

test('641px through 1023px use one bounded effective dock width', () => {
  const compact = blockAfter(css, '@media (min-width: 641px) and (max-width: 1023px)');
  const compactRoot = blockAfter(compact, ':root');
  const clampMatch = compactRoot.match(
    /--app-effective-dock-width:\s*clamp\(\s*(\d+)px,\s*([\d.]+)vw,\s*var\(--app-dock-width\)\s*\)/
  );
  assert.ok(clampMatch, 'effective dock width must use a bounded clamp');

  const minimum = Number(clampMatch[1]);
  const fluid = Number(clampMatch[2]);
  const desktopClamp = blockAfter(css, ':root').match(
    /--app-dock-width:\s*clamp\(\s*(\d+)px,\s*([\d.]+)vw,\s*(\d+)px\s*\)/
  );
  assert.ok(desktopClamp, 'effective dock maximum must be bounded by the desktop dock token');
  const maximum = Number(desktopClamp[3]);
  assert.ok(minimum >= 300 && minimum < maximum, 'dock minimum must remain usable and bounded');
  assert.ok(fluid > 0, 'dock preferred width must respond to viewport width');
  assert.ok(maximum <= 390, 'compact dock must not exceed the Pencil desktop dock');

  assert.match(blockAfter(compact, '#side-panel'), /width:\s*var\(--app-effective-dock-width\)/);
  assert.match(
    blockAfter(compact, 'body.panel-visible #panel-toggle'),
    /right:\s*var\(--app-effective-dock-width\)/
  );

  const compactDesktop = blockAfter(css, '@media (min-width: 900px) and (max-width: 1023px)');
  assert.match(compactDesktop, /#map[\s\S]*?right:\s*var\(--app-effective-dock-width\)/);
});

test('visible compact panels reserve dock toggle and gap before sizing search', () => {
  const compact = blockAfter(css, '@media (min-width: 641px) and (max-width: 1023px)');
  const compactRoot = blockAfter(compact, ':root');
  const visibleSearch = blockAfter(compact, 'body.panel-visible #map-search-bar');
  const hiddenSearch = blockAfter(compact, 'body.panel-hidden #map-search-bar');

  assert.match(compactRoot, /--app-compact-search-left:\s*var\(--sp-6\)/);
  assert.match(compactRoot, /--app-compact-control-gap:\s*var\(--sp-3\)/);
  assert.match(visibleSearch, /width:\s*min\(/);
  assert.match(visibleSearch, /500px/);
  assert.match(visibleSearch, /100vw/);
  assert.match(visibleSearch, /var\(--app-effective-dock-width\)/);
  assert.match(visibleSearch, /36px/);
  assert.match(visibleSearch, /var\(--app-compact-search-left\)/);
  assert.match(visibleSearch, /var\(--app-compact-control-gap\)/);
  assert.match(hiddenSearch, /width:\s*min\(\s*500px/);
  assert.doesNotMatch(hiddenSearch, /app-effective-dock-width/);

  const clampMatch = compactRoot.match(
    /--app-effective-dock-width:\s*clamp\(\s*(\d+)px,\s*([\d.]+)vw,\s*var\(--app-dock-width\)\s*\)/
  );
  const minimum = Number(clampMatch[1]);
  const fluid = Number(clampMatch[2]);
  const desktopClamp = blockAfter(css, ':root').match(
    /--app-dock-width:\s*clamp\(\s*(\d+)px,\s*([\d.]+)vw,\s*(\d+)px\s*\)/
  );
  const desktopMinimum = Number(desktopClamp[1]);
  const desktopFluid = Number(desktopClamp[2]);
  const desktopMaximum = Number(desktopClamp[3]);
  const left = Number(css.match(/--sp-6:\s*(\d+)px/)[1]);
  const gap = Number(css.match(/--sp-3:\s*(\d+)px/)[1]);
  const clamp = (low, preferred, high) => Math.min(high, Math.max(low, preferred));
  const desktopDockAt = (viewportWidth) => clamp(
    desktopMinimum,
    viewportWidth * desktopFluid / 100,
    desktopMaximum
  );

  [641, 719, 720, 768, 899, 900, 1023].forEach((viewportWidth) => {
    const dockWidth = clamp(minimum, viewportWidth * fluid / 100, desktopDockAt(viewportWidth));
    const searchWidth = Math.min(500, viewportWidth - dockWidth - 36 - left - gap);
    const searchRight = left + searchWidth;
    const toggleLeft = viewportWidth - dockWidth - 36;
    const panelLeft = viewportWidth - dockWidth;
    const hiddenSearchWidth = Math.min(500, viewportWidth - (left * 2));
    const hiddenSearchRight = left + hiddenSearchWidth;
    const hiddenToggleLeft = viewportWidth - 36;
    assert.ok(searchRight <= toggleLeft - gap, `${viewportWidth}px controls must keep their gap`);
    assert.ok(toggleLeft >= 0, `${viewportWidth}px toggle must stay inside the viewport`);
    assert.equal(toggleLeft + 36, panelLeft, `${viewportWidth}px toggle must meet the dock edge`);
    assert.ok(panelLeft + dockWidth <= viewportWidth, `${viewportWidth}px dock must not overflow`);
    assert.ok(
      hiddenSearchRight <= hiddenToggleLeft - gap,
      `${viewportWidth}px hidden-panel search must leave the toggle operable`
    );
  });

  const compactAt1023 = clamp(minimum, 1023 * fluid / 100, desktopDockAt(1023));
  const desktopAt1024 = desktopDockAt(1024);
  assert.ok(Math.abs(desktopAt1024 - compactAt1023) <= 2, '1023px to 1024px dock jump must stay subtle');
});

test('responsive geometry leaves JavaScript display and access thresholds intact', () => {
  assert.match(indexHtml, /const narrowViewMedia = window\.matchMedia\('\(max-width: 640px\)'\);/);
  assert.match(
    indexHtml,
    /let isPanelVisible = narrowViewMedia\.matches \|\| window\.matchMedia\('\(min-width: 900px\)'\)\.matches;/
  );
  assert.match(functionBody('setNarrowView'), /state\.narrowView\s*=\s*narrow/);
  assert.match(functionBody('setNarrowView'), /window\.matchMedia\('\(min-width: 900px\)'\)\.matches/);
  assert.match(functionBody('canEdit'), /state\.narrowView\s*!==\s*true/);
});

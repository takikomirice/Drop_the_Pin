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

function numericToken(source, name, unit) {
  const match = source.match(new RegExp(`${name}:\\s*([\\d.]+)${unit}`));
  assert.ok(match, `missing ${name} ${unit} token`);
  return Number(match[1]);
}

test('expanded dock uses bounded height-aware pin and route tracks', () => {
  const rootBlock = blockAfter(css, ':root');
  const expanded = blockAfter(css, '#side-panel.route-dock-expanded');

  assert.doesNotMatch(expanded, /minmax\(\s*300px\s*,\s*450px\s*\)/);
  assert.match(expanded, /grid-template-rows:/);
  assert.match(expanded, /var\(--dock-pin-region-height\s*,\s*clamp\(/);
  assert.match(expanded, /var\(--dock-pin-region-min-height\)/);
  assert.match(expanded, /var\(--dock-pin-region-preferred-height\)/);
  assert.match(expanded, /var\(--dock-pin-region-max-height\)/);
  assert.match(expanded, /var\(--dock-route-resizer-size\)/);
  assert.match(
    expanded,
    /minmax\(\s*var\(--dock-route-region-effective-min-height,\s*var\(--dock-route-region-min-height\)\)\s*,\s*1fr\s*\)/
  );

  const pinMinimum = numericToken(rootBlock, '--dock-pin-region-min-height', 'px');
  const pinPreferred = numericToken(rootBlock, '--dock-pin-region-preferred-height', '%');
  const pinMaximum = numericToken(rootBlock, '--dock-pin-region-max-height', 'px');
  const routeMinimum = numericToken(rootBlock, '--dock-route-region-min-height', 'px');
  const resizerSize = numericToken(rootBlock, '--dock-route-resizer-size', 'px');
  assert.ok(pinMinimum >= 192 && pinMinimum <= 240, 'pin minimum must keep a header, tabs, and one row usable');
  assert.ok(pinPreferred >= 55 && pinPreferred <= 68, 'pin preferred share must balance both dock regions');
  assert.ok(pinMaximum <= 450, 'normal-height pin allocation must not exceed the current maximum');
  assert.ok(routeMinimum >= 150, 'route minimum must keep its header and one route row usable');
  assert.ok(resizerSize >= 8 && resizerSize <= 16, 'resizer must be easy to target without wasting dock height');
});

test('height allocation preserves both minimums from 520px through 900px', () => {
  const rootBlock = blockAfter(css, ':root');
  const headerHeight = numericToken(rootBlock, '--app-header-height', 'px');
  const pinMinimum = numericToken(rootBlock, '--dock-pin-region-min-height', 'px');
  const pinPreferred = numericToken(rootBlock, '--dock-pin-region-preferred-height', '%');
  const pinMaximum = numericToken(rootBlock, '--dock-pin-region-max-height', 'px');
  const routeMinimum = numericToken(rootBlock, '--dock-route-region-min-height', 'px');
  const resizerSize = numericToken(rootBlock, '--dock-route-resizer-size', 'px');
  const clamp = (low, preferred, high) => Math.min(high, Math.max(low, preferred));

  [520, 600, 768, 900].forEach((viewportHeight) => {
    const availableHeight = viewportHeight - headerHeight;
    const pinHeight = Math.min(
      availableHeight - resizerSize - routeMinimum,
      clamp(pinMinimum, availableHeight * pinPreferred / 100, pinMaximum)
    );
    const routeHeight = availableHeight - pinHeight - resizerSize;
    assert.ok(pinHeight >= pinMinimum, `${viewportHeight}px pin region must meet its minimum`);
    assert.ok(routeHeight >= routeMinimum, `${viewportHeight}px route region must meet its minimum`);
    assert.ok(pinHeight + resizerSize + routeHeight <= availableHeight, `${viewportHeight}px tracks must fit the dock`);
  });

  const normalAvailable = 768 - headerHeight;
  const normalPin = clamp(pinMinimum, normalAvailable * pinPreferred / 100, pinMaximum);
  assert.ok(Math.abs(normalPin - 450) <= 1, '768px height must retain the established pin allocation');
});

test('dock scroll containers and mobile geometry remain unchanged', () => {
  assert.match(blockAfter(css, '#dock-pin-region, #dock-route-region'), /min-height:\s*0/);
  assert.match(blockAfter(css, '#side-placed, #side-unplaced'), /overflow-y:\s*auto/);
  assert.match(blockAfter(css, '#route-list'), /overflow-y:\s*auto/);

  const mobileSource = css.slice(css.indexOf('/* Phase 9.5-F UI-7'));
  const mobile = blockAfter(mobileSource, '@media (max-width: 640px)');
  const mobilePanel = blockAfter(mobile, 'body.narrow-view #side-panel');
  assert.match(mobilePanel, /height:\s*220px/);
  assert.match(mobilePanel, /min-height:\s*220px/);
  assert.match(mobilePanel, /grid-template:\s*auto minmax\(0, 1fr\)/);
  assert.match(mobile, /body\.narrow-view\.mobile-sheet-expanded #side-panel\s*\{\s*height:\s*min\(70dvh, 586px\)/);
});

test('RWD-001 width geometry and route expansion wiring remain intact', () => {
  const compact = blockAfter(css, '@media (min-width: 641px) and (max-width: 1023px)');
  assert.match(compact, /--app-effective-dock-width:\s*clamp\(/);
  assert.match(compact, /#side-panel\s*\{\s*width:\s*var\(--app-effective-dock-width\)/);

  const renderState = functionBody('renderRouteDockState');
  assert.match(renderState, /classList\.toggle\('route-dock-expanded', expanded\)/);
  assert.match(renderState, /list\.hidden\s*=\s*!expanded/);
  assert.match(
    indexHtml,
    /getElementById\('route-dock-toggle'\)\.addEventListener\('click',[\s\S]*?setRouteDockExpanded\(!state\.routeDockExpanded\)/
  );
  assert.match(indexHtml, /const narrowViewMedia = window\.matchMedia\('\(max-width: 640px\)'\)/);
});

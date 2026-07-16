const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const belowDesktopQuery = '@media not all and (min-width: 900px)';

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

test('below-desktop media rules exactly complement the 900px desktop threshold', () => {
  const integerMaxQueries = css.match(/@media\s*\(max-width:\s*899px\)/g) || [];
  assert.equal(
    integerMaxQueries.length,
    0,
    'max-width: 899px leaves fractional CSS viewport widths below 900px unmatched'
  );

  const complementQueries = css.match(/@media\s+not all and \(min-width:\s*900px\)/g) || [];
  assert.equal(complementQueries.length, 5, 'every below-desktop rule must use the exact complement');
});

test('the below-desktop panel stays laid out in both visibility states', () => {
  const belowDesktop = blockAfter(css, belowDesktopQuery);
  const visiblePanel = blockAfter(belowDesktop, 'body.panel-visible #side-panel');
  const hiddenPanel = blockAfter(belowDesktop, 'body.panel-hidden #side-panel');

  assert.match(visiblePanel, /display:\s*grid/);
  assert.match(visiblePanel, /transform:\s*translateX\(0\)/);
  assert.match(visiblePanel, /visibility:\s*visible/);
  assert.match(hiddenPanel, /display:\s*grid/);
  assert.match(hiddenPanel, /transform:\s*translateX\(100%\)/);
  assert.match(hiddenPanel, /visibility:\s*hidden/);
});

test('breakpoint continuity does not change narrow view or panel state wiring', () => {
  assert.match(indexHtml, /const narrowViewMedia = window\.matchMedia\('\(max-width: 640px\)'\);/);
  assert.match(
    indexHtml,
    /let isPanelVisible = narrowViewMedia\.matches \|\| window\.matchMedia\('\(min-width: 900px\)'\)\.matches;/
  );
  assert.match(
    indexHtml,
    /document\.getElementById\('panel-toggle'\)\.addEventListener\('click', \(\) => \{[\s\S]*?isPanelVisible = !isPanelVisible;[\s\S]*?renderPanelToggle\(\);/
  );
});

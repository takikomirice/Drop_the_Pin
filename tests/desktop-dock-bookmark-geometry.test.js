const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];

function blockAfter(source, selector) {
  const start = source.indexOf(selector);
  assert.notEqual(start, -1, `Expected ${selector}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${selector}`);
}

test('desktop dock bookmark uses a compact 44 by 44 target', () => {
  const bookmark = blockAfter(css, '#panel-toggle');

  assert.match(bookmark, /top:\s*calc\(var\(--app-header-height\) \+ 18px\)/);
  assert.match(bookmark, /min-width:\s*44px/);
  assert.match(bookmark, /width:\s*44px/);
  assert.match(bookmark, /min-height:\s*44px/);
  assert.match(bookmark, /height:\s*44px/);
  assert.match(bookmark, /border-radius:\s*8px 0 0 8px/);
  assert.match(blockAfter(css, '.panel-toggle-icon'), /width:\s*16px/);
  assert.match(blockAfter(css, '.panel-toggle-icon'), /height:\s*16px/);
});

test('bookmark keeps theme tokens, panel state wiring, and mobile exclusion', () => {
  const bookmark = blockAfter(css, '#panel-toggle');

  assert.match(bookmark, /background:\s*var\(--color-surface-raised\)/);
  assert.match(bookmark, /color:\s*var\(--color-text-muted\)/);
  assert.match(bookmark, /box-shadow:\s*var\(--shadow-sm\)/);
  assert.match(css, /body\.panel-visible #panel-toggle\s*\{\s*right:\s*var\(--app-dock-width\)/);
  assert.match(css, /body\.panel-hidden #panel-toggle\s*\{\s*right:\s*0/);
  assert.match(
    css,
    /@media \(max-width:\s*640px\)[\s\S]*?body\.narrow-view #panel-toggle,[\s\S]*?display:\s*none !important/
  );
  assert.match(
    indexHtml,
    /class="panel-toggle-icon panel-toggle-icon-close"[\s\S]*?class="panel-toggle-icon panel-toggle-icon-open"/
  );
  assert.match(indexHtml, /button\.setAttribute\('aria-expanded', String\(isPanelVisible\)\)/);
  assert.doesNotMatch(indexHtml, /button\.textContent\s*=\s*isPanelVisible/);
});

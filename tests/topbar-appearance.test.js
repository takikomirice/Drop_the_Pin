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

function elementWithId(id) {
  const match = indexHtml.match(new RegExp(`<([a-z]+)\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/\\1>`));
  assert.ok(match, `Expected element #${id}`);
  return match[0];
}

test('topbar uses the compact height token and keeps a visible search gap', () => {
  assert.match(blockAfter(css, ':root'), /--app-header-height:\s*56px/);
  assert.match(blockAfter(css, '#app-shell'), /inset:\s*var\(--app-header-height\) 0 0/);
  assert.match(blockAfter(css, '#panel-toggle'), /top:\s*calc\(var\(--app-header-height\) \+ 18px\)/);
  assert.match(blockAfter(css, '#map-search-bar'), /top:\s*calc\(var\(--app-header-height\) \+ 16px\)/);
  assert.match(
    css,
    /body\.narrow-view #map-search-bar\s*\{[\s\S]*?top:\s*calc\(var\(--app-header-height\) \+ 12px\)/
  );
});

test('persistent topbar actions are borderless and transparent until hover', () => {
  const button = blockAfter(css, '#topbar .text-btn');
  const hover = blockAfter(css, '#topbar .text-btn:hover');

  assert.match(button, /min-height:\s*44px/);
  assert.match(button, /border:\s*0/);
  assert.match(button, /background:\s*transparent/);
  assert.match(hover, /background:\s*var\(--color-surface-muted\)/);
});

test('topbar and its menu use inline line icons without emoji or icon dependencies', () => {
  const topbar = indexHtml.slice(
    indexHtml.indexOf('<div id="topbar">'),
    indexHtml.indexOf('<div class="share-banner">')
  );
  [
    'topbar-brand-icon',
    'topbar-share-icon',
    'topbar-data-icon',
    'topbar-more-icon',
    'topbar-preview-icon',
    'panel-toggle-icon',
    'topbar-settings-icon',
    'topbar-help-icon',
    'theme-icon-moon',
    'theme-icon-sun',
    'topbar-drive-icon'
  ].forEach((className) => {
    assert.match(topbar, new RegExp(`<svg[^>]*class="[^"]*${className}[^"]*"`), `missing ${className}`);
  });
  assert.doesNotMatch(topbar, /📍|⚙|🌙|☀|↗/u);
  assert.doesNotMatch(topbar, /<img\b|<use\b|iconify|font-awesome/i);
});

test('icon-only topbar controls retain accessible names and standard actions retain 44px targets', () => {
  const panelToggle = elementWithId('panel-toggle');
  const moreToggle = elementWithId('more-menu-toggle');

  assert.match(panelToggle, /aria-label="一覧を閉じる"/);
  assert.match(panelToggle, /title="一覧を閉じる"/);
  assert.match(moreToggle, /aria-label="その他の操作"/);
  assert.match(moreToggle, /title="その他の操作"/);
  assert.match(blockAfter(css, '#topbar .icon-btn'), /min-width:\s*44px/);
  assert.match(blockAfter(css, '#topbar .icon-btn'), /min-height:\s*44px/);
  assert.match(blockAfter(css, 'body.narrow-view .mobile-icon-action'), /width:\s*44px/);
  assert.match(blockAfter(css, 'body.narrow-view .mobile-icon-action'), /min-width:\s*44px/);
});

test('theme, preview and panel rendering preserve inline SVG nodes and existing wiring', () => {
  assert.doesNotMatch(indexHtml, /getElementById\('theme-toggle-icon'\)\.textContent/);
  assert.match(indexHtml, /toggle\.querySelector\('\.topbar-action-label'\)\.textContent\s*=/);
  assert.doesNotMatch(indexHtml, /button\.textContent\s*=\s*isPanelVisible/);

  assert.match(indexHtml, /getElementById\('theme-toggle'\)\.addEventListener\('click'/);
  assert.match(indexHtml, /getElementById\('more-menu-toggle'\)\.addEventListener\('click'/);
  assert.match(indexHtml, /getElementById\('edit-toggle'\)\.addEventListener\('click'/);
  assert.match(indexHtml, /getElementById\('panel-toggle'\)\.addEventListener\('click'/);
});

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

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

test('map search keeps its wiring IDs with a concise single-line field', () => {
  ['map-search-bar', 'map-search-bar-row1', 'map-search-input', 'map-search-toggle']
    .forEach((id) => assert.equal(countId(id), 1, `${id} must exist exactly once`));
  assert.equal(countId('map-search-hint'), 0, 'the redundant search hint must be removed');

  assert.match(
    indexHtml,
    /<span class="map-search-leading-icon" aria-hidden="true">[\s\S]*?<svg[\s\S]*?<\/svg>[\s\S]*?<\/span>/
  );
  assert.match(
    indexHtml,
    /<div class="map-search-field">[\s\S]*?id="map-search-input"[\s\S]*?placeholder="ピン・タグ・場所を検索"[\s\S]*?<\/div>/
  );
  assert.doesNotMatch(indexHtml, /入力で絞り込み • Enterで場所へ移動/);
  assert.doesNotMatch(indexHtml, /aria-describedby="map-search-hint"/);
  assert.match(
    indexHtml,
    /id="map-search-toggle"[^>]*aria-label="フィルタ・ソート"[^>]*aria-controls="map-search-expanded"[^>]*aria-expanded="false"/
  );
});

test('desktop map search uses a low-profile bordered input field', () => {
  const desktopSource = css.slice(css.indexOf('/* Desktop map search hierarchy */'));
  const desktop = blockAfter(desktopSource, '@media (min-width: 641px) {');
  const shell = blockAfter(desktop, '#map-search-bar');
  const row = blockAfter(desktop, '#map-search-bar-row1');
  const icon = blockAfter(desktop, '.map-search-leading-icon');
  const field = blockAfter(desktop, '.map-search-field');
  const input = blockAfter(desktop, '#map-search-input');
  const placeholder = blockAfter(desktop, '#map-search-input::placeholder');
  const toggle = blockAfter(desktop, '#map-search-toggle');

  assert.match(shell, /padding:\s*4px 7px 4px 12px/);
  assert.match(shell, /border-radius:\s*12px/);
  assert.match(row, /min-height:\s*34px/);
  assert.match(row, /gap:\s*8px/);
  assert.match(icon, /display:\s*inline-flex/);
  assert.match(icon, /width:\s*18px/);
  assert.match(icon, /height:\s*18px/);
  assert.match(field, /height:\s*32px/);
  assert.match(field, /padding:\s*0 8px/);
  assert.match(field, /border:\s*1px solid color-mix/);
  assert.match(field, /border-radius:\s*8px/);
  assert.match(input, /height:\s*30px/);
  assert.match(input, /padding:\s*0/);
  assert.match(input, /border:\s*0/);
  assert.match(input, /font-size:\s*14px/);
  assert.match(input, /font-weight:\s*500/);
  assert.match(placeholder, /color:\s*var\(--color-text-muted\)/);
  assert.match(toggle, /width:\s*44px/);
  assert.match(toggle, /height:\s*44px/);
  assert.match(toggle, /flex:\s*0 0 44px/);
});

test('expanded search controls are lower and gain one pixel of horizontal room', () => {
  assert.match(indexHtml, /id="map-sort"[^>]*style="padding:7px 9px;font-size:12px;"/);

  const filterButtons = blockAfter(css, '#map-search-controls .map-filter-btn');
  assert.match(filterButtons, /min-height:\s*44px/);
  assert.match(filterButtons, /padding:\s*5px 13px/);

  const iconButtons = blockAfter(css, '#map-icon-filter .icon-filter-chip');
  assert.match(iconButtons, /min-width:\s*43px/);
});

test('mobile search remains single-line and existing search wiring stays intact', () => {
  const leading = blockAfter(css, '.map-search-leading-icon');
  const row = blockAfter(css, '#map-search-bar-row1');
  const field = blockAfter(css, '.map-search-field');
  const mobileBar = blockAfter(css, 'body.narrow-view #map-search-bar');
  const mobileToggle = blockAfter(css, 'body.narrow-view #map-search-toggle');

  assert.match(leading, /display:\s*none/);
  assert.match(mobileBar, /top:\s*calc\(var\(--app-header-height\) \+ 12px\)/);
  assert.match(mobileBar, /left:\s*16px/);
  assert.match(mobileBar, /width:\s*calc\(100vw - 32px\)/);
  assert.match(mobileBar, /padding:\s*5px 7px/);
  assert.match(row, /display:\s*flex/);
  assert.match(field, /flex:\s*1/);
  assert.match(field, /min-width:\s*0/);
  assert.match(indexHtml, /body\.narrow-view #map-search-bar-row1\s*\{[\s\S]*?min-height:\s*32px/);
  assert.match(indexHtml, /body\.narrow-view #map-search-input\s*\{[\s\S]*?font-size:\s*13px/);
  assert.match(mobileToggle, /flex:\s*0 0 44px/);
  assert.match(mobileToggle, /width:\s*44px/);
  assert.match(mobileToggle, /min-width:\s*44px/);
  assert.match(mobileToggle, /height:\s*44px/);
  assert.match(mobileToggle, /min-height:\s*44px/);
  assert.doesNotMatch(mobileToggle, /(?:width|height):\s*32px/);
  assert.match(
    indexHtml,
    /document\.getElementById\('map-search-input'\)\.addEventListener\('input',[\s\S]*?state\.listQuery = this\.value/
  );
  assert.match(
    indexHtml,
    /document\.getElementById\('map-search-toggle'\)\.addEventListener\('click', function\(\) \{\s*const expanded = document\.getElementById\('map-search-expanded'\)/
  );
  assert.match(indexHtml, /this\.setAttribute\('aria-expanded', String\(!isOpen\)\)/);
  assert.match(indexHtml, /setupGeoSearch\([\s\S]*?document\.getElementById\('map-search-input'\)/);
});

test('pin add is a compact header control and the old map FAB is absent', () => {
  const addButton = blockAfter(css, '#pin-add-btn {');
  assert.match(addButton, /min-width:\s*44px/);
  assert.match(addButton, /min-height:\s*44px/);
  assert.match(addButton, /border:\s*0/);
  assert.doesNotMatch(css, /#fab\s*\{/);
});

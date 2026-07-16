const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];

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

function cssBlock(selector, source = css) {
  const start = source.indexOf(selector);
  assert.notEqual(start, -1, `Expected CSS selector ${selector}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(open + 1, index);
  }
  assert.fail(`Could not parse CSS selector ${selector}`);
}

test('geocode results are native 44px buttons with one activation path', () => {
  const render = functionSource('renderGeocodeResults');
  assert.match(render, /document\.createElement\('button'\)/);
  assert.match(render, /item\.type\s*=\s*'button'/);
  assert.match(render, /item\.setAttribute\('aria-label'/);
  assert.match(render, /item\.addEventListener\('click'/);
  assert.doesNotMatch(render, /item\.addEventListener\('keydown'/);
  assert.doesNotMatch(render, /var item = document\.createElement\('div'\)/);

  const item = cssBlock('.geocode-result-item');
  assert.match(item, /min-height:\s*44px/);
  assert.match(item, /width:\s*100%/);
  assert.match(css, /\.geocode-result-item:(?:hover|focus-visible)[^{]*\{|\.geocode-result-item\.is-selected/);
});

test('pin and imported route cards keep expansion and visibility as sibling buttons', () => {
  for (const name of ['buildRouteItem', 'buildUnifiedTrackRouteItem']) {
    const source = functionSource(name);
    assert.match(source, /header\.className\s*=\s*'route-card-header'/, name);
    assert.match(source, /summary\.type\s*=\s*'button'/, name);
    assert.match(source, /summary\.setAttribute\('aria-expanded'/, name);
    assert.match(source, /summary\.setAttribute\('aria-label'/, name);
    assert.match(source, /visibility\.type\s*=\s*'button'/, name);
    assert.match(source, /visibility\.setAttribute\('aria-pressed'/, name);
    assert.match(source, /visibility\.setAttribute\('aria-label'/, name);
    assert.match(source, /header\.appendChild\(summary\)[\s\S]*header\.appendChild\(visibility\)/, name);
    assert.doesNotMatch(source, /summary\.appendChild\(visibility\)/, name);
    assert.doesNotMatch(source, /visibility\.setAttribute\('role',\s*'button'\)/, name);
    assert.doesNotMatch(source, /visibility\.setAttribute\('tabindex'/, name);
    assert.doesNotMatch(source, /visibility\.addEventListener\('keydown'/, name);
  }
  assert.match(functionSource('attachRouteGroupSortable'), /handle:\s*'\.route-summary'/);
  assert.match(functionSource('attachRoutePinSortable'), /filter:\s*'\.route-pin-remove'/);
});

test('all audited controls expose a real 44 by 44 minimum hit area', () => {
  const expectations = [
    ['#panel-toggle', /min-width:\s*44px/, /min-height:\s*44px/],
    ['.dock-pin-tab {', /min-height:\s*44px/],
    ['.imported-route-actions .ghost-btn', /min-height:\s*44px/],
    ['.route-setting-select', /height:\s*44px/],
    ['.route-pin-remove', /width:\s*44px/, /height:\s*44px/],
    ['#map-search-controls .map-filter-btn', /min-height:\s*44px/]
  ];
  for (const [selector, ...patterns] of expectations) {
    const block = cssBlock(selector);
    for (const pattern of patterns) assert.match(block, pattern, selector);
  }
  const desktopSearch = css.slice(css.indexOf('/* Desktop map search hierarchy */'));
  const toggle = cssBlock('#map-search-toggle', desktopSearch);
  assert.match(toggle, /width:\s*44px/);
  assert.match(toggle, /height:\s*44px/);
  assert.match(css, /\.route-pin-row\s*\{[^}]*grid-template-columns:\s*26px minmax\(0, 1fr\) 44px/);
});

test('rendered route line-style selects use the audited 44px control class', () => {
  const routeSettingSelects = Array.from(
    indexHtml.matchAll(/<select\b[^>]*class="[^"]*\broute-setting-select\b[^"]*"[^>]*>/g),
    (match) => match[0]
  );

  assert.ok(routeSettingSelects.length >= 1, 'expected at least one rendered route-setting-select');
  for (const select of routeSettingSelects) {
    assert.match(select, /\bclass="[^"]*\bform-select\b[^"]*"/, select);
  }
  assert.match(cssBlock('.route-setting-select'), /height:\s*44px/);
});

test('dynamic operation renderers use decorative SVG instead of character icons', () => {
  assert.match(functionSource('createActionIconElement'), /createElementNS/);
  assert.match(functionSource('createActionIconElement'), /aria-hidden/);
  assert.match(functionSource('renderFolders'), /createActionIconElement[\s\S]*folder[\s\S]*chevron-right/);
  assert.match(functionSource('loadUploadFolder'), /createActionIconElement[\s\S]*folder[\s\S]*chevron-right/);
  assert.match(functionSource('buildRoutePinList'), /createActionIconElement[\s\S]*close/);
  assert.doesNotMatch(indexHtml, /📁|›|✖/u);
  assert.match(functionSource('renderFolders'), /button\.setAttribute\('aria-label'/);
  assert.match(functionSource('loadUploadFolder'), /button\.setAttribute\('aria-label'/);
  assert.match(functionSource('buildRoutePinList'), /removeButton\.setAttribute\('aria-label'/);
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionBody(name) {
  const start = indexHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${name}`);
}

function cssRule(selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = indexHtml.match(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`));
  assert.ok(match, `Expected CSS rule for ${selector}`);
  return match[1];
}

test('desktop pin tabs use a slightly lower segmented surface', () => {
  const tabs = cssRule('.dock-pin-tabs');
  const tab = cssRule('.dock-pin-tab');
  assert.match(tabs, /background:\s*var\(--color-surface-muted\)/);
  assert.match(tabs, /border-radius:\s*var\(--radius-md\)/);
  assert.match(tabs, /padding:\s*3px var\(--sp-1\)/);
  assert.match(tab, /min-height:\s*44px/);
  assert.match(tab, /font-size:\s*14px/);
});

test('dock headings stay compact and the pin title starts with an inline pin icon', () => {
  const routeHeader = cssRule('#route-panel-header');
  const expandedRouteHeader = cssRule('#side-panel.route-dock-expanded #route-panel-header');
  const pinTitle = cssRule('.dock-pin-title');
  const pinTitleIcon = cssRule('.dock-pin-title-icon');

  assert.match(routeHeader, /min-height:\s*63px/);
  assert.match(expandedRouteHeader, /min-height:\s*45px/);
  assert.match(indexHtml, /body\.narrow-view #route-panel-header\s*\{[^}]*min-height:\s*45px/);
  assert.match(pinTitle, /display:\s*inline-flex/);
  assert.match(pinTitle, /align-items:\s*center/);
  assert.match(pinTitleIcon, /width:\s*18px/);
  assert.match(pinTitleIcon, /height:\s*20px/);
  assert.match(
    indexHtml,
    /<h2 id="dock-pin-title"[^>]*>\s*<svg class="dock-pin-title-icon"[^>]*aria-hidden="true">[\s\S]*?<\/svg>\s*<span>ピン<\/span>\s*<\/h2>/
  );
});

test('pin rows accommodate the refreshed marker on desktop and mobile', () => {
  const row = cssRule('#side-panel .pin-list-row');
  assert.match(row, /min-height:\s*62px/);
  assert.match(row, /margin-bottom:\s*5px/);
  assert.match(row, /border-radius:\s*10px/);
  const main = cssRule('#side-panel .pin-list-row .list-item');
  assert.match(main, /padding:\s*6px 5px/);
  assert.match(main, /border:\s*0/);
  assert.match(cssRule('#side-panel .list-title'), /font-size:\s*14px/);
  assert.match(cssRule('#side-panel .list-subtitle'), /font-size:\s*11px/);
  assert.match(indexHtml, /body\.narrow-view #side-panel \.pin-list-row\s*\{[^}]*min-height:\s*66px[^}]*margin-bottom:\s*5px/);
  assert.match(indexHtml, /body\.narrow-view #side-panel \.pin-list-row \.list-item\s*\{[^}]*min-height:\s*64px[^}]*padding:\s*7px 5px/);
});

test('desktop bulk toolbar fits one row without horizontal scrolling and keeps 44px actions', () => {
  const toolbar = cssRule('#bulk-action-bar');
  const action = cssRule('#bulk-action-bar button');
  const count = cssRule('#bulk-count');
  assert.match(toolbar, /min-height:\s*56px/);
  assert.match(toolbar, /overflow-x:\s*hidden/);
  assert.doesNotMatch(toolbar, /overflow-x:\s*auto/);
  assert.match(toolbar, /border-block:\s*var\(--border-width\) solid var\(--border\)/);
  assert.match(toolbar, /border-radius:\s*var\(--radius-sm\)/);
  assert.match(toolbar, /background:\s*var\(--color-surface-muted\)/);
  assert.match(action, /width:\s*44px/);
  assert.match(action, /height:\s*44px/);
  assert.match(action, /flex:\s*0 0 44px/);
  assert.match(action, /padding:\s*0/);
  assert.match(count, /flex:\s*1 1 auto/);
  assert.match(count, /min-width:\s*0/);
  assert.doesNotMatch(indexHtml, /@media \(min-width:\s*641px\) and \(max-width:\s*1279px\)[\s\S]*?#bulk-action-bar/);
  assert.ok(24 + 16 + (3 * 44) + (2 * 2) <= 340, 'toolbar fixed geometry fits at 340px');
});

test('editable pin cards expose and visually synchronize their pressed state', () => {
  const build = functionBody('buildListItem');
  const sync = functionBody('syncListItemSelection');

  assert.match(build, /button\.setAttribute\('aria-pressed',\s*String\(isSelected\)\)/);
  assert.match(build, /row\.className = 'pin-list-row' \+ \(isSelected \? ' is-selected' : ''\)/);
  assert.match(build, /class="pin-check"[^>]*aria-hidden="true"/);
  assert.doesNotMatch(build, /class="pin-check"[^>]*style=/);
  assert.match(sync, /button\.setAttribute\('aria-pressed',\s*String\(isSelected\)\)/);
  assert.match(sync, /button\.closest\('\.pin-list-row'\)/);
  assert.match(sync, /row\.classList\.toggle\('is-selected',\s*isSelected\)/);
  assert.match(indexHtml, /#side-panel \.pin-list-row\.is-selected\s*\{[^}]*border-color:\s*var\(--accent\)[^}]*border-left-width:\s*4px[^}]*background:\s*var\(--accent-soft\)/);
  assert.match(indexHtml, /#side-panel \.pin-list-row\.is-selected \.pin-check\s*\{[^}]*border-color:\s*var\(--accent\)[^}]*color:\s*var\(--accent\)/);
  assert.match(cssRule('.pin-check'), /border-radius:\s*3px/);
  assert.match(cssRule('.pin-check'), /border:\s*2px solid var\(--color-text-muted\)/);
});

test('selection synchronizer derives semantics and visuals from selectedPinIds', () => {
  function run(selectedIds) {
    const attributes = {};
    const classes = {};
    const check = { textContent: 'stale' };
    const row = { classList: { toggle(name, active) { classes[name] = active; } } };
    const button = {
      setAttribute(name, value) { attributes[name] = value; },
      closest(selector) { return selector === '.pin-list-row' ? row : null; },
      querySelector(selector) { return selector === '.pin-check' ? check : null; }
    };
    const sync = vm.runInNewContext(
      `(function(button, pinId) {${functionBody('syncListItemSelection')}})`,
      { state: { selectedPinIds: new Set(selectedIds) } }
    );
    sync(button, 'pin-a');
    return { attributes, classes, check };
  }

  const selected = run(['pin-a']);
  assert.equal(selected.attributes['aria-pressed'], 'true');
  assert.equal(selected.classes['is-selected'], true);
  assert.equal(selected.check.textContent, '✓');

  const unselected = run([]);
  assert.equal(unselected.attributes['aria-pressed'], 'false');
  assert.equal(unselected.classes['is-selected'], false);
  assert.equal(unselected.check.textContent, '');
});

test('selection hierarchy preserves the existing selectedPinIds behavior and edit guards', () => {
  const build = functionBody('buildListItem');
  assert.match(build, /state\.selectedPinIds\.has\(pin\.id\)/);
  assert.match(build, /state\.selectedPinIds\.delete\(pin\.id\)/);
  assert.match(build, /state\.selectedPinIds\.add\(pin\.id\)/);
  assert.match(build, /updateBulkBar\(\)/);
  assert.match(build, /if \(canEdit\(\)\)[\s\S]*aria-pressed/);
  assert.match(build, /else[\s\S]*button\.setAttribute\('aria-label'/);
});

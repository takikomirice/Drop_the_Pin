const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');
const css = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const bodyMarkup = sharedHtml.slice(sharedHtml.indexOf('<body'), sharedHtml.indexOf('<script>', sharedHtml.indexOf('<body')));

function functionSource(name) {
  const start = sharedHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const bodyStart = sharedHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < sharedHtml.length; index += 1) {
    const character = sharedHtml[index];
    if (quote) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'" || character === '`') {
      quote = character;
      continue;
    }
    if (character === '{') depth += 1;
    if (character === '}') {
      depth -= 1;
      if (depth === 0) return sharedHtml.slice(start, index + 1);
    }
  }
  assert.fail(`Could not parse ${name}`);
}

function elementWithId(id) {
  const match = bodyMarkup.match(new RegExp(`<button\\b[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?<\\/button>`));
  assert.ok(match, `Expected button #${id}`);
  return match[0];
}

function visibleText(markup) {
  return markup.replace(/<svg\b[\s\S]*?<\/svg>/g, '').replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function hexToRgb(hex) {
  const value = Number.parseInt(hex.slice(1), 16);
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255];
}

function luminance(hex) {
  const channels = hexToRgb(hex).map((channel) => {
    const normalized = channel / 255;
    return normalized <= 0.04045 ? normalized / 12.92 : ((normalized + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}

function contrast(foreground, background) {
  const light = Math.max(luminance(foreground), luminance(background));
  const dark = Math.min(luminance(foreground), luminance(background));
  return (light + 0.05) / (dark + 0.05);
}

function themeBlock(selector) {
  const start = css.indexOf(selector);
  assert.notEqual(start, -1, `Expected theme selector ${selector}`);
  const open = css.indexOf('{', start);
  const close = css.indexOf('}', open);
  return css.slice(open + 1, close);
}

function token(block, name) {
  const match = block.match(new RegExp(`${name}:\\s*(#[0-9a-fA-F]{6})`));
  assert.ok(match, `Expected ${name} to be a six-digit hex token`);
  return match[1].toLowerCase();
}

test('shared viewport and shell consume safe-area insets without fixed fallback padding', () => {
  assert.match(sharedHtml, /<meta name="viewport" content="width=device-width, initial-scale=1\.0, viewport-fit=cover">/);
  ['top', 'right', 'bottom', 'left'].forEach((edge) => {
    assert.match(css, new RegExp(`--safe-area-${edge}:\\s*env\\(safe-area-inset-${edge},\\s*0px\\)`));
  });
  assert.match(css, /#shared-topbar[\s\S]*var\(--safe-area-top\)[\s\S]*var\(--safe-area-right\)[\s\S]*var\(--safe-area-left\)/);
  assert.match(css, /#shared-app-shell[\s\S]*var\(--safe-area-left\)[\s\S]*var\(--safe-area-right\)/);
  assert.match(css, /#shared-pin-list,[\s\S]*#shared-route-list[\s\S]*var\(--safe-area-bottom\)/);
});

test('shared light and dark themes expose the viewing token contract', () => {
  const light = themeBlock(':root, [data-theme="light"]');
  const dark = themeBlock('[data-theme="dark"]');
  [light, dark].forEach((block) => {
    [
      '--color-surface', '--color-surface-raised', '--color-surface-muted',
      '--color-text', '--color-text-muted', '--accent', '--accent-soft', '--accent-strong',
      '--on-accent', '--border', '--form-bg', '--form-border',
      '--route-type-pin-foreground', '--route-type-pin-background',
      '--route-type-gpx-foreground', '--route-type-gpx-background',
      '--route-type-geojson-foreground', '--route-type-geojson-background'
    ].forEach((name) => assert.match(block, new RegExp(`${name}:`), `${name} missing from theme`));
  });

  assert.match(themeBlock(':root'), /--focus:/);
  assert.match(themeBlock(':root'), /--selected-swatch-ring-light:/);
  assert.match(themeBlock(':root'), /--selected-swatch-ring-dark:/);

  ['pin', 'gpx', 'geojson'].forEach((kind) => {
    assert.ok(contrast(token(light, `--route-type-${kind}-foreground`), token(light, `--route-type-${kind}-background`)) >= 4.5);
    assert.ok(contrast(token(dark, `--route-type-${kind}-foreground`), token(dark, `--route-type-${kind}-background`)) >= 4.5);
  });
});

test('topbar and search icon controls use decorative inline SVGs and 44px icon buttons', () => {
  ['shared-help-open-btn', 'shared-theme-toggle', 'shared-panel-toggle', 'shared-search-toggle'].forEach((id) => {
    const markup = elementWithId(id);
    assert.match(markup, /class="[^"]*shared-icon-btn[^"]*"/);
    assert.match(markup, /aria-label="[^"]+"/);
    assert.match(markup, /title="[^"]+"/);
    assert.match(markup, /<svg\b[^>]*aria-hidden="true"/);
    assert.equal(visibleText(markup), '');
  });
  assert.doesNotMatch(bodyMarkup, /[?🌙☀️🗺️📋▼▲]/u);
  assert.match(css, /\.shared-icon-btn\s*\{[\s\S]*?width:\s*44px[\s\S]*?height:\s*44px[\s\S]*?padding:\s*0/);
  assert.match(css, /\.shared-icon-btn svg\s*\{[\s\S]*?(?:width|inline-size):\s*(?:18|19|20)px[\s\S]*?(?:height|block-size):\s*(?:18|19|20)px/);
});

test('all static button SVGs are decorative and text buttons share the common target contract', () => {
  for (const button of bodyMarkup.matchAll(/<button\b[^>]*>[\s\S]*?<\/button>/g)) {
    for (const svg of button[0].matchAll(/<svg\b[^>]*>/g)) {
      assert.match(svg[0], /aria-hidden="true"/);
    }
  }
  assert.match(css, /\.shared-text-btn\s*\{[\s\S]*?min-height:\s*44px[\s\S]*?border-radius:\s*10px[\s\S]*?padding:\s*[^;]*12px[\s\S]*?white-space:\s*nowrap/);
  assert.doesNotMatch(css, /\.shared-control-btn\.is-off\s*\{[^}]*opacity:/);
  assert.doesNotMatch(css, /\.shared-route-display-mode\.is-off\s*\{[^}]*opacity:/);
});

test('theme panel and search state updates preserve the inline SVG markup', () => {
  assert.doesNotMatch(functionSource('setTheme'), /textContent/);
  assert.doesNotMatch(functionSource('renderSharedPanelUi'), /toggle\.textContent/);
  assert.doesNotMatch(functionSource('renderSharedSearchUi'), /searchToggle\.textContent/);
  assert.match(functionSource('setTheme'), /setAttribute\('aria-label'/);
  assert.match(functionSource('renderSharedPanelUi'), /setAttribute\('aria-expanded'/);
  assert.match(functionSource('renderSharedSearchUi'), /setAttribute\('aria-expanded'/);
});

test('geocode results are native full-width buttons with an accessible combined name', () => {
  const renderer = functionSource('renderGeocodeResults');
  assert.match(renderer, /document\.createElement\('button'\)/);
  assert.match(renderer, /item\.type\s*=\s*'button'/);
  assert.match(renderer, /item\.setAttribute\('aria-label',\s*name\s*\+\s*'[^']*'\s*\+\s*address\)/);
  assert.doesNotMatch(renderer, /item\.addEventListener\('keydown'/);
  assert.match(css, /\.geocode-result-item\s*\{[\s\S]*?width:\s*100%[\s\S]*?min-height:\s*44px[\s\S]*?text-align:\s*left/);
  assert.match(css, /#shared-search-input\s*\{[\s\S]*?min-width:\s*0/);
});

test('tag and color filters expose pressed state and a non-color selected indicator', () => {
  const tags = functionSource('renderSharedTagChips');
  const colors = functionSource('renderSharedColorFilter');
  assert.match(tags, /setAttribute\('aria-pressed'/);
  assert.match(tags, /setAttribute\('aria-label'/);
  assert.match(colors, /setAttribute\('aria-pressed'/);
  assert.match(colors, /setAttribute\('aria-label'/);
  assert.match(css, /\.shared-color-swatch\.active::after\s*\{[\s\S]*?content:\s*['"]✓['"][\s\S]*?var\(--selected-swatch-ring-light\)[\s\S]*?var\(--selected-swatch-ring-dark\)/u);
});

test('route cards use sibling native controls and index-matching route order text', () => {
  const pinBuilder = functionSource('buildSharedPinRouteCard');
  const trackBuilder = functionSource('buildSharedTrackRouteCard');
  const pinListBuilder = functionSource('buildSharedRoutePinList');
  [pinBuilder, trackBuilder].forEach((builder) => {
    assert.doesNotMatch(builder, /item\.tabIndex\s*=/);
    assert.doesNotMatch(builder, /item\.setAttribute\('role',\s*'button'\)/);
    assert.match(builder, /fitButton\s*=\s*document\.createElement\('button'\)/);
    assert.match(builder, /fitButton\.className\s*=\s*'route-fit'/);
    assert.match(builder, /visibility\s*=\s*document\.createElement\('button'\)/);
    assert.match(builder, /header\.appendChild\(fitButton\);[\s\S]*header\.appendChild\(visibility\);/);
  });
  assert.match(pinListBuilder, /action\s*=\s*document\.createElement\('button'\)/);
  assert.doesNotMatch(pinListBuilder, /order\.style\.(?:backgroundColor|color)/);
  assert.match(css, /\.route-fit\s*\{[\s\S]*?min-height:\s*44px/);
  assert.match(css, /\.route-pin-order\s*\{[^}]*color:\s*var\(--text-sub\)[^}]*font-size:\s*10px[^}]*text-align:\s*right/);
});

test('mobile viewing shell keeps tabs and one contained body scroller above safe area', () => {
  ['shared-mobile-pins-tab', 'shared-mobile-routes-tab'].forEach((id) => {
    const markup = elementWithId(id);
    assert.match(markup, /role="tab"/);
    assert.match(markup, /aria-selected="(?:true|false)"/);
  });
  assert.match(css, /#shared-mobile-sheet-tabs button\s*\{[^}]*min-height:\s*44px/);
  assert.match(css, /#shared-side-panel\s*\{[^}]*overflow:\s*hidden/);
  assert.match(css, /@media \(max-width:\s*640px\)[\s\S]*?#shared-pin-list,\s*#shared-route-list\s*\{[^}]*overflow-y:\s*auto[^}]*var\(--safe-area-bottom\)/);
});

test('reduced motion disables shared highlight and panel transitions', () => {
  const reduced = css.slice(css.lastIndexOf('@media (prefers-reduced-motion: reduce)'));
  assert.match(reduced, /animation-duration:\s*0\.01ms\s*!important/);
  assert.match(reduced, /transition-duration:\s*0\.01ms\s*!important/);
  assert.match(reduced, /#shared-map[^}]*transition:\s*none/);
});

test('shared read-only IDs and event wiring remain present exactly once', () => {
  [
    'shared-topbar', 'shared-map', 'shared-map-search-bar', 'shared-search-input',
    'shared-search-toggle', 'shared-geocode-results', 'shared-list-panel', 'shared-route-list',
    'shared-list', 'shared-help-overlay', 'shared-detail-overlay', 'shared-theme-toggle',
    'shared-panel-toggle', 'shared-mobile-pins-tab', 'shared-mobile-routes-tab'
  ].forEach((id) => {
    assert.equal((bodyMarkup.match(new RegExp(`\\bid="${id}"`, 'g')) || []).length, 1, `#${id} changed`);
  });
  ['shared-search-toggle', 'shared-theme-toggle', 'shared-panel-toggle', 'shared-mobile-pins-tab', 'shared-mobile-routes-tab'].forEach((id) => {
    assert.match(sharedHtml, new RegExp(`getElementById\\('${id}'\\)\\.addEventListener\\('click'`));
  });
  assert.doesNotMatch(sharedHtml, /shared-(?:edit|delete|save|import|preset|bulk)-/);
});

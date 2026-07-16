const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');

function functionSource(html, name) {
  const start = html.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const bodyStart = html.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < html.length; index += 1) {
    const character = html[index];
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
      if (depth === 0) return html.slice(start, index + 1);
    }
  }
  assert.fail(`Could not parse ${name}`);
}

function constantArray(html, name) {
  const match = html.match(new RegExp(`const ${name} = \\[([\\s\\S]*?)\\n\\s*\\];`));
  assert.ok(match, `Expected ${name}`);
  return vm.runInNewContext(`[${match[1]}]`);
}

function escHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function loadIndexRenderer() {
  const context = {
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: constantArray(indexHtml, 'PIN_COLORS'),
    PIN_ICONS: constantArray(indexHtml, 'PIN_ICONS'),
    escHtml,
    L: { divIcon: (options) => options }
  };
  vm.runInNewContext([
    'safeColor', 'hexToRgb', 'relativeLuminance', 'contrastRatio', 'getReadableTextColor',
    'normalizeIcon', 'pinIconGlyph', 'createPinMarkerSvg', 'createRouteNumberBadge',
    'createPinIcon', 'createNumberedPinIcon'
  ].map((name) => functionSource(indexHtml, name)).join('\n'), context);
  return context;
}

function loadSharedRenderer() {
  const context = {
    SHARED_DEFAULT_COLOR: '#e53935',
    SHARED_SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    SHARED_PIN_ICONS: constantArray(sharedHtml, 'SHARED_PIN_ICONS'),
    escHtml,
    L: { divIcon: (options) => options }
  };
  vm.runInNewContext([
    'safeColor', 'hexToRgb', 'relativeLuminance', 'contrastRatio', 'getReadableTextColor',
    'normalizeSharedIcon', 'sharedPinIconGlyph', 'createPinMarkerSvg', 'createRouteNumberBadge',
    'sharedPinListIconMarkup', 'createPinIcon', 'createSharedNumberedPinIcon'
  ].map((name) => functionSource(sharedHtml, name)).join('\n'), context);
  return context;
}

test('shared and index render identical teardrops for all 15 colors and 8 icons', () => {
  const index = loadIndexRenderer();
  const shared = loadSharedRenderer();
  const colors = Array.from(index.PIN_COLORS, (item) => item.hex);
  const icons = Array.from(index.PIN_ICONS, (item) => item.id);

  assert.equal(colors.length, 15);
  assert.equal(icons.length, 8);
  for (const color of colors) {
    for (const icon of icons) {
      const expectedSvg = index.createPinMarkerSvg(color, icon, false, 'map-pin-svg');
      assert.equal(shared.createPinMarkerSvg(color, icon, false, 'map-pin-svg'), expectedSvg, `${color} ${icon}`);
      assert.equal(
        shared.sharedPinListIconMarkup({ color, icon }, 'list-thumb-icon'),
        index.createPinMarkerSvg(color, icon, false, 'list-thumb-icon'),
        `list ${color} ${icon}`
      );
      const expected = index.createPinIcon(color, false, false, icon);
      const actual = shared.createPinIcon(color, false, false, icon);
      assert.equal(actual.html, expected.html, `map ${color} ${icon}`);
      assert.equal(JSON.stringify(actual.iconSize), JSON.stringify(expected.iconSize));
      assert.equal(JSON.stringify(actual.iconAnchor), JSON.stringify(expected.iconAnchor));
      assert.equal(actual.className, expected.className);
    }
  }
});

test('shared route numbers use the identical left-top shoulder badge without hiding the category icon', () => {
  const index = loadIndexRenderer();
  const shared = loadSharedRenderer();
  const colors = Array.from(index.PIN_COLORS, (item) => item.hex);
  const icons = Array.from(index.PIN_ICONS, (item) => item.id);

  colors.forEach((pinColor, colorIndex) => {
    icons.forEach((icon, iconIndex) => {
      const routeColor = colors[(colorIndex + iconIndex + 1) % colors.length];
      [1, 12, 123].forEach((number) => {
        const expected = index.createNumberedPinIcon(pinColor, number, false, false, icon, routeColor);
        const actual = shared.createSharedNumberedPinIcon(pinColor, number, false, false, icon, routeColor);
        assert.equal(actual.html, expected.html, `${pinColor} ${icon} route ${routeColor} #${number}`);
        assert.match(actual.html, /class="route-number-badge map-route-number-badge"/);
        assert.match(actual.html, /class="pin-category-glyph"/);
        assert.doesNotMatch(actual.html, /<text\b/);
        assert.equal(JSON.stringify(actual.iconSize), '[43,38]');
        assert.equal(JSON.stringify(actual.iconAnchor), '[30,37]');
        assert.equal(actual.className, 'pin-marker-root');
      });
    });
  });
});

test('shared marker geometry and theme-independent glyph styling copy the index rules', () => {
  const indexCss = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
  const sharedCss = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
  const selectors = [
    '.pin-marker-root', '.pin-icon-visual', '.pin-marker-svg',
    '.pin-icon-visual .map-pin-svg', '.pin-icon-visual svg',
    '.route-number-badge', '.map-route-number-badge', '.list-thumb', '.list-thumb-icon'
  ];
  selectors.forEach((selector) => {
    const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(`${escaped}\\s*\\{([^}]*)\\}`);
    const expected = indexCss.match(pattern);
    const actual = sharedCss.match(pattern);
    assert.ok(expected, `Missing index rule ${selector}`);
    assert.ok(actual, `Missing shared rule ${selector}`);
    assert.equal(actual[1].replace(/\s+/g, ' ').trim(), expected[1].replace(/\s+/g, ' ').trim(), selector);
  });
});

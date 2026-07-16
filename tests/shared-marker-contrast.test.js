const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');
const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

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

function indexConstantArray(name) {
  const match = indexHtml.match(new RegExp(`const ${name} = \\[([\\s\\S]*?)\\n\\s*\\];`));
  assert.ok(match, `Expected ${name}`);
  return vm.runInNewContext(`[${match[1]}]`);
}

function luminance(hex) {
  const normalized = hex.toLowerCase().slice(1);
  const channels = [0, 2, 4].map((offset) => {
    const value = Number.parseInt(normalized.slice(offset, offset + 2), 16) / 255;
    return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
  });
  return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2];
}

function contrast(first, second) {
  const lighter = Math.max(luminance(first), luminance(second));
  const darker = Math.min(luminance(first), luminance(second));
  return (lighter + 0.05) / (darker + 0.05);
}

function attribute(markup, element, name) {
  const elementMatch = markup.match(new RegExp(`<${element}\\b[^>]*>`));
  assert.ok(elementMatch, `Expected <${element}> in ${markup}`);
  const attributeMatch = elementMatch[0].match(new RegExp(`\\b${name}="([^"]+)"`));
  assert.ok(attributeMatch, `Expected ${name} on <${element}>`);
  return attributeMatch[1].toLowerCase();
}

const context = {
  SHARED_DEFAULT_COLOR: '#e53935',
  SHARED_DEFAULT_ROUTE_COLOR: '#1e88e5',
  SHARED_SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
  SHARED_SVG_XMLNS: 'http://www.w3.org/2000/svg',
  SHARED_PIN_ICONS: [
    { id: 'default', label: '標準' },
    { id: 'photo', label: '写真' },
    { id: 'food', label: '食事' },
    { id: 'hotel', label: '宿' },
    { id: 'nature', label: '自然' },
    { id: 'shop', label: '店' },
    { id: 'transit', label: '交通' },
    { id: 'warning', label: '注意' }
  ],
  escHtml: (value) => String(value).replace(/[&<>"']/g, (character) => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[character]),
  L: { divIcon: (options) => options }
};

vm.runInNewContext([
  functionSource('safeColor'),
  functionSource('safeRouteColor'),
  functionSource('hexToRgb'),
  functionSource('relativeLuminance'),
  functionSource('contrastRatio'),
  functionSource('getReadableTextColor'),
  functionSource('getSharedContrastForeground'),
  functionSource('normalizeSharedIcon'),
  functionSource('sharedPinIconGlyph'),
  functionSource('createPinMarkerSvg'),
  functionSource('createRouteNumberBadge'),
  functionSource('sharedPinListIconMarkup'),
  functionSource('createPinIcon'),
  functionSource('createSharedNumberedPinIcon')
].join('\n'), context);

const sharedPalette = indexConstantArray('PIN_COLORS').map((color) => color.hex.toLowerCase());
const contrastFixtures = Array.from(new Set([
  ...sharedPalette,
  '#ffeb3b', // bright yellow
  '#8bc34a', // bright green
  '#f7f7f2', // near white
  '#0d2a5b', // dark blue
  '#101211' // near black
]));

test('shared badge foreground makes the same best black-or-white contrast choice as index', () => {
  for (const fill of contrastFixtures) {
    const foreground = context.getSharedContrastForeground(fill);
    const opposite = foreground.toLowerCase() === '#000000' ? '#ffffff' : '#000000';
    assert.ok(
      contrast(foreground, fill) >= contrast(opposite, fill),
      `${foreground} is not the best foreground for ${fill}`
    );
  }
});

test('every normal shared marker preserves the canonical white category glyph', () => {
  for (const fill of ['#ffeb3b', '#e53935', '#0d2a5b']) {
    for (const icon of context.SHARED_PIN_ICONS) {
      const marker = context.createPinIcon(fill, false, icon.id === 'photo', icon.id).html;
      assert.equal(attribute(marker, 'g', 'color'), '#fff', `${icon.id} marker on ${fill}`);
      assert.equal(attribute(marker, 'g', 'opacity'), '1');
      assert.match(context.sharedPinIconGlyph(icon.id), /currentColor/, `${icon.id} glyph must inherit the computed foreground`);
    }
  }
});

test('numbered shared markers use a shoulder badge for one two and three digit labels', () => {
  for (const fill of ['#ffeb3b', '#e53935', '#0d2a5b']) {
    const expected = context.getSharedContrastForeground(fill);
    for (const number of [7, 42, 123]) {
      const marker = context.createSharedNumberedPinIcon('#e53935', number, false, false, 'warning', fill).html;
      assert.match(marker, new RegExp(`--route-badge-foreground: ${expected}`, 'i'));
      assert.match(marker, new RegExp(`>${number}<\\/span>`));
      assert.match(marker, /class="pin-category-glyph"/);
      assert.doesNotMatch(marker, /<text\b/);
    }
  }
});

test('list map and route membership follow the canonical non-destructive marker hierarchy', () => {
  const routePinRenderer = functionSource('buildSharedRoutePinList');
  assert.doesNotMatch(routePinRenderer, /order\.style\.(?:color|backgroundColor)/);

  for (const fill of contrastFixtures) {
    const listMarker = context.sharedPinListIconMarkup({ color: fill, icon: 'warning' });
    const mapMarker = context.createPinIcon(fill, false, false, 'warning').html;
    const numberedMarker = context.createSharedNumberedPinIcon(fill, 12, false, false, 'warning', fill).html;
    assert.equal(attribute(listMarker, 'svg', 'viewBox'), '0 0 34 48');
    assert.equal(attribute(mapMarker, 'svg', 'viewBox'), '0 0 34 48');
    assert.equal(attribute(numberedMarker, 'svg', 'viewBox'), '0 0 34 48');
    assert.match(numberedMarker, /map-route-number-badge/);
  }
});

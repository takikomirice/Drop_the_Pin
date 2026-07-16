const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

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

function constantArray(name) {
  const match = indexHtml.match(new RegExp(`const ${name} = \\[([\\s\\S]*?)\\n\\s*\\];`));
  assert.ok(match, `Expected ${name}`);
  return vm.runInNewContext(`[${match[1]}]`);
}

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

test('pin category and color catalogs keep fixed display and CSV names', () => {
  assert.deepEqual(
    JSON.parse(JSON.stringify(constantArray('PIN_COLORS'))),
    [
      { hex: '#e53935', csvName: 'red', label: '赤' },
      { hex: '#e91e63', csvName: 'pink', label: 'ピンク' },
      { hex: '#9c27b0', csvName: 'purple', label: '紫' },
      { hex: '#3f51b5', csvName: 'indigo', label: '藍' },
      { hex: '#2196f3', csvName: 'blue', label: '青' },
      { hex: '#00bcd4', csvName: 'cyan', label: '水色' },
      { hex: '#009688', csvName: 'teal', label: 'ティール' },
      { hex: '#4caf50', csvName: 'green', label: '緑' },
      { hex: '#8bc34a', csvName: 'lime', label: '黄緑' },
      { hex: '#ffeb3b', csvName: 'yellow', label: '黄' },
      { hex: '#ff9800', csvName: 'orange', label: '橙' },
      { hex: '#ff5722', csvName: 'deep-orange', label: '朱' },
      { hex: '#795548', csvName: 'brown', label: '茶' },
      { hex: '#607d8b', csvName: 'gray', label: 'グレー' },
      { hex: '#212121', csvName: 'black', label: '黒' }
    ]
  );
  assert.deepEqual(
    JSON.parse(JSON.stringify(constantArray('PIN_ICONS'))),
    [
      { id: 'default', label: '標準' }, { id: 'photo', label: '写真' },
      { id: 'food', label: '食事' }, { id: 'hotel', label: '宿' },
      { id: 'nature', label: '自然' }, { id: 'shop', label: '店' },
      { id: 'transit', label: '交通' }, { id: 'warning', label: '注意' }
    ]
  );
});

test('one shared inline SVG renders the teardrop and category glyph', () => {
  const render = vm.runInNewContext(`(${functionSource('createPinMarkerSvg')})`, {
    safeColor: (value) => value,
    pinIconGlyph: (icon) => `<path data-icon="${icon}"/>`,
    escHtml: escapeHtml
  });
  const svg = render('#2196f3', 'photo', false, 'list-thumb-icon');

  assert.match(svg, /class="pin-marker-svg list-thumb-icon"/);
  assert.match(svg, /width="26" height="36"/);
  assert.match(svg, /viewBox="0 0 34 48"/);
  assert.match(svg, /class="pin-drop"[^>]*fill="#2196f3"/);
  assert.match(svg, /class="pin-category-glyph"/);
  assert.match(svg, /data-icon="photo"/);
  assert.doesNotMatch(svg, /rgba\(255,255,255,0\.42\)/);
  assert.doesNotMatch(svg, /<img|\shref=/);
});

test('map route numbers use a round route-colored shoulder badge and preserve the pin category', () => {
  const badgeContext = {
    escHtml: escapeHtml,
    safeColor: (value) => value,
    getReadableTextColor: (value) => value === '#2196f3' ? '#000000' : '#FFFFFF'
  };
  const badge = vm.runInNewContext(`(${functionSource('createRouteNumberBadge')})`, badgeContext);
  assert.equal(badge(null, 'map-route-number-badge'), '');
  assert.match(badge(1, 'map-route-number-badge', '#2196f3'), /class="route-number-badge map-route-number-badge"[^>]*--route-badge-color:\s*#2196f3;\s*--route-badge-foreground:\s*#000000[^>]*>1<\/span>/);
  assert.match(badge(12, 'map-route-number-badge', '#e53935'), /--route-badge-color:\s*#e53935;\s*--route-badge-foreground:\s*#FFFFFF[^>]*>12<\/span>/);

  const createPin = functionSource('createPinIcon');
  assert.match(createPin, /createPinMarkerSvg/);
  assert.match(createPin, /createRouteNumberBadge/);
  assert.match(createPin, /createRouteNumberBadge\(routeNumber,[\s\S]*routeColor\)/);
  assert.match(createPin, /iconSize:\s*\[43,\s*38\]/);
  assert.match(createPin, /iconAnchor:\s*\[30,\s*37\]/);

  const display = functionSource('createPinIconForDisplay');
  assert.match(display, /createNumberedPinIcon\(pin\.color,[\s\S]*pin\.icon,[\s\S]*routeNumberDisplay\.color/);
  assert.doesNotMatch(functionSource('createNumberedPinIcon'), /<text/);
});

test('side-panel rows reuse the same pin SVG, route color, and number-slot alignment', () => {
  assert.match(functionSource('createListPinHeadSvg'), /createPinMarkerSvg/);

  const build = functionSource('buildListItem');
  assert.match(build, /getRouteNumberDisplayForPin\(pin\.id\)/);
  assert.match(build, /createRouteNumberBadge\(routeNumberDisplay\.number,[\s\S]*routeNumberDisplay\.color\)/);
  assert.match(build, /list-route-number-spacer/);
  assert.ok(
    build.indexOf('${routeBadgeHtml}') < build.indexOf('${renderListThumb(pin)}'),
    'route number must precede the category pin in reading order'
  );
});

test('round route badges and compact pins fit one and two digit labels on desktop and mobile', () => {
  assert.match(css, /\.route-number-badge\s*\{[\s\S]*?width:\s*18px[\s\S]*?height:\s*18px[\s\S]*?border-radius:\s*50%/);
  assert.match(css, /\.route-number-badge\s*\{[\s\S]*?background:\s*var\(--route-badge-color,[\s\S]*?color:\s*var\(--route-badge-foreground/);
  assert.match(css, /\.map-route-number-badge\s*\{[\s\S]*?position:\s*absolute[\s\S]*?left:\s*0[\s\S]*?top:\s*2px/);
  assert.match(css, /\.pin-icon-visual\s*\{[\s\S]*?width:\s*43px[\s\S]*?height:\s*38px[\s\S]*?transform-origin:\s*30px bottom/);
  assert.match(css, /\.pin-icon-visual \.map-pin-svg\s*\{[\s\S]*?left:\s*17px[\s\S]*?width:\s*26px[\s\S]*?height:\s*36px/);
  assert.match(css, /#side-panel \.pin-list-row\s*\{[\s\S]*?min-height:\s*62px[\s\S]*?margin-bottom:\s*5px/);
  assert.match(css, /\.list-route-number-spacer\s*\{[\s\S]*?width:\s*18px/);
  assert.match(css, /\.list-thumb\s*\{[\s\S]*?width:\s*30px[\s\S]*?height:\s*36px[\s\S]*?overflow:\s*visible/);
  assert.match(css, /\.list-thumb-icon\s*\{[\s\S]*?width:\s*26px[\s\S]*?height:\s*36px/);
  assert.match(css, /body\.narrow-view #side-panel \.pin-list-row\s*\{[\s\S]*?min-height:\s*66px[\s\S]*?margin-bottom:\s*5px/);
});

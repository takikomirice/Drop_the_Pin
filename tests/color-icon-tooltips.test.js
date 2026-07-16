const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

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

test('color palettes use PIN_COLORS labels for title and aria-label', () => {
  const body = functionBody('renderColorPaletteButtons');
  assert.match(body, /PIN_COLORS\.forEach/);
  assert.match(body, /button\.title = color\.label/);
  assert.match(body, /button\.setAttribute\('aria-label', color\.label\)/);
  assert.match(indexHtml, /paletteSelect\('upload-color-palette'/);
  assert.match(indexHtml, /paletteSelect\('edit-color-palette'/);
  assert.match(indexHtml, /paletteSelect\('route-edit-color-palette'/);
});

test('single-pin icon pickers and icon filters use PIN_ICONS labels for title and aria-label', () => {
  const pickerBody = functionBody('renderIconPickerButtons');
  assert.match(pickerBody, /PIN_ICONS\.forEach/);
  assert.match(pickerBody, /button\.title = icon\.label/);
  assert.match(pickerBody, /button\.setAttribute\('aria-label', icon\.label\)/);

  const filterBody = functionBody('renderIconFilterUI');
  assert.match(filterBody, /btn\.title = icon\.label/);
  assert.match(filterBody, /btn\.setAttribute\('aria-label', icon\.label\)/);
  assert.match(filterBody, /btn\.innerHTML = createInlineIconSvg\(icon\.id\)/);
  assert.doesNotMatch(filterBody, /escHtml\(icon\.label\)/);
});

test('map color filters use compact teardrop pins with PIN_COLORS tooltips', () => {
  const body = functionBody('renderColorFilterUI');
  assert.match(body, /btn\.title = c\.label/);
  assert.match(body, /btn\.setAttribute\('aria-label', c\.label\)/);
  assert.match(body, /btn\.innerHTML = createPinMarkerSvg\(c\.hex, 'default', false, 'color-filter-pin-svg'\)/);
  assert.doesNotMatch(body, /btn\.style\.background/);
});

test('share color filters reuse the same compact teardrop pins and labels', () => {
  const body = functionBody('renderShareFilterUi');
  assert.match(body, /btn\.title = color\.label/);
  assert.match(body, /btn\.setAttribute\('aria-label', color\.label\)/);
  assert.match(body, /btn\.innerHTML = createPinMarkerSvg\(color\.hex, 'default', false, 'color-filter-pin-svg'\)/);
  assert.doesNotMatch(body, /btn\.style\.background/);
});

test('bulk metadata icon options and select title use labels without a duplicate mapping', () => {
  const optionBody = functionBody('renderBulkMetadataIconOptions');
  assert.match(optionBody, /unchanged\.textContent = '変更しない'/);
  assert.match(optionBody, /unchanged\.title = '変更しない'/);
  assert.match(optionBody, /unchanged\.setAttribute\('aria-label', '変更しない'\)/);
  assert.match(optionBody, /PIN_ICONS\.forEach/);
  assert.match(optionBody, /option\.textContent = icon\.label/);
  assert.match(optionBody, /option\.title = icon\.label/);
  assert.match(optionBody, /option\.setAttribute\('aria-label', icon\.label\)/);

  const controlsBody = functionBody('updateBulkMetadataControls');
  assert.match(controlsBody, /iconSelect\.title = bulkMetadata\.icon \? getIconLabel\(bulkMetadata\.icon\) : '変更しない'/);
  assert.equal((indexHtml.match(/function getIconLabel\(/g) || []).length, 1);
  assert.equal(indexHtml.includes('const BULK_ICON_LABELS'), false);
});

test('bulk icon selection changes update the select title through existing controls', () => {
  const eventRegion = indexHtml.slice(indexHtml.indexOf("document.getElementById('bulk-icon-select')"));
  assert.match(eventRegion, /addEventListener\('change', updateBulkMetadataControls\)/);
  const openBody = functionBody('openBulkMetadataOverlay');
  assert.match(openBody, /renderBulkMetadataIconOptions\(\)/);
  assert.match(openBody, /updateBulkMetadataControls\(\)/);
});

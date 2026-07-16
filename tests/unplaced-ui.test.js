const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

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

test('map unplaced badge and dedicated dialog are absent with no dedicated dispatcher path', () => {
  for (const removed of [
    'id="unplaced-badge"', 'id="unplaced-overlay"', 'id="unplaced-list"',
    'id="unplaced-close"', 'function updateUnplacedBadge(',
    "getElementById('unplaced-badge')", "getElementById('unplaced-close')",
    "openOverlay('unplaced-overlay')", "id === 'unplaced-overlay'"
  ]) {
    assert.equal(indexHtml.includes(removed), false, `${removed} must be removed`);
  }
  assert.doesNotMatch(indexHtml, /#unplaced-badge\b/);
  assert.doesNotMatch(functionSource('dispatchEscape'), /unplaced-overlay/);
  assert.doesNotMatch(functionSource('dismissOverlayById'), /unplaced-overlay/);
});

test('desktop and narrow side panels retain unplaced pins and their normal edit operations', () => {
  assert.match(indexHtml, /id="pin-tab-unplaced"[^>]*aria-controls="side-unplaced"/);
  assert.match(indexHtml, /id="side-unplaced"[^>]*role="tabpanel"/);
  assert.match(indexHtml, /id="side-unplaced-count"/);
  assert.match(indexHtml, /id="mobile-pin-list"/);

  const panel = functionSource('renderSidePanel');
  assert.match(panel, /entry\.pin\.lat == null \|\| entry\.pin\.lng == null/);
  assert.match(panel, /unplacedContainer\.appendChild\(buildListItem/);
  assert.match(panel, /mobilePinList\.appendChild\(buildListItem/);
  assert.match(panel, /onMenu:\s*canEdit\(\) \? \(\) => openPinMenu\(pin\) : null/);

  for (const id of ['pin-menu-edit', 'pin-menu-place', 'pin-menu-unplace', 'pin-menu-delete']) {
    assert.match(indexHtml, new RegExp(`getElementById\\('${id}'\\)\\.addEventListener\\('click'`));
  }
  assert.match(functionSource('openPinMenu'), /isUnplaced[\s\S]*pin-menu-place/);
  assert.match(functionSource('beginPlacement'), /state\.placement/);
});

test('unplaced state, GPS-less saving, and unplace mutation paths remain intact', () => {
  const save = functionSource('saveNewPin');
  assert.match(save, /lat:\s*coords && coords\.lat != null \? coords\.lat : null/);
  assert.match(save, /lng:\s*coords && coords\.lng != null \? coords\.lng : null/);
  assert.match(indexHtml, /positionMode === 'unplaced'/);
  assert.match(indexHtml, /setUploadPositionMode\('unplaced'\)/);
  assert.match(indexHtml, /withGAS\('unplacePin',\s*withEditToken/);
  assert.match(indexHtml, /withGAS\('movePin',\s*withEditToken/);
  assert.match(indexHtml, /withGAS\('deletePin',\s*withEditToken/);
});

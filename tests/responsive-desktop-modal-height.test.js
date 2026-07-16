const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const css = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];

function blockAfter(source, marker) {
  const markerIndex = source.indexOf(marker);
  assert.notEqual(markerIndex, -1, `missing ${marker}`);
  const openIndex = source.indexOf('{', markerIndex + marker.length);
  assert.notEqual(openIndex, -1, `missing block for ${marker}`);
  let depth = 0;
  for (let index = openIndex; index < source.length; index += 1) {
    if (source[index] === '{') depth += 1;
    if (source[index] === '}') depth -= 1;
    if (depth === 0) return source.slice(openIndex + 1, index);
  }
  assert.fail(`unterminated block for ${marker}`);
}

function pixelToken(source, name) {
  const match = source.match(new RegExp(`${name}:\\s*(\\d+)px`));
  assert.ok(match, `missing ${name} pixel token`);
  return Number(match[1]);
}

test('desktop generic and workbench modals use the dynamic available height', () => {
  const rootBlock = blockAfter(css, ':root');
  const marker = '/* Desktop modal available-height contract. */';
  const markerIndex = css.indexOf(marker);
  assert.notEqual(markerIndex, -1, 'missing desktop modal height contract');
  const desktop = blockAfter(css.slice(markerIndex), '@media (min-width: 900px)');
  const genericSheet = blockAfter(desktop, '.sheet-body');
  const workbenchSheet = blockAfter(desktop, '.workbench-sheet');

  assert.match(rootBlock, /--app-modal-viewport-inset:\s*28px/);
  [genericSheet, workbenchSheet].forEach((rule) => {
    assert.match(rule, /max-height:\s*min\(/);
    assert.match(rule, /calc\(var\(--dialog-viewport-height,\s*100dvh\) - var\(--app-modal-viewport-inset\)\)/);
  });
  assert.match(genericSheet, /860px/);
  assert.match(workbenchSheet, /780px/);
});

test('modal caps fit all short and normal desktop heights', () => {
  const rootBlock = blockAfter(css, ':root');
  const inset = pixelToken(rootBlock, '--app-modal-viewport-inset');
  const overlayPadding = pixelToken(blockAfter(css, '.sheet-overlay'), 'padding');
  const genericCap = 860;
  const workbenchCap = 780;

  assert.equal(inset, overlayPadding * 2);
  [520, 600, 720, 768, 900].forEach((viewportHeight) => {
    const availableHeight = viewportHeight - inset;
    const genericHeight = Math.min(availableHeight, genericCap);
    const workbenchHeight = Math.min(availableHeight, workbenchCap);

    assert.ok(genericHeight > 0 && genericHeight <= availableHeight);
    assert.ok(workbenchHeight > 0 && workbenchHeight <= availableHeight);
    assert.ok(viewportHeight - genericHeight >= inset);
    assert.ok(viewportHeight - workbenchHeight >= inset);
  });
});

test('generic and workbench content retain the correct scroll owner', () => {
  const overlaySource = css.slice(css.indexOf('.sheet-overlay {'));
  const genericSheet = blockAfter(overlaySource, '.sheet-body');
  const workbenchSheet = blockAfter(css, '.workbench-sheet');
  const workbenchContent = blockAfter(css, '.workbench-content');

  assert.match(genericSheet, /overflow-y:\s*auto/);
  assert.match(workbenchSheet, /overflow:\s*hidden/);
  assert.match(workbenchSheet, /display:\s*flex/);
  assert.match(workbenchSheet, /flex-direction:\s*column/);
  assert.match(workbenchContent, /min-height:\s*0/);
  assert.match(workbenchContent, /flex:\s*1 1 auto/);
  assert.match(workbenchContent, /overflow-y:\s*auto/);
});

test('desktop modal constraint preserves responsive thresholds while mobile detail follows the visible viewport', () => {
  const mobileSource = css.slice(css.indexOf('/* Phase 9.5-F UI-7'));
  const mobile = blockAfter(mobileSource, '@media (max-width: 640px)');
  const pinDetail = blockAfter(mobile, 'body.narrow-view #pin-detail-overlay .sheet-body');

  assert.match(pinDetail, /height:\s*min\(75dvh,\s*calc\(var\(--dialog-viewport-height,\s*100dvh\)/);
  assert.match(pinDetail, /max-height:\s*min\(75dvh,\s*calc\(var\(--dialog-viewport-height,\s*100dvh\)/);
  assert.match(pinDetail, /env\(safe-area-inset-bottom\)/);
  assert.match(indexHtml, /const narrowViewMedia = window\.matchMedia\('\(max-width: 640px\)'\);/);
  assert.match(
    indexHtml,
    /let isPanelVisible = narrowViewMedia\.matches \|\| window\.matchMedia\('\(min-width: 900px\)'\)\.matches;/
  );
});

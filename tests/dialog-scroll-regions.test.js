const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
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

test('dialog frames contain scrolling and generic sheets own one bounded scroll region', () => {
  const overlay = blockAfter(css, '.sheet-overlay');
  const genericSheet = blockAfter(css.slice(css.indexOf('.sheet-overlay {')), '.sheet-body');

  assert.match(overlay, /overflow:\s*hidden/);
  assert.match(overlay, /overscroll-behavior:\s*contain/);
  assert.match(genericSheet, /min-height:\s*0/);
  assert.match(genericSheet, /overflow-y:\s*auto/);
  assert.match(genericSheet, /overscroll-behavior:\s*contain/);
});

test('generic headers and primary actions remain visible without introducing another scroller', () => {
  assert.match(css, /\.sheet-body\s*>\s*\.sheet-title[^\{]*\{[^}]*position:\s*sticky[^}]*top:\s*0/);
  assert.match(
    css,
    /\.sheet-body\s*>\s*:is\(\.inline-actions,\s*\.menu-actions\):last-child[^\{]*\{[^}]*position:\s*sticky[^}]*bottom:\s*0/
  );
  assert.match(css, /\.single-pin-header[^\{]*\{[^}]*position:\s*sticky[^}]*top:\s*0/);
  assert.match(css, /\.single-pin-footer[^\{]*\{[^}]*position:\s*sticky[^}]*bottom:\s*0/);
  assert.match(css, /\.dialog-scroll-frame\s*\{[^}]*display:\s*flex[^}]*overflow:\s*hidden/);
  assert.match(css, /\.dialog-scroll-content\s*\{[^}]*min-height:\s*0[^}]*overflow-y:\s*auto[^}]*overscroll-behavior:\s*contain/);
  assert.match(indexHtml, /function setupDialogScrollRegions\(\)[\s\S]*?dialog-scroll-frame[\s\S]*?dialog-scroll-content[\s\S]*?dialog-scroll-footer/);
  assert.match(indexHtml, /setupDialogScrollRegions\(\);[\s\S]*?setupDialogViewportTracking\(\);/);
});

test('workbenches keep fixed headers and one contained body scroller', () => {
  const sheet = blockAfter(css, '.workbench-sheet');
  const header = blockAfter(css, '.workbench-header');
  const content = blockAfter(css, '.workbench-content');

  assert.match(sheet, /min-height:\s*0/);
  assert.match(sheet, /overflow:\s*hidden/);
  assert.match(header, /flex:\s*0 0 auto/);
  assert.match(content, /min-height:\s*0/);
  assert.match(content, /overflow-y:\s*auto/);
  assert.match(content, /overscroll-behavior:\s*contain/);
  assert.match(css, /\.settings-list\s*>\s*\.inline-actions:last-child[^\{]*\{[^}]*position:\s*sticky[^}]*bottom:\s*0/);
});

test('route and import previews keep their shell hidden and panes independently contained', () => {
  ['.route-preview-sheet', '.drive-photo-import-sheet', '.import-preview-sheet', '.input-presets-sheet']
    .forEach((selector) => {
      const rule = blockAfter(css, selector);
      assert.match(rule, /min-height:\s*0/, selector);
      assert.match(rule, /overflow:\s*hidden/, selector);
      assert.match(rule, /display:\s*flex/, selector);
      assert.match(rule, /flex-direction:\s*column/, selector);
    });

  [
    '.route-preview-editor', '.route-preview-summary', '.import-preview-list-pane',
    '.import-preview-editor', '.import-preview-photo-pane', '#drive-photo-import-folder-list',
    '#drive-photo-import-photo-list', '.input-preset-list'
  ].forEach((selector) => {
    const rule = blockAfter(css, selector);
    assert.match(rule, /min-height:\s*0/, selector);
    assert.match(rule, /overflow-y:\s*auto/, selector);
    assert.match(rule, /overscroll-behavior:\s*contain/, selector);
  });
  assert.match(css, /\.input-presets-editor-pane\s*\{[^}]*min-height:\s*0[^}]*overflow:\s*hidden/);
  assert.match(
    css,
    /\.input-preset-editor-content\s*\{[^}]*min-height:\s*0[^}]*overflow-y:\s*auto[^}]*overscroll-behavior:\s*contain/
  );
});

test('mobile route and preset layouts remove the outer double scroll', () => {
  const presetSource = css.slice(css.indexOf('.input-presets-sheet {'));
  const mobile = blockAfter(presetSource, '@media not all and (min-width: 900px)');

  assert.doesNotMatch(mobile, /\.input-presets-sheet\s*,\s*\.route-preview-sheet\s*\{[^}]*overflow-y:\s*auto/);
  assert.match(mobile, /\.input-presets-sheet\s*,\s*\.route-preview-sheet\s*\{[^}]*overflow:\s*hidden/);
  assert.match(mobile, /\.route-preview-grid\s*\{[^}]*display:\s*grid[^}]*overflow:\s*hidden/);
  assert.match(mobile, /\.route-preview-editor\s*,\s*\.route-preview-summary\s*\{[^}]*overflow-y:\s*auto/);
  assert.match(mobile, /\.input-presets-workbench\s*\{[^}]*display:\s*grid[^}]*overflow:\s*hidden/);
  assert.match(mobile, /grid-template-rows:\s*minmax\(0,\s*1fr\)\s+minmax\(0,\s*1fr\)/);
});

test('low-height Drive and mobile import layouts keep footer actions outside scrolling panes', () => {
  const mobileImport = blockAfter(css.slice(css.indexOf('.import-preview-sheet {')), '@media (max-width: 900px)');
  assert.match(mobileImport, /\.import-preview-sheet\[data-import-variant="photo"\]\s+\.import-preview-layout\s*\{[^}]*grid-template-rows:\s*repeat\(3,\s*minmax\(0,\s*1fr\)\)/);
  assert.match(mobileImport, /\.import-preview-sheet\[data-import-variant="csv"\][\s\S]*?\.import-preview-layout\s*\{[^}]*grid-template-rows:\s*repeat\(2,\s*minmax\(0,\s*1fr\)\)/);
  assert.match(css, /@media \(max-width:\s*900px\) and \(max-height:\s*600px\)[\s\S]*?\.drive-photo-import-workbench\s*\{[^}]*grid-template-rows:\s*minmax\(0,\s*1fr\)/);
  assert.match(css, /\.drive-photo-import-footer\s*\{[^}]*flex:\s*0 0 auto/);
  assert.match(css, /\.import-preview-actions\s*\{[^}]*flex:\s*0 0 auto/);
  assert.match(css, /\.route-preview-actions\s*\{[^}]*flex:\s*0 0 auto/);
});

test('scroll-region changes preserve the modal interaction and accessibility contracts', () => {
  assert.match(indexHtml, /function openOverlay\(id\)/);
  assert.match(indexHtml, /function trapOverlayFocus\(event\)/);
  assert.match(indexHtml, /function dismissOverlayById\(id/);
  assert.match(indexHtml, /inert/);
  assert.match(indexHtml, /overscroll-behavior:\s*contain/);
});

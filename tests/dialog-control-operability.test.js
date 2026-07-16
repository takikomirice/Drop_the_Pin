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
  let quote = null;
  let escaped = false;
  for (let index = open; index < indexHtml.length; index += 1) {
    const character = indexHtml[index];
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
    if (character === '}' && --depth === 0) return indexHtml.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${name}`);
}

test('photo selection is a native button with click and drop support', () => {
  const overlay = indexHtml.match(/<div id="upload-overlay"[\s\S]*?<div id="drive-photo-import-overlay"/);
  assert.ok(overlay);
  assert.match(overlay[0], /<button id="file-drop" class="file-drop" type="button">/);
  assert.doesNotMatch(overlay[0], /<div id="file-drop"/);

  const setup = functionBody('setupUploadPhotoPicker');
  assert.match(setup, /fileDrop\.addEventListener\('click'/);
  assert.match(setup, /fileDrop\.addEventListener\('dragover'/);
  assert.match(setup, /fileDrop\.addEventListener\('drop'/);
  assert.match(setup, /event\.dataTransfer\.files/);
  assert.match(setup, /handleUploadPhotoSelected/);
  assert.match(setup, /fileInput\.addEventListener\('change', handleUploadPhotoSelected\)/);
});

test('photo selection exposes its unavailable state with native disabled controls', () => {
  const refresh = functionBody('refreshUploadSubmitState');
  assert.match(refresh, /fileDrop\.disabled = photoSelectionDisabled/);
  assert.match(refresh, /fileInput\.disabled = photoSelectionDisabled/);
  assert.match(refresh, /state\.upload\.converting/);
  assert.match(refresh, /state\.upload\.saving/);
});

test('color and icon buttons expose readable labels and synchronized pressed state', () => {
  const palette = functionBody('renderColorPaletteButtons');
  assert.match(palette, /button\.type = 'button'/);
  assert.match(palette, /button\.setAttribute\('aria-label', color\.label\)/);
  assert.match(palette, /button\.setAttribute\('aria-pressed', String\(isSelected\)\)/);
  assert.match(palette, /button\.style\.backgroundColor = color\.hex/);
  assert.match(palette, /button\.disabled = disabled/);

  const picker = functionBody('renderIconPickerButtons');
  assert.match(picker, /button\.type = 'button'/);
  assert.match(picker, /button\.setAttribute\('aria-label', icon\.label\)/);
  assert.match(picker, /button\.setAttribute\('aria-pressed', String\(isSelected\)\)/);
  assert.match(picker, /button\.disabled = disabled/);

  const presetControls = functionBody('updateInputPresetEditorControls');
  assert.match(presetControls, /querySelectorAll\('button'\)/);
  assert.match(presetControls, /button\.disabled = saving \|\| !visible/);
});

test('rerendered color and icon selections retain keyboard focus', () => {
  const restore = functionBody('restorePickerSelectionFocus');
  assert.match(restore, /!shouldRestoreFocus \|\| button\.isConnected/);
  assert.match(restore, /selector \+ '\[aria-pressed="true"\]'/);
  assert.match(restore, /selectedButton\.focus\(\)/);

  const palette = functionBody('renderColorPaletteButtons');
  assert.match(palette, /const shouldRestoreFocus = document\.activeElement === button/);
  assert.match(palette, /restorePickerSelectionFocus\(container, '\.color-swatch', button, shouldRestoreFocus\)/);

  const picker = functionBody('renderIconPickerButtons');
  assert.match(picker, /const shouldRestoreFocus = document\.activeElement === button/);
  assert.match(picker, /restorePickerSelectionFocus\(container, '\.icon-choice', button, shouldRestoreFocus\)/);
});

test('dialog swatches, icon choices, and compact photo controls meet 44px targets', () => {
  assert.match(indexHtml, /(?:^|\n)\s*\.color-swatch\s*\{[^}]*width:\s*44px;[^}]*height:\s*44px;[^}]*background-clip:\s*content-box;/s);
  assert.match(indexHtml, /\.single-pin-options \.color-swatch\s*\{[^}]*width:\s*44px;[^}]*height:\s*44px;[^}]*flex:\s*0 0 44px;/s);
  assert.match(indexHtml, /\.icon-choice, \.icon-filter-chip\s*\{[^}]*min-width:\s*44px;[^}]*min-height:\s*44px;/s);
  assert.match(indexHtml, /\.single-pin-options \.icon-choice\s*\{[^}]*min-width:\s*44px;[^}]*min-height:\s*44px;[^}]*flex:\s*0 0 44px;/s);
  assert.match(indexHtml, /\.single-pin-photo-field \.file-drop\s*\{[^}]*min-height:\s*44px;/s);
  assert.match(indexHtml, /\.single-pin-photo-actions \.ghost-btn\s*\{[^}]*min-height:\s*44px;/s);
  assert.match(indexHtml, /\.single-pin-date-control \.ghost-btn\s*\{[^}]*min-height:\s*44px;/s);
  assert.match(indexHtml, /\.single-pin-folder > summary\s*\{[^}]*min-height:\s*44px;/s);
  assert.match(indexHtml, /\.workbench-close\s*\{[^}]*min-width:\s*44px;[^}]*min-height:\s*44px;/s);
});

test('preset drag handle remains a labeled button with a 44px touch target', () => {
  const render = functionBody('renderInputPresetList');
  assert.match(render, /dragHandle\.type = 'button'/);
  assert.match(render, /dragHandle\.innerHTML = '[^']*action-icon-sort/);
  assert.match(render, /dragHandle\.setAttribute\('aria-label', '並べ替え'\)/);
  assert.match(render, /dragHandle\.disabled = controlsDisabled/);
  assert.match(render, /row\.appendChild\(main\);[\s\S]*row\.appendChild\(dragHandle\);/);
  assert.match(indexHtml, /\.input-preset-row\s*\{[^}]*grid-template-columns:\s*minmax\(0,\s*1fr\) 44px;/s);
  assert.match(indexHtml, /\.input-preset-drag-handle\s*\{[^}]*min-width:\s*44px;[^}]*min-height:\s*44px;/s);
  assert.match(indexHtml, /\.input-preset-actions button\s*\{[^}]*min-width:\s*44px;[^}]*min-height:\s*44px;/s);
  assert.match(indexHtml, /handle:\s*'\.input-preset-drag-handle'/);
  assert.match(render, /'プリセットを上へ移動'/);
  assert.match(render, /'プリセットを下へ移動'/);
});

test('preset rows use an unframed grip and compact single-line summary', () => {
  const render = functionBody('renderInputPresetList');
  assert.match(indexHtml, /\.input-preset-row\s*\{[^}]*gap:\s*8px;[^}]*padding:\s*8px 10px;/s);
  assert.match(indexHtml, /\.input-preset-drag-handle\s*\{[^}]*align-self:\s*center;[^}]*border:\s*0;[^}]*background:\s*transparent;/s);
  assert.match(indexHtml, /\.input-preset-main\s*\{[^}]*min-width:\s*0;[^}]*display:\s*grid;/s);
  assert.match(indexHtml, /\.input-preset-summary\s*\{[^}]*white-space:\s*nowrap;[^}]*overflow:\s*hidden;[^}]*text-overflow:\s*ellipsis;/s);
  assert.match(indexHtml, /\.input-preset-actions\s*\{[^}]*grid-column:\s*1 \/ -1;[^}]*margin-top:\s*4px;/s);
  assert.match(render, /const summaryText = getInputPresetSummary\(preset\)/);
  assert.match(render, /summaryElement\.textContent = summaryText/);
  assert.match(render, /summaryElement\.title = summaryText/);
});

test('preset action buttons keep 44px targets around a compact visual face', () => {
  assert.match(indexHtml, /\.input-preset-actions button\s*\{[^}]*min-height:\s*44px;[^}]*position:\s*relative;[^}]*border:\s*0;[^}]*background:\s*transparent;/s);
  assert.match(indexHtml, /\.input-preset-actions button::before\s*\{[^}]*inset:\s*5px 3px;[^}]*border:\s*1px solid var\(--border\);[^}]*background:\s*transparent;/s);
  assert.match(indexHtml, /\.input-preset-actions \.danger-btn::before\s*\{[^}]*background:\s*var\(--danger\);/s);
});

test('preset actions wrap without horizontal scrolling and retain compact 44px targets', () => {
  assert.match(indexHtml, /\.input-preset-actions\s*\{[^}]*flex-wrap:\s*wrap;[^}]*gap:\s*2px;[^}]*overflow-x:\s*hidden;/s);
  assert.match(indexHtml, /\.input-preset-actions button\s*\{[^}]*min-width:\s*44px;[^}]*padding:\s*0 6px;/s);
  assert.match(indexHtml, /\.input-preset-actions button::before\s*\{[^}]*inset:\s*5px 3px;/s);
  assert.match(indexHtml, /@media \(max-width:\s*520px\)[\s\S]*\.input-preset-actions button\s*\{[^}]*flex:\s*0 0 44px;/s);
  assert.match(indexHtml, /\.input-preset-list\s*\{[^}]*overflow-y:\s*auto;[^}]*scrollbar-gutter:\s*stable;[^}]*padding-right:\s*8px;/s);
  assert.match(indexHtml, /\.input-preset-editor \.form-group\[hidden\]\s*\{[^}]*display:\s*none;/s);
});

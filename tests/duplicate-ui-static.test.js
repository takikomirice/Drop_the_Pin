const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

function assertIncludes(source, text) {
  assert.ok(source.includes(text), `Expected source to include: ${text}`);
}

function assertNotIncludes(source, text) {
  assert.equal(source.includes(text), false, `Expected source not to include: ${text}`);
}

test('index pin menu exposes duplicate actions only in the normal app', () => {
  assertIncludes(indexHtml, 'id="pin-menu-duplicate-unplaced"');
  assertIncludes(indexHtml, '未配置で複製');
  assertIncludes(indexHtml, 'id="pin-menu-duplicate-same"');
  assertIncludes(indexHtml, '同じ場所に複製');
  assertIncludes(indexHtml, 'id="pin-menu-duplicate-point"');
  assertIncludes(indexHtml, '場所を指定して複製');

  assertNotIncludes(sharedHtml, '未配置で複製');
  assertNotIncludes(sharedHtml, '同じ場所に複製');
  assertNotIncludes(sharedHtml, '場所を指定して複製');
  assertNotIncludes(sharedHtml, 'duplicatePin');
});

test('index keeps duplicate target state separate from bulk selection', () => {
  assertIncludes(indexHtml, 'activePinId: null');
  assertIncludes(indexHtml, 'copiedPinSourceId: null');
  assertIncludes(indexHtml, 'duplicatePlacement: null');
  assertIncludes(indexHtml, 'state.selectedPinIds.size');
});

test('index has duplicate placement and shortcut guards for typing targets', () => {
  assertIncludes(indexHtml, 'function isTypingTarget(target)');
  assertIncludes(indexHtml, "target.tagName === 'INPUT'");
  assertIncludes(indexHtml, "target.tagName === 'TEXTAREA'");
  assertIncludes(indexHtml, "target.tagName === 'SELECT'");
  assertIncludes(indexHtml, 'target.isContentEditable');
  assertIncludes(indexHtml, 'function handleDuplicateShortcut(event)');
  const shortcutOverlayStart = indexHtml.indexOf('function isShortcutOverlayOpen()');
  const shortcutOverlayEnd = indexHtml.indexOf('function hasMultipleSelectedPins()', shortcutOverlayStart);
  assertIncludes(indexHtml.slice(shortcutOverlayStart, shortcutOverlayEnd), "'audio-editor-overlay'");
  assertIncludes(indexHtml, "event.key.toLowerCase() === 'd'");
  assertIncludes(indexHtml, "event.key.toLowerCase() === 'c'");
  assertIncludes(indexHtml, "event.key.toLowerCase() === 'v'");
  assertIncludes(indexHtml, 'if (isTypingTarget(event.target) || isTypingTarget(document.activeElement)) return;');
});

test('index calls duplicatePin for unplaced same point and paste flows', () => {
  assertIncludes(indexHtml, "withGAS('duplicatePin'");
  assertIncludes(indexHtml, 'payload = { sourcePinId: source.id, mode: mode }');
  assertIncludes(indexHtml, "duplicatePinFromSource(pin.id, 'unplaced'");
  assertIncludes(indexHtml, "duplicatePinFromSource(pin.id, 'same'");
  assertIncludes(indexHtml, "'point'");
  assertIncludes(indexHtml, "beginDuplicatePlacement(pin.id, 'menu')");
  assertIncludes(indexHtml, "beginDuplicatePlacement(pin.id, 'paste')");
  assertIncludes(indexHtml, '複製先の場所を地図上で選んでください');
  assertIncludes(indexHtml, '貼り付ける場所を地図上でクリックしてください');
  assertIncludes(indexHtml, '複製は1件ずつ行ってください');
  assertIncludes(indexHtml, 'コピー元のピンが見つかりません');
});

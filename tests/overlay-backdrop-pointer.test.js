const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function sourceFunction(source, name) {
  const index = source.indexOf(`function ${name}`);
  assert.notEqual(index, -1, `Expected function ${name} to exist`);
  const openIndex = source.indexOf('{', index);
  assert.notEqual(openIndex, -1, `Expected function ${name} to have a body`);
  let depth = 0;
  for (let cursor = openIndex; cursor < source.length; cursor += 1) {
    if (source[cursor] === '{') depth += 1;
    if (source[cursor] === '}') depth -= 1;
    if (depth === 0) return source.slice(index, cursor + 1);
  }
  assert.fail(`Could not parse function ${name}`);
}

function createPointerHarness(options) {
  const config = Object.assign({ dismissResult: undefined }, options);
  const listeners = {};
  const overlay = {
    addEventListener(type, listener) { listeners[type] = listener; },
    closest() { return null; }
  };
  let hitTarget = overlay;
  let dismissCount = 0;
  let preventCount = 0;
  let restoredFocusCount = 0;
  const focusedControl = { focus() { restoredFocusCount += 1; } };
  const context = {
    Math: Math,
    OVERLAY_BACKDROP_TAP_MAX_MOVEMENT_PX: 9,
    document: {
      activeElement: focusedControl,
      elementFromPoint() { return hitTarget; }
    },
    restoreSurfaceFocus(target) {
      if (target && typeof target.focus === 'function') target.focus();
    }
  };
  const source = [
    sourceFunction(indexHtml, 'isOverlayBackdropStartTarget'),
    sourceFunction(indexHtml, 'isOverlayBackdropEndTarget'),
    sourceFunction(indexHtml, 'setupOverlayBackdropDismissal')
  ].join('\n');
  vm.runInNewContext(`${source}\nthis.setup = setupOverlayBackdropDismissal;`, context);
  context.setup(overlay, function() {
    dismissCount += 1;
    return config.dismissResult;
  });

  function emit(type, event) {
    const nextEvent = Object.assign({
      pointerId: 1,
      clientX: 100,
      clientY: 100,
      target: overlay,
      preventDefault() { preventCount += 1; }
    }, event);
    listeners[type](nextEvent);
    return nextEvent;
  }

  return {
    overlay: overlay,
    emit: emit,
    setHitTarget(target) { hitTarget = target; },
    dismissCount() { return dismissCount; },
    preventCount() { return preventCount; },
    restoredFocusCount() { return restoredFocusCount; }
  };
}

function interactiveTarget(selector) {
  return {
    closest(query) {
      return query.includes(selector) ? this : null;
    }
  };
}

test('backdrop dismissal requires a same-background pointer tap', () => {
  const harness = createPointerHarness();
  harness.emit('pointerdown');
  harness.emit('pointerup');
  assert.equal(harness.dismissCount(), 1);
  assert.equal(harness.preventCount(), 1, 'backdrop pointerdown keeps focus in the dialog');
});

test('input text selection and sheet-body drags never dismiss the backdrop', () => {
  const inputHarness = createPointerHarness();
  inputHarness.emit('pointerdown', { target: interactiveTarget('input') });
  inputHarness.emit('pointermove', { target: inputHarness.overlay, clientX: 140 });
  inputHarness.emit('pointerup');
  assert.equal(inputHarness.dismissCount(), 0);
  assert.equal(inputHarness.preventCount(), 0);

  const sheetHarness = createPointerHarness();
  sheetHarness.emit('pointerdown', { target: interactiveTarget('.sheet-body') });
  sheetHarness.emit('pointermove', { target: sheetHarness.overlay, clientY: 140 });
  sheetHarness.emit('pointerup');
  assert.equal(sheetHarness.dismissCount(), 0);
  assert.equal(sheetHarness.preventCount(), 0);
});

test('a refused backdrop dismissal restores the focus held at pointerdown', () => {
  const harness = createPointerHarness({ dismissResult: false });
  harness.emit('pointerdown');
  harness.emit('pointerup');

  assert.equal(harness.dismissCount(), 1);
  assert.equal(harness.restoredFocusCount(), 1);
});

test('interactive controls never begin a backdrop-dismiss gesture', () => {
  ['input', 'textarea', 'select', 'button', 'a', '[contenteditable]', '.sheet-body'].forEach(function(selector) {
    const harness = createPointerHarness();
    harness.emit('pointerdown', { target: interactiveTarget(selector) });
    harness.emit('pointerup');
    assert.equal(harness.dismissCount(), 0, selector);
  });
});

test('movement, a different pointer, and pointercancel do not dismiss the backdrop', () => {
  const moved = createPointerHarness();
  moved.emit('pointerdown');
  moved.emit('pointermove', { clientX: 110 });
  moved.emit('pointerup', { clientX: 110 });
  assert.equal(moved.dismissCount(), 0);

  const differentPointer = createPointerHarness();
  differentPointer.emit('pointerdown');
  differentPointer.emit('pointerdown', { pointerId: 2 });
  differentPointer.emit('pointerup', { pointerId: 2 });
  assert.equal(differentPointer.dismissCount(), 0);
  differentPointer.emit('pointercancel');
  assert.equal(differentPointer.dismissCount(), 0);

  const cancelled = createPointerHarness();
  cancelled.emit('pointerdown');
  cancelled.emit('pointercancel');
  cancelled.emit('pointerup');
  assert.equal(cancelled.dismissCount(), 0);

  const capturedPointer = createPointerHarness();
  capturedPointer.emit('pointerdown');
  capturedPointer.setHitTarget(interactiveTarget('.sheet-body'));
  capturedPointer.emit('pointerup');
  assert.equal(capturedPointer.dismissCount(), 0);
});

test('backdrop dismissal delegates the active modal to the common dispatcher', () => {
  const source = sourceFunction(indexHtml, 'closeOverlayFromBackdrop');
  const body = source.slice(source.indexOf('{') + 1);
  assert.match(source, /getTopModalSheetOverlayRecord\(\)/);
  assert.match(source, /dismissOverlayById\(record\.id\)/);
  assert.doesNotMatch(body, /closeOverlay\(/);
  assert.doesNotMatch(body, /close[A-Z][A-Za-z]+\(/);
});

test('all dialogs track backdrop pointers while the dismissal allowlist remains unchanged', () => {
  assert.match(indexHtml, /\['help-overlay'\]\.concat\(MAIN_DISMISSIBLE_OVERLAY_IDS\)\.forEach\(function\(id\)/);
  assert.match(indexHtml, /const overlay = document\.getElementById\(id\);\s*if \(!overlay\) return;\s*setupOverlayBackdropDismissal\(overlay, function\(\) \{\s*return closeOverlayFromBackdrop\(id\);\s*}\);/);
  const idsStart = indexHtml.indexOf('const BACKDROP_DISMISS_OVERLAY_IDS =');
  const idsEnd = indexHtml.indexOf('];', idsStart);
  assert.notEqual(idsStart, -1);
  assert.notEqual(idsEnd, -1);
  assert.equal(indexHtml.slice(idsStart, idsEnd).includes("'edit-overlay'"), false);
  assert.equal(indexHtml.slice(idsStart, idsEnd).includes("'multi-photo-preparation-overlay'"), false);
  assert.equal(indexHtml.slice(idsStart, idsEnd).includes("'import-preview-overlay'"), false);
  assert.equal(indexHtml.slice(idsStart, idsEnd).includes("'track-import-preview-overlay'"), false);
  assert.equal(indexHtml.slice(idsStart, idsEnd).includes("'drive-photo-import-overlay'"), true);
  assert.equal(indexHtml.slice(idsStart, idsEnd).includes("'unplaced-overlay'"), false);
  assert.doesNotMatch(indexHtml, /setupOverlayBackdropDismissal\(document\.getElementById\('drive-photo-import-overlay'\)/);
  assert.match(indexHtml, /getElementById\('add-menu-close'\)[\s\S]*dismissOverlayById\('add-menu-overlay'\)/);
  assert.match(indexHtml, /const OVERLAY_BACKDROP_TAP_MAX_MOVEMENT_PX = (?:8|9|10);/);
});

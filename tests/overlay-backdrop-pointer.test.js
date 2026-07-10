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

function createPointerHarness() {
  const listeners = {};
  const overlay = {
    addEventListener(type, listener) { listeners[type] = listener; },
    closest() { return null; }
  };
  let hitTarget = overlay;
  let dismissCount = 0;
  const context = {
    Math: Math,
    OVERLAY_BACKDROP_TAP_MAX_MOVEMENT_PX: 9,
    document: { elementFromPoint() { return hitTarget; } }
  };
  const source = [
    sourceFunction(indexHtml, 'isOverlayBackdropStartTarget'),
    sourceFunction(indexHtml, 'isOverlayBackdropEndTarget'),
    sourceFunction(indexHtml, 'setupOverlayBackdropDismissal')
  ].join('\n');
  vm.runInNewContext(`${source}\nthis.setup = setupOverlayBackdropDismissal;`, context);
  context.setup(overlay, function() { dismissCount += 1; });

  function emit(type, event) {
    listeners[type](Object.assign({
      pointerId: 1,
      clientX: 100,
      clientY: 100,
      target: overlay
    }, event));
  }

  return {
    overlay: overlay,
    emit: emit,
    setHitTarget(target) { hitTarget = target; },
    dismissCount() { return dismissCount; }
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
});

test('input text selection and sheet-body drags never dismiss the backdrop', () => {
  const inputHarness = createPointerHarness();
  inputHarness.emit('pointerdown', { target: interactiveTarget('input') });
  inputHarness.emit('pointermove', { target: inputHarness.overlay, clientX: 140 });
  inputHarness.emit('pointerup');
  assert.equal(inputHarness.dismissCount(), 0);

  const sheetHarness = createPointerHarness();
  sheetHarness.emit('pointerdown', { target: interactiveTarget('.sheet-body') });
  sheetHarness.emit('pointermove', { target: sheetHarness.overlay, clientY: 140 });
  sheetHarness.emit('pointerup');
  assert.equal(sheetHarness.dismissCount(), 0);
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

test('backdrop dismissal preserves overlay-specific cleanup', () => {
  const calls = [];
  const context = {
    state: { bulkDeletePending: true, shownPinId: 'pin-1' },
    closeHelpPanel() { calls.push('help'); },
    closeShareDialog() { calls.push('share'); },
    closeShareQr() { calls.push('share-qr'); },
    closeOverlay(id) { calls.push(`close:${id}`); }
  };
  vm.runInNewContext(`${sourceFunction(indexHtml, 'closeOverlayFromBackdrop')}\nthis.close = closeOverlayFromBackdrop;`, context);

  context.close('help-overlay');
  context.close('share-overlay');
  context.close('share-qr-overlay');
  context.close('delete-overlay');
  context.close('pin-detail-overlay');

  assert.deepEqual(calls, [
    'help',
    'share',
    'share-qr',
    'close:delete-overlay',
    'close:pin-detail-overlay'
  ]);
  assert.equal(context.state.bulkDeletePending, false);
  assert.equal(context.state.shownPinId, null);
});

test('pointer dismissal is attached to closable overlays but never edit-overlay', () => {
  assert.match(indexHtml, /setupOverlayBackdropDismissal\(document\.getElementById\(id\), function\(\) \{\s*closeOverlayFromBackdrop\(id\);\s*}\);/);
  const idsStart = indexHtml.indexOf('const BACKDROP_DISMISS_OVERLAY_IDS =');
  const idsEnd = indexHtml.indexOf('];', idsStart);
  assert.notEqual(idsStart, -1);
  assert.notEqual(idsEnd, -1);
  assert.equal(indexHtml.slice(idsStart, idsEnd).includes("'edit-overlay'"), false);
  assert.match(indexHtml, /const OVERLAY_BACKDROP_TAP_MAX_MOVEMENT_PX = (?:8|9|10);/);
});

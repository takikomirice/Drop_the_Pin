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
  const bodyStart = indexHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < indexHtml.length; index += 1) {
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
    if (character === '}') {
      depth -= 1;
      if (depth === 0) return indexHtml.slice(start, index + 1);
    }
  }
  assert.fail(`Could not parse ${name}`);
}

function listenerTarget() {
  const listeners = new Map();
  return {
    listeners,
    addEventListener(type, listener) {
      if (!listeners.has(type)) listeners.set(type, new Set());
      listeners.get(type).add(listener);
    },
    removeEventListener(type, listener) {
      if (listeners.has(type)) listeners.get(type).delete(listener);
    },
    count(type) { return listeners.has(type) ? listeners.get(type).size : 0; }
  };
}

function createHarness(options = {}) {
  const viewportEvents = listenerTarget();
  const windowEvents = listenerTarget();
  const documentEvents = listenerTarget();
  const styleValues = new Map();
  const frames = [];
  const viewport = Object.assign(viewportEvents, {
    width: 390,
    height: 500,
    offsetTop: 24,
    offsetLeft: 0
  });
  const documentApi = Object.assign(documentEvents, {
    activeElement: null,
    documentElement: {
      style: {
        setProperty(name, value) { styleValues.set(name, value); },
        removeProperty(name) { styleValues.delete(name); }
      }
    }
  });
  const windowApi = Object.assign(windowEvents, {
    innerHeight: options.innerHeight || 844,
    visualViewport: options.withoutVisualViewport ? null : viewport,
    matchMedia(query) {
      return { matches: query === '(max-width: 899px)' ? options.mobile !== false : options.docked === true };
    }
  });
  const context = {
    window: windowApi,
    document: documentApi,
    dialogViewportCleanup: null,
    dialogViewportFrame: 0,
    requestAnimationFrame(callback) { frames.push(callback); return frames.length; },
    cancelAnimationFrame() {},
    isDockedPinDetailOverlay() { return options.docked === true; }
  };
  [
    'isMobileDialogViewport',
    'getFocusedDialogControl',
    'ensureFocusedDialogControlVisible',
    'syncDialogViewport',
    'scheduleDialogViewportSync',
    'setupDialogViewportTracking',
    'teardownDialogViewportTracking'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });
  return {
    context,
    viewport,
    windowEvents,
    documentEvents,
    styleValues,
    flushFrame() {
      const callback = frames.shift();
      if (callback) callback();
    }
  };
}

function focusedControl({ top = 650, bottom = 694, open = true } = {}) {
  const overlay = { classList: { contains(name) { return name === 'open' && open; } } };
  return {
    scrollCalls: [],
    matches(selector) { return selector.includes('input:not'); },
    closest(selector) { return selector === '.sheet-overlay' ? overlay : null; },
    getBoundingClientRect() { return { top, bottom }; },
    scrollIntoView(options) { this.scrollCalls.push(options); }
  };
}

test('mobile dialogs opt into safe-area layout and keep a dynamic viewport fallback', () => {
  assert.match(indexHtml, /name="viewport"[^>]*viewport-fit=cover/);
  assert.match(css, /--dialog-viewport-height:\s*100dvh/);
  assert.match(css, /--dialog-viewport-offset-top:\s*0px/);
  assert.match(css, /\.sheet-overlay\s*\{[\s\S]*?height:\s*var\(--dialog-viewport-height,\s*100dvh\)/);
  ['top', 'right', 'bottom', 'left'].forEach((side) => {
    assert.match(css, new RegExp(`env\\(safe-area-inset-${side}\\)`), side);
  });
  assert.match(css, /padding:[^;]*env\(safe-area-inset-bottom\)/);
  ['route-preview-sheet', 'drive-photo-import-sheet', 'import-preview-sheet', 'input-presets-sheet']
    .forEach((name) => assert.match(
      css,
      new RegExp(`\\.${name}\\s*\\{[^}]*var\\(--dialog-viewport-height,\\s*100dvh\\)`),
      name
    ));
});

test('visualViewport metrics are synchronized after resize, scroll, and orientation changes', () => {
  const harness = createHarness();
  const cleanup = harness.context.setupDialogViewportTracking();
  harness.flushFrame();

  assert.equal(harness.styleValues.get('--dialog-viewport-height'), '500px');
  assert.equal(harness.styleValues.get('--dialog-viewport-offset-top'), '24px');
  assert.equal(harness.viewport.count('resize'), 1);
  assert.equal(harness.viewport.count('scroll'), 1);
  assert.equal(harness.windowEvents.count('resize'), 1);
  assert.equal(harness.windowEvents.count('orientationchange'), 1);

  harness.context.setupDialogViewportTracking();
  assert.equal(harness.viewport.count('resize'), 1, 'setup must be idempotent');
  assert.equal(harness.documentEvents.count('focusin'), 1, 'focus listener must not be duplicated');

  cleanup();
  assert.equal(harness.viewport.count('resize'), 0);
  assert.equal(harness.viewport.count('scroll'), 0);
  assert.equal(harness.windowEvents.count('resize'), 0);
  assert.equal(harness.windowEvents.count('orientationchange'), 0);
  assert.equal(harness.documentEvents.count('focusin'), 0);
});

test('only a keyboard-obscured mobile dialog field scrolls by the nearest amount', () => {
  const harness = createHarness();
  const control = focusedControl();
  harness.context.document.activeElement = control;
  assert.equal(harness.context.ensureFocusedDialogControlVisible(control), true);
  assert.equal(control.scrollCalls.length, 1);
  assert.deepEqual(
    Object.assign({}, control.scrollCalls[0]),
    { block: 'nearest', inline: 'nearest', behavior: 'auto' }
  );

  const visible = focusedControl({ top: 100, bottom: 144 });
  assert.equal(harness.context.ensureFocusedDialogControlVisible(visible), false);
  assert.equal(visible.scrollCalls.length, 0);

  harness.context.window.innerHeight = 500;
  const keyboardClosed = focusedControl();
  assert.equal(harness.context.ensureFocusedDialogControlVisible(keyboardClosed), false);
  assert.equal(keyboardClosed.scrollCalls.length, 0);
});

test('desktop and the docked pin detail never receive keyboard compensation scrolling', () => {
  const desktop = createHarness({ mobile: false });
  const desktopControl = focusedControl();
  assert.equal(desktop.context.ensureFocusedDialogControlVisible(desktopControl), false);
  assert.equal(desktopControl.scrollCalls.length, 0);

  const docked = createHarness({ docked: true });
  const dockedControl = focusedControl();
  assert.equal(docked.context.ensureFocusedDialogControlVisible(dockedControl), false);
  assert.equal(dockedControl.scrollCalls.length, 0);
});

test('the viewport tracker is initialized once and has a page lifecycle cleanup', () => {
  assert.match(indexHtml, /setupDialogViewportTracking\(\);[\s\S]*?initializeApp\(\);/);
  assert.match(functionSource('setupDialogViewportTracking'), /pagehide/);
  assert.match(functionSource('teardownDialogViewportTracking'), /dialogViewportCleanup/);
});

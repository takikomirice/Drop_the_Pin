const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');

function functionSource(name) {
  const start = sharedHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const bodyStart = sharedHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < sharedHtml.length; index += 1) {
    const character = sharedHtml[index];
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
      if (depth === 0) return sharedHtml.slice(start, index + 1);
    }
  }
  assert.fail(`Could not parse ${name}`);
}

function classList() {
  const values = new Set();
  return {
    add(value) { values.add(value); },
    remove(value) { values.delete(value); },
    contains(value) { return values.has(value); }
  };
}

function opener(label) {
  return {
    label,
    connected: true,
    disabled: false,
    focusCount: 0,
    getAttribute() { return null; },
    closest() { return null; },
    focus() { this.focusCount += 1; }
  };
}

function createHarness() {
  const elements = {
    'shared-help-overlay': { classList: classList(), style: {}, innerHTML: '', closest() { return null; } },
    'shared-detail-overlay': { classList: classList(), style: {}, innerHTML: '', closest() { return null; } },
    'shared-photo-viewer-overlay': { classList: classList(), style: {}, innerHTML: '', closest() { return null; } }
  };
  const context = {
    sharedSurfaceOpenRecords: [],
    sharedBackgroundInteractionStateRecords: new Map(),
    SHARED_DISMISSIBLE_SURFACE_IDS: ['shared-help-overlay', 'shared-detail-overlay', 'shared-photo-viewer-overlay'],
    captureSharedInteractionState() { return {}; },
    restoreSharedInteractionState() {},
    syncSharedSurfaceInteractionState() {},
    focusSharedSurfaceInitial() {},
    closeSharedPhotoViewerForTrigger() { return false; },
    updateSharedPhotoViewerTrigger() { return false; },
    closeSharedPhotoViewer() { return context.closeSharedSurface('shared-photo-viewer-overlay'); },
    sharedPinAudioPlayer: { close() { return true; } },
    document: {
      activeElement: null,
      body: { children: [], classList: classList() },
      getElementById(id) { return elements[id]; },
      contains(node) { return !!node && node.connected !== false; }
    }
  };
  [
    'removeSharedSurfaceRecord', 'getSharedSurfaceRecord', 'getOpenSharedSurfaceRecords', 'openSharedSurface',
    'closeSharedSurface', 'canRestoreSharedSurfaceFocus', 'restoreSharedSurfaceFocus',
    'getTopSharedSurfaceRecord', 'closeSharedHelpPanel', 'closeSharedDetail',
    'dispatchSharedEscape', 'handleSharedGlobalKeydown'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });
  return { context, elements };
}

function pressEscape(context) {
  const event = {
    key: 'Escape', preventCount: 0, stopCount: 0,
    preventDefault() { this.preventCount += 1; },
    stopPropagation() { this.stopCount += 1; }
  };
  context.handleSharedGlobalKeydown(event);
  return event;
}

test('shared dispatcher closes only the newest viewer/help/detail surface and restores focus', () => {
  const inventoryMatch = sharedHtml.match(/const SHARED_DISMISSIBLE_SURFACE_IDS = Object\.freeze\(\[([\s\S]*?)\]\);/);
  assert.ok(inventoryMatch, 'shared surface inventory');
  const inventoryIds = [...inventoryMatch[1].matchAll(/'([^']+-overlay)'/g)].map((match) => match[1]).sort();
  assert.deepEqual(inventoryIds, ['shared-detail-overlay', 'shared-help-overlay', 'shared-photo-viewer-overlay']);

  const harness = createHarness();
  const helpOpener = opener('help');
  const detailOpener = opener('detail');
  const photoOpener = opener('photo');
  harness.context.document.activeElement = helpOpener;
  harness.context.openSharedSurface('shared-help-overlay');
  harness.context.document.activeElement = detailOpener;
  harness.context.openSharedSurface('shared-detail-overlay');
  harness.context.document.activeElement = photoOpener;
  harness.context.openSharedSurface('shared-photo-viewer-overlay');

  let event = pressEscape(harness.context);
  assert.equal(harness.elements['shared-photo-viewer-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['shared-detail-overlay'].classList.contains('open'), true);
  assert.equal(photoOpener.focusCount, 1);
  assert.equal(detailOpener.focusCount, 0);
  assert.equal(event.preventCount, 1);
  assert.equal(event.stopCount, 1);

  event = pressEscape(harness.context);
  assert.equal(harness.elements['shared-detail-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['shared-help-overlay'].classList.contains('open'), true);
  assert.equal(detailOpener.focusCount, 1);
  assert.equal(helpOpener.focusCount, 0);
  assert.equal(event.preventCount, 1);
  assert.equal(event.stopCount, 1);

  event = pressEscape(harness.context);
  assert.equal(harness.elements['shared-help-overlay'].classList.contains('open'), false);
  assert.equal(helpOpener.focusCount, 1);
  assert.equal(event.preventCount, 1);
  assert.equal(event.stopCount, 1);
});

test('shared detail closes without destroying its static shell and ignores a removed opener safely', () => {
  const harness = createHarness();
  const detailOpener = opener('removed');
  harness.elements['shared-detail-overlay'].innerHTML = '<div>detail</div>';
  harness.elements['shared-detail-overlay'].style.display = '';
  harness.context.document.activeElement = detailOpener;
  harness.context.openSharedSurface('shared-detail-overlay');
  detailOpener.connected = false;

  assert.doesNotThrow(() => pressEscape(harness.context));
  assert.equal(harness.elements['shared-detail-overlay'].classList.contains('open'), false);
  assert.equal(harness.elements['shared-detail-overlay'].innerHTML, '<div>detail</div>');
  assert.equal(harness.elements['shared-detail-overlay'].style.display, '');
  assert.equal(detailOpener.focusCount, 0);
});

test('shared global handler is unique and ignores ordinary search/list keys', () => {
  const harness = createHarness();
  const ordinary = {
    key: 'Enter', preventCount: 0, stopCount: 0,
    preventDefault() { this.preventCount += 1; },
    stopPropagation() { this.stopCount += 1; }
  };
  harness.context.handleSharedGlobalKeydown(ordinary);
  assert.equal(ordinary.preventCount, 0);
  assert.equal(ordinary.stopCount, 0);
  assert.equal(sharedHtml.match(/document\.addEventListener\('keydown'/g).length, 1);
  assert.match(functionSource('dispatchSharedEscape'), /closeSharedDetail/);
  assert.match(functionSource('dispatchSharedEscape'), /closeSharedHelpPanel/);
});

test('shared geocode results consume Escape before the surface dispatcher', () => {
  assert.match(
    sharedHtml,
    /if \(panelEl\.style\.display !== 'none'\) \{\s*closeCurrentPanel\(\);\s*e\.preventDefault\(\);\s*e\.stopPropagation\(\);/,
  );
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

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

function createClassList(initial) {
  const values = new Set(initial || []);
  return {
    add(...items) { items.forEach((item) => values.add(item)); },
    remove(...items) { items.forEach((item) => values.delete(item)); },
    toggle(item, force) {
      if (force === true) values.add(item);
      else if (force === false) values.delete(item);
      else if (values.has(item)) values.delete(item);
      else values.add(item);
      return values.has(item);
    },
    contains(item) { return values.has(item); }
  };
}

function createHarness(options) {
  const config = Object.assign({ desktop: true, panelVisible: true }, options);
  const elements = {};
  const body = {
    children: [],
    classList: createClassList(config.panelVisible ? ['panel-visible'] : ['panel-hidden'])
  };
  const documentApi = {
    activeElement: null,
    body,
    getElementById(id) { return elements[id] || null; },
    contains(element) { return !!element && element.connected !== false; }
  };

  function createElement(elementOptions) {
    const elementConfig = elementOptions || {};
    const attributes = Object.assign({}, elementConfig.attributes);
    const element = {
      id: elementConfig.id || '',
      tagName: String(elementConfig.tagName || 'DIV').toUpperCase(),
      connected: elementConfig.connected !== false,
      disabled: !!elementConfig.disabled,
      hidden: !!elementConfig.hidden,
      inert: elementConfig.inert === true || Object.prototype.hasOwnProperty.call(attributes, 'inert'),
      classList: createClassList(elementConfig.classes),
      style: { zIndex: elementConfig.zIndex || '' },
      getAttribute(name) {
        return Object.prototype.hasOwnProperty.call(attributes, name) ? attributes[name] : null;
      },
      setAttribute(name, value) {
        attributes[name] = String(value);
        if (name === 'inert') this.inert = true;
      },
      removeAttribute(name) {
        delete attributes[name];
        if (name === 'inert') this.inert = false;
      },
      hasAttribute(name) { return Object.prototype.hasOwnProperty.call(attributes, name); },
      focus() { documentApi.activeElement = this; }
    };
    elements[element.id] = element;
    if (elementConfig.bodyChild !== false) body.children.push(element);
    return element;
  }

  function createBackground(id, backgroundOptions) {
    return createElement(Object.assign({ id }, backgroundOptions));
  }

  function createOverlay(id, overlayOptions) {
    const overlayConfig = overlayOptions || {};
    const overlay = createElement({
      id,
      classes: ['sheet-overlay'],
      attributes: Object.assign({ 'aria-modal': 'true' }, overlayConfig.attributes),
      inert: overlayConfig.inert,
      zIndex: overlayConfig.zIndex
    });
    const control = createElement({ id: `${id}-control`, tagName: 'BUTTON', bodyChild: false });
    overlay.control = control;
    overlay.querySelector = function(selector) {
      return selector === '.sheet-body' ? null : null;
    };
    overlay.querySelectorAll = function() { return [control]; };
    overlay.contains = function(element) { return element === this || element === control; };
    return overlay;
  }

  const context = {
    document: documentApi,
    window: {
      matchMedia(query) {
        return { matches: query === '(min-width: 900px)' ? config.desktop : false };
      }
    },
    Map,
    WeakSet,
    overlayOpenRecords: [],
    overlayBackgroundStateRecords: new Map(),
    OVERLAY_STACK_Z_INDEX_BASE: 1400,
    OVERLAY_STACK_Z_INDEX_STEP: 10,
    focusOverlayInitial(overlay) { overlay.control.focus(); },
    clearOverlayFallbackFocus() {},
    restoreSurfaceFocus(opener) {
      if (!opener || typeof opener.focus !== 'function') return false;
      opener.focus();
      return true;
    }
  };

  [
    'removeOverlayOpenRecord', 'getOverlayOpenRecord', 'getTopOpenSheetOverlayRecord',
    'captureOverlayInteractionState', 'restoreOverlayInteractionState',
    'setOverlayInteractionAttribute', 'isDockedPinDetailOverlay',
    'getOpenSheetOverlayRecords', 'getTopModalSheetOverlayRecord',
    'restoreOverlayBackgroundState', 'syncOverlayInteractionState',
    'openOverlay', 'closeOverlay'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });

  return { context, document: documentApi, body, createBackground, createOverlay };
}

test('a single modal makes normal background inert and keeps only itself modal', () => {
  const harness = createHarness();
  const topbar = harness.createBackground('topbar');
  const app = harness.createBackground('app-shell');
  const dialog = harness.createOverlay('settings-overlay');

  harness.context.openOverlay('settings-overlay');

  assert.equal(topbar.inert, true);
  assert.equal(app.inert, true);
  assert.equal(topbar.getAttribute('aria-hidden'), 'true');
  assert.equal(app.getAttribute('aria-hidden'), 'true');
  assert.equal(dialog.inert, false);
  assert.equal(dialog.getAttribute('aria-hidden'), null);
  assert.equal(dialog.getAttribute('aria-modal'), 'true');
});

test('nested dialogs make the parent inert and stack by open order instead of DOM order', () => {
  const harness = createHarness();
  harness.createBackground('topbar');
  const child = harness.createOverlay('share-qr-overlay');
  const parent = harness.createOverlay('share-overlay');

  harness.context.openOverlay('share-overlay');
  harness.context.openOverlay('share-qr-overlay');

  assert.equal(parent.inert, true);
  assert.equal(parent.getAttribute('aria-hidden'), 'true');
  assert.equal(parent.getAttribute('aria-modal'), 'false');
  assert.equal(child.inert, false);
  assert.equal(child.getAttribute('aria-hidden'), null);
  assert.equal(child.getAttribute('aria-modal'), 'true');
  assert.equal(parent.style.zIndex, '1400');
  assert.equal(child.style.zIndex, '1410');
});

test('closing a child restores its parent as the only interactive modal', () => {
  const harness = createHarness();
  const topbar = harness.createBackground('topbar');
  const parent = harness.createOverlay('settings-overlay');
  const child = harness.createOverlay('input-presets-overlay');

  harness.context.openOverlay('settings-overlay');
  harness.context.openOverlay('input-presets-overlay');
  harness.context.closeOverlay('input-presets-overlay', { restoreFocus: false });

  assert.equal(topbar.inert, true);
  assert.equal(parent.inert, false);
  assert.equal(parent.getAttribute('aria-hidden'), null);
  assert.equal(parent.getAttribute('aria-modal'), 'true');
  assert.equal(parent.style.zIndex, '1400');
  assert.equal(child.classList.contains('open'), false);
});

test('closing every modal restores background and overlay attribute values exactly', () => {
  const harness = createHarness();
  const preservedBackground = harness.createBackground('preserved', {
    attributes: { inert: 'legacy', 'aria-hidden': 'false' }
  });
  const ordinaryBackground = harness.createBackground('ordinary');
  const dialog = harness.createOverlay('dialog', {
    inert: true,
    attributes: { inert: 'dialog-legacy', 'aria-hidden': 'false', 'aria-modal': 'false' },
    zIndex: '77'
  });

  harness.context.openOverlay('dialog');
  assert.equal(dialog.inert, false, 'the active dialog is temporarily operable');
  assert.equal(dialog.getAttribute('aria-modal'), 'true');

  harness.context.closeOverlay('dialog', { restoreFocus: false });

  assert.equal(preservedBackground.inert, true);
  assert.equal(preservedBackground.getAttribute('inert'), 'legacy');
  assert.equal(preservedBackground.getAttribute('aria-hidden'), 'false');
  assert.equal(ordinaryBackground.inert, false);
  assert.equal(ordinaryBackground.getAttribute('inert'), null);
  assert.equal(ordinaryBackground.getAttribute('aria-hidden'), null);
  assert.equal(dialog.inert, true);
  assert.equal(dialog.getAttribute('inert'), 'dialog-legacy');
  assert.equal(dialog.getAttribute('aria-hidden'), 'false');
  assert.equal(dialog.getAttribute('aria-modal'), 'false');
  assert.equal(dialog.style.zIndex, '77');
});

test('closed dialogs are temporarily non-modal and recover their original attributes afterward', () => {
  const harness = createHarness();
  harness.createBackground('topbar');
  const inactive = harness.createOverlay('inactive-overlay', {
    attributes: { 'aria-modal': 'true', 'aria-hidden': 'false' }
  });
  harness.createOverlay('active-overlay');

  harness.context.openOverlay('active-overlay');
  assert.equal(inactive.inert, true);
  assert.equal(inactive.getAttribute('aria-hidden'), 'true');
  assert.equal(inactive.getAttribute('aria-modal'), 'false');

  harness.context.closeOverlay('active-overlay', { restoreFocus: false });
  assert.equal(inactive.inert, false);
  assert.equal(inactive.getAttribute('aria-hidden'), 'false');
  assert.equal(inactive.getAttribute('aria-modal'), 'true');
});

test('desktop docked pin detail remains non-modal and leaves map and topbar operable', () => {
  const harness = createHarness({ desktop: true, panelVisible: true });
  const topbar = harness.createBackground('topbar');
  const map = harness.createBackground('app-shell');
  const detail = harness.createOverlay('pin-detail-overlay');

  harness.context.openOverlay('pin-detail-overlay');

  assert.equal(topbar.inert, false);
  assert.equal(map.inert, false);
  assert.equal(topbar.getAttribute('aria-hidden'), null);
  assert.equal(detail.inert, false);
  assert.equal(detail.getAttribute('aria-modal'), 'false');
  assert.equal(harness.context.getTopModalSheetOverlayRecord(), null);
});

test('mobile and hidden-panel pin detail are isolated as normal modals', () => {
  [
    { label: 'mobile', desktop: false, panelVisible: true },
    { label: 'desktop hidden panel', desktop: true, panelVisible: false }
  ].forEach((scenario) => {
    const harness = createHarness(scenario);
    const topbar = harness.createBackground('topbar');
    const detail = harness.createOverlay('pin-detail-overlay');

    harness.context.openOverlay('pin-detail-overlay');

    assert.equal(topbar.inert, true, scenario.label);
    assert.equal(detail.inert, false, scenario.label);
    assert.equal(detail.getAttribute('aria-modal'), 'true', scenario.label);
    assert.equal(harness.context.getTopModalSheetOverlayRecord().id, 'pin-detail-overlay', scenario.label);
  });
});

test('a modal above docked pin detail isolates it and closing the modal restores dock behavior', () => {
  const harness = createHarness({ desktop: true, panelVisible: true });
  const topbar = harness.createBackground('topbar');
  const detail = harness.createOverlay('pin-detail-overlay');
  const share = harness.createOverlay('share-overlay');

  harness.context.openOverlay('pin-detail-overlay');
  harness.context.openOverlay('share-overlay');
  assert.equal(topbar.inert, true);
  assert.equal(detail.inert, true);
  assert.equal(detail.getAttribute('aria-hidden'), 'true');
  assert.equal(share.getAttribute('aria-modal'), 'true');

  harness.context.closeOverlay('share-overlay', { restoreFocus: false });
  assert.equal(topbar.inert, false);
  assert.equal(detail.inert, false);
  assert.equal(detail.getAttribute('aria-hidden'), null);
  assert.equal(detail.getAttribute('aria-modal'), 'false');
});

test('panel and desktop breakpoint changes re-synchronize an open pin detail surface', () => {
  assert.match(functionSource('renderPanelToggle'), /syncOverlayInteractionState\(\)/);
  assert.match(indexHtml, /const handleDesktopViewChange = function\(\) \{[\s\S]{0,240}syncOverlayInteractionState\(\)/);
  assert.match(indexHtml, /desktopViewMedia\.(?:addEventListener|addListener)\([^\n]*handleDesktopViewChange/);
});

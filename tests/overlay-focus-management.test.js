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
    contains(item) { return values.has(item); }
  };
}

function createHarness() {
  const elements = {};
  const body = { children: [], classList: createClassList(['panel-hidden']) };
  const documentApi = {
    activeElement: null,
    body,
    getElementById(id) { return elements[id] || null; },
    contains(element) { return !!element && element.connected !== false; }
  };

  function createElement(options) {
    const config = options || {};
    const attributes = Object.assign({}, config.attributes);
    const element = {
      id: config.id || '',
      tagName: String(config.tagName || 'BUTTON').toUpperCase(),
      type: config.type || '',
      disabled: !!config.disabled,
      hidden: !!config.hidden,
      connected: config.connected !== false,
      visible: config.visible !== false,
      display: config.display || '',
      visibility: config.visibility || '',
      focusCount: 0,
      inert: config.inert === true || Object.prototype.hasOwnProperty.call(attributes, 'inert'),
      classList: createClassList(config.classes),
      style: { zIndex: config.zIndex || '' },
      ownerOverlay: config.ownerOverlay || null,
      hiddenAncestor: config.hiddenAncestor || null,
      ariaHiddenAncestor: config.ariaHiddenAncestor || null,
      getAttribute(name) {
        if (name === 'type' && this.type) return this.type;
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
      matches(selector) {
        if (selector === ':disabled') return this.disabled;
        if (selector === 'input[type="hidden"]') return this.tagName === 'INPUT' && this.type === 'hidden';
        return false;
      },
      closest(selector) {
        if (selector === '.sheet-overlay') return this.ownerOverlay;
        if (selector === '[hidden]') return this.hidden ? this : this.hiddenAncestor;
        if (selector === '[aria-hidden="true"]') {
          return this.getAttribute('aria-hidden') === 'true' ? this : this.ariaHiddenAncestor;
        }
        return null;
      },
      getClientRects() { return this.visible && !this.hidden && !this.hiddenAncestor ? [{}] : []; },
      focus() {
        this.focusCount += 1;
        documentApi.activeElement = this;
      }
    };
    return element;
  }

  function createOverlay(id, focusCandidates) {
    const overlay = createElement({ id, tagName: 'DIV', classes: ['sheet-overlay'] });
    const body = createElement({ id: `${id}-body`, tagName: 'DIV', ownerOverlay: overlay });
    overlay.body = body;
    overlay.focusCandidates = focusCandidates || [];
    overlay.focusCandidates.forEach((element) => { element.ownerOverlay = overlay; });
    overlay.querySelector = function(selector) {
      if (selector === '.sheet-body') return body;
      if (selector === '[data-overlay-initial-focus]') {
        return this.focusCandidates.find((element) => element.hasAttribute('data-overlay-initial-focus')) || null;
      }
      return null;
    };
    overlay.querySelectorAll = function(selector) {
      if (selector === '[data-overlay-initial-focus]') {
        return this.focusCandidates.filter((element) => element.hasAttribute('data-overlay-initial-focus'));
      }
      return this.focusCandidates.slice();
    };
    overlay.contains = function(element) {
      return element === this || element === body || this.focusCandidates.includes(element);
    };
    elements[id] = overlay;
    documentApi.body.children.push(overlay);
    return overlay;
  }

  const context = {
    document: documentApi,
    window: {
      getComputedStyle(element) {
        return { display: element.display, visibility: element.visibility };
      },
      matchMedia() { return { matches: false }; }
    },
    Map,
    WeakSet,
    overlayOpenRecords: [],
    overlayBackgroundStateRecords: new Map(),
    overlayFallbackFocusBodies: new WeakSet(),
    OVERLAY_STACK_Z_INDEX_BASE: 1400,
    OVERLAY_STACK_Z_INDEX_STEP: 10,
    OVERLAY_FOCUSABLE_SELECTOR: [
      'a[href]', 'button', 'input:not([type="hidden"])', 'select', 'textarea',
      'summary', '[contenteditable="true"]', '[tabindex]'
    ].join(',')
  };

  [
    'setOverlayInteractionAttribute', 'captureOverlayInteractionState',
    'restoreOverlayInteractionState', 'removeOverlayOpenRecord', 'getOverlayOpenRecord',
    'isDockedPinDetailOverlay', 'getOpenSheetOverlayRecords', 'getTopOpenSheetOverlayRecord',
    'getTopModalSheetOverlayRecord', 'restoreOverlayBackgroundState', 'syncOverlayInteractionState',
    'isOverlayElementVisible', 'isOverlayFocusableElement', 'getOverlayFocusableElements',
    'clearOverlayFallbackFocus', 'focusOverlayInitial', 'canRestoreSurfaceFocus',
    'restoreSurfaceFocus', 'openOverlay', 'closeOverlay', 'trapOverlayFocus'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });

  return {
    context,
    document: documentApi,
    createElement,
    createOverlay,
    keydown(shiftKey) {
      const event = {
        key: 'Tab',
        shiftKey: !!shiftKey,
        preventCount: 0,
        stopCount: 0,
        preventDefault() { this.preventCount += 1; },
        stopPropagation() { this.stopCount += 1; }
      };
      context.trapOverlayFocus(event);
      return event;
    }
  };
}

test('initial focus prefers a valid marker, then the first valid interactive element', () => {
  const harness = createHarness();
  const hiddenMarker = harness.createElement({ hidden: true, attributes: { 'data-overlay-initial-focus': '' } });
  const disabled = harness.createElement({ disabled: true });
  const negativeTabIndex = harness.createElement({ attributes: { tabindex: '-1' } });
  const ariaHidden = harness.createElement({ attributes: { 'aria-hidden': 'true' } });
  const firstInput = harness.createElement({ tagName: 'INPUT' });
  const preferred = harness.createElement({ tagName: 'A', attributes: { href: '#', 'data-overlay-initial-focus': '' } });
  harness.createOverlay('dialog', [hiddenMarker, disabled, negativeTabIndex, ariaHidden, firstInput, preferred]);

  harness.context.openOverlay('dialog');
  assert.equal(harness.document.activeElement, preferred);

  preferred.hidden = true;
  harness.context.closeOverlay('dialog', { restoreFocus: false });
  harness.context.openOverlay('dialog');
  assert.equal(harness.document.activeElement, firstInput);
});

test('initial focus excludes CSS-hidden and hidden-ancestor controls', () => {
  const harness = createHarness();
  const displayNone = harness.createElement({ display: 'none' });
  const visibilityHidden = harness.createElement({ visibility: 'hidden' });
  const clippedByParent = harness.createElement({ visible: false });
  const parentAriaHidden = harness.createElement({ ariaHiddenAncestor: {} });
  const valid = harness.createElement({ tagName: 'SELECT' });
  harness.createOverlay('dialog', [displayNone, visibilityHidden, clippedByParent, parentAriaHidden, valid]);

  harness.context.openOverlay('dialog');
  assert.equal(harness.document.activeElement, valid);
});

test('a dialog without interactive controls focuses its sheet body with only a temporary tabindex', () => {
  const harness = createHarness();
  const overlay = harness.createOverlay('empty-dialog', []);

  harness.context.openOverlay('empty-dialog');
  assert.equal(harness.document.activeElement, overlay.body);
  assert.equal(overlay.body.getAttribute('tabindex'), '-1');

  harness.context.closeOverlay('empty-dialog', { restoreFocus: false });
  assert.equal(overlay.body.getAttribute('tabindex'), null);
});

test('Tab and Shift+Tab wrap between the last and first controls', () => {
  const harness = createHarness();
  const first = harness.createElement({ tagName: 'INPUT' });
  const middle = harness.createElement({ tagName: 'SUMMARY' });
  const last = harness.createElement({ tagName: 'BUTTON' });
  harness.createOverlay('dialog', [first, middle, last]);
  harness.context.openOverlay('dialog');

  last.focus();
  let event = harness.keydown(false);
  assert.equal(harness.document.activeElement, first);
  assert.equal(event.preventCount, 1);

  first.focus();
  event = harness.keydown(true);
  assert.equal(harness.document.activeElement, last);
  assert.equal(event.preventCount, 1);
});

test('zero and one focusable control trap Tab without throwing', () => {
  const emptyHarness = createHarness();
  const emptyOverlay = emptyHarness.createOverlay('empty', []);
  emptyHarness.context.openOverlay('empty');
  assert.doesNotThrow(() => emptyHarness.keydown(false));
  assert.equal(emptyHarness.document.activeElement, emptyOverlay.body);

  const singleHarness = createHarness();
  const only = singleHarness.createElement({ tagName: 'TEXTAREA' });
  singleHarness.createOverlay('single', [only]);
  singleHarness.context.openOverlay('single');
  const event = singleHarness.keydown(false);
  assert.equal(singleHarness.document.activeElement, only);
  assert.equal(event.preventCount, 1);
});

test('only the top open dialog traps focus and focus outside it is pulled back in', () => {
  const harness = createHarness();
  const lowerFirst = harness.createElement({ tagName: 'BUTTON' });
  const lowerLast = harness.createElement({ tagName: 'A', attributes: { href: '#' } });
  const upperFirst = harness.createElement({ tagName: 'SELECT' });
  const upperLast = harness.createElement({ tagName: 'BUTTON' });
  harness.createOverlay('lower', [lowerFirst, lowerLast]);
  harness.createOverlay('upper', [upperFirst, upperLast]);
  harness.context.openOverlay('lower');
  harness.context.openOverlay('upper');

  upperLast.focus();
  harness.keydown(false);
  assert.equal(harness.document.activeElement, upperFirst);
  assert.equal(lowerFirst.focusCount, 1, 'lower dialog only received its own initial focus');

  const outside = harness.createElement({ tagName: 'BUTTON' });
  outside.focus();
  harness.keydown(true);
  assert.equal(harness.document.activeElement, upperLast);
});

test('closing restores its live opener', () => {
  const harness = createHarness();
  const opener = harness.createElement({ tagName: 'BUTTON' });
  harness.createOverlay('dialog', [harness.createElement({ tagName: 'INPUT' })]);
  opener.focus();

  harness.context.openOverlay('dialog');
  harness.context.closeOverlay('dialog');
  assert.equal(harness.document.activeElement, opener);
  assert.equal(opener.focusCount, 2);
});

test('closing safely ignores removed, disabled, hidden, aria-hidden and closed-dialog openers', () => {
  const cases = [
    { label: 'removed', configure(opener) { opener.connected = false; } },
    { label: 'disabled', configure(opener) { opener.disabled = true; } },
    { label: 'hidden', configure(opener) { opener.hidden = true; } },
    { label: 'aria-hidden', configure(opener) { opener.setAttribute('aria-hidden', 'true'); } },
    {
      label: 'closed parent dialog',
      configure(opener, harness) {
        const parent = harness.createOverlay('parent', []);
        opener.ownerOverlay = parent;
      }
    }
  ];

  cases.forEach(({ label, configure }) => {
    const harness = createHarness();
    const opener = harness.createElement({ tagName: 'BUTTON' });
    harness.createOverlay('dialog', [harness.createElement({ tagName: 'BUTTON' })]);
    opener.focus();
    harness.context.openOverlay('dialog');
    configure(opener, harness);
    harness.context.closeOverlay('dialog');
    assert.equal(opener.focusCount, 1, label);
  });
});

test('restoreFocus false suppresses the old opener during a dialog transition', () => {
  const harness = createHarness();
  const opener = harness.createElement({ tagName: 'BUTTON' });
  const oldControl = harness.createElement({ tagName: 'BUTTON' });
  const nextControl = harness.createElement({ tagName: 'INPUT' });
  harness.createOverlay('old-dialog', [oldControl]);
  harness.createOverlay('next-dialog', [nextControl]);
  opener.focus();
  harness.context.openOverlay('old-dialog');

  harness.context.closeOverlay('old-dialog', { restoreFocus: false });
  assert.equal(opener.focusCount, 1);
  harness.context.openOverlay('next-dialog');
  assert.equal(harness.document.activeElement, nextControl);
});

test('Escape, backdrop and busy dismissal contracts remain routed through existing guards', () => {
  const escape = functionSource('dispatchEscape');
  const backdrop = functionSource('closeOverlayFromBackdrop');
  const dismiss = functionSource('dismissOverlayById');
  assert.match(escape, /dismissOverlayById\(record\.id\)/);
  assert.doesNotMatch(escape, /restoreSurfaceFocus\(record\.opener\)/);
  assert.match(backdrop, /dismissOverlayById\(record\.id\)/);
  assert.match(dismiss, /settings-overlay[\s\S]*closeSettingsModal\(\)/);
  assert.match(dismiss, /delete-overlay[\s\S]*cancelDeleteConfirmation\(\)/);
  assert.match(dismiss, /state\.trackImport\.saving/);
  assert.match(dismiss, /state\.multiPhotoImport\.registering/);
  assert.match(functionSource('closeSettingsModal'), /settingsSavePending|isProductionImportBusy\(\)/);
  assert.match(functionSource('cancelDeleteConfirmation'), /state\.deleteMutationPending/);
});

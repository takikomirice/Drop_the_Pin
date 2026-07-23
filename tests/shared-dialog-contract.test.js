const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');
const css = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const mainScriptStart = sharedHtml.lastIndexOf('<script>', sharedHtml.indexOf('const SHARED_DEFAULT_COLOR'));
const bodyMarkup = sharedHtml.slice(sharedHtml.indexOf('<body'), mainScriptStart)
  .replace(/<script(?:\s[^>]*)?>[\s\S]*?<\/script>/gi, '');

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

function elementMarkup(id) {
  const start = bodyMarkup.indexOf(`id="${id}"`);
  assert.notEqual(start, -1, `Expected #${id}`);
  const tagStart = bodyMarkup.lastIndexOf('<', start);
  return bodyMarkup.slice(tagStart, bodyMarkup.indexOf('>', start) + 1);
}

function classList(initial = []) {
  const values = new Set(initial);
  return {
    add(...items) { items.forEach((item) => values.add(item)); },
    remove(...items) { items.forEach((item) => values.delete(item)); },
    contains(item) { return values.has(item); },
  };
}

function createHarness() {
  const elements = {};
  const body = { children: [], classList: classList() };
  let desktop = false;
  const documentApi = {
    activeElement: null,
    body,
    getElementById(id) { return elements[id] || null; },
    contains(element) { return Boolean(element && element.connected !== false); },
  };

  function createElement(options = {}) {
    const attributes = { ...(options.attributes || {}) };
    const element = {
      id: options.id || '',
      tagName: String(options.tagName || 'DIV').toUpperCase(),
      type: options.type || '',
      disabled: false,
      hidden: false,
      inert: Boolean(options.inert),
      connected: true,
      classList: classList(options.classes),
      style: {},
      focusCount: 0,
      focus() { this.focusCount += 1; documentApi.activeElement = this; },
      getAttribute(name) { return Object.prototype.hasOwnProperty.call(attributes, name) ? attributes[name] : null; },
      setAttribute(name, value) { attributes[name] = String(value); if (name === 'inert') this.inert = true; },
      removeAttribute(name) { delete attributes[name]; if (name === 'inert') this.inert = false; },
      matches(selector) { return selector === ':disabled' ? this.disabled : false; },
      closest(selector) {
        if (selector === '.sheet-overlay') return this.ownerOverlay || null;
        if (selector === '[hidden]' || selector === '[aria-hidden="true"]') return null;
        return null;
      },
      contains(target) { return this === target || (this.focusables || []).includes(target); },
      querySelectorAll(selector) {
        if (selector === '[data-shared-initial-focus]') {
          return (this.focusables || []).filter((item) => item.getAttribute('data-shared-initial-focus') !== null);
        }
        return this.focusables || [];
      },
    };
    if (element.id) elements[element.id] = element;
    return element;
  }

  const appShell = createElement({ id: 'shared-app-shell', attributes: { 'aria-hidden': 'false' } });
  const help = createElement({
    id: 'shared-help-overlay', classes: ['sheet-overlay'],
    attributes: { role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'shared-help-title' },
  });
  const detail = createElement({
    id: 'shared-detail-overlay', classes: ['sheet-overlay'],
    attributes: { role: 'dialog', 'aria-modal': 'true', 'aria-labelledby': 'shared-detail-title' },
  });
  const helpClose = createElement({ tagName: 'BUTTON', attributes: { 'data-shared-initial-focus': '' } });
  const helpLink = createElement({ tagName: 'A', attributes: { href: 'https://example.test' } });
  helpClose.ownerOverlay = help;
  helpLink.ownerOverlay = help;
  help.focusables = [helpClose, helpLink];
  detail.focusables = [createElement({ tagName: 'BUTTON', attributes: { 'data-shared-initial-focus': '' } })];
  detail.focusables[0].ownerOverlay = detail;
  body.children = [appShell, help, detail];

  const context = {
    document: documentApi,
    window: { matchMedia() { return { matches: desktop }; } },
    SHARED_DISMISSIBLE_SURFACE_IDS: ['shared-help-overlay', 'shared-detail-overlay'],
    SHARED_SURFACE_FOCUSABLE_SELECTOR: 'button, a[href]',
    sharedSurfaceOpenRecords: [],
    sharedBackgroundInteractionStateRecords: new Map(),
  };
  [
    'setSharedInteractionAttribute', 'captureSharedInteractionState', 'restoreSharedInteractionState',
    'removeSharedSurfaceRecord', 'getSharedSurfaceRecord', 'isSharedDockedPinDetailOverlay', 'getOpenSharedSurfaceRecords',
    'restoreSharedBackgroundInteractionState', 'syncSharedSurfaceInteractionState',
    'isSharedSurfaceFocusable', 'getSharedSurfaceFocusableElements', 'focusSharedSurfaceInitial',
    'canRestoreSharedSurfaceFocus', 'restoreSharedSurfaceFocus', 'getTopSharedSurfaceRecord', 'getTopSharedModalSurfaceRecord',
    'openSharedSurface', 'closeSharedSurface', 'trapSharedSurfaceFocus',
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });
  return {
    context, document: documentApi, appShell, help, detail, helpClose, helpLink,
    setDesktop(value) { desktop = Boolean(value); }
  };
}

test('help and detail dialogs expose complete labelled read-only dialog markup', () => {
  const help = elementMarkup('shared-help-overlay');
  const detail = elementMarkup('shared-detail-overlay');
  assert.match(help, /role="dialog"/);
  assert.match(help, /aria-modal="true"/);
  assert.match(help, /aria-labelledby="shared-help-title"/);
  assert.match(help, /aria-describedby="shared-help-description"/);
  assert.match(detail, /role="dialog"/);
  assert.match(detail, /aria-modal="true"/);
  assert.match(detail, /aria-labelledby="shared-detail-title"/);
  assert.match(detail, /aria-describedby="shared-detail-description"/);
  assert.match(bodyMarkup, /id="shared-help-title"/);
  assert.match(bodyMarkup, /id="shared-help-description"/);
});

test('help and detail use index ghost close controls with 44px targets', () => {
  const helpClose = bodyMarkup.match(/<button\b[^>]*id="shared-help-close"[^>]*>[\s\S]*?<\/button>/);
  assert.ok(helpClose);
  assert.match(helpClose[0], /class="ghost-btn"/);
  assert.match(helpClose[0], /title="閉じる"/);
  assert.match(helpClose[0], /data-shared-initial-focus/);
  assert.match(helpClose[0], />閉じる<\/button>/);

  const detailClose = bodyMarkup.match(/<button\b[^>]*id="shared-detail-close"[^>]*>[\s\S]*?<\/button>/);
  assert.ok(detailClose);
  assert.match(detailClose[0], /class="ghost-btn"/);
  assert.match(detailClose[0], /data-shared-initial-focus/);
  assert.match(detailClose[0], />閉じる<\/button>/);
  assert.match(css, /\.ghost-btn\s*\{[^}]*min-height:\s*44px/);
});

test('dialog sheets use the index sheet body scroller and responsive handle contracts', () => {
  assert.equal((bodyMarkup.match(/class="sheet-handle" aria-hidden="true"/g) || []).length, 2);
  assert.equal((bodyMarkup.match(/class="sheet-body"/g) || []).length, 2);
  assert.match(css, /\.sheet-body\s*\{[^}]*overflow-y:\s*auto[^}]*overscroll-behavior:\s*contain[^}]*scroll-padding-block:\s*12px 64px/);
  assert.match(css, /\.sheet-handle\s*\{[^}]*width:\s*42px[^}]*height:\s*4px/);
  assert.match(css, /@media \(min-width:\s*900px\)[\s\S]*?\.sheet-handle\s*\{[^}]*display:\s*none/);
  assert.match(css, /body\.shared-narrow-view #shared-detail-overlay \.sheet-body\s*\{[^}]*height:\s*min\(75dvh[^}]*var\(--safe-area-bottom\)|body\.shared-narrow-view #shared-detail-overlay \.sheet-body\s*\{[^}]*height:\s*min\(75dvh[^}]*env\(safe-area-inset-bottom\)/);
});

test('opening a dialog focuses its marked control and makes the background inert', () => {
  const harness = createHarness();
  const opener = { connected: true, disabled: false, focusCount: 0, focus() { this.focusCount += 1; }, getAttribute() { return null; }, closest() { return null; } };
  harness.document.activeElement = opener;
  harness.context.openSharedSurface('shared-help-overlay');

  assert.equal(harness.help.classList.contains('open'), true);
  assert.equal(harness.document.activeElement, harness.helpClose);
  assert.equal(harness.appShell.inert, true);
  assert.equal(harness.appShell.getAttribute('aria-hidden'), 'true');
  assert.equal(harness.help.inert, false);
  assert.equal(harness.help.getAttribute('aria-hidden'), null);
  assert.equal(harness.help.getAttribute('aria-modal'), 'true');
});

test('Tab and Shift+Tab cycle within only the top shared dialog', () => {
  const harness = createHarness();
  harness.context.openSharedSurface('shared-help-overlay');
  harness.helpLink.focus();
  let event = { key: 'Tab', shiftKey: false, preventCount: 0, stopCount: 0, preventDefault() { this.preventCount += 1; }, stopPropagation() { this.stopCount += 1; } };
  assert.equal(harness.context.trapSharedSurfaceFocus(event), true);
  assert.equal(harness.document.activeElement, harness.helpClose);
  assert.equal(event.preventCount, 1);

  event = { key: 'Tab', shiftKey: true, preventCount: 0, stopCount: 0, preventDefault() { this.preventCount += 1; }, stopPropagation() { this.stopCount += 1; } };
  assert.equal(harness.context.trapSharedSurfaceFocus(event), true);
  assert.equal(harness.document.activeElement, harness.helpLink);
  assert.equal(event.stopCount, 1);
});

test('desktop docked detail is non-modal while mobile detail remains modal', () => {
  const harness = createHarness();
  harness.setDesktop(true);
  harness.document.body.classList.add('shared-panel-visible');
  harness.context.openSharedSurface('shared-detail-overlay');

  assert.equal(harness.detail.getAttribute('aria-modal'), 'false');
  assert.equal(harness.appShell.inert, false);
  const tabEvent = { key: 'Tab', preventDefault() {}, stopPropagation() {} };
  assert.equal(harness.context.trapSharedSurfaceFocus(tabEvent), false);

  harness.context.closeSharedSurface('shared-detail-overlay');
  harness.setDesktop(false);
  harness.context.openSharedSurface('shared-detail-overlay');
  assert.equal(harness.detail.getAttribute('aria-modal'), 'true');
  assert.equal(harness.appShell.inert, true);
});

test('closing restores background attributes and focus to the live opener', () => {
  const harness = createHarness();
  const opener = { connected: true, disabled: false, focusCount: 0, focus() { this.focusCount += 1; harness.document.activeElement = this; }, getAttribute() { return null; }, closest() { return null; } };
  harness.document.activeElement = opener;
  harness.context.openSharedSurface('shared-help-overlay');
  harness.context.closeSharedSurface('shared-help-overlay');

  assert.equal(harness.help.classList.contains('open'), false);
  assert.equal(harness.appShell.inert, false);
  assert.equal(harness.appShell.getAttribute('inert'), null);
  assert.equal(harness.appShell.getAttribute('aria-hidden'), 'false');
  assert.equal(opener.focusCount, 1);
  assert.equal(harness.context.sharedSurfaceOpenRecords.length, 0);
});

test('surface helpers avoid duplicate records and global listeners', () => {
  const harness = createHarness();
  harness.context.openSharedSurface('shared-help-overlay');
  harness.context.openSharedSurface('shared-help-overlay');
  assert.equal(harness.context.sharedSurfaceOpenRecords.length, 1);
  assert.doesNotMatch(functionSource('openSharedSurface'), /addEventListener/);
  assert.equal((sharedHtml.match(/document\.addEventListener\('keydown',\s*handleSharedGlobalKeydown\)/g) || []).length, 1);
  assert.equal((sharedHtml.match(/setupSharedOverlayBackdropDismissal\(document\.getElementById\('shared-help-overlay'\)/g) || []).length, 1);
  assert.equal((sharedHtml.match(/setupSharedOverlayBackdropDismissal\(document\.getElementById\('shared-detail-overlay'\)/g) || []).length, 1);
});

test('pin detail remains read-only and keeps protected photos plus route context', () => {
  const detail = functionSource('openSharedDetail');
  assert.doesNotMatch(detail, /sharedPinListIconMarkup|shared-detail-icon|shared-detail-pin-icon/);
  assert.doesNotMatch(bodyMarkup, /id="shared-detail-icon"|class="shared-detail-pin-icon"/);
  assert.match(detail, /pin\.status/);
  assert.match(detail, /getSharedPinRouteLabels\(pin\.id\)/);
  assert.match(bodyMarkup, /id="shared-detail-image" class="photo-fit-cover protected-photo" draggable="false"/);
  assert.match(detail, /getElementById\('shared-detail-image'\)/);
  assert.doesNotMatch(detail, /編集|削除|保存|共有作成/);
  assert.doesNotMatch(sharedHtml, /visualViewport/);
});

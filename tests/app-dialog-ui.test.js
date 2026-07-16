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

function createHarness() {
  const elements = {};
  const stack = [];
  const documentApi = {
    activeElement: null,
    getElementById(id) { return elements[id] || null; }
  };

  function element(id, tagName) {
    const attributes = {};
    const classes = new Set();
    const node = {
      id,
      tagName: String(tagName || 'DIV').toUpperCase(),
      textContent: '',
      className: '',
      disabled: false,
      dataset: {},
      focusCount: 0,
      classList: {
        add(...items) { items.forEach((item) => classes.add(item)); },
        remove(...items) { items.forEach((item) => classes.delete(item)); },
        contains(item) { return classes.has(item); }
      },
      setAttribute(name, value) { attributes[name] = String(value); },
      getAttribute(name) { return Object.prototype.hasOwnProperty.call(attributes, name) ? attributes[name] : null; },
      removeAttribute(name) { delete attributes[name]; },
      focus() {
        this.focusCount += 1;
        documentApi.activeElement = this;
      }
    };
    elements[id] = node;
    return node;
  }

  [
    ['app-notification-overlay', 'div'], ['app-notification-title', 'div'],
    ['app-notification-message', 'div'], ['app-notification-close', 'button'],
    ['app-confirmation-overlay', 'div'], ['app-confirmation-title', 'div'],
    ['app-confirmation-message', 'div'], ['app-confirmation-confirm', 'button'],
    ['app-confirmation-cancel', 'button']
  ].forEach(([id, tagName]) => element(id, tagName));

  function openOverlay(id) {
    const opener = documentApi.activeElement;
    stack.push({ id, opener });
    elements[id].classList.add('open');
    const initial = id === 'app-notification-overlay'
      ? elements['app-notification-close']
      : elements['app-confirmation-confirm'];
    initial.focus();
  }

  function closeOverlay(id) {
    const index = stack.map((entry) => entry.id).lastIndexOf(id);
    if (index === -1) return;
    const [record] = stack.splice(index, 1);
    elements[id].classList.remove('open');
    if (record.opener && typeof record.opener.focus === 'function') record.opener.focus();
  }

  const context = {
    document: documentApi,
    Promise,
    String,
    Array,
    Object,
    appDialogQueue: [],
    activeAppDialog: null,
    appDialogRequestSequence: 0,
    openOverlay,
    closeOverlay,
    getTopModalSheetOverlayRecord() { return stack.length ? stack[stack.length - 1] : null; }
  };
  [
    'normalizeAppDialogConfig', 'hasPendingAppConfirmation', 'renderAppDialog',
    'showNextAppDialog', 'enqueueAppDialog', 'showAppNotification',
    'requestAppConfirmation', 'settleAppDialog', 'dismissAppDialog',
    'handleAppDialogEnter'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });

  return { context, document: documentApi, elements, stack, element };
}

test('notification and confirmation dialogs expose configurable in-app markup', () => {
  assert.match(indexHtml, /id="app-notification-overlay"[^>]*role="dialog"[^>]*aria-modal="true"/);
  assert.match(indexHtml, /id="app-confirmation-overlay"[^>]*role="alertdialog"[^>]*aria-modal="true"/);
  assert.match(indexHtml, /\.app-dialog-sheet[\s\S]*max-height:[^;]*(?:dvh|dialog-viewport-height)/);
  assert.match(indexHtml, /\.app-dialog-message[\s\S]*white-space:\s*pre-line/);
});

test('confirmation approval and cancellation settle once and reject a duplicate request', async () => {
  const harness = createHarness();
  const first = harness.context.requestAppConfirmation({
    title: '削除確認', message: '削除しますか？', confirmLabel: '削除する', cancelLabel: '戻る', danger: true
  });
  const duplicate = harness.context.requestAppConfirmation({ message: '重複' });

  assert.equal(await duplicate, false);
  assert.equal(harness.elements['app-confirmation-title'].textContent, '削除確認');
  assert.equal(harness.elements['app-confirmation-message'].textContent, '削除しますか？');
  assert.equal(harness.elements['app-confirmation-confirm'].textContent, '削除する');
  assert.equal(harness.elements['app-confirmation-cancel'].textContent, '戻る');
  assert.equal(harness.elements['app-confirmation-confirm'].className, 'danger-btn');

  assert.equal(harness.context.settleAppDialog(true), true);
  assert.equal(harness.context.settleAppDialog(true), false);
  assert.equal(await first, true);

  const cancelled = harness.context.requestAppConfirmation({ message: '通常確認' });
  assert.equal(harness.context.dismissAppDialog('app-confirmation-overlay'), true);
  assert.equal(await cancelled, false);
});

test('nested confirmation uses the overlay stack and Escape-style dismissal restores its opener', async () => {
  const harness = createHarness();
  const baseOpener = harness.element('base-opener', 'button');
  const nestedOpener = harness.element('nested-opener', 'button');
  baseOpener.focus();
  harness.context.openOverlay('app-notification-overlay');
  nestedOpener.focus();

  const confirmation = harness.context.requestAppConfirmation({ message: 'nested' });
  assert.deepEqual(harness.stack.map((entry) => entry.id), ['app-notification-overlay', 'app-confirmation-overlay']);

  assert.equal(harness.context.dismissAppDialog('app-confirmation-overlay'), true);
  assert.equal(await confirmation, false);
  assert.deepEqual(harness.stack.map((entry) => entry.id), ['app-notification-overlay']);
  assert.equal(harness.document.activeElement, nestedOpener);
});

test('Enter settles only the top app dialog and notification resolves through the same queue', async () => {
  const harness = createHarness();
  const notice = harness.context.showAppNotification({ title: '読込エラー', message: '詳細', confirmLabel: '閉じる' });
  const event = {
    key: 'Enter', preventCount: 0, stopCount: 0,
    preventDefault() { this.preventCount += 1; },
    stopPropagation() { this.stopCount += 1; }
  };
  assert.equal(harness.context.handleAppDialogEnter(event), true);
  assert.equal(await notice, true);
  assert.equal(event.preventCount, 1);
  assert.equal(event.stopCount, 1);

  const confirmation = harness.context.requestAppConfirmation({ message: 'キャンセルしますか？' });
  const cancelEvent = {
    key: 'Enter', target: harness.elements['app-confirmation-cancel'],
    preventDefault() {}, stopPropagation() {}
  };
  assert.equal(harness.context.handleAppDialogEnter(cancelEvent), true);
  assert.equal(await confirmation, false);
});

test('app dialogs are integrated with overlay dismissal, focus trapping, and beforeunload remains native', () => {
  const inventory = indexHtml.slice(
    indexHtml.indexOf('const MAIN_DISMISSIBLE_OVERLAY_IDS'),
    indexHtml.indexOf('const overlayOpenRecords')
  );
  assert.match(inventory, /'app-notification-overlay'/);
  assert.match(inventory, /'app-confirmation-overlay'/);
  const dismiss = functionSource('dismissOverlayById');
  assert.match(dismiss, /app-notification-overlay[\s\S]*dismissAppDialog/);
  assert.match(dismiss, /app-confirmation-overlay[\s\S]*dismissAppDialog/);
  const keydown = functionSource('handleGlobalKeydown');
  assert.match(keydown, /handleAppDialogEnter\(event\)/);
  assert.match(keydown, /trapOverlayFocus\(event\)/);
  assert.match(indexHtml, /beforeunload[\s\S]*hasPendingMutationWork\(\)[\s\S]*event\.returnValue = ''/);
});

test('no user-facing native alert or confirm calls remain', () => {
  assert.doesNotMatch(indexHtml, /(^|[^\w.])(?:alert|confirm)\s*\(/m);
});

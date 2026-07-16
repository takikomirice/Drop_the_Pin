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

function classList(initial) {
  const values = new Set(initial || []);
  return {
    add(...items) { items.forEach((item) => values.add(item)); },
    remove(...items) { items.forEach((item) => values.delete(item)); },
    contains(item) { return values.has(item); }
  };
}

function createHarness() {
  const elements = {};
  const documentApi = {
    activeElement: null,
    getElementById(id) { return elements[id] || null; },
    createElement(tagName) { return createElement({ tagName }); }
  };

  function createElement(options) {
    const config = options || {};
    const attributes = {};
    const children = [];
    const node = {
      id: config.id || '',
      tagName: String(config.tagName || 'DIV').toUpperCase(),
      value: config.value || '',
      hidden: false,
      style: {},
      dataset: {},
      className: '',
      classList: classList(config.classes),
      parentElement: null,
      children,
      focusCount: 0,
      scrollCount: 0,
      textWriteCount: 0,
      _textContent: '',
      get firstChild() { return children[0] || null; },
      get textContent() { return this._textContent; },
      set textContent(value) {
        this._textContent = String(value);
        this.textWriteCount += 1;
      },
      setAttribute(name, value) {
        attributes[name] = String(value);
        if (name === 'id') {
          this.id = String(value);
          elements[this.id] = this;
        }
      },
      getAttribute(name) {
        if (name === 'id') return this.id || null;
        return Object.prototype.hasOwnProperty.call(attributes, name) ? attributes[name] : null;
      },
      removeAttribute(name) { delete attributes[name]; },
      appendChild(child) {
        if (child.parentElement) child.remove();
        child.parentElement = this;
        children.push(child);
        return child;
      },
      insertBefore(child, reference) {
        if (child.parentElement) child.remove();
        child.parentElement = this;
        const index = reference ? children.indexOf(reference) : -1;
        if (index === -1) children.push(child);
        else children.splice(index, 0, child);
        return child;
      },
      remove() {
        if (!this.parentElement) return;
        const index = this.parentElement.children.indexOf(this);
        if (index !== -1) this.parentElement.children.splice(index, 1);
        this.parentElement = null;
      },
      contains(target) {
        if (target === this) return true;
        return children.some((child) => child.contains(target));
      },
      querySelector(selector) {
        if (selector === '[data-dialog-validation-summary]') {
          return findNode(this, (candidate) => candidate.getAttribute('data-dialog-validation-summary') !== null);
        }
        if (selector === '.dialog-scroll-content') {
          return findNode(this, (candidate) => candidate.classList.contains('dialog-scroll-content'));
        }
        return null;
      },
      closest(selector) {
        if (selector.includes('.form-group') && this.formGroup) return this.formGroup;
        if (selector.includes('.dialog-scroll-content') && this.scrollRegion) return this.scrollRegion;
        let current = this;
        while (current) {
          if (selector === '.sheet-overlay' && current.classList.contains('sheet-overlay')) return current;
          current = current.parentElement;
        }
        return null;
      },
      focus() {
        this.focusCount += 1;
        documentApi.activeElement = this;
      },
      scrollIntoView() { this.scrollCount += 1; }
    };
    if (node.id) elements[node.id] = node;
    return node;
  }

  function findNode(root, predicate) {
    for (const child of root.children) {
      if (predicate(child)) return child;
      const nested = findNode(child, predicate);
      if (nested) return nested;
    }
    return null;
  }

  const overlay = createElement({ id: 'test-overlay', classes: ['sheet-overlay'] });
  const scrollRegion = createElement({ classes: ['dialog-scroll-content'] });
  overlay.appendChild(scrollRegion);
  const group = createElement({ classes: ['form-group'] });
  scrollRegion.appendChild(group);
  const field = createElement({ id: 'test-field', tagName: 'INPUT', value: '' });
  field.formGroup = group;
  field.scrollRegion = scrollRegion;
  field.setAttribute('aria-describedby', 'existing-help');
  group.appendChild(field);

  const context = {
    document: documentApi,
    Map,
    Set,
    dialogValidationStates: new Map(),
    dialogValidationFieldSequence: 0
  };
  [
    'resolveDialogValidationElement', 'getDialogValidationContent',
    'ensureDialogValidationSummary', 'ensureDialogValidationFieldId',
    'setDialogValidationDescription', 'clearDialogValidationField',
    'updateDialogValidationSummary', 'focusDialogValidationField',
    'showDialogValidation', 'clearDialogValidation', 'handleDialogValidationChange'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });

  return { context, document: documentApi, overlay, scrollRegion, group, field };
}

test('common dialog validation connects field errors through ARIA and focuses the first invalid field', () => {
  const harness = createHarness();
  const originalValue = harness.field.value;

  const valid = harness.context.showDialogValidation({
    container: harness.overlay,
    errors: [{
      field: harness.field,
      message: '入力してください。',
      validate() { return harness.field.value.trim().length > 0; }
    }]
  });

  assert.equal(valid, false);
  assert.equal(harness.field.getAttribute('aria-invalid'), 'true');
  assert.equal(harness.field.getAttribute('aria-describedby'), 'existing-help test-field-validation-error');
  assert.equal(harness.field.focusCount, 1);
  assert.equal(harness.field.scrollCount, 1);
  assert.equal(harness.field.value, originalValue, 'validation must preserve the entered value');

  const error = harness.document.getElementById('test-field-validation-error');
  assert.equal(error.textContent, '入力してください。');
  assert.equal(error.getAttribute('role'), null, 'field errors must not create a second live announcement');

  const summary = harness.overlay.querySelector('[data-dialog-validation-summary]');
  assert.equal(summary.getAttribute('role'), 'alert');
  assert.equal(summary.getAttribute('aria-atomic'), 'true');
  assert.match(summary.textContent, /入力内容を確認してください。/);
});

test('resolved errors clear ARIA and display state without changing the input value', () => {
  const harness = createHarness();
  harness.context.showDialogValidation({
    container: harness.overlay,
    errors: [{
      field: harness.field,
      message: '入力してください。',
      validate() { return harness.field.value.trim().length > 0; }
    }]
  });

  harness.field.value = 'kept value';
  harness.context.handleDialogValidationChange({ target: harness.field });

  assert.equal(harness.field.getAttribute('aria-invalid'), null);
  assert.equal(harness.field.getAttribute('aria-describedby'), 'existing-help');
  assert.equal(harness.document.getElementById('test-field-validation-error').parentElement, null);
  assert.equal(harness.overlay.querySelector('[data-dialog-validation-summary]').hidden, true);
  assert.equal(harness.field.value, 'kept value');
});

test('submitting the same unresolved error does not mutate the alert summary twice', () => {
  const harness = createHarness();
  const options = {
    container: harness.overlay,
    focus: false,
    errors: [{ field: harness.field, message: '入力してください。', validate() { return false; } }]
  };
  harness.context.showDialogValidation(options);
  const summary = harness.overlay.querySelector('[data-dialog-validation-summary]');
  const writes = summary.textWriteCount;

  harness.context.showDialogValidation(options);

  assert.equal(summary.textWriteCount, writes);
});

test('target forms delegate existing rules to the common dialog validation helper', () => {
  assert.match(functionSource('validatePinDialogForm'), /showDialogValidation/);
  assert.match(functionSource('validateInputPresetEditor'), /showDialogValidation/);
  assert.match(functionSource('validateBulkMetadataForm'), /showDialogValidation/);
  assert.match(functionSource('validateRouteNameDialog'), /showDialogValidation/);
  assert.match(functionSource('showImportPreviewValidationError'), /showDialogValidation/);
  assert.match(functionSource('validateTrackImportPreviewForm'), /showDialogValidation/);
  assert.match(indexHtml, /share-overlay[\s\S]*?showDialogValidation/);
});

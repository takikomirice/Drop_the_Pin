const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionSource(name) {
  const start = indexHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

function createHarness() {
  const elements = Object.create(null);
  const openRecords = [];
  const closed = [];
  const queued = [];
  const updates = [];
  let routeSequence = 0;
  let renderCount = 0;

  const documentApi = {
    activeElement: null,
    getElementById(id) { return elements[id] || null; }
  };

  function element(id) {
    const node = {
      id,
      value: '',
      textContent: '',
      disabled: false,
      focusCount: 0,
      attributes: Object.create(null),
      classList: {
        values: new Set(),
        add(value) { this.values.add(value); },
        remove(value) { this.values.delete(value); },
        contains(value) { return this.values.has(value); }
      },
      setAttribute(name, value) { this.attributes[name] = String(value); },
      getAttribute(name) { return Object.hasOwn(this.attributes, name) ? this.attributes[name] : null; },
      removeAttribute(name) { delete this.attributes[name]; },
      focus() { documentApi.activeElement = this; this.focusCount += 1; }
    };
    elements[id] = node;
    return node;
  }

  [
    'route-name-overlay', 'route-name-title', 'route-name-help', 'route-name-input',
    'route-name-submit', 'route-name-cancel', 'route-name-form', 'route-edit-name-action'
  ].forEach(element);

  const state = {
    routeGroups: [{ id: 'route-existing', routeId: 'route-existing', name: '旧名称', pinIds: [] }],
    routeUiOpenGroupId: null,
    routeUiOpenKey: null,
    routeNameDialog: { mode: null, routeId: null, submitting: false }
  };

  const context = {
    state,
    document: documentApi,
    canEdit: () => true,
    canEditRouteControls: () => true,
    getRouteGroupById(routeId) {
      return state.routeGroups.find((group) => (group.id || group.routeId) === routeId) || null;
    },
    createClientRouteId() { routeSequence += 1; return `route-new-${routeSequence}`; },
    cloneRouteGroupForState(group) { return Object.assign({}, group, { pinIds: (group.pinIds || []).slice() }); },
    setRouteDockExpanded() {},
    renderRoutePanel() { renderCount += 1; },
    queueRouteGroupSave(routeId, snapshot) { queued.push([routeId, snapshot]); },
    updateRouteGroupOptimistic(routeId, patch) { updates.push([routeId, patch]); },
    clearDialogValidation() {
      elements['route-name-input'].removeAttribute('aria-invalid');
      elements['route-name-input'].setAttribute('aria-describedby', 'route-name-help');
    },
    showDialogValidation(options) {
      const errors = options.errors || [];
      if (!errors.length) {
        elements['route-name-input'].removeAttribute('aria-invalid');
        elements['route-name-input'].setAttribute('aria-describedby', 'route-name-help');
        return true;
      }
      const field = errors[0].field;
      field.setAttribute('aria-invalid', 'true');
      field.setAttribute('aria-describedby', 'route-name-help route-name-input-validation-error');
      field.focus();
      return false;
    },
    setActionButtonLabel(button, label) { button.textContent = label; },
    openOverlay(id) {
      openRecords.push({ id, opener: documentApi.activeElement });
      elements[id].classList.add('open');
      elements['route-name-input'].focus();
    },
    closeOverlay(id) {
      const index = openRecords.map((record) => record.id).lastIndexOf(id);
      const record = index === -1 ? null : openRecords.splice(index, 1)[0];
      elements[id].classList.remove('open');
      closed.push(id);
      if (record && record.opener) record.opener.focus();
    },
    Date,
    console
  };

  [
    'setRouteNameDialogBusy', 'validateRouteNameDialog', 'commitNewRouteGroup',
    'openRouteNameDialog', 'closeRouteNameDialog', 'submitRouteNameDialog',
    'createRouteGroup', 'editRouteGroupName'
  ].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(name)})`, context);
  });

  return {
    context, state, elements, document: documentApi, openRecords, closed, queued, updates,
    renderCount: () => renderCount
  };
}

test('route name form is one accessible app overlay and native prompts are absent', () => {
  assert.doesNotMatch(indexHtml, /(^|[^\w.])prompt\s*\(/m);
  assert.match(indexHtml, /id="route-name-overlay"[^>]*role="dialog"[^>]*aria-modal="true"[^>]*aria-labelledby="route-name-title"[^>]*aria-describedby="route-name-help"/);
  assert.match(indexHtml, /<form id="route-name-form"[^>]*novalidate/);
  assert.match(indexHtml, /id="route-name-input"[^>]*maxlength="100"[^>]*aria-describedby="route-name-help"[^>]*data-overlay-initial-focus/);
  assert.match(indexHtml, /data-dialog-validation-summary/);
  assert.match(indexHtml, /id="route-name-submit"[^>]*type="submit"/);
  assert.match(indexHtml, /id="route-name-cancel"[^>]*type="button"/);
});

test('new route approval preserves creation defaults and queue arguments', () => {
  const harness = createHarness();
  const opener = harness.elements['route-edit-name-action'];
  opener.focus();

  assert.equal(harness.context.createRouteGroup(), true);
  assert.equal(harness.state.routeNameDialog.mode, 'create');
  assert.equal(harness.elements['route-name-input'].value, '新しいルート');
  harness.elements['route-name-input'].value = ' 新規ルート ';

  assert.equal(harness.context.submitRouteNameDialog(), true);
  const created = harness.state.routeGroups[1];
  assert.equal(created.name, '新規ルート');
  assert.equal(created.visible, true);
  assert.equal(created.color, '#1e88e5');
  assert.equal(Array.from(created.pinIds).length, 0);
  assert.deepEqual(harness.queued, [[created.id, null]]);
  assert.equal(harness.state.routeUiOpenGroupId, created.id);
  assert.equal(harness.state.routeUiOpenKey, `pin:${created.id}`);
  assert.deepEqual(harness.closed, ['route-name-overlay']);
  assert.equal(harness.document.activeElement, opener);
});

test('new route cancellation and validation do not mutate route state', () => {
  const harness = createHarness();
  const before = harness.state.routeGroups.length;
  harness.context.createRouteGroup();
  harness.elements['route-name-input'].value = '   ';

  assert.equal(harness.context.submitRouteNameDialog(), false);
  assert.equal(harness.state.routeGroups.length, before);
  assert.equal(harness.elements['route-name-input'].getAttribute('aria-invalid'), 'true');
  assert.match(harness.elements['route-name-input'].getAttribute('aria-describedby'), /route-name-input-validation-error/);
  assert.equal(harness.document.activeElement, harness.elements['route-name-input']);
  assert.equal(harness.elements['route-name-overlay'].classList.contains('open'), true);

  assert.equal(harness.context.closeRouteNameDialog(), true);
  assert.equal(harness.state.routeGroups.length, before);
  assert.deepEqual(harness.queued, []);
});

test('existing route approval and cancellation preserve update arguments', () => {
  const harness = createHarness();
  assert.equal(harness.context.editRouteGroupName('route-existing'), true);
  assert.equal(harness.state.routeNameDialog.mode, 'edit');
  assert.equal(harness.elements['route-name-input'].value, '旧名称');
  harness.elements['route-name-input'].value = '新名称';
  assert.equal(harness.context.submitRouteNameDialog(), true);
  assert.equal(harness.updates.length, 1);
  assert.equal(harness.updates[0][0], 'route-existing');
  assert.equal(harness.updates[0][1].name, '新名称');

  harness.context.editRouteGroupName('route-existing');
  assert.equal(harness.context.closeRouteNameDialog(), true);
  assert.equal(harness.updates.length, 1);
});

test('busy guard prevents duplicate submission and dismissal', () => {
  const harness = createHarness();
  harness.context.createRouteGroup();
  harness.elements['route-name-input'].value = '二重防止';
  harness.state.routeNameDialog.submitting = true;
  harness.context.setRouteNameDialogBusy(true);

  assert.equal(harness.context.submitRouteNameDialog(), false);
  assert.equal(harness.context.closeRouteNameDialog(), false);
  assert.equal(harness.state.routeGroups.length, 1);
  assert.equal(harness.elements['route-name-submit'].disabled, true);
  assert.equal(harness.elements['route-name-cancel'].disabled, true);
});

test('Escape dispatcher and route detail rename use the shared nested overlay path', () => {
  const dismiss = functionSource('dismissOverlayById');
  const renderEdit = functionSource('renderRouteEditOverlay');
  assert.match(dismiss, /route-name-overlay[\s\S]*closeRouteNameDialog/);
  assert.match(renderEdit, /route-edit-name-action[\s\S]*editRouteGroupName\(routeId\)/);
  assert.match(indexHtml, /MAIN_DISMISSIBLE_OVERLAY_IDS[\s\S]*'route-name-overlay'/);
  assert.match(indexHtml, /BACKDROP_DISMISS_OVERLAY_IDS[\s\S]*'route-name-overlay'/);
  assert.match(indexHtml, /route-name-form[\s\S]*addEventListener\('submit'/);
});

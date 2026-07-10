const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

function assertIncludes(source, needle, message) {
  assert.ok(source.includes(needle), message || `Expected source to include ${needle}`);
}

function sourceFunctionBody(source, name) {
  const index = source.indexOf(`function ${name}`);
  assert.notEqual(index, -1, `Expected function ${name} to exist`);
  const openIndex = source.indexOf('{', index);
  assert.notEqual(openIndex, -1, `Expected function ${name} to have a body`);
  let depth = 0;
  for (let i = openIndex; i < source.length; i += 1) {
    const char = source[i];
    if (char === '{') depth += 1;
    if (char === '}') depth -= 1;
    if (depth === 0) return source.slice(openIndex + 1, i);
  }
  assert.fail(`Could not parse function ${name}`);
}

test('index reads the injected edit token and exposes withEditToken', () => {
  assertIncludes(indexHtml, "const editToken = String(window.__EDIT_TOKEN__ || '');");
  assertIncludes(indexHtml, 'const hasEditToken = !!editToken;');

  const body = sourceFunctionBody(indexHtml, 'withEditToken');
  assertIncludes(body, 'Object.assign({}, payload || {})');
  assertIncludes(body, 'next.__editToken = editToken;');
  assertIncludes(body, 'return next;');
});

test('canEdit and edit mode entry require an edit token', () => {
  const canEditBody = sourceFunctionBody(indexHtml, 'canEdit');
  assertIncludes(canEditBody, 'hasEditToken');
  assertIncludes(canEditBody, 'state.editMode === true');
  assertIncludes(canEditBody, 'state.shareMode !== true');

  const setEditModeBody = sourceFunctionBody(indexHtml, 'setEditMode');
  assertIncludes(setEditModeBody, 'if (next && !hasEditToken)');
  assertIncludes(setEditModeBody, 'showEditAccessNotice();');
  assertIncludes(setEditModeBody, 'return;');
});

test('readonly and invalid edit URL guidance exists without exposing editKey', () => {
  assertIncludes(indexHtml, 'id="readonly-banner"');
  assertIncludes(indexHtml, '閲覧専用　編集するにはconfigシートの EDIT_URL から開いてください');
  assertIncludes(indexHtml, '編集URLとして開かれていますが、編集キーが確認できません。configシートの EDIT_KEY とURLの editKey を確認してください。');
  assertIncludes(indexHtml, 'function isEditUrlRequested()');
  assert.equal(indexHtml.includes('editKey='), false);
});

test('edit-only controls are hidden unless an edit token is available', () => {
  [
    'body:not(.has-edit-token) #edit-toggle',
    'body:not(.has-edit-token) #settings-toggle',
    'body:not(.has-edit-token) #share-open-btn',
    'body:not(.has-edit-token) #fab',
    'body:not(.has-edit-token) #bulk-action-bar',
    'body:not(.has-edit-token) #route-add-btn',
    'body:not(.has-edit-token) #pin-detail-drive',
    'body:not(.has-edit-token) #pin-detail-route-add'
  ].forEach((needle) => assertIncludes(indexHtml, needle));
  assertIncludes(indexHtml, "document.body.classList.toggle('has-edit-token', hasEditToken);");
});

test('shared view remains read-only and independent from edit token injection', () => {
  assertIncludes(sharedHtml, '共有ビューは読み取り専用。編集・追加・削除・D&D・ルート編集は持たない。');
  assertIncludes(sharedHtml, "withGAS('getSharedViewData', state.token)");
  assert.equal(sharedHtml.includes('__EDIT_TOKEN__'), false);
  assert.equal(sharedHtml.includes('withEditToken'), false);
});

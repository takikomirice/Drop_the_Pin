const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(__dirname, '..', 'shared.html'), 'utf8');

function functionSource(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const bodyStart = source.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < source.length; index += 1) {
    const character = source[index];
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
      if (depth === 0) return source.slice(start, index + 1);
    }
  }
  throw new Error(`Unterminated function ${name}`);
}

function body(name) {
  return functionSource(indexHtml, name);
}

function selectOptions(id) {
  const match = indexHtml.match(new RegExp(`<select id="${id}"[^>]*>([\\s\\S]*?)<\\/select>`));
  assert.ok(match, `Expected select #${id}`);
  return Array.from(match[1].matchAll(/<option value="([^"]+)">([^<]+)<\/option>/g), (option) => (
    [option[1], option[2]]
  ));
}

function classList() {
  const values = new Set();
  return {
    toggle(name, force) {
      if (force) values.add(name);
      else values.delete(name);
    },
    contains(name) { return values.has(name); }
  };
}

test('settings exposes an edit-only input preset manager and a complete editor overlay', () => {
  assert.match(indexHtml, /id="input-presets-manage"[^>]*>入力プリセットを管理</);
  assert.match(indexHtml, /id="input-presets-overlay"[^>]*role="dialog"/);
  [
    'input-presets-list', 'input-preset-add', 'input-presets-error', 'input-presets-close',
    'input-preset-editor', 'input-preset-name', 'input-preset-enabled',
    'input-preset-tags-mode', 'input-preset-tags', 'input-preset-color-mode',
    'input-preset-color-palette', 'input-preset-icon-mode', 'input-preset-icon-picker',
    'input-preset-status-mode', 'input-preset-status', 'input-preset-save', 'input-preset-cancel'
  ].forEach((id) => assert.match(indexHtml, new RegExp(`id="${id}"`), id));
  assert.match(indexHtml, /id="input-preset-tags-mode"[\s\S]*value="keep"[\s\S]*value="set"[\s\S]*value="clear"/);
  assert.match(indexHtml, /id="input-preset-color-mode"[\s\S]*value="keep"[\s\S]*value="set"/);
  assert.match(indexHtml, /id="input-preset-icon-mode"[\s\S]*value="keep"[\s\S]*value="set"/);
  assert.deepEqual(selectOptions('input-preset-tags-mode'), [
    ['keep', '変更しない'], ['set', '設定する'], ['clear', '空にする']
  ]);
  assert.deepEqual(selectOptions('input-preset-color-mode'), [
    ['keep', '変更しない'], ['set', '設定する']
  ]);
  assert.deepEqual(selectOptions('input-preset-icon-mode'), [
    ['keep', '変更しない'], ['set', '設定する']
  ]);
  assert.deepEqual(selectOptions('input-preset-status-mode'), [
    ['keep', '変更しない'], ['set', '設定する'], ['clear', '空にする']
  ]);
  for (const [id, label] of [
    ['input-preset-tags-mode', 'タグ'],
    ['input-preset-color-mode', '色'],
    ['input-preset-icon-mode', 'アイコン'],
    ['input-preset-status-mode', '状態']
  ]) {
    assert.match(indexHtml, new RegExp(`<label[^>]*for="${id}"[^>]*>${label}<\\/label>`));
  }
  assert.doesNotMatch(indexHtml, /(?:タグ|色|アイコン|状態)の動作/);
});

test('the manager opens and loads only while editable with an edit token', () => {
  const source = body('openInputPresetManager');
  assert.match(source, /if \(!canEdit\(\)\) return/);
  assert.match(source, /reloadInputPresets\(\)/);
  assert.match(body('loadInputPresetCatalog'), /withGAS\('listInputPresets', withEditToken\(\{\}\)\)/);
  assert.match(source, /openOverlay\('input-presets-overlay'\)/);
  assert.match(body('updateInputPresetManagerButton'), /disabled = !canEdit\(\)/);
});

test('preset list builds user content with textContent and exposes all management actions', () => {
  const source = body('renderInputPresetList');
  assert.match(source, /nameElement\.textContent = preset\.name/);
  assert.match(source, /summaryElement\.textContent =/);
  assert.match(source, /editInputPreset/);
  assert.match(source, /duplicateInputPreset/);
  assert.match(source, /deleteInputPresetFromUi/);
  assert.match(source, /toggleInputPresetEnabled/);
  assert.match(source, /moveInputPreset/);
  for (const label of [
    'プリセットを編集', 'プリセットを複製', 'プリセットを有効にする',
    'プリセットを無効にする', 'プリセットを上へ移動',
    'プリセットを下へ移動', 'プリセットを削除'
  ]) assert.ok(source.includes(`'${label}'`), label);
  assert.match(source, /dragHandle\.title = '並べ替え'/);
  assert.match(source, /dragHandle\.setAttribute\('aria-label', '並べ替え'\)/);
  assert.doesNotMatch(source, /preset\.name[^;\n]*innerHTML|preset\.tags[^;\n]*innerHTML/);
});

test('only the enable toggle has visible text while every preset action stays labelled and operable', () => {
  function element() {
    return {
      attributes: {}, children: [], dataset: {}, style: {},
      appendChild(child) { this.children.push(child); return child; },
      removeChild(child) { this.children.splice(this.children.indexOf(child), 1); },
      get firstChild() { return this.children[0] || null; },
      setAttribute(name, value) { this.attributes[name] = String(value); },
      addEventListener() {}
    };
  }
  const elements = {
    'input-presets-list': element(),
    'input-presets-count': element(),
    'input-presets-empty': element(),
    'input-preset-add': element(),
    'input-presets-close': element()
  };
  const context = {
    document: {
      createElement: () => element(),
      getElementById: (id) => elements[id]
    },
    state: { inputPresets: { items: [
      { presetId: 'enabled', name: '有効', enabled: true },
      { presetId: 'disabled', name: '無効', enabled: false }
    ], loading: false } },
    destroyInputPresetSortable() {},
    isInputPresetManagerBusy: () => false,
    hasLoadedInputPresets: () => true,
    canEdit: () => true,
    getInputPresetSummary: () => '',
    toggleInputPresetEnabled() {}, editInputPreset() {}, duplicateInputPreset() {},
    deleteInputPresetFromUi() {}, moveInputPreset() {}, attachInputPresetSortable() {}
  };
  vm.runInNewContext([
    body('createInputPresetActionButton'),
    body('renderInputPresetList'),
    'this.render = renderInputPresetList;'
  ].join('\n'), context);
  context.render();

  const rows = elements['input-presets-list'].children;
  assert.equal(rows.length, 2);
  rows.forEach((row, rowIndex) => {
    const actions = row.children[0].children[2].children;
    assert.equal(actions.length, 6);
    assert.equal(actions[0].children[0].textContent, rowIndex === 0 ? '無効にする' : '有効にする');
    actions.slice(1).forEach((button) => assert.equal(button.children.length, 0));
    actions.forEach((button) => {
      assert.ok(button.title);
      assert.ok(button.attributes['aria-label']);
    });
  });
  assert.match(indexHtml, /\.input-preset-actions button\s*\{[^}]*min-width:\s*44px;[^}]*min-height:\s*44px;/);
});

test('duplicate copies values but never identity or timestamps', () => {
  const cloneSource = body('cloneInputPreset');
  const duplicateSource = body('createInputPresetDuplicate');
  const context = {};
  vm.runInNewContext(`${cloneSource}\n${duplicateSource}\nthis.copy = createInputPresetDuplicate({
    presetId: 'original-id', name: '植物', enabled: false, orderIndex: 3,
    tagsMode: 'set', tags: ['植物'], colorMode: 'set', color: '#4caf50',
    iconMode: 'set', icon: 'nature', statusMode: 'clear', status: null,
    createdAt: 'created', updatedAt: 'updated'
  });`, context);
  assert.equal(context.copy.name, '植物 のコピー');
  assert.equal(Object.hasOwn(context.copy, 'presetId'), false);
  assert.equal(Object.hasOwn(context.copy, 'createdAt'), false);
  assert.equal(Object.hasOwn(context.copy, 'updatedAt'), false);
  assert.deepEqual(Array.from(context.copy.tags), ['植物']);
});

test('new presets start after the greatest current order and tag input follows existing normalization', () => {
  const context = {
    state: { inputPresets: { items: [{ orderIndex: 10 }, { orderIndex: 20 }] } },
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'default' }],
    PIN_STATUSES: ['未対応']
  };
  vm.runInNewContext([
    body('normalizeTags'),
    body('normalizeInputPresetTags'),
    body('createBlankInputPreset'),
    "this.blank = createBlankInputPreset();",
    "this.tags = normalizeInputPresetTags(' #植物, 観察、植物 ');"
  ].join('\n'), context);
  assert.equal(context.blank.orderIndex, 21);
  assert.deepEqual(Array.from(context.tags), ['植物', '観察']);
});

test('editor mode controls reject no-op and reuse existing color and icon definitions', () => {
  const controls = body('updateInputPresetEditorControls');
  assert.match(controls, /inputPresetHasAction/);
  assert.match(controls, /saveButton\.disabled =/);
  for (const field of ['tags', 'color', 'icon', 'status']) {
    assert.match(controls, new RegExp(`input-preset-${field}-field`));
  }
  assert.match(controls, /const visible = modeField\.value === 'set'/);
  assert.match(controls, /field\.hidden = !visible/);

  const editor = body('renderInputPresetEditor');
  assert.match(editor, /paletteSelect\('input-preset-color-palette'/);
  assert.match(editor, /iconPickerSelect\('input-preset-icon-picker'/);
  assert.doesNotMatch(indexHtml, /INPUT_PRESET_(?:COLOR|ICON)_LABELS/);
  const palette = body('renderColorPaletteButtons');
  const picker = body('renderIconPickerButtons');
  assert.match(palette, /button\.title = color\.label/);
  assert.match(palette, /setAttribute\('aria-label', color\.label\)/);
  assert.match(picker, /button\.title = icon\.label/);
  assert.match(picker, /setAttribute\('aria-label', icon\.label\)/);
});

test('dependent editor fields are visible and enabled only for set mode and restore focus safely', () => {
  const elements = {};
  function element(id, value = '') {
    const current = {
      id, value, checked: true, disabled: false, hidden: false, classList: classList(),
      buttons: [], contains(node) { return node === current || current.buttons.includes(node); },
      querySelectorAll(selector) { return selector === 'button' ? current.buttons : []; },
      focus() { documentApi.activeElement = current; }
    };
    elements[id] = current;
    return current;
  }
  element('input-preset-name', 'Preset');
  element('input-preset-enabled');
  const tagsMode = element('input-preset-tags-mode', 'keep');
  const colorMode = element('input-preset-color-mode', 'set');
  const iconMode = element('input-preset-icon-mode', 'keep');
  const statusMode = element('input-preset-status-mode', 'clear');
  const tagsField = element('input-preset-tags-field');
  const colorField = element('input-preset-color-field');
  const iconField = element('input-preset-icon-field');
  const statusField = element('input-preset-status-field');
  const tags = element('input-preset-tags');
  const status = element('input-preset-status', '完了');
  const colorPalette = element('input-preset-color-palette');
  const iconPicker = element('input-preset-icon-picker');
  const colorButton = { disabled: false };
  const iconButton = { disabled: false };
  tagsField.buttons.push(tags);
  colorPalette.buttons.push(colorButton);
  iconPicker.buttons.push(iconButton);
  element('input-preset-save');
  element('input-preset-cancel');
  const documentApi = {
    activeElement: tags,
    getElementById(id) { return elements[id]; }
  };
  const context = {
    state: { inputPresets: { editing: {}, saving: false } },
    document: documentApi,
    hasLoadedInputPresets: () => true,
    inputPresetHasAction: () => true,
    setActionButtonLabel() {}
  };
  vm.runInNewContext(`${body('updateInputPresetEditorControls')}\nthis.update = updateInputPresetEditorControls;`, context);
  context.update();
  assert.equal(tagsField.hidden, true);
  assert.equal(tags.disabled, true);
  assert.equal(documentApi.activeElement, tagsMode);
  assert.equal(colorField.hidden, false);
  assert.equal(colorButton.disabled, false);
  assert.equal(iconField.hidden, true);
  assert.equal(iconButton.disabled, true);
  assert.equal(statusField.hidden, true);
  assert.equal(status.disabled, true);

  tagsMode.value = 'set';
  colorMode.value = 'keep';
  iconMode.value = 'set';
  statusMode.value = 'set';
  context.update();
  assert.equal(tagsField.hidden, false);
  assert.equal(tags.disabled, false);
  assert.equal(colorField.hidden, true);
  assert.equal(colorButton.disabled, true);
  assert.equal(iconField.hidden, false);
  assert.equal(iconButton.disabled, false);
  assert.equal(statusField.hidden, false);
  assert.equal(status.disabled, false);
});

test('save payload omits values unused by keep or clear modes', () => {
  const values = {
    'input-preset-name': { value: 'Preset' },
    'input-preset-enabled': { checked: true },
    'input-preset-tags-mode': { value: 'clear' },
    'input-preset-tags': { value: '#stored' },
    'input-preset-color-mode': { value: 'keep' },
    'input-preset-icon-mode': { value: 'set' },
    'input-preset-status-mode': { value: 'clear' },
    'input-preset-status': { value: '完了' }
  };
  const context = {
    state: { inputPresets: { editing: {
      presetId: 'preset-1', orderIndex: 2, color: '#e53935', icon: 'photo'
    } } },
    document: { getElementById(id) { return values[id]; } }
  };
  vm.runInNewContext([
    body('normalizeTags'),
    body('normalizeInputPresetTags'),
    body('inputPresetHasAction'),
    body('compactInputPresetPayload'),
    body('buildInputPresetPayload'),
    'this.payload = buildInputPresetPayload();'
  ].join('\n'), context);
  const payload = JSON.parse(JSON.stringify(context.payload));
  assert.equal(payload.icon, 'photo');
  assert.equal(Object.hasOwn(payload, 'tags'), false);
  assert.equal(Object.hasOwn(payload, 'color'), false);
  assert.equal(Object.hasOwn(payload, 'status'), false);
  assert.deepEqual(payload, {
    presetId: 'preset-1', name: 'Preset', enabled: true, orderIndex: 2,
    tagsMode: 'clear', colorMode: 'keep', iconMode: 'set', icon: 'photo', statusMode: 'clear'
  });
  assert.match(body('toggleInputPresetEnabled'), /compactInputPresetPayload\(optimistic\)/);
});

test('new edit and duplicate paths share the same dependency visibility update', () => {
  assert.match(body('renderInputPresetEditor'), /updateInputPresetEditorControls\(\)/);
  assert.match(body('addInputPreset'), /renderInputPresetEditor\(\)/);
  assert.match(body('editInputPreset'), /renderInputPresetEditor\(\)/);
  assert.match(body('duplicateInputPreset'), /renderInputPresetEditor\(\)/);
  for (const id of ['tags', 'color', 'icon', 'status']) {
    assert.match(indexHtml, new RegExp(`id="input-preset-${id}-field"[^>]*hidden`));
  }
  assert.match(indexHtml, /id="input-preset-tags"[^>]*disabled/);
  assert.match(indexHtml, /id="input-preset-status"[^>]*disabled/);
});

test('save, delete, enable toggle, and reorder use the four preset GAS APIs', () => {
  assert.match(body('saveInputPresetEditor'), /withGAS\('saveInputPreset', withEditToken\(payload\)\)/);
  assert.match(body('deleteInputPresetFromUi'), /withGAS\('deleteInputPreset', withEditToken\(\{ presetId:/);
  assert.match(body('toggleInputPresetEnabled'), /withGAS\('saveInputPreset', withEditToken\(payload\)\)/);
  assert.match(body('persistInputPresetOrder'), /withGAS\('updateInputPresetOrder', withEditToken\(\{ presetIds:/);
});

test('reordering supports Sortable and buttons and rolls back only on failure', () => {
  const sortable = body('attachInputPresetSortable');
  assert.match(sortable, /window\.Sortable/);
  assert.match(sortable, /handle: '.input-preset-drag-handle'/);
  assert.match(sortable, /persistInputPresetOrder/);
  const persist = body('persistInputPresetOrder');
  assert.match(persist, /previousItems/);
  assert.match(persist, /state\.inputPresets\.items = previousItems/);
  assert.match(persist, /result\.presets/);
  assert.match(body('renderInputPresetList'), /上へ/);
  assert.match(body('renderInputPresetList'), /下へ/);
});

test('busy state blocks duplicate operations and closing while errors stay text-safe', () => {
  ['saveInputPresetEditor', 'deleteInputPresetFromUi', 'toggleInputPresetEnabled', 'persistInputPresetOrder'].forEach((name) => {
    assert.match(body(name), /state\.inputPresets\.saving|isInputPresetManagerBusy\(\)/, name);
  });
  assert.match(body('closeInputPresetManager'), /isInputPresetManagerBusy\(\)/);
  assert.match(body('openInputPresetManager'), /isInputPresetManagerBusy\(\)/);
  assert.match(body('cancelInputPresetEditor'), /clearInputPresetError\(\)/);
  const errorSource = body('setInputPresetError');
  assert.match(errorSource, /\.textContent =/);
  assert.doesNotMatch(errorSource, /innerHTML|alert\(/);
  assert.match(body('closeOverlayFromBackdrop'), /dismissOverlayById\(record\.id\)/);
  assert.match(body('dismissOverlayById'), /input-presets-overlay[\s\S]*closeInputPresetManager/);
});

test('management never calls preset APIs in shared view or changes the single-add draft', () => {
  ['listInputPresets', 'saveInputPreset', 'deleteInputPreset', 'updateInputPresetOrder'].forEach((method) => {
    assert.equal(sharedHtml.includes(method), false, method);
  });
  const managerFunctions = [
    'openInputPresetManager', 'closeInputPresetManager', 'renderInputPresetList',
    'editInputPreset', 'duplicateInputPreset', 'saveInputPresetEditor',
    'deleteInputPresetFromUi', 'toggleInputPresetEnabled', 'persistInputPresetOrder'
  ].map(body).join('\n');
  assert.doesNotMatch(managerFunctions, /state\.upload|upload-(?:title|tags|status|color|icon)/);
});

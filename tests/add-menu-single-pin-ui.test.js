const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

function functionBody(name) {
  const start = indexHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = open; index < indexHtml.length; index += 1) {
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
    if (character === '}' && --depth === 0) return indexHtml.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${name}`);
}

test('pin header add button opens the edit-only menu containing exactly the five approved actions', () => {
  assert.equal(countId('fab'), 0);
  assert.equal(countId('pin-add-btn'), 1);
  const pinHeader = indexHtml.match(/<div class="dock-pin-header">([\s\S]*?)<\/div>\s*<div class="dock-pin-tabs"/);
  assert.ok(pinHeader, 'pin header must exist');
  assert.match(pinHeader[1], /id="pin-count"/);
  assert.match(pinHeader[1], /<button id="pin-add-btn"[^>]*aria-label="ピンを追加"[^>]*title="ピンを追加"[^>]*>[\s\S]*?action-icon-add[\s\S]*?<\/button>/);
  assert.equal((pinHeader[1].match(/<button\b/g) || []).length, 1);
  assert.equal(countId('pin-panel-toggle'), 0);
  assert.doesNotMatch(indexHtml, /pin-panel-toggle/);
  assert.equal(countId('add-menu-overlay'), 1);

  const menu = indexHtml.match(/<div id="add-menu-overlay"[\s\S]*?<div class="add-menu-actions"[^>]*>([\s\S]*?)<\/div>\s*<div id="multi-photo-conflict-message"/);
  assert.ok(menu, 'add menu action list must exist');
  const actions = menu[1].match(/<button\b[^>]*data-add-action=/g) || [];
  assert.equal(actions.length, 5);
  [
    'ピンを追加',
    '端末から写真を取込',
    'Driveから写真を取込',
    '端末から音声を取込',
    'Driveから音声を取込'
  ].forEach((label) => assert.ok(menu[1].includes(label), `${label} must be in the add menu`));
  assert.doesNotMatch(menu[1], /CSV|GeoJSON|GPX/);

  assert.match(indexHtml, /getElementById\('pin-add-btn'\)\.addEventListener\('click',\s*openAddMenu\)/);
  assert.match(indexHtml, /getElementById\('panel-toggle'\)\.addEventListener\('click',[\s\S]*?renderPanelToggle\(\)/);
  assert.doesNotMatch(indexHtml, /#fab|body[^\n{]*#fab|getElementById\('fab'\)/);
  assert.match(indexHtml, /#pin-add-btn\s*\{[^}]*min-width:\s*44px[^}]*min-height:\s*44px/);
  assert.match(indexHtml, /#pin-add-btn\s*\{[^}]*width:\s*44px[^}]*height:\s*44px/);
  assert.match(indexHtml, /\.pin-panel-actions::after\s*\{[^}]*content:\s*['"]['"][^}]*width:\s*44px[^}]*height:\s*44px[^}]*flex:\s*0 0 44px/);
  assert.match(indexHtml, /body\.narrow-view \.pin-panel-actions::after\s*\{\s*display:\s*none/);
  assert.match(indexHtml, /body:not\(\.edit-mode\) #pin-add-btn/);
  assert.doesNotMatch(indexHtml, /body\.narrow-view \.dock-pin-header[\s\S]*?display:\s*none/);
  assert.match(functionBody('startSinglePinFromAddMenu'), /openUploadModal\(\)/);
  assert.doesNotMatch(functionBody('prepareAddMenuPhotoImport'), /resetUploadState\(\)/);
  assert.match(functionBody('prepareAddMenuPhotoImport'), /startAction\(\)/);
});

test('busy and permission guards prevent every add action from starting', () => {
  const guard = functionBody('canStartAddAction');
  assert.match(guard, /canEdit\(\)/);
  assert.match(guard, /isProductionImportBusy\(\)/);
  assert.match(guard, /state\.upload\.saving/);
  assert.match(guard, /state\.addMenuPreparing/);

  assert.match(functionBody('openAddMenu'), /if \(!canEdit\(\)\)/);
  assert.match(functionBody('startSinglePinFromAddMenu'), /if \(!canStartAddAction\(\)\)/);
  assert.match(functionBody('prepareAddMenuPhotoImport'), /if \(!canStartAddAction\(\)\)/);
  assert.match(functionBody('handleMultiPhotoButtonClick'), /prepareAddMenuPhotoImport/);
  assert.match(functionBody('handleDrivePhotoImportButtonClick'), /prepareAddMenuPhotoImport/);

  assert.match(indexHtml, /body:not\(\.has-edit-token\) #add-menu-overlay/);
  assert.match(indexHtml, /body:not\(\.edit-mode\) #add-menu-overlay/);
});

test('single-pin overlay keeps legacy IDs in the Pencil 07 form structure', () => {
  const requiredIds = [
    'upload-overlay', 'upload-preset-select', 'upload-preset-apply',
    'upload-title', 'file-input', 'file-drop', 'upload-preview', 'clear-photo-btn',
    'upload-desc', 'upload-event-at', 'upload-tags', 'upload-color-palette',
    'upload-icon-picker', 'upload-status', 'upload-links',
    'upload-position-photo', 'upload-position-map', 'upload-save-unplaced',
    'upload-submit', 'upload-cancel'
  ];
  requiredIds.forEach((id) => assert.equal(countId(id), 1, `${id} must exist exactly once`));

  const overlay = indexHtml.match(/<div id="upload-overlay"[\s\S]*?<div id="drive-photo-import-overlay"/);
  assert.ok(overlay);
  ['入力プリセット', 'タイトル', '写真（任意）', '説明', '日時', 'タグ', '色', 'アイコン', '状態', 'URL', '位置']
    .forEach((label) => assert.ok(overlay[0].includes(label), `${label} must be visible in the form`));
  assert.match(overlay[0], /id="upload-submit"[^>]*>[\s\S]*?action-icon-save[\s\S]*?<span class="action-label">保存<\/span><\/button>/);
  assert.match(overlay[0], /id="upload-cancel"[^>]*>キャンセル<\/button>/);
  assert.match(indexHtml, /\.single-pin-sheet\s*\{[\s\S]*?width:\s*min\(calc\(100vw - 28px\),\s*996px\)/);
  assert.match(indexHtml, /\.single-pin-main-grid\s*\{[\s\S]*?display:\s*grid/);
});

test('preset selection remains inert until Apply and only changes four metadata fields', () => {
  assert.match(indexHtml, /getElementById\('upload-preset-select'\)\.addEventListener\('change',\s*renderUploadInputPresetControls\)/);
  const apply = functionBody('applyInputPresetToUpload');
  assert.match(apply, /InputPresetApplyCore\.apply\(\{[\s\S]*?tags:[\s\S]*?color:[\s\S]*?icon:[\s\S]*?status:/);
  assert.match(apply, /setUploadColor\(applied\.color\)/);
  assert.match(apply, /setUploadIcon\(applied\.icon\)/);
  assert.doesNotMatch(apply, /upload-title|upload-desc|upload-event-at|upload-links/);
  assert.match(functionBody('renderUploadInputPresetControls'), /preset\.enabled === true/);
});

test('position controls disable photo location without usable GPS and protect explicit choices', () => {
  const usable = functionBody('hasUsableUploadGps');
  assert.match(usable, /state\.upload\.uploadFile/);
  assert.match(usable, /metadataStatus === 'success'/);
  assert.match(usable, /Number\.isFinite\(state\.upload\.lat\)/);
  assert.match(usable, /Number\.isFinite\(state\.upload\.lng\)/);

  const refresh = functionBody('refreshUploadPositionControls');
  assert.match(refresh, /photoButton\.disabled = !hasUsableUploadGps\(\)/);
  assert.match(refresh, /aria-pressed/);

  const select = functionBody('setUploadPositionMode');
  assert.match(select, /mode === 'photo' && !hasUsableUploadGps\(\)/);
  assert.match(select, /positionModeManual/);

  const submit = functionBody('handleUploadSubmit');
  assert.match(submit, /positionMode === 'unplaced'/);
  assert.match(submit, /saveNewPin\(draft, null\)/);
  assert.match(submit, /positionMode === 'map'/);
  assert.match(submit, /beginPlacement\(\{ mode: 'new', draft: draft \}\)/);
  assert.match(submit, /hasUsableUploadGps\(\)/);

  const photoHandler = functionBody('handleUploadPhotoSelected');
  assert.match(photoHandler, /!state\.upload\.positionModeManual && hasUsableUploadGps\(\)/);
  assert.match(photoHandler, /setUploadPositionMode\('photo',\s*\{ manual: false \}\)/);
  assert.match(functionBody('setupUploadPhotoPicker'), /fileInput\.addEventListener\('change',\s*handleUploadPhotoSelected\)/);
});

test('position selection runtime keeps photo, map, and unplaced modes explicit', () => {
  const attributes = {};
  const elements = {
    'upload-position-photo': { disabled: false, setAttribute: (name, value) => { attributes[`photo:${name}`] = value; } },
    'upload-position-map': { disabled: false, setAttribute: (name, value) => { attributes[`map:${name}`] = value; } },
    'upload-save-unplaced': { disabled: false, setAttribute: (name, value) => { attributes[`unplaced:${name}`] = value; } },
    'upload-position-note': { textContent: '' }
  };
  const context = {
    state: {
      upload: {
        uploadFile: null, metadataStatus: 'idle', lat: null, lng: null,
        positionMode: 'map', positionModeManual: false,
        saving: false, converting: false, conversionError: ''
      }
    },
    document: { getElementById: (id) => elements[id] },
    canEdit: () => true,
    Number
  };
  vm.runInNewContext(`
    function hasUsableUploadGps() {${functionBody('hasUsableUploadGps')}}
    function refreshUploadPositionControls() {${functionBody('refreshUploadPositionControls')}}
    function setUploadPositionMode(mode, options) {${functionBody('setUploadPositionMode')}}
    this.api = { hasUsableUploadGps, refreshUploadPositionControls, setUploadPositionMode };
  `, context);

  context.api.refreshUploadPositionControls();
  assert.equal(elements['upload-position-photo'].disabled, true);
  assert.equal(context.api.setUploadPositionMode('photo'), false);
  assert.equal(context.state.upload.positionMode, 'map');

  Object.assign(context.state.upload, {
    uploadFile: { name: 'gps.jpg' }, metadataStatus: 'success', lat: 35, lng: 139
  });
  context.api.refreshUploadPositionControls();
  assert.equal(elements['upload-position-photo'].disabled, false);
  assert.equal(context.api.setUploadPositionMode('photo'), true);
  assert.equal(context.state.upload.positionMode, 'photo');
  assert.equal(context.state.upload.positionModeManual, true);
  assert.equal(attributes['photo:aria-pressed'], 'true');

  assert.equal(context.api.setUploadPositionMode('unplaced'), true);
  Object.assign(context.state.upload, { lat: 36, lng: 140 });
  context.api.refreshUploadPositionControls();
  assert.equal(context.state.upload.positionMode, 'unplaced');
  assert.equal(attributes['unplaced:aria-pressed'], 'true');
  assert.match(elements['upload-position-note'].textContent, /未配置/);
});

test('cancel and backdrop preserve File/Object URL cleanup and visible save states', () => {
  const clear = functionBody('clearUploadPhotoState');
  assert.match(clear, /URL\.revokeObjectURL\(state\.upload\.previewUrl\)/);
  assert.match(clear, /state\.upload\.originalFile = null/);
  assert.match(clear, /state\.upload\.uploadFile = null/);

  const cancel = functionBody('cancelUpload');
  assert.match(cancel, /if \(state\.upload\.saving\) return/);
  assert.match(cancel, /clearUploadPhotoState\(\)/);
  assert.match(cancel, /closeOverlay\('upload-overlay'\)/);
  assert.match(functionBody('closeOverlayFromBackdrop'), /dismissOverlayById\(record\.id\)/);
  assert.match(functionBody('dismissOverlayById'), /id === 'upload-overlay'[\s\S]*cancelUpload\(\)/);

  assert.equal(countId('upload-form-status'), 1);
  const formState = functionBody('refreshUploadFormState');
  assert.match(formState, /保存しています/);
  assert.match(formState, /state\.upload\.saveError/);
  const save = functionBody('saveNewPin');
  assert.match(save, /state\.upload\.saving = true/);
  assert.match(save, /state\.upload\.saveError/);
  assert.match(save, /finally/);
});

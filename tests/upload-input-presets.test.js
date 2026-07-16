const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function functionSource(name) {
  const functionStart = indexHtml.indexOf(`function ${name}(`);
  assert.notEqual(functionStart, -1, `Expected function ${name}`);
  const start = indexHtml.slice(Math.max(0, functionStart - 6), functionStart) === 'async '
    ? functionStart - 6
    : functionStart;
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
    if (character === '}' && --depth === 0) return indexHtml.slice(start, index + 1);
  }
  throw new Error(`Unterminated function ${name}`);
}

function applyCoreSource() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const ImportJobCore = (function() {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  return indexHtml.slice(start, end);
}

class FakeClassList {
  constructor(initial = []) { this.values = new Set(initial); }
  add(value) { this.values.add(value); }
  remove(value) { this.values.delete(value); }
  contains(value) { return this.values.has(value); }
}

class FakeElement {
  constructor(id) {
    this.id = id;
    this.value = '';
    this.textContent = '';
    this.disabled = false;
    this.style = {};
    this.children = [];
    this.listeners = {};
    this.classList = new FakeClassList();
  }
  addEventListener(type, listener) {
    if (!this.listeners[type]) this.listeners[type] = [];
    this.listeners[type].push(listener);
  }
  dispatch(type) {
    (this.listeners[type] || []).forEach((listener) => listener({ target: this }));
  }
  replaceChildren(...children) { this.children = children; }
  appendChild(child) { this.children.push(child); return child; }
  setAttribute(name, value) { this[name] = String(value); }
}

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function settle() {
  return new Promise((resolve) => setImmediate(resolve));
}

function preset(id, overrides = {}) {
  return {
    presetId: id,
    name: `Preset ${id}`,
    enabled: true,
    orderIndex: 0,
    tagsMode: 'keep', tags: [],
    colorMode: 'keep', color: null,
    iconMode: 'keep', icon: null,
    statusMode: 'keep', status: null,
    ...overrides
  };
}

function createHarness() {
  const ids = [
    'upload-overlay', 'upload-preset-select', 'upload-preset-apply',
    'upload-preset-status', 'upload-preset-error', 'upload-preset-retry',
    'upload-title', 'upload-desc', 'upload-event-at', 'upload-tags', 'upload-status',
    'upload-links', 'folder-breadcrumb'
  ];
  const elements = Object.fromEntries(ids.map((id) => [id, new FakeElement(id)]));
  elements['upload-overlay'].classList.add('open');
  elements['upload-title'].value = 'Keep title';
  elements['upload-desc'].value = 'Keep description';
  elements['upload-event-at'].value = '2026-07-11T12:34';
  elements['upload-tags'].value = '#before';
  elements['upload-status'].value = '未対応';
  elements['upload-links'].textContent = 'https://example.com/keep';
  elements['folder-breadcrumb'].textContent = 'Keep folder';
  const loads = [];
  let candidates = [];
  let colorSetCalls = 0;
  let iconSetCalls = 0;
  const state = {
    upload: {
      originalFile: { name: 'keep.jpg' },
      uploadFile: { name: 'keep.jpg' },
      previewUrl: 'blob:keep', lat: 35, lng: 139, capturedAt: '2026-07-11',
      metadataStatus: 'success', converting: false, conversionError: '', selectionToken: 1,
      color: '#e53935', icon: 'default', presetOptions: [], presetLoading: false,
      presetApplying: false, presetError: '', presetRequestToken: 0
    }
  };
  const context = {
    state,
    document: {
      getElementById(id) { return elements[id]; },
      createElement() { return new FakeElement('option'); }
    },
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_ICONS: [
      { id: 'default' }, { id: 'photo' }, { id: 'food' }, { id: 'hotel' },
      { id: 'nature' }, { id: 'shop' }, { id: 'transit' }, { id: 'warning' }
    ],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    canEdit: () => true,
    ensureInputPresetsLoaded: () => {
      const next = loads.shift();
      return next ? next.promise : Promise.resolve();
    },
    reloadInputPresets: () => {
      const next = loads.shift();
      return next ? next.promise : Promise.resolve();
    },
    getEnabledInputPresets: () => candidates,
    normalizeTags(value) {
      return String(value || '').split(/\s+/).map((tag) => tag.replace(/^#/, '')).filter(Boolean);
    },
    setUploadColor(value) { colorSetCalls += 1; state.upload.color = value; },
    setUploadIcon(value) { iconSetCalls += 1; state.upload.icon = value; },
    Promise,
    console
  };
  const sources = [
    applyCoreSource(),
    functionSource('uploadPresetViewIsActive'),
    functionSource('invalidateUploadPresetView'),
    functionSource('renderUploadInputPresetControls'),
    functionSource('loadUploadInputPresets'),
    functionSource('applyInputPresetToUpload'),
    "document.getElementById('upload-preset-apply').addEventListener('click', applyInputPresetToUpload);",
    "document.getElementById('upload-preset-select').addEventListener('change', renderUploadInputPresetControls);",
    "document.getElementById('upload-preset-retry').addEventListener('click', function() { loadUploadInputPresets(true); });"
  ];
  vm.runInNewContext(`${sources.join('\n')}\nthis.api = { loadUploadInputPresets, applyInputPresetToUpload, invalidateUploadPresetView };`, context);
  return {
    api: context.api,
    elements,
    state,
    enqueue(load) { loads.push(load); },
    setCandidates(next) { candidates = next; },
    colorSetCalls() { return colorSetCalls; },
    iconSetCalls() { return iconSetCalls; }
  };
}

test('upload UI exposes preset controls and an explicit blank registration status', () => {
  for (const id of ['upload-preset-select', 'upload-preset-apply', 'upload-preset-retry']) {
    assert.match(indexHtml, new RegExp(`id=["']${id}["']`));
  }
  const status = indexHtml.match(/<select id="upload-status"[^>]*>([\s\S]*?)<\/select>/);
  assert.ok(status);
  assert.match(status[1], /<option value="">未設定<\/option>/);
  assert.ok(status[1].indexOf('value=""') < status[1].indexOf('value="未対応"'));
  assert.match(functionSource('resetUploadState'), /upload-status'\)\.value = '未対応'/);
});

test('upload catalog loads asynchronously, excludes disabled presets, and selection does not apply', async () => {
  const harness = createHarness();
  const pending = deferred();
  harness.enqueue(pending);
  harness.setCandidates([
    preset('enabled', { tagsMode: 'set', tags: ['after'] }),
    preset('disabled', { enabled: false, tagsMode: 'clear' })
  ]);
  const loading = harness.api.loadUploadInputPresets(false);
  assert.equal(harness.elements['upload-preset-select'].disabled, true);
  pending.resolve();
  await loading;

  assert.equal(harness.elements['upload-preset-select'].children.length, 2);
  assert.equal(harness.elements['upload-preset-select'].children[0].textContent, 'プリセットを選択');
  assert.equal(harness.elements['upload-preset-select'].children[1].value, 'enabled');
  harness.elements['upload-preset-select'].value = 'enabled';
  harness.elements['upload-preset-select'].dispatch('change');
  assert.equal(harness.elements['upload-tags'].value, '#before');
  assert.equal(harness.state.upload.color, '#e53935');
});

test('upload applies sequential presets only to tags, color, icon, and status', async () => {
  const harness = createHarness();
  harness.setCandidates([
    preset('first', { tagsMode: 'set', tags: ['one'], colorMode: 'set', color: '#009688' }),
    preset('second', { name: '<img onerror=alert(1)>', tagsMode: 'clear', iconMode: 'set', icon: 'photo', statusMode: 'clear' })
  ]);
  await harness.api.loadUploadInputPresets(false);
  const uploadBefore = { ...harness.state.upload };

  harness.elements['upload-preset-select'].value = 'first';
  harness.elements['upload-preset-select'].dispatch('change');
  harness.elements['upload-preset-apply'].dispatch('click');
  harness.elements['upload-preset-apply'].dispatch('click');
  await settle();
  harness.elements['upload-preset-select'].value = 'second';
  harness.elements['upload-preset-select'].dispatch('change');
  harness.elements['upload-preset-apply'].dispatch('click');
  await settle();

  assert.equal(harness.elements['upload-tags'].value, '');
  assert.equal(harness.state.upload.color, '#009688');
  assert.equal(harness.state.upload.icon, 'photo');
  assert.equal(harness.elements['upload-status'].value, '');
  assert.equal(harness.elements['upload-preset-status'].textContent, '<img onerror=alert(1)>を適用しました');
  assert.equal(harness.colorSetCalls(), 2);
  assert.equal(harness.iconSetCalls(), 2);
  assert.equal(harness.elements['upload-title'].value, 'Keep title');
  assert.equal(harness.elements['upload-desc'].value, 'Keep description');
  assert.equal(harness.elements['upload-event-at'].value, '2026-07-11T12:34');
  assert.equal(harness.elements['upload-links'].textContent, 'https://example.com/keep');
  assert.equal(harness.elements['folder-breadcrumb'].textContent, 'Keep folder');
  for (const field of ['originalFile', 'uploadFile', 'previewUrl', 'lat', 'lng', 'capturedAt', 'metadataStatus', 'selectionToken']) {
    assert.equal(harness.state.upload[field], uploadBefore[field], field);
  }
});

test('upload empty catalog points to management without disabling normal form fields', async () => {
  const harness = createHarness();
  await harness.api.loadUploadInputPresets(false);
  assert.match(harness.elements['upload-preset-status'].textContent, /管理画面で作成/);
  assert.equal(harness.elements['upload-title'].disabled, false);
  assert.equal(harness.elements['upload-desc'].disabled, false);
});

test('closed upload view ignores stale load results while the reopened view can render shared results', async () => {
  const harness = createHarness();
  const first = deferred();
  harness.enqueue(first);
  const oldLoad = harness.api.loadUploadInputPresets(false);
  harness.api.invalidateUploadPresetView();
  harness.elements['upload-overlay'].classList.remove('open');
  harness.setCandidates([preset('stale')]);
  first.resolve();
  await oldLoad;
  assert.equal(harness.elements['upload-preset-select'].children.length, 1);
  assert.equal(harness.elements['upload-preset-select'].children[0].value, '');

  harness.elements['upload-overlay'].classList.add('open');
  harness.setCandidates([preset('fresh')]);
  await harness.api.loadUploadInputPresets(false);
  assert.equal(harness.elements['upload-preset-select'].children[1].value, 'fresh');
});

test('upload load errors are sanitized and retry restores applying without changing photo state', async () => {
  const harness = createHarness();
  const failed = deferred();
  harness.enqueue(failed);
  const photo = harness.state.upload.originalFile;
  const loading = harness.api.loadUploadInputPresets(false);
  failed.reject(new Error('<script>private detail</script>'));
  await loading;
  assert.match(harness.elements['upload-preset-error'].textContent, /読み込みに失敗/);
  assert.doesNotMatch(harness.elements['upload-preset-error'].textContent, /private detail/);

  const recovered = deferred();
  harness.enqueue(recovered);
  harness.setCandidates([preset('recovered', { iconMode: 'set', icon: 'food' })]);
  harness.elements['upload-preset-retry'].dispatch('click');
  recovered.resolve();
  await settle();
  assert.equal(harness.elements['upload-preset-select'].disabled, false);
  harness.elements['upload-preset-select'].value = 'recovered';
  harness.elements['upload-preset-select'].dispatch('change');
  assert.equal(harness.elements['upload-preset-apply'].disabled, false);
  assert.equal(harness.state.upload.originalFile, photo);
});

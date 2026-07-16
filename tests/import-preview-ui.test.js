const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function extractImportScript() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1, 'Expected ImportJobCore');
  assert.notEqual(end, -1, 'Expected application state after import modules');
  return indexHtml.slice(start, end);
}

class FakeClassList {
  constructor(initial) {
    this.values = new Set(initial || []);
  }
  add(value) { this.values.add(value); }
  remove(value) { this.values.delete(value); }
  contains(value) { return this.values.has(value); }
  toggle(value, force) {
    const enabled = force == null ? !this.contains(value) : !!force;
    if (enabled) this.add(value); else this.remove(value);
    return enabled;
  }
}

class FakeElement {
  constructor(tagName, id) {
    this.tagName = String(tagName || 'div').toUpperCase();
    this.id = id || '';
    this.classList = new FakeClassList();
    this.children = [];
    this.parentNode = null;
    this.dataset = {};
    this.style = {};
    this.attributes = {};
    this.listeners = Object.create(null);
    this.disabled = false;
    this.value = '';
    this.src = '';
    this._textContent = '';
  }
  set textContent(value) {
    this._textContent = value == null ? '' : String(value);
    this.children = [];
  }
  get textContent() {
    return this._textContent + this.children.map((child) => child.textContent).join('');
  }
  set innerHTML(_value) {
    throw new Error('Import preview must not write innerHTML');
  }
  get innerHTML() { return ''; }
  appendChild(child) {
    child.parentNode = this;
    this.children.push(child);
    return child;
  }
  replaceChildren(...children) {
    this.children.forEach((child) => { child.parentNode = null; });
    this.children = [];
    children.forEach((child) => this.appendChild(child));
  }
  setAttribute(name, value) {
    this.attributes[name] = String(value);
    if (name === 'src') this.src = String(value);
  }
  removeAttribute(name) {
    delete this.attributes[name];
    if (name === 'src') this.src = '';
  }
  addEventListener(type, listener) {
    if (!this.listeners[type]) this.listeners[type] = [];
    this.listeners[type].push(listener);
  }
  listenerCount(type) {
    return (this.listeners[type] || []).length;
  }
  dispatch(type, target) {
    const event = { type, target: target || this, preventDefault() {} };
    (this.listeners[type] || []).forEach((listener) => listener(event));
  }
  closest(selector) {
    if (selector === '[data-import-item-id]' && this.dataset.importItemId) return this;
    return this.parentNode ? this.parentNode.closest(selector) : null;
  }
  focus() {}
}

function createPreviewDocument() {
  const ids = [
    'import-preview-overlay', 'import-preview-sheet', 'import-preview-title', 'import-preview-source',
    'import-preview-count-total', 'import-preview-count-waiting',
    'import-preview-count-processing', 'import-preview-count-succeeded',
    'import-preview-count-failed', 'import-preview-filter-all',
    'import-preview-filter-needs-review', 'import-preview-filter-processing',
    'import-preview-filter-succeeded', 'import-preview-filter-failed',
    'import-preview-list', 'import-preview-empty',
    'import-preview-job-status', 'import-preview-operation-note', 'import-preview-operation-error',
    'import-preview-completion-message',
    'import-preview-presets', 'import-preview-preset-select', 'import-preview-preset-apply-selected',
    'import-preview-preset-apply-all', 'import-preview-preset-status',
    'import-preview-preset-error', 'import-preview-preset-retry',
    'import-preview-editor', 'import-preview-photo-pane', 'import-preview-image-trigger', 'import-preview-image',
    'import-preview-location-summary', 'import-preview-item-status',
    'import-preview-item-error', 'import-preview-edit-title', 'import-preview-edit-description',
    'import-preview-edit-lat', 'import-preview-edit-lng', 'import-preview-time-field-label',
    'import-preview-edit-captured-at', 'import-preview-edit-links',
    'import-preview-edit-tags', 'import-preview-edit-color',
    'import-preview-edit-icon', 'import-preview-edit-status',
    'import-preview-edit-metadata-status', 'import-preview-edit-conversion-status',
    'import-preview-delete', 'import-preview-primary', 'import-preview-cancel',
    'import-preview-resume', 'import-preview-retry', 'import-preview-close', 'import-preview-discard',
    'multi-photo-track-match-panel', 'multi-photo-track-select', 'multi-photo-track-utc-offset',
    'multi-photo-track-clock-correction', 'multi-photo-track-max-gap',
    'multi-photo-track-endpoint-tolerance', 'multi-photo-track-run',
    'multi-photo-track-status', 'multi-photo-track-error', 'multi-photo-track-counts',
    'multi-photo-track-warnings', 'multi-photo-track-results', 'multi-photo-track-apply',
    'multi-photo-track-clear'
  ];
  const elements = Object.create(null);
  ids.forEach((id) => {
    const tag = id === 'import-preview-image' ? 'img'
      : id === 'import-preview-image-trigger' ? 'button'
      : id === 'import-preview-preset-select' || id === 'import-preview-edit-status' ? 'select'
      : id === 'import-preview-edit-color' || id === 'import-preview-edit-icon' ? 'div'
      : id.startsWith('import-preview-edit-') ? 'input'
        : id.includes('delete') || id.includes('primary') || id.includes('cancel')
          || id.includes('resume') || id.includes('retry') || id.includes('apply')
          || id.includes('close') || id.includes('discard') ? 'button'
          : 'div';
    elements[id] = new FakeElement(tag, id);
  });
  elements['import-preview-overlay'].classList.add('sheet-overlay');
  const fieldMap = {
    'import-preview-edit-title': 'title',
    'import-preview-edit-description': 'description',
    'import-preview-edit-lat': 'lat',
    'import-preview-edit-lng': 'lng',
    'import-preview-edit-captured-at': 'capturedAt',
    'import-preview-edit-links': 'links',
    'import-preview-edit-tags': 'tags',
    'import-preview-edit-color': 'color',
    'import-preview-edit-icon': 'icon',
    'import-preview-edit-status': 'status',
    'import-preview-edit-metadata-status': 'metadataStatus',
    'import-preview-edit-conversion-status': 'conversionStatus'
  };
  Object.entries(fieldMap).forEach(([id, field]) => {
    elements[id].dataset.importField = field;
  });
  return {
    elements,
    document: {
      getElementById(id) { return elements[id] || null; },
      createElement(tagName) { return new FakeElement(tagName); }
    }
  };
}

function loadModules(options = {}) {
  const harness = createPreviewDocument();
  const context = {
    document: harness.document,
    URL: options.urlApi || { revokeObjectURL() {} },
    console,
    Number,
    Object,
    Array,
    String,
    Error,
    createInlineIconSvg(icon) { return `<svg data-icon="${icon}"></svg>`; }
  };
  context.SAFE_COLOR_RE = /^#[0-9a-fA-F]{6}$/;
  context.PIN_COLORS = [
    { hex: '#e53935', label: '赤' }, { hex: '#009688', label: 'ティール' },
    { hex: '#2196f3', label: '青' }, { hex: '#4caf50', label: '緑' }
  ];
  context.PIN_ICONS = [
    { id: 'default', label: '標準' }, { id: 'photo', label: '写真' },
    { id: 'food', label: '食事' }, { id: 'hotel', label: '宿' },
    { id: 'nature', label: '自然' }, { id: 'shop', label: '店' },
    { id: 'transit', label: '交通' }, { id: 'warning', label: '注意' }
  ];
  context.PIN_STATUSES = ['未対応', '対応中', '完了', '保留'];
  vm.createContext(context);
  vm.runInContext(
    `${extractImportScript()}\nglobalThis.__applyCore = InputPresetApplyCore;\nglobalThis.__core = ImportJobCore;\nglobalThis.__ui = ImportPreviewUI;\nglobalThis.__openMulti = typeof openMultiPhotoImportPreview === 'undefined' ? null : openMultiPhotoImportPreview;`,
    context
  );
  return {
    applyCore: context.__applyCore,
    core: context.__core,
    ui: context.__ui,
    openMulti: context.__openMulti,
    builder: context.__builder,
    elements: harness.elements
  };
}

function createSampleJob(core) {
  return core.createJob({
    id: 'job-1',
    sourceType: 'photo',
    items: [
      {
        id: 'item-1', sourceType: 'photo', sourceRef: '<img src=x onerror=1>.jpg',
        title: '<b>東京</b>', description: '説明', lat: 35.6812, lng: 139.7671,
        tags: ['旅', '<script>'], color: '#2196f3', icon: 'photo',
        links: ['https://example.com', 'https://example.org'],
        status: '',
        runtime: { previewUrl: 'blob:preview-1' }
      },
      {
        id: 'item-2', sourceType: 'csv', title: '失敗項目', lat: null, lng: null,
        uploadStatus: 'failed', error: '<img src=x onerror=alert(1)>'
      }
    ]
  });
}

test('import preview DOM exists and is hidden initially', () => {
  const requiredIds = [
    'import-preview-overlay', 'import-preview-title', 'import-preview-source',
    'import-preview-count-total', 'import-preview-count-waiting',
    'import-preview-count-processing', 'import-preview-count-succeeded',
    'import-preview-count-failed', 'import-preview-list', 'import-preview-editor',
    'import-preview-preset-select', 'import-preview-preset-apply-selected',
    'import-preview-preset-apply-all', 'import-preview-preset-retry', 'import-preview-edit-status',
    'import-preview-time-field-label', 'import-preview-edit-links',
    'import-preview-primary', 'import-preview-close', 'import-preview-discard'
    , 'import-preview-completion-message'
  ];
  requiredIds.forEach((id) => assert.match(indexHtml, new RegExp(`id=["']${id}["']`)));
  const openingTag = indexHtml.match(/<div id="import-preview-overlay"[^>]*>/);
  assert.ok(openingTag);
  assert.equal(openingTag[0].includes('open'), false);
});

test('metadata pickers reuse the single-pin palette UI and render idempotently from shared constants', () => {
  assert.doesNotMatch(indexHtml, /<select id="import-preview-edit-(?:color|icon)"/);
  assert.doesNotMatch(indexHtml, /<input[^>]*id="import-preview-edit-(?:color|icon)"/);
  const colorPicker = indexHtml.match(/<div id="import-preview-edit-color"[^>]*>/)[0];
  const iconPicker = indexHtml.match(/<div id="import-preview-edit-icon"[^>]*>/)[0];
  const statusSelect = indexHtml.match(/<select id="import-preview-edit-status"[\s\S]*?<\/select>/)[0];
  assert.match(colorPicker, /class="color-palette/);
  assert.match(colorPicker, /role="group"/);
  assert.match(colorPicker, /aria-labelledby="import-preview-color-label"/);
  assert.match(iconPicker, /class="icon-picker/);
  assert.match(iconPicker, /role="group"/);
  assert.match(iconPicker, /aria-labelledby="import-preview-icon-label"/);
  assert.equal(statusSelect.includes('<option'), false);
  assert.doesNotMatch(indexHtml, /import-preview-(?:color|icon)-(?:options|preview)|<datalist/);
  assert.match(indexHtml, /\.import-preview-metadata-picker[\s\S]*?min-height:\s*44px/);

  const { core, ui, elements } = loadModules();
  const job = core.createJob({ id: 'metadata-options', items: [{ id: 'one', color: '#2196f3', icon: 'hotel' }] });
  ui.open({ job });
  ui.close();
  ui.open({ job });

  assert.deepEqual(
    elements['import-preview-edit-color'].children.map((button) => ({
      value: button.dataset.pinColor, title: button.title,
      ariaLabel: button.attributes['aria-label'], pressed: button.attributes['aria-pressed']
    })),
    [
      { value: '#e53935', title: '赤', ariaLabel: '赤', pressed: 'false' },
      { value: '#009688', title: 'ティール', ariaLabel: 'ティール', pressed: 'false' },
      { value: '#2196f3', title: '青', ariaLabel: '青', pressed: 'true' },
      { value: '#4caf50', title: '緑', ariaLabel: '緑', pressed: 'false' }
    ]
  );
  assert.equal(elements['import-preview-edit-icon'].children.length, 8);
  assert.equal(elements['import-preview-edit-color'].tagName, 'DIV');
  assert.equal(elements['import-preview-edit-icon'].tagName, 'DIV');
  assert.equal(elements['import-preview-edit-icon'].children.filter((button) => button.attributes['aria-pressed'] === 'true').length, 1);
  assert.equal(elements['import-preview-edit-icon'].children[3].dataset.pinIcon, 'hotel');
  assert.equal(elements['import-preview-edit-icon'].children[3].attributes['aria-pressed'], 'true');
  assert.deepEqual(
    elements['import-preview-edit-status'].children.map((option) => [option.value, option.textContent]),
    [['', '未設定'], ['未対応', '未対応'], ['対応中', '対応中'], ['完了', '完了'], ['保留', '保留']]
  );
});

test('multi-photo discard uses the builder URL owner and deduplicates repeated URLs', async () => {
  const harness = createPreviewDocument();
  const globalRevoked = [];
  const customRevoked = [];
  const context = {
    document: harness.document,
    URL: { revokeObjectURL(url) { globalRevoked.push(url); } },
    console,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: [{ id: 'default' }],
    PIN_STATUSES: ['未対応'],
    createInlineIconSvg(icon) { return `<svg data-icon="${icon}"></svg>`; }
  };
  vm.createContext(context);
  vm.runInContext(
    `${extractImportScript()}\n`
      + 'globalThis.__builder = MultiPhotoImportBuilder;\n'
      + 'globalThis.__ui = ImportPreviewUI;\n'
      + 'globalThis.__openMulti = openMultiPhotoImportPreview;',
    context
  );
  let id = 0;
  const session = context.__builder.create({
    preparePhoto: async (file) => ({
      originalFile: file, uploadFile: file, metadataStatus: 'no-gps', conversionStatus: 'not-needed'
    }),
    createObjectURL: () => 'blob:shared',
    revokeObjectURL: (url) => customRevoked.push(url),
    createId: () => `id-${++id}`
  });
  const files = [
    { name: 'one.jpg', type: 'image/jpeg' },
    { name: 'two.jpg', type: 'image/jpeg' }
  ];
  const job = await session.start(files, {
    tags: [], color: '#e53935', icon: 'default', status: ''
  });

  let latest = null;
  context.__openMulti(job, { onDraftChange(value) { latest = value; } });
  harness.elements['import-preview-edit-title'].value = 'edited';
  harness.elements['import-preview-editor'].dispatch(
    'change', harness.elements['import-preview-edit-title']
  );
  assert.equal(latest.items[0].title, 'edited');
  context.__ui.close();
  context.__openMulti(latest);
  context.__ui.close({ discard: true });

  assert.deepEqual(customRevoked, ['blob:shared']);
  assert.deepEqual(globalRevoked, []);
  assert.equal(latest.items[0].runtime.originalFile, null);
  assert.equal(latest.items[1].runtime.originalFile, null);
});

test('hidePrimaryAction hides only opted-in preview primary actions and resets on close', () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);

  ui.open({ job, onPrimaryAction() {}, hidePrimaryAction: true });
  assert.equal(elements['import-preview-primary'].style.display, 'none');
  assert.equal(elements['import-preview-primary'].disabled, true);
  ui.close();

  ui.open({ job, onPrimaryAction() {} });
  assert.equal(elements['import-preview-primary'].style.display, '');
  assert.equal(elements['import-preview-primary'].disabled, true);
});

test('multi-photo preview opener supplies safe copy, hides registration, and forwards the latest draft', () => {
  const { core, ui, openMulti, elements } = loadModules();
  const job = createSampleJob(core);
  let latest = null;

  assert.equal(typeof openMulti, 'function');
  openMulti(job, { onDraftChange(value) { latest = value; } });
  assert.equal(elements['import-preview-title'].textContent, '複数写真を確認');
  assert.equal(elements['import-preview-source'].textContent, '写真 2件');
  assert.equal(elements['import-preview-primary'].style.display, 'none');
  elements['import-preview-edit-title'].value = '更新';
  elements['import-preview-editor'].dispatch('change', elements['import-preview-edit-title']);
  assert.equal(latest.items[0].title, '更新');
  assert.equal(ui.getJob().items[0].title, '更新');
  const failedButton = elements['import-preview-list'].children[1];
  elements['import-preview-list'].dispatch('click', failedButton);
  elements['import-preview-delete'].dispatch('click');
  assert.equal(latest.items.length, 1);
  assert.equal(latest.items[0].id, 'item-1');
});

test('discard-only policy hides normal close and always releases resources through one action', () => {
  const revoked = [];
  const { core, ui, elements } = loadModules({
    urlApi: { revokeObjectURL(url) { revoked.push(url); } }
  });
  const job = core.createJob({
    id: 'production-close',
    items: [{
      id: 'one', title: 'One', uploadStatus: 'queued',
      runtime: { originalFile: {}, uploadFile: {}, previewUrl: 'blob:production' }
    }]
  });
  let closeInfo = null;
  ui.open({ job, closePolicy: 'discard-only', onClose(info) { closeInfo = info; } });
  assert.equal(elements['import-preview-close'].style.display, 'none');
  assert.equal(elements['import-preview-discard'].style.display, '');
  assert.equal(elements['import-preview-discard'].textContent, '取込を破棄');

  assert.equal(ui.close(), true);
  assert.equal(closeInfo.discarded, true);
  assert.deepEqual(revoked, ['blob:production']);
  assert.equal(job.items[0].runtime.uploadFile, null);

  ui.open({ job: createSampleJob(core) });
  assert.equal(elements['import-preview-close'].style.display, '');
  assert.equal(elements['import-preview-discard'].textContent, '破棄して閉じる');
});

test('production preview renders safe completion, partial failure, and cancelled copy', () => {
  const { core, ui, elements } = loadModules();
  const callbacks = { onResumeAction() {}, onRetryAction() {} };
  const completed = Object.assign({}, createSampleJob(core), {
    status: 'completed',
    items: createSampleJob(core).items.map((item) => ({ ...item, uploadStatus: 'succeeded', error: null })),
    counts: { total: 2, waiting: 0, processing: 0, succeeded: 2, failed: 0 }
  });
  ui.open({ job: completed, closePolicy: 'discard-only', ...callbacks });
  assert.equal(elements['import-preview-completion-message'].textContent, '2件の写真を登録しました');
  assert.equal(elements['import-preview-discard'].textContent, '閉じる');

  const partial = Object.assign({}, completed, {
    items: [completed.items[0], { ...completed.items[1], uploadStatus: 'failed', error: 'safe' }],
    counts: { total: 2, waiting: 0, processing: 0, succeeded: 1, failed: 1 }
  });
  ui.setJob(partial);
  assert.match(elements['import-preview-completion-message'].textContent, /1件成功、1件失敗/);
  assert.match(elements['import-preview-completion-message'].textContent, /失敗した項目だけ再試行できます/);
  assert.equal(elements['import-preview-discard'].textContent, '取込を破棄');

  const cancelled = Object.assign({}, partial, {
    status: 'cancelled',
    items: [{ ...partial.items[0], uploadStatus: 'succeeded' }, { ...partial.items[1], uploadStatus: 'queued' }],
    counts: { total: 2, waiting: 1, processing: 0, succeeded: 1, failed: 0 }
  });
  ui.setJob(cancelled);
  assert.match(elements['import-preview-completion-message'].textContent, /登録をキャンセルしました/);
  assert.match(elements['import-preview-completion-message'].textContent, /未処理の項目を再開できます/);
});

test('completion item label is scoped to one open and resets to the photo default', () => {
  const { core, ui, elements } = loadModules();
  const completed = Object.assign({}, createSampleJob(core), {
    status: 'completed',
    items: createSampleJob(core).items.map((item) => ({ ...item, uploadStatus: 'succeeded', error: null })),
    counts: { total: 2, waiting: 0, processing: 0, succeeded: 2, failed: 0 }
  });
  ui.open({ job: completed, closePolicy: 'discard-only', completionItemLabel: 'ピン' });
  assert.equal(elements['import-preview-completion-message'].textContent, '2件のピンを登録しました');
  ui.close();
  ui.open({ job: completed, closePolicy: 'discard-only' });
  assert.equal(elements['import-preview-completion-message'].textContent, '2件の写真を登録しました');
});

test('completion item label accepts only fixed product copy', () => {
  const { core, ui, elements } = loadModules();
  const completed = Object.assign({}, createSampleJob(core), {
    status: 'completed',
    items: createSampleJob(core).items.map((item) => ({ ...item, uploadStatus: 'succeeded' })),
    counts: { total: 2, waiting: 0, processing: 0, succeeded: 2, failed: 0 }
  });
  ui.open({ job: completed, closePolicy: 'discard-only', completionItemLabel: '<利用者入力>' });
  assert.equal(elements['import-preview-completion-message'].textContent, '2件の写真を登録しました');
});

test('close callback exceptions are isolated after all preview references are cleared', () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  ui.open({
    job,
    closePolicy: 'discard-only',
    onClose() { throw new Error('consumer callback failed'); }
  });
  assert.doesNotThrow(() => ui.close());
  assert.equal(ui.getJob(), null);
  assert.equal(elements['import-preview-overlay'].classList.contains('open'), false);
  ui.open({ job: createSampleJob(core) });
  assert.equal(ui.getJob().id, 'job-1');
});

test('close render failures cannot prevent the ownership callback or leave a retained preview', () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  let callbackCount = 0;
  ui.open({
    job,
    closePolicy: 'discard-only',
    onClose(info) {
      callbackCount += 1;
      assert.equal(info.job, job);
      assert.equal(info.discarded, true);
    }
  });
  elements['import-preview-list'].replaceChildren = function() {
    throw new Error('close render failed');
  };

  assert.doesNotThrow(() => ui.close());
  assert.equal(callbackCount, 1);
  assert.equal(ui.getJob(), null);
  assert.equal(elements['import-preview-overlay'].classList.contains('open'), false);
});

test('preview initialization failure rolls back retained job and callbacks', () => {
  const { core, ui, elements } = loadModules();
  const replaceChildren = elements['import-preview-list'].replaceChildren.bind(elements['import-preview-list']);
  elements['import-preview-list'].replaceChildren = function() {
    throw new Error('render failed');
  };
  assert.throws(() => ui.open({
    job: createSampleJob(core),
    closePolicy: 'discard-only',
    onClose() { throw new Error('stale callback'); }
  }), /render failed/);
  assert.equal(ui.getJob(), null);
  assert.equal(elements['import-preview-overlay'].classList.contains('open'), false);

  elements['import-preview-list'].replaceChildren = replaceChildren;
  assert.doesNotThrow(() => ui.open({ job: createSampleJob(core) }));
  assert.equal(ui.close(), true);
});

test('non-retryable failures stay visible with actionable guidance and cannot retry', () => {
  const { core, ui, elements } = loadModules();
  const base = core.createJob({
    id: 'permanent-failure',
    items: [{
      id: 'invalid', title: 'Invalid', uploadStatus: 'failed',
      error: '入力内容が不正です。', errorCode: 'INVALID_IMPORT_PAYLOAD', retryable: false
    }]
  });
  const completed = Object.assign({}, base, {
    status: 'completed',
    counts: core.getCounts(base)
  });

  ui.open({ job: completed, closePolicy: 'discard-only', onRetryAction() {} });

  assert.match(elements['import-preview-item-error'].textContent, /入力内容が不正です。/);
  assert.match(elements['import-preview-item-error'].textContent, /入力を修正してください。/);
  assert.equal(elements['import-preview-retry'].style.display, '');
  assert.equal(elements['import-preview-retry'].disabled, true);
  assert.match(elements['import-preview-completion-message'].textContent, /再試行できない項目があります/);

  const corrupted = Object.assign({}, completed, {
    items: completed.items.map((item) => Object.assign({}, item, {
      errorCode: 'IMPORT_RECEIPT_CORRUPTED'
    }))
  });
  ui.setJob(corrupted);
  assert.match(elements['import-preview-item-error'].textContent, /管理者へ連絡してください。/);

  ui.close();
  const cancelled = Object.assign({}, completed, {
    status: 'cancelled', cancelRequested: true,
    items: completed.items.map((item) => Object.assign({}, item, {
      errorCode: 'IMPORT_RECEIPT_CORRUPTED', retryable: false
    }))
  });
  ui.open({ job: cancelled, closePolicy: 'discard-only', onResumeAction() {} });
  assert.equal(elements['import-preview-resume'].style.display, 'none');
  assert.equal(elements['import-preview-resume'].disabled, true);
});

test('failed CSV rows are read-only but remain removable', () => {
  const { core, ui, elements } = loadModules();
  const job = core.createJob({
    id: 'csv-failed',
    sourceType: 'csv',
    items: [{
      id: 'bad-row', sourceType: 'csv', sourceRef: 'CSV 2行目', title: 'Bad',
      uploadStatus: 'failed', error: 'CSV 2行目: URLを確認してください。',
      errorCode: 'CSV_ROW_LINKS_INVALID', retryable: false
    }]
  });
  ui.open({ job, hidePrimaryAction: true, closePolicy: 'discard-only' });
  assert.equal(elements['import-preview-edit-title'].disabled, true);
  assert.equal(elements['import-preview-edit-links'].disabled, true);
  assert.equal(elements['import-preview-delete'].disabled, false);
  assert.match(elements['import-preview-item-error'].textContent, /この項目を削除してください/);
});

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

test('preview loads enabled presets asynchronously and selection alone never changes drafts', async () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  const pending = deferred();
  let loadCalls = 0;
  let draftChanges = 0;
  const candidates = [
    preset('enabled', { tagsMode: 'set', tags: ['適用後'] }),
    preset('disabled', { enabled: false, tagsMode: 'clear' })
  ];

  ui.open({
    job,
    loadInputPresets() { loadCalls += 1; return pending.promise; },
    getInputPresets() { return candidates; },
    onDraftChange() { draftChanges += 1; }
  });
  assert.equal(elements['import-preview-preset-select'].disabled, true);
  assert.equal(elements['import-preview-preset-apply-selected'].disabled, true);
  pending.resolve();
  await pending.promise;
  await settle();

  assert.equal(loadCalls, 1);
  assert.equal(elements['import-preview-preset-select'].children.length, 1);
  assert.equal(elements['import-preview-preset-select'].children[0].value, 'enabled');
  elements['import-preview-preset-select'].value = 'enabled';
  elements['import-preview-preset-select'].dispatch('change');
  assert.equal(draftChanges, 0);
  assert.deepEqual(plain(ui.getJob()), plain(job));
});

test('selected preset applications compose from current values and notify once per click', async () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  const beforeSecond = plain(job.items[1]);
  let draftChanges = 0;
  const candidates = [
    preset('first', { tagsMode: 'set', tags: ['最初'], colorMode: 'set', color: '#009688' }),
    preset('second', { tagsMode: 'clear', iconMode: 'set', icon: 'food', statusMode: 'set', status: '完了' })
  ];
  ui.open({
    job,
    loadInputPresets: () => Promise.resolve(),
    getInputPresets: () => candidates,
    onDraftChange() { draftChanges += 1; }
  });
  await settle();

  elements['import-preview-preset-select'].value = 'first';
  elements['import-preview-preset-apply-selected'].dispatch('click');
  await settle();
  elements['import-preview-preset-select'].value = 'second';
  elements['import-preview-preset-apply-selected'].dispatch('click');
  await settle();

  const applied = ui.getJob().items[0];
  assert.deepEqual(plain(applied.tags), []);
  assert.equal(applied.color, '#009688');
  assert.equal(applied.icon, 'food');
  assert.equal(applied.status, '完了');
  assert.equal(applied.title, job.items[0].title);
  assert.equal(applied.runtime.previewUrl, 'blob:preview-1');
  assert.deepEqual(plain(applied.links), ['https://example.com', 'https://example.org']);
  assert.deepEqual(plain(ui.getJob().items[1]), beforeSecond);
  assert.equal(draftChanges, 2);
  assert.equal(elements['import-preview-edit-color'].value, '#009688');
  assert.equal(elements['import-preview-edit-color'].children[1].attributes['aria-pressed'], 'true');
  assert.equal(elements['import-preview-edit-icon'].value, 'food');
  assert.equal(elements['import-preview-edit-icon'].children[2].attributes['aria-pressed'], 'true');
});

test('color and icon buttons follow item switching and preserve the exact clicked icon id', () => {
  const { core, ui, elements } = loadModules();
  const job = core.createJob({
    id: 'metadata-switch', sourceType: 'multi-photo', items: [
      { id: 'one', sourceType: 'photo', title: 'One', color: '#2196f3', icon: 'photo' },
      { id: 'two', sourceType: 'photo', title: 'Two', color: '#4caf50', icon: 'nature' }
    ]
  });
  let latest = null;
  ui.open({ job, onDraftChange(value) { latest = value; } });
  assert.equal(elements['import-preview-edit-color'].value, '#2196f3');
  assert.equal(elements['import-preview-edit-icon'].value, 'photo');
  assert.equal(elements['import-preview-edit-icon'].children[1].attributes['aria-pressed'], 'true');

  elements['import-preview-list'].dispatch('click', elements['import-preview-list'].children[1]);
  assert.equal(elements['import-preview-edit-color'].value, '#4caf50');
  assert.equal(elements['import-preview-edit-icon'].value, 'nature');
  assert.equal(elements['import-preview-edit-icon'].children[4].attributes['aria-pressed'], 'true');
  assert.equal(elements['import-preview-edit-icon'].children.filter((button) => button.attributes['aria-pressed'] === 'true').length, 1);

  elements['import-preview-edit-color'].children[1].dispatch('click');
  elements['import-preview-edit-icon'].children[7].dispatch('click');
  assert.equal(latest.items[1].color, '#009688');
  assert.equal(latest.items[1].icon, 'warning');
  assert.equal(elements['import-preview-edit-color'].value, '#009688');
  assert.equal(elements['import-preview-edit-icon'].value, 'warning');
  assert.equal(elements['import-preview-edit-icon'].children[7].attributes['aria-pressed'], 'true');
  assert.equal(elements['import-preview-edit-icon'].children.filter((button) => button.attributes['aria-pressed'] === 'true').length, 1);
});

test('metadata palettes re-enable when switching from a read-only item to a queued item', () => {
  const { core, ui, elements } = loadModules();
  const job = core.createJob({
    id: 'metadata-disabled-switch', sourceType: 'multi-photo', items: [
      { id: 'failed', sourceType: 'photo', title: 'Failed', uploadStatus: 'failed', color: '#2196f3', icon: 'photo' },
      { id: 'queued', sourceType: 'photo', title: 'Queued', color: '#4caf50', icon: 'nature' }
    ]
  });

  ui.open({ job });
  assert.equal(elements['import-preview-edit-color'].children.every((button) => button.disabled), true);
  assert.equal(elements['import-preview-edit-icon'].children.every((button) => button.disabled), true);

  elements['import-preview-list'].dispatch('click', elements['import-preview-list'].children[1]);
  assert.equal(elements['import-preview-edit-color'].children.every((button) => !button.disabled), true);
  assert.equal(elements['import-preview-edit-icon'].children.every((button) => !button.disabled), true);
  assert.equal(elements['import-preview-edit-icon'].children[4].attributes['aria-pressed'], 'true');
});

test('apply all performs one atomic job update and one draft notification', async () => {
  const { core, ui, elements } = loadModules();
  const job = core.createJob({
    id: 'bulk',
    items: [
      { id: 'one', title: 'One', tags: ['A'], color: '#e53935', icon: 'default', status: '', uploadStatus: 'queued', attempts: 2 },
      { id: 'two', title: 'Two', tags: ['B'], color: '#2196f3', icon: 'photo', status: '未対応', uploadStatus: 'failed', error: 'keep me', attempts: 3 },
      { id: 'done', title: 'Done', tags: ['C'], color: '#4caf50', icon: 'nature', status: '完了', uploadStatus: 'succeeded', attempts: 1 }
    ]
  });
  let draftChanges = 0;
  ui.open({
    job,
    loadInputPresets: () => Promise.resolve(),
    getInputPresets: () => [preset('all', { tagsMode: 'clear', statusMode: 'set', status: '保留' })],
    onDraftChange() { draftChanges += 1; }
  });
  await settle();
  elements['import-preview-preset-apply-all'].dispatch('click');
  await settle();

  const updated = ui.getJob();
  assert.notEqual(updated, job);
  assert.deepEqual(plain(updated.items.map((item) => item.tags)), [[], [], []]);
  assert.deepEqual(plain(updated.items.map((item) => item.status)), ['保留', '保留', '保留']);
  assert.deepEqual(plain(updated.items.map((item) => item.uploadStatus)), ['queued', 'failed', 'succeeded']);
  assert.deepEqual(plain(updated.items.map((item) => item.attempts)), [2, 3, 1]);
  assert.equal(updated.items[1].error, 'keep me');
  assert.equal(draftChanges, 1);
});

test('apply all validates every item before update and retry recovers a failed catalog load', async () => {
  const { core, ui, elements } = loadModules();
  const job = core.createJob({
    id: 'atomic-failure',
    items: [
      { id: 'valid', tags: ['A'], color: '#e53935', icon: 'default', status: '' },
      { id: 'invalid', tags: ['B'], color: '#2196f3', icon: 'unknown', status: '' }
    ]
  });
  const before = plain(job);
  const firstLoad = deferred();
  let loadCalls = 0;
  let draftChanges = 0;
  const candidates = [preset('clear', { tagsMode: 'clear' })];
  ui.open({
    job,
    loadInputPresets() {
      loadCalls += 1;
      return loadCalls === 1 ? firstLoad.promise : Promise.resolve();
    },
    getInputPresets: () => candidates,
    onDraftChange() { draftChanges += 1; }
  });
  firstLoad.reject(new Error('<script>private detail</script>'));
  await firstLoad.promise.catch(() => {});
  await settle();
  assert.match(elements['import-preview-preset-error'].textContent, /読み込みに失敗/);
  assert.doesNotMatch(elements['import-preview-preset-error'].textContent, /private detail/);
  elements['import-preview-preset-retry'].dispatch('click');
  await settle();
  assert.equal(loadCalls, 2);

  elements['import-preview-preset-apply-all'].dispatch('click');
  await settle();
  assert.deepEqual(plain(ui.getJob()), before);
  assert.equal(draftChanges, 0);
  assert.match(elements['import-preview-preset-error'].textContent, /適用できません/);
  assert.equal(elements['import-preview-preset-retry'].style.display, 'none');
  assert.equal(elements['import-preview-preset-apply-all'].disabled, false);
});

test('draft item updates are immutable, allowlisted, and validate coordinates', () => {
  const { core } = loadModules();
  const job = createSampleJob(core);
  const before = plain(job);
  const updated = core.updateDraftItem(job, 'item-1', {
    title: '更新後', lat: null, lng: 135.5, tags: ['更新', 'タグ']
  });

  assert.notEqual(updated, job);
  assert.notEqual(updated.items, job.items);
  assert.equal(updated.items[0].title, '更新後');
  assert.equal(updated.items[0].lat, null);
  assert.equal(updated.items[0].lng, 135.5);
  assert.deepEqual(plain(updated.items[0].tags), ['更新', 'タグ']);
  assert.deepEqual(plain(job), before);

  for (const field of ['id', 'uploadStatus', 'attempts', 'runtime', 'unknown']) {
    assert.throws(
      () => core.updateDraftItem(job, 'item-1', { [field]: 'bad' }),
      (error) => error.code === 'INVALID_IMPORT_DRAFT_PATCH'
    );
  }
  for (const invalid of [NaN, Infinity, -Infinity, '35.0', undefined]) {
    assert.throws(
      () => core.updateDraftItem(job, 'item-1', { lat: invalid }),
      (error) => error.code === 'INVALID_IMPORT_COORDINATE'
    );
  }
  assert.throws(
    () => core.updateDraftItem(job, 'missing', { title: 'x' }),
    (error) => error.code === 'IMPORT_ITEM_NOT_FOUND'
  );
});

test('draft updates and removals are allowed only for idle jobs', () => {
  const { core } = loadModules();
  const job = createSampleJob(core);
  for (const status of ['ready', 'running', 'completed', 'cancelled']) {
    const nonIdle = Object.assign({}, job, { status });
    assert.throws(
      () => core.updateDraftItem(nonIdle, 'item-1', { title: 'x' }),
      (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
    );
    assert.throws(
      () => core.removeDraftItem(nonIdle, 'item-1'),
      (error) => error.code === 'INVALID_IMPORT_JOB_TRANSITION'
    );
  }
});

test('draft removal revokes preview URL, releases returned references, and recounts without mutating the job', () => {
  const revoked = [];
  const { core } = loadModules();
  const job = createSampleJob(core);
  const originalFile = new Blob(['original']);
  const uploadFile = new Blob(['upload']);
  job.items[0].runtime.originalFile = originalFile;
  job.items[0].runtime.uploadFile = uploadFile;
  const before = plain(job);

  const updated = core.removeDraftItem(job, 'item-1', {
    revokeObjectURL(url) { revoked.push(url); }
  });

  assert.deepEqual(revoked, ['blob:preview-1']);
  assert.equal(updated.status, 'idle');
  assert.equal(updated.items.length, 1);
  assert.equal(updated.counts.total, 1);
  assert.equal(updated.counts.failed, 1);
  assert.equal(updated.items.some((item) => item.runtime.originalFile === originalFile), false);
  assert.deepEqual(plain(job), before);
  assert.throws(
    () => core.removeDraftItem(job, 'missing'),
    (error) => error.code === 'IMPORT_ITEM_NOT_FOUND'
  );

  const empty = core.removeDraftItem(core.createJob({
    id: 'empty-after-remove', items: [{ id: 'only' }]
  }), 'only');
  assert.equal(empty.status, 'idle');
  assert.deepEqual(plain(empty.counts), {
    total: 0, succeeded: 0, failed: 0, processing: 0, waiting: 0
  });
});

test('preview renders job counts, item details, statuses, and errors with textContent', () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  ui.open({ job, title: '<i>取込確認</i>', sourceLabel: '複数写真' });

  assert.equal(elements['import-preview-overlay'].classList.contains('open'), true);
  assert.equal(elements['import-preview-title'].textContent, '<i>取込確認</i>');
  assert.equal(elements['import-preview-source'].textContent, '複数写真');
  assert.equal(elements['import-preview-count-total'].textContent, '2');
  assert.equal(elements['import-preview-count-waiting'].textContent, '1');
  assert.equal(elements['import-preview-count-failed'].textContent, '1');
  const listText = elements['import-preview-list'].textContent;
  assert.match(listText, /<b>東京<\/b>/);
  assert.match(listText, /35\.6812/);
  assert.match(listText, /139\.7671/);
  assert.match(listText, /要確認/);
  assert.match(listText, /失敗/);
  assert.match(listText, /<img src=x onerror=alert\(1\)>/);
  assert.equal(extractImportScript().includes('.innerHTML'), false);

  const allStatuses = core.createJob({
    id: 'job-all-statuses',
    items: [
      { id: 'queued', uploadStatus: 'queued' },
      { id: 'processing', uploadStatus: 'processing' },
      { id: 'succeeded', uploadStatus: 'succeeded' },
      { id: 'failed', uploadStatus: 'failed', error: '失敗理由' }
    ]
  });
  ui.setJob(allStatuses);
  assert.equal(elements['import-preview-count-total'].textContent, '4');
  assert.equal(elements['import-preview-count-waiting'].textContent, '1');
  assert.equal(elements['import-preview-count-processing'].textContent, '1');
  assert.equal(elements['import-preview-count-succeeded'].textContent, '1');
  assert.equal(elements['import-preview-count-failed'].textContent, '1');
  assert.match(elements['import-preview-list'].textContent, /要確認/);
  assert.match(elements['import-preview-list'].textContent, /登録中/);
  assert.match(elements['import-preview-list'].textContent, /登録済み/);
  assert.match(elements['import-preview-list'].textContent, /失敗/);
});

test('preview variants separate photo-only pane and status filters select visible items', () => {
  const { core, ui, elements } = loadModules();
  const photoJob = core.createJob({
    id: 'photo-variant', sourceType: 'multi-photo', items: [
      { id: 'queued', title: '要確認項目', uploadStatus: 'queued' },
      { id: 'processing', title: '登録中項目', uploadStatus: 'processing' },
      { id: 'done', title: '登録済み項目', uploadStatus: 'succeeded' },
      { id: 'failed', title: '失敗項目', uploadStatus: 'failed', error: '失敗' }
    ]
  });
  ui.open({ job: photoJob });
  assert.equal(elements['import-preview-sheet'].attributes['data-import-variant'], 'photo');
  elements['import-preview-filter-failed'].dispatch('click');
  assert.match(elements['import-preview-list'].textContent, /失敗項目/);
  assert.doesNotMatch(elements['import-preview-list'].textContent, /要確認項目|登録中項目|登録済み項目/);
  assert.equal(elements['import-preview-filter-failed'].attributes['aria-pressed'], 'true');
  assert.equal(elements['import-preview-edit-title'].value, '失敗項目');
  const settledItems = photoJob.items.map((item) => item.id === 'failed'
    ? { ...item, uploadStatus: 'succeeded', error: null }
    : item);
  ui.setJob({ ...photoJob, items: settledItems, counts: core.getCounts({ items: settledItems }) });
  assert.equal(elements['import-preview-list'].children.length, 0);
  assert.equal(elements['import-preview-edit-title'].value, '');

  ui.close();
  const csvJob = core.createJob({
    id: 'csv-variant', sourceType: 'csv', items: [{ id: 'csv', title: 'CSV', uploadStatus: 'queued' }]
  });
  ui.open({ job: csvJob });
  assert.equal(elements['import-preview-sheet'].attributes['data-import-variant'], 'csv');
  assert.equal(elements['import-preview-filter-all'].attributes['aria-pressed'], 'true');
});

test('retryable uncertain save result uses safe reconfirmation copy', () => {
  const { core, ui, elements } = loadModules();
  const base = core.createJob({
    id: 'uncertain-save', sourceType: 'multi-photo', items: [{
      id: 'photo', title: 'Photo', uploadStatus: 'failed',
      error: 'transport detail', errorCode: 'IMPORT_ITEM_SAVE_FAILED', retryable: true
    }]
  });
  ui.open({ job: Object.assign({}, base, { status: 'completed', counts: core.getCounts(base) }) });
  assert.equal(
    elements['import-preview-item-error'].textContent,
    '保存結果を確認できませんでした。安全に再確認します。'
  );
  assert.doesNotMatch(elements['import-preview-item-error'].textContent, /transport|response|idempotency/);
});

test('known Preview variants reject mixed item sources before retaining the job', () => {
  const { core, ui } = loadModules();
  const mixed = core.createJob({
    id: 'mixed-photo-csv', sourceType: 'multi-photo', items: [
      { id: 'photo', sourceType: 'photo', title: 'Photo' },
      { id: 'csv', sourceType: 'csv', title: 'CSV' }
    ]
  });
  assert.throws(() => ui.open({ job: mixed }), /mixed source variants/);
  assert.equal(ui.getJob(), null);
});

test('primary action is disabled without a callback and invokes only the injected callback', () => {
  const { core, ui, elements } = loadModules();
  const job = core.createJob({
    id: 'ready-compatible',
    sourceType: 'csv',
    items: [
      { id: 'queued', uploadStatus: 'queued' },
      { id: 'succeeded', uploadStatus: 'succeeded' }
    ]
  });
  ui.open({ job, title: '確認', sourceLabel: 'CSV' });
  assert.equal(elements['import-preview-primary'].disabled, true);

  let received = null;
  ui.open({
    job,
    title: '確認',
    sourceLabel: 'CSV',
    onPrimaryAction(currentJob) { received = currentJob; }
  });
  assert.equal(elements['import-preview-primary'].disabled, false);
  assert.doesNotThrow(() => core.readyJob(ui.getJob()));
  elements['import-preview-primary'].dispatch('click');
  assert.equal(received, job);
});

test('primary action remains disabled for idle jobs that Core cannot make ready', () => {
  const { core, ui, elements } = loadModules();
  const invalid = createSampleJob(core);
  ui.open({
    job: invalid,
    title: '確認',
    sourceLabel: 'CSV',
    onPrimaryAction() {}
  });

  assert.throws(
    () => core.readyJob(invalid),
    (error) => error.code === 'INVALID_IMPORT_JOB_ITEMS'
  );
  assert.equal(elements['import-preview-primary'].disabled, true);
});

test('opening repeatedly registers each preview event only once', () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  ui.open({ job, title: '1', sourceLabel: 'CSV', onPrimaryAction() {} });
  ui.open({ job, title: '2', sourceLabel: 'GeoJSON', onPrimaryAction() {} });
  ui.open({ job, title: '3', sourceLabel: 'GPX', onPrimaryAction() {} });

  assert.equal(elements['import-preview-list'].listenerCount('click'), 1);
  assert.equal(elements['import-preview-editor'].listenerCount('change'), 1);
  assert.equal(elements['import-preview-delete'].listenerCount('click'), 1);
  assert.equal(elements['import-preview-primary'].listenerCount('click'), 1);
  assert.equal(elements['import-preview-cancel'].listenerCount('click'), 1);
  assert.equal(elements['import-preview-resume'].listenerCount('click'), 1);
  assert.equal(elements['import-preview-retry'].listenerCount('click'), 1);
  assert.equal(elements['import-preview-close'].listenerCount('click'), 1);
  assert.equal(elements['import-preview-discard'].listenerCount('click'), 1);
  elements['import-preview-edit-icon'].children.forEach((button) => {
    assert.equal(button.listenerCount('click'), 1);
  });
});

test('idle items can be selected, edited, and removed through the UI', () => {
  const revoked = [];
  const harness = loadModules({ urlApi: { revokeObjectURL(url) { revoked.push(url); } } });
  const { core, ui, elements } = harness;
  const job = createSampleJob(core);
  ui.open({ job, title: '確認', sourceLabel: '写真', onPrimaryAction() {} });

  const firstButton = elements['import-preview-list'].children[0];
  elements['import-preview-list'].dispatch('click', firstButton);
  assert.equal(elements['import-preview-edit-title'].value, '<b>東京</b>');
  elements['import-preview-edit-title'].value = '編集済み';
  elements['import-preview-editor'].dispatch('change', elements['import-preview-edit-title']);
  assert.equal(ui.getJob().items[0].title, '編集済み');
  elements['import-preview-edit-status'].value = '保留';
  elements['import-preview-editor'].dispatch('change', elements['import-preview-edit-status']);
  assert.equal(ui.getJob().items[0].status, '保留');
  assert.equal(ui.getJob().items[0].uploadStatus, 'queued');
  assert.equal(elements['import-preview-edit-links'].value, 'https://example.com\nhttps://example.org');
  elements['import-preview-edit-links'].value = ' https://a.example \n\ninvalid-value\n';
  elements['import-preview-editor'].dispatch('change', elements['import-preview-edit-links']);
  assert.deepEqual(plain(ui.getJob().items[0].links), ['https://a.example', 'invalid-value']);

  elements['import-preview-delete'].dispatch('click');
  assert.equal(ui.getJob().items.length, 1);
  assert.deepEqual(revoked, ['blob:preview-1']);
});

test('CSV preview customizes and resets the time label while hiding registration and clearing on close', () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  let closedJob = null;
  ui.open({
    job,
    title: 'CSVインポートを確認',
    sourceLabel: 'CSV 2件',
    hidePrimaryAction: true,
    closePolicy: 'discard-only',
    timeFieldLabel: 'イベント時刻',
    onClose(result) { closedJob = result.job; }
  });
  assert.equal(elements['import-preview-title'].textContent, 'CSVインポートを確認');
  assert.equal(elements['import-preview-source'].textContent, 'CSV 2件');
  assert.equal(elements['import-preview-time-field-label'].textContent, 'イベント時刻');
  assert.equal(elements['import-preview-primary'].style.display, 'none');
  assert.equal(elements['import-preview-close'].style.display, 'none');
  elements['import-preview-discard'].dispatch('click');
  assert.equal(closedJob, job);
  assert.equal(ui.getJob(), null);

  ui.open({ job: createSampleJob(core), title: '写真' });
  assert.equal(elements['import-preview-time-field-label'].textContent, '撮影日時');
});

test('ready jobs disable editing and deletion while remaining startable', () => {
  const { core, ui, elements } = loadModules();
  const base = createSampleJob(core);
  const items = base.items.map((item, index) => Object.assign({}, item, {
    uploadStatus: index === 0 ? 'succeeded' : 'queued'
  }));
  const job = Object.assign({}, base, {
    status: 'ready', items, counts: core.getCounts({ items })
  });
  ui.open({ job, title: '確認', sourceLabel: 'Drive', onPrimaryAction() {} });

  assert.equal(elements['import-preview-delete'].disabled, true);
  assert.equal(elements['import-preview-edit-title'].disabled, true);
  assert.equal(elements['import-preview-edit-status'].disabled, true);
  assert.equal(elements['import-preview-primary'].disabled, false);
  assert.match(elements['import-preview-list'].textContent, /登録済み/);
});

test('running or processing jobs cannot close', () => {
  const { core, ui, elements } = loadModules();
  const base = createSampleJob(core);
  const items = base.items.map((item, index) => Object.assign({}, item, {
    uploadStatus: index === 0 ? 'processing' : 'queued'
  }));
  const running = Object.assign({}, base, {
    status: 'running', items, counts: core.getCounts({ items })
  });
  ui.open({ job: running, title: '処理中', sourceLabel: '写真' });

  assert.equal(elements['import-preview-close'].disabled, true);
  assert.equal(elements['import-preview-discard'].disabled, true);
  assert.equal(ui.close(), false);
  assert.equal(elements['import-preview-overlay'].classList.contains('open'), true);
});

test('explicit discard releases job resources and clears retained references', () => {
  const revoked = [];
  const { core, ui, elements } = loadModules({
    urlApi: { revokeObjectURL(url) { revoked.push(url); } }
  });
  const job = createSampleJob(core);
  let closeInfo = null;
  ui.open({
    job, title: '確認', sourceLabel: '写真',
    onClose(info) { closeInfo = info; }
  });

  assert.equal(ui.close({ discard: true }), true);
  assert.deepEqual(revoked, ['blob:preview-1']);
  assert.equal(job.items[0].runtime.previewUrl, '');
  assert.equal(ui.getJob(), null);
  assert.equal(elements['import-preview-overlay'].classList.contains('open'), false);
  assert.equal(closeInfo.discarded, true);
});

test('normal close preserves job resources while clearing controller callbacks and references', () => {
  const { core, ui, elements } = loadModules();
  const job = core.createJob({
    id: 'normal-close',
    items: [{ id: 'queued', runtime: { previewUrl: 'blob:keep-preview' } }]
  });
  let closeInfo = null;
  ui.open({
    job,
    title: '確認',
    sourceLabel: '写真',
    onPrimaryAction() {},
    onClose(info) { closeInfo = info; }
  });

  assert.equal(ui.close(), true);
  assert.equal(job.items[0].runtime.previewUrl, 'blob:keep-preview');
  assert.equal(ui.getJob(), null);
  assert.equal(closeInfo.discarded, false);
  assert.equal(elements['import-preview-overlay'].classList.contains('open'), false);

  ui.open({ job, title: '再表示', sourceLabel: '写真' });
  assert.equal(elements['import-preview-primary'].disabled, true);
});

test('only safe preview URLs can be assigned to image src', () => {
  const { core, ui, elements } = loadModules();
  const job = createSampleJob(core);
  job.items[0].runtime.previewUrl = 'javascript:alert(1)';
  ui.open({ job, title: '確認', sourceLabel: '写真' });
  assert.equal(elements['import-preview-image'].src, '');
  assert.equal(elements['import-preview-image'].style.display, 'none');

  job.items[0].runtime.previewUrl = 'blob:safe-preview';
  ui.setJob(job);
  assert.equal(elements['import-preview-image'].src, 'blob:safe-preview');
  assert.equal(elements['import-preview-image'].style.display, 'block');
});

test('preview overlay is responsive, blocks shortcuts, and is excluded from backdrop dismissal', () => {
  assert.match(indexHtml, /@media \(max-width: 900px\)[\s\S]*\.import-preview-sheet\[data-import-variant\][\s\S]*\.import-preview-layout/);

  const shortcutStart = indexHtml.indexOf('function isShortcutOverlayOpen()');
  const shortcutEnd = indexHtml.indexOf('function hasMultipleSelectedPins()', shortcutStart);
  assert.notEqual(shortcutStart, -1);
  assert.match(indexHtml.slice(shortcutStart, shortcutEnd), /'import-preview-overlay'/);

  const backdropStart = indexHtml.indexOf('const BACKDROP_DISMISS_OVERLAY_IDS = [');
  const backdropEnd = indexHtml.indexOf('];', backdropStart);
  assert.notEqual(backdropStart, -1);
  assert.notEqual(backdropEnd, -1);
  assert.equal(indexHtml.slice(backdropStart, backdropEnd).includes("'import-preview-overlay'"), false);
});

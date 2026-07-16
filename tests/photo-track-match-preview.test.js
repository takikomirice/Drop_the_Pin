const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const readme = fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8');

class FakeClassList {
  constructor() { this.values = new Set(); }
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
  constructor(tagName = 'div', id = '') {
    this.tagName = tagName.toUpperCase();
    this.id = id;
    this.classList = new FakeClassList();
    this.children = [];
    this.parentNode = null;
    this.dataset = {};
    this.style = {};
    this.attributes = {};
    this.listeners = Object.create(null);
    this.disabled = false;
    this.checked = false;
    this.value = '';
    this.type = '';
    this._textContent = '';
    this.focused = false;
  }
  set textContent(value) { this._textContent = value == null ? '' : String(value); this.children = []; }
  get textContent() { return this._textContent + this.children.map((child) => child.textContent).join(''); }
  set innerHTML(_value) { throw new Error('innerHTML is forbidden'); }
  appendChild(child) { child.parentNode = this; this.children.push(child); return child; }
  replaceChildren(...children) { this.children = []; children.forEach((child) => this.appendChild(child)); }
  setAttribute(name, value) { this.attributes[name] = String(value); }
  removeAttribute(name) { delete this.attributes[name]; }
  addEventListener(type, listener) { (this.listeners[type] ||= []).push(listener); }
  dispatch(type, target = this) {
    const event = { type, target, preventDefault() {} };
    (this.listeners[type] || []).forEach((listener) => listener(event));
  }
  closest(selector) {
    if (selector === '[data-import-item-id]' && this.dataset.importItemId) return this;
    return this.parentNode ? this.parentNode.closest(selector) : null;
  }
  focus() { this.focused = true; }
}

function createDocument() {
  const baseIds = [
    'import-preview-overlay', 'import-preview-sheet', 'import-preview-title', 'import-preview-source',
    'import-preview-count-total', 'import-preview-count-waiting', 'import-preview-count-processing',
    'import-preview-count-succeeded', 'import-preview-count-failed',
    'import-preview-filter-all', 'import-preview-filter-needs-review',
    'import-preview-filter-processing', 'import-preview-filter-succeeded',
    'import-preview-filter-failed', 'import-preview-list',
    'import-preview-empty', 'import-preview-job-status', 'import-preview-operation-note',
    'import-preview-operation-error', 'import-preview-completion-message', 'import-preview-presets',
    'import-preview-preset-select', 'import-preview-preset-apply-selected',
    'import-preview-preset-apply-all', 'import-preview-preset-status', 'import-preview-preset-error',
    'import-preview-preset-retry', 'import-preview-editor', 'import-preview-photo-pane',
    'import-preview-image-trigger', 'import-preview-image', 'import-preview-location-summary',
    'import-preview-item-status', 'import-preview-item-error', 'import-preview-edit-title',
    'import-preview-edit-description', 'import-preview-edit-lat', 'import-preview-edit-lng',
    'import-preview-time-field-label', 'import-preview-edit-captured-at', 'import-preview-edit-links',
    'import-preview-edit-tags', 'import-preview-edit-color', 'import-preview-color-preview',
    'import-preview-edit-icon', 'import-preview-icon-preview', 'import-preview-edit-status',
    'import-preview-edit-metadata-status', 'import-preview-edit-conversion-status',
    'import-preview-delete', 'import-preview-primary', 'import-preview-cancel', 'import-preview-resume',
    'import-preview-retry', 'import-preview-close', 'import-preview-discard'
  ];
  const trackIds = [
    'multi-photo-track-match-panel', 'multi-photo-track-select', 'multi-photo-track-utc-offset',
    'multi-photo-track-clock-correction', 'multi-photo-track-max-gap',
    'multi-photo-track-endpoint-tolerance', 'multi-photo-track-run',
    'multi-photo-track-status', 'multi-photo-track-error', 'multi-photo-track-counts',
    'multi-photo-track-warnings', 'multi-photo-track-results', 'multi-photo-track-apply',
    'multi-photo-track-clear'
  ];
  const elements = Object.create(null);
  baseIds.concat(trackIds).forEach((id) => {
    const tag = id.includes('select') || id === 'import-preview-edit-status'
        || id === 'import-preview-edit-color' || id === 'import-preview-edit-icon' ? 'select'
      : id.includes('button') || id.includes('apply') || id.includes('run') || id.includes('clear')
        || id.includes('delete') || id.includes('primary') || id.includes('cancel')
        || id.includes('resume') || id.includes('retry') || id.includes('close') || id.includes('discard')
        ? 'button' : id.includes('offset') || id.includes('correction') || id.includes('gap')
          || id.startsWith('import-preview-edit-') ? 'input' : 'div';
    elements[id] = new FakeElement(tag, id);
  });
  const fields = {
    'import-preview-edit-title': 'title', 'import-preview-edit-description': 'description',
    'import-preview-edit-lat': 'lat', 'import-preview-edit-lng': 'lng',
    'import-preview-edit-captured-at': 'capturedAt', 'import-preview-edit-links': 'links',
    'import-preview-edit-tags': 'tags', 'import-preview-edit-color': 'color',
    'import-preview-edit-icon': 'icon', 'import-preview-edit-status': 'status',
    'import-preview-edit-metadata-status': 'metadataStatus',
    'import-preview-edit-conversion-status': 'conversionStatus'
  };
  Object.entries(fields).forEach(([id, field]) => { elements[id].dataset.importField = field; });
  return {
    elements,
    document: {
      getElementById(id) { return elements[id] || null; },
      createElement(tagName) { return new FakeElement(tagName); }
    }
  };
}

function loadUi() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  const harness = createDocument();
  const context = {
    document: harness.document,
    URL: { revokeObjectURL() {} },
    console,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#2196f3', label: '青' }],
    PIN_ICONS: [{ id: 'photo', label: '写真' }],
    PIN_STATUSES: ['未対応'],
    Date,
    Math,
    Map,
    Set
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__core = ImportJobCore; globalThis.__ui = ImportPreviewUI;',
    context
  );
  return { core: context.__core, ui: context.__ui, elements: harness.elements };
}

function createJob(core, sourceType = 'multi-photo') {
  const itemSourceType = sourceType === 'multi-photo' ? 'photo' : sourceType;
  return core.createJob({
    id: 'job-1', sourceType,
    items: [
      { id: 'a', sourceType: itemSourceType, title: '<b>GPS写真</b>', lat: 35, lng: 139, uploadStatus: 'queued' },
      { id: 'b', sourceType: itemSourceType, title: '<img src=x onerror=1>', lat: null, lng: null, uploadStatus: 'queued' }
    ]
  });
}

function trackMatchController(audit) {
  const model = {
    enabled: true, open: true,
    tracks: [{
      trackId: 'track-1', revisionId: 'rev-1', name: '<img src=x onerror=1>',
      startTime: '2026-07-12T00:00:00.000Z', endTime: '2026-07-12T01:00:00.000Z',
      pointCount: 1245, timedPointCount: 1245
    }],
    trackId: 'track-1', trackRevisionId: 'rev-1', utcOffsetText: '+09:00',
    clockCorrectionSeconds: 0, maxInterpolationGapSeconds: 300, endpointToleranceSeconds: 0,
    result: {
      counts: { total: 2, matched: 1, exact: 0, interpolated: 1, endpoint: 0,
        skippedExistingGps: 1, missingTime: 0, invalidTime: 0, outsideRange: 0,
        gapTooLarge: 0, ambiguous: 0, invalidInput: 0 },
      warnings: [{ code: 'TRACK_MATCH_POINTS_WITHOUT_TIME', count: 2 }],
      results: [
        { itemId: 'a', status: 'skipped-existing-gps' },
        { itemId: 'b', status: 'matched-interpolated' }
      ]
    },
    selectedItemIds: ['b'], appliedItemIds: [], stale: false, running: false,
    errorCode: '', errorMessage: '', noticeCode: 'matched', canRun: true, canApply: true
  };
  return {
    model,
    getViewModel() { audit.views += 1; return model; },
    setTrack(value) { audit.trackIds.push(value); model.trackId = value; },
    setOption(key, value) { audit.options.push([key, value]); model[key] = value; },
    run() { audit.runs += 1; return model.result; },
    setSelected(itemId, selected) { audit.selections.push([itemId, selected]); },
    apply(job) {
      audit.applies += 1;
      model.appliedItemIds = ['b'];
      model.noticeCode = 'applied';
      return { ...job, items: job.items.map((item) => item.id === 'b' ? { ...item, lat: 15, lng: 25 } : { ...item }) };
    },
    clearResult() { audit.clears += 1; model.result = null; model.selectedItemIds = []; },
    onDraftChange() { audit.drafts += 1; },
    cleanup() { audit.cleanups += 1; }
  };
}

test('fixed track-match panel is accessible and contains every required control', () => {
  const ids = [
    'multi-photo-track-match-panel', 'multi-photo-track-select', 'multi-photo-track-utc-offset',
    'multi-photo-track-clock-correction', 'multi-photo-track-max-gap',
    'multi-photo-track-endpoint-tolerance', 'multi-photo-track-run',
    'multi-photo-track-status', 'multi-photo-track-error', 'multi-photo-track-counts',
    'multi-photo-track-warnings', 'multi-photo-track-results', 'multi-photo-track-apply',
    'multi-photo-track-clear'
  ];
  ids.forEach((id) => assert.match(indexHtml, new RegExp(`id=["']${id}["']`)));
  assert.match(indexHtml, /id="multi-photo-track-status"[^>]*aria-live="polite"/);
  assert.match(indexHtml, /id="multi-photo-track-error"[^>]*role="alert"/);
  assert.match(indexHtml, /for="multi-photo-track-select"/);
  assert.match(indexHtml, /for="multi-photo-track-utc-offset"/);
  assert.match(indexHtml, /for="multi-photo-track-clock-correction"/);
  assert.match(indexHtml, /候補のルート時刻は、このブラウザのローカル時刻で表示/);
  assert.match(readme, /候補に表示するトラック時刻は現在のブラウザのローカル時刻/);
});

test('panel is hidden without the dedicated extension and renders tracks, counts, warnings, and ordered results safely', () => {
  const audit = { views: 0, trackIds: [], options: [], runs: 0, selections: [], applies: 0, clears: 0, drafts: 0, cleanups: 0 };
  const { core, ui, elements } = loadUi();
  const csv = createJob(core, 'csv');
  ui.open({ job: csv });
  assert.equal(elements['multi-photo-track-match-panel'].style.display, 'none');
  ui.close();

  const job = createJob(core);
  const controller = trackMatchController(audit);
  ui.open({ job, trackMatch: controller });
  assert.equal(elements['multi-photo-track-match-panel'].style.display, '');
  assert.equal(elements['multi-photo-track-select'].children.length, 1);
  assert.match(elements['multi-photo-track-select'].children[0].textContent, /<img src=x onerror=1>/);
  assert.match(elements['multi-photo-track-select'].children[0].textContent, /1,245地点/);
  assert.match(elements['multi-photo-track-counts'].textContent, /位置候補あり 1/);
  assert.match(elements['multi-photo-track-counts'].textContent, /既存GPS 1/);
  assert.equal(elements['multi-photo-track-warnings'].textContent, '時刻のないトラック地点が2件あります。');
  assert.equal(elements['multi-photo-track-results'].children.length, 2);
  assert.match(elements['multi-photo-track-results'].children[0].textContent, /既存GPSを使用/);
  assert.match(elements['multi-photo-track-results'].children[1].textContent, /位置候補あり：ルート上を補間/);
  const checkbox = elements['multi-photo-track-results'].children[1].children[0];
  assert.equal(checkbox.type, 'checkbox');
  assert.equal(checkbox.checked, true);
  assert.match(checkbox.attributes['aria-label'], /<img src=x onerror=1>/);
  assert.equal(elements['multi-photo-track-results'].textContent.includes('15, 25'), false);
});

test('unknown matched-like statuses never render an apply checkbox', () => {
  const audit = { views: 0, trackIds: [], options: [], runs: 0, selections: [], applies: 0, clears: 0, drafts: 0, cleanups: 0 };
  const { core, ui, elements } = loadUi();
  const controller = trackMatchController(audit);
  controller.model.result.results[1].status = 'matched-evil';
  controller.model.selectedItemIds = ['b'];
  ui.open({ job: createJob(core), trackMatch: controller });
  const resultRow = elements['multi-photo-track-results'].children[1];
  assert.equal(resultRow.children.some((child) => child.type === 'checkbox'), false);
});

test('run, selection, apply, option changes, clear, draft edits, and close use the dedicated controller', () => {
  const audit = { views: 0, trackIds: [], options: [], runs: 0, selections: [], applies: 0, clears: 0, drafts: 0, cleanups: 0 };
  const { core, ui, elements } = loadUi();
  const controller = trackMatchController(audit);
  let changedJob = null;
  ui.open({ job: createJob(core), trackMatch: controller, onDraftChange(job) { changedJob = job; } });

  elements['multi-photo-track-run'].dispatch('click');
  assert.equal(audit.runs, 1);
  const checkbox = elements['multi-photo-track-results'].children[1].children[0];
  checkbox.checked = false;
  elements['multi-photo-track-results'].dispatch('change', checkbox);
  assert.deepEqual(audit.selections, [['b', false]]);
  elements['multi-photo-track-utc-offset'].value = '+05:45';
  elements['multi-photo-track-utc-offset'].dispatch('change');
  assert.deepEqual(audit.options.at(-1), ['utcOffsetText', '+05:45']);
  elements['multi-photo-track-apply'].dispatch('click');
  assert.equal(audit.applies, 1);
  assert.equal(changedJob.items[1].lat, 15);
  assert.equal(elements['multi-photo-track-status'].focused, true);
  elements['import-preview-edit-captured-at'].value = '2026-07-12T09:10:00';
  elements['import-preview-editor'].dispatch('change', elements['import-preview-edit-captured-at']);
  assert.equal(audit.drafts, 1);
  elements['multi-photo-track-clear'].dispatch('click');
  assert.equal(audit.clears, 1);
  ui.close({ discard: true });
  assert.equal(audit.cleanups, 1);
});

test('Preview initialization failure cleans the dedicated controller before releasing retained state', () => {
  const { core, ui } = loadUi();
  let cleanups = 0;
  const controller = {
    getViewModel() {
      return {
        enabled: true,
        get tracks() { throw new Error('render failed'); }
      };
    },
    cleanup() { cleanups += 1; }
  };
  assert.throws(() => ui.open({ job: createJob(core), trackMatch: controller }), /render failed/);
  assert.equal(cleanups, 1);
  assert.equal(ui.getJob(), null);
});

test('no timed tracks shows fixed guidance and disables matching', () => {
  const { core, ui, elements } = loadUi();
  const controller = {
    getViewModel() {
      return {
        enabled: true, open: true, tracks: [], trackId: '', trackRevisionId: '',
        utcOffsetText: '+09:00', clockCorrectionSeconds: 0,
        maxInterpolationGapSeconds: 300, endpointToleranceSeconds: 0,
        result: null, selectedItemIds: [], appliedItemIds: [], stale: false,
        running: false, errorCode: '', errorMessage: '', noticeCode: '',
        canRun: false, canApply: false
      };
    },
    cleanup() {}
  };
  ui.open({ job: createJob(core), trackMatch: controller });
  assert.equal(elements['multi-photo-track-run'].disabled, true);
  assert.equal(elements['multi-photo-track-select'].disabled, true);
  assert.equal(
    elements['multi-photo-track-status'].textContent,
    '時刻付きの保存済みルートがありません。先にGPXルートを取り込んでください。'
  );
});

test('dedicated extension remains hidden for every representative non-multi-photo source', () => {
  ['csv', 'geojson', 'gpx', 'photo'].forEach((sourceType) => {
    const audit = { views: 0, trackIds: [], options: [], runs: 0, selections: [], applies: 0, clears: 0, drafts: 0, cleanups: 0 };
    const { core, ui, elements } = loadUi();
    ui.open({ job: createJob(core, sourceType), trackMatch: trackMatchController(audit) });
    assert.equal(elements['multi-photo-track-match-panel'].style.display, 'none', sourceType);
    assert.equal(audit.views, 0, sourceType);
    ui.close();
  });
});

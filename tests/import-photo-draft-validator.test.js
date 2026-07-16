const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const importUrlVectors = require('./fixtures/import-url-vectors');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function loadValidator(options = {}) {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = {
    console,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: options.icons || [{ id: 'default' }, { id: 'photo' }],
    PIN_STATUSES: options.statuses || ['未対応', '対応中', '完了', '保留'],
    URL: Object.prototype.hasOwnProperty.call(options, 'URL')
      ? options.URL : { createObjectURL() {}, revokeObjectURL() {} },
    crypto: { randomUUID: () => 'uuid' }
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__validator = typeof ImportPhotoDraftValidator === "undefined" ? null : ImportPhotoDraftValidator;',
    context
  );
  return context.__validator;
}

function loadValidators(options = {}) {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const state = {', start);
  const context = {
    console,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935' }],
    PIN_ICONS: options.icons || [{ id: 'default' }, { id: 'photo' }],
    PIN_STATUSES: options.statuses || ['未対応', '対応中', '完了', '保留'],
    URL: Object.prototype.hasOwnProperty.call(options, 'URL')
      ? options.URL : { createObjectURL() {}, revokeObjectURL() {} },
    crypto: { randomUUID: () => 'uuid' }
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__validators = {'
      + 'common: typeof ImportDraftValidationCore === "undefined" ? null : ImportDraftValidationCore,'
      + 'photo: typeof ImportPhotoDraftValidator === "undefined" ? null : ImportPhotoDraftValidator,'
      + 'pin: typeof ImportPinDraftValidator === "undefined" ? null : ImportPinDraftValidator};',
    context
  );
  return context.__validators;
}

function item(overrides = {}) {
  return {
    id: 'item-1', sourceRef: 'photo.jpg', title: '写真', description: '',
    lat: null, lng: null, tags: ['観察'], links: ['https://example.com'], color: '#e53935', icon: 'photo',
    status: '', uploadStatus: 'queued', error: null, attempts: 0,
    runtime: { uploadFile: { name: 'photo.jpg' }, originalFile: {}, previewUrl: 'blob:one' },
    ...overrides
  };
}

function job(items, overrides = {}) {
  return { id: 'job-1', status: 'idle', items, ...overrides };
}

test('validator accepts a valid queued item and job without mutation', () => {
  const validator = loadValidator();
  assert.ok(validator);
  const source = item({ lat: 35, lng: 139, tags: [] });
  const before = JSON.stringify(source);
  assert.equal(validator.validateItem(source), true);
  assert.equal(validator.validateJob(job([source])), true);
  assert.equal(JSON.stringify(source), before);
});

test('item validation rejects every unsafe registration field with item-safe errors', () => {
  const validator = loadValidator();
  const invalid = [
    [{ title: '' }, 'IMPORT_DRAFT_TITLE_REQUIRED'],
    [{ title: 'x'.repeat(81) }, 'IMPORT_DRAFT_TITLE_TOO_LONG'],
    [{ lat: 35, lng: null }, 'IMPORT_DRAFT_COORDINATES_INVALID'],
    [{ lat: '35', lng: '139' }, 'IMPORT_DRAFT_COORDINATES_INVALID'],
    [{ lat: 91, lng: 139 }, 'IMPORT_DRAFT_COORDINATES_INVALID'],
    [{ lat: 35, lng: 181 }, 'IMPORT_DRAFT_COORDINATES_INVALID'],
    [{ tags: ['1', '2', '3', '4', '5', '6'] }, 'IMPORT_DRAFT_TAGS_INVALID'],
    [{ links: 'https://example.com' }, 'IMPORT_DRAFT_LINKS_INVALID'],
    [{ links: ['javascript:alert(1)'] }, 'IMPORT_DRAFT_LINKS_INVALID'],
    [{ color: 'red' }, 'IMPORT_DRAFT_COLOR_INVALID'],
    [{ icon: 'unknown' }, 'IMPORT_DRAFT_ICON_INVALID'],
    [{ status: 'unknown' }, 'IMPORT_DRAFT_STATUS_INVALID'],
    [{ runtime: { uploadFile: null } }, 'IMPORT_DRAFT_UPLOAD_FILE_REQUIRED']
  ];
  invalid.forEach(([changes, code]) => {
    assert.throws(
      () => validator.validateItem(item({ sourceRef: '<unsafe>.jpg', ...changes })),
      (error) => error.code === code
        && error.itemId === 'item-1'
        && error.message.includes('<unsafe>.jpg')
    );
  });
});

test('job validation blocks preparation failures, empty jobs, and jobs without queued work', () => {
  const validator = loadValidator();
  assert.throws(() => validator.validateJob(job([])), (error) => error.code === 'IMPORT_DRAFT_JOB_EMPTY');
  assert.throws(
    () => validator.validateJob(job([item({ uploadStatus: 'failed', error: '変換失敗' })])),
    (error) => error.code === 'IMPORT_DRAFT_PREPARATION_FAILED'
      && error.message.includes('photo.jpg')
  );
  assert.throws(
    () => validator.validateJob(job([item({ uploadStatus: 'succeeded' })])),
    (error) => error.code === 'IMPORT_DRAFT_NO_QUEUED_ITEMS'
  );
});

test('job validation ignores settled registration failures while validating retryable work', () => {
  const validator = loadValidator();
  const permanentFailure = item({
    id: 'permanent', uploadStatus: 'failed', attempts: 1,
    error: '入力が不正です。', errorCode: 'INVALID_IMPORT_PAYLOAD', retryable: false
  });
  const queued = item({ id: 'retry', sourceRef: 'retry.jpg' });
  assert.equal(validator.validateJob(job([permanentFailure, queued], { status: 'ready' })), true);
});

test('validator derives allowed icon and status values from production constants', () => {
  const validator = loadValidator({
    icons: [{ id: 'future-icon' }],
    statuses: ['後日']
  });
  assert.equal(validator.validateItem(item({ icon: 'future-icon', status: '後日' })), true);
  assert.throws(
    () => validator.validateItem(item({ icon: 'photo', status: '後日' })),
    (error) => error.code === 'IMPORT_DRAFT_ICON_INVALID'
  );
});

test('validator accepts only colors declared by PIN_COLORS', () => {
  const validator = loadValidator();
  assert.equal(validator.validateItem(item({ color: '#e53935' })), true);
  assert.throws(
    () => validator.validateItem(item({ color: '#abcdef' })),
    (error) => error.code === 'IMPORT_DRAFT_COLOR_INVALID'
  );
});

test('common validator is shared while only the photo validator requires uploadFile', () => {
  const validators = loadValidators();
  assert.ok(validators.common);
  assert.ok(validators.photo);
  assert.ok(validators.pin);
  const pinItem = item({ runtime: null, capturedAt: '2026-07-11T10:30:59' });
  assert.equal(validators.common.validateCommonItem(pinItem), true);
  assert.equal(validators.pin.validateItem(pinItem), true);
  assert.throws(
    () => validators.photo.validateItem(pinItem),
    (error) => error.code === 'IMPORT_DRAFT_UPLOAD_FILE_REQUIRED'
  );
});

test('common validator rejects description, URL, and capturedAt values that strict server validation rejects', () => {
  const { common } = loadValidators();
  const invalid = [
    [{ description: 'x'.repeat(401) }, 'IMPORT_DRAFT_DESCRIPTION_TOO_LONG'],
    [{ links: ['https://'] }, 'IMPORT_DRAFT_LINKS_INVALID'],
    [{ links: [42] }, 'IMPORT_DRAFT_LINKS_INVALID'],
    [{ capturedAt: '2026-02-30T10:00' }, 'IMPORT_DRAFT_EVENT_AT_INVALID'],
    [{ capturedAt: 'not-a-date' }, 'IMPORT_DRAFT_EVENT_AT_INVALID']
  ];
  invalid.forEach(([changes, code]) => {
    assert.throws(() => common.validateCommonItem(item(changes)), (error) => error.code === code);
  });
});

test('common validator uses the shared URL vectors with native and fallback parsing', () => {
  [URL, undefined].forEach((UrlCtor) => {
    const { common } = loadValidators({ URL: UrlCtor });
    importUrlVectors.allowed.forEach((link) => {
      assert.equal(common.validateCommonItem(item({ links: [link] })), true, link);
    });
    importUrlVectors.rejected.forEach((link) => {
      assert.throws(
        () => common.validateCommonItem(item({ links: [link] })),
        (error) => error.code === 'IMPORT_DRAFT_LINKS_INVALID',
        link
      );
    });
  });
});

test('common validator accepts real year boundaries and rejects year zero', () => {
  const { common } = loadValidators();
  assert.equal(common.validateCommonItem(item({ capturedAt: '0001-01-01T00:00:00' })), true);
  assert.equal(common.validateCommonItem(item({ capturedAt: '2000-02-29T23:59:59' })), true);
  assert.throws(
    () => common.validateCommonItem(item({ capturedAt: '0000-12-31T23:59' })),
    (error) => error.code === 'IMPORT_DRAFT_EVENT_AT_INVALID'
  );
});

test('pin job validator blocks initial CSV failures until deletion but ignores settled permanent failures', () => {
  const { pin } = loadValidators();
  assert.throws(
    () => pin.validateJob(job([
      item({ id: 'bad', uploadStatus: 'failed', attempts: 0, retryable: false, errorCode: 'CSV_ROW_TITLE_REQUIRED' }),
      item({ id: 'good', runtime: null })
    ])),
    (error) => error.code === 'IMPORT_DRAFT_PREPARATION_FAILED' && /削除/.test(error.message)
  );
  assert.equal(pin.validateJob(job([
    item({ id: 'saved-failure', uploadStatus: 'failed', attempts: 1, retryable: false, runtime: null }),
    item({ id: 'queued', runtime: null })
  ])), true);
});

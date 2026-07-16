const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function loadCore() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const ImportJobCore = (function() {', start);
  assert.notEqual(start, -1, 'Expected InputPresetApplyCore');
  assert.notEqual(end, -1, 'Expected ImportJobCore after InputPresetApplyCore');
  const context = {};
  vm.runInNewContext(`
    const SAFE_COLOR_RE = /^#[0-9a-fA-F]{6}$/;
    const PIN_ICONS = [
      { id: 'default' }, { id: 'photo' }, { id: 'food' }, { id: 'hotel' },
      { id: 'nature' }, { id: 'shop' }, { id: 'transit' }, { id: 'warning' }
    ];
    const PIN_STATUSES = ['未対応', '対応中', '完了', '保留'];
    ${indexHtml.slice(start, end)}
    globalThis.__core = InputPresetApplyCore;
  `, context);
  return context.__core;
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function current(overrides = {}) {
  return {
    title: '変更しないタイトル',
    description: '変更しない説明',
    lat: 35,
    tags: ['既存'],
    color: '#e53935',
    icon: 'default',
    status: '未対応',
    runtime: { previewUrl: 'blob:keep' },
    ...overrides
  };
}

function preset(overrides = {}) {
  return {
    presetId: 'preset-1',
    name: '基本',
    enabled: true,
    tagsMode: 'keep',
    tags: [],
    colorMode: 'keep',
    color: null,
    iconMode: 'keep',
    icon: null,
    statusMode: 'keep',
    status: null,
    ...overrides
  };
}

test('apply is immutable and changes only tags, color, icon, and status', () => {
  const core = loadCore();
  const source = current();
  const before = plain(source);
  const result = core.apply(source, preset({
    tagsMode: 'set', tags: ['旅行', '写真'],
    colorMode: 'set', color: '#2196f3',
    iconMode: 'set', icon: 'photo',
    statusMode: 'set', status: '対応中'
  }));

  assert.notEqual(result, source);
  assert.notEqual(result.tags, source.tags);
  assert.deepEqual(plain(result), {
    ...before,
    tags: ['旅行', '写真'],
    color: '#2196f3',
    icon: 'photo',
    status: '対応中'
  });
  assert.deepEqual(plain(source), before);
  assert.equal(result.runtime, source.runtime);
});

test('keep, set, and clear modes compose sequentially from current values', () => {
  const core = loadCore();
  const first = core.apply(current({ tags: ['A'], status: '' }), preset({
    tagsMode: 'set', tags: ['B'], colorMode: 'set', color: '#009688'
  }));
  const second = core.apply(first, preset({
    tagsMode: 'clear', statusMode: 'set', status: '完了'
  }));
  const third = core.apply(second, preset({ statusMode: 'clear' }));

  assert.deepEqual(plain(first.tags), ['B']);
  assert.equal(first.status, '');
  assert.deepEqual(plain(second.tags), []);
  assert.equal(second.color, '#009688');
  assert.equal(second.status, '完了');
  assert.equal(third.status, '');
});

test('apply rejects disabled presets, unknown modes, and unsafe values without mutation', () => {
  const core = loadCore();
  const source = current();
  const before = plain(source);
  const invalidPresets = [
    preset({ enabled: false }),
    preset({ tagsMode: 'append' }),
    preset({ colorMode: 'clear' }),
    preset({ iconMode: 'clear' }),
    preset({ statusMode: 'unknown' }),
    preset({ colorMode: 'set', color: 'red' }),
    preset({ iconMode: 'set', icon: '<img>' }),
    preset({ statusMode: 'set', status: '未知' }),
    preset({ tagsMode: 'set', tags: 'not-an-array' }),
    preset({ tagsMode: 'set', tags: ['1', '2', '3', '4', '5', '6'] })
  ];

  invalidPresets.forEach((invalid) => assert.throws(() => core.apply(source, invalid)));
  assert.throws(() => core.apply(current({ color: 'javascript:alert(1)' }), preset()));
  assert.throws(() => core.apply(current({ icon: 'unknown' }), preset()));
  assert.throws(() => core.apply(current({ status: 'unknown' }), preset()));
  assert.deepEqual(plain(source), before);
});

test('blank status is valid current input and preset references are never retained', () => {
  const core = loadCore();
  const tags = ['入力'];
  const sourcePreset = preset({ tagsMode: 'set', tags });
  const result = core.apply(current({ status: '' }), sourcePreset);

  tags.push('後から追加');
  sourcePreset.tags[0] = '後から変更';
  assert.deepEqual(plain(result.tags), ['入力']);
  assert.equal(result.status, '');
  assert.equal(Object.prototype.hasOwnProperty.call(result, 'presetId'), false);
});

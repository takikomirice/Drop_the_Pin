const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function loadPreviewCore() {
  const start = indexHtml.indexOf('const InputPresetApplyCore = (function(');
  const end = indexHtml.indexOf('    const GeoJsonTrackInterchangeCore =', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = {
    PIN_COLORS: [{ hex: '#2196f3', label: '青' }],
    PIN_ICONS: [{ id: 'photo', label: '写真' }],
    PIN_STATUSES: ['未対応'],
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    Date,
    Math,
    Map,
    Set
  };
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__previewCore = typeof PhotoTrackMatchPreviewCore === "undefined" ? null : PhotoTrackMatchPreviewCore;',
    context
  );
  assert.ok(context.__previewCore, 'Expected PhotoTrackMatchPreviewCore');
  return context.__previewCore;
}

test('UTC offset parser accepts strict signed HH:MM values through the supported range', () => {
  const core = loadPreviewCore();
  const accepted = new Map([
    ['+09:00', 540], ['-08:00', -480], ['+05:30', 330], ['+05:45', 345],
    ['+14:00', 840], ['-14:00', -840], ['+00:00', 0], ['-00:30', -30]
  ]);
  accepted.forEach((minutes, text) => {
    assert.equal(core.parseUtcOffsetText(text), minutes, text);
    assert.equal(core.formatUtcOffsetMinutes(minutes), text === '-00:30' ? '-00:30' : text);
  });
  assert.equal(core.parseUtcOffsetText('-00:00'), 0);
  assert.equal(core.formatUtcOffsetMinutes(core.parseUtcOffsetText('-00:00')), '+00:00');
});

test('UTC offset parser rejects non-canonical, out-of-range, and padded values', () => {
  const core = loadPreviewCore();
  [
    '+14:01', '-14:01', '+15:00', '-15:00', '09:00', '540', '+9:00',
    '+09:0', '+09:60', ' +09:00', '+09:00 ', '', null, 540
  ].forEach((value) => assert.equal(core.parseUtcOffsetText(value), null, String(value)));
  [841, -841, 1.5, NaN, Infinity, '540', null].forEach((value) => {
    assert.equal(core.formatUtcOffsetMinutes(value), '');
  });
});

test('local offset default is formatted only inside the supported range and otherwise uses +09:00', () => {
  const core = loadPreviewCore();
  assert.equal(core.defaultUtcOffsetText(() => -540), '+09:00');
  assert.equal(core.defaultUtcOffsetText(() => 480), '-08:00');
  assert.equal(core.defaultUtcOffsetText(() => -900), '+09:00');
  assert.equal(core.defaultUtcOffsetText(() => NaN), '+09:00');
  assert.equal(core.defaultUtcOffsetText(() => { throw new Error('unavailable'); }), '+09:00');
});

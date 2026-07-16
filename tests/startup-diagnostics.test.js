const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const readme = fs.readFileSync(path.join(root, 'README.md'), 'utf8');

function functionSource(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let cursor = open; cursor < source.length; cursor += 1) {
    const character = source[cursor];
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
    if (character === '}' && --depth === 0) return source.slice(start, cursor + 1);
  }
  assert.fail(`Could not parse function ${name}`);
}

function functionBody(source, name) {
  const text = functionSource(source, name);
  return text.slice(text.indexOf('{') + 1, -1);
}

function noArgGasCalls(name) {
  return Array.from(functionBody(indexHtml, name).matchAll(/withGASNoArg\('([^']+)'\)/g), (match) => match[1]);
}

function createDiagnosticsHarness(options = {}) {
  let clock = 0;
  const logs = [];
  const context = {
    Promise,
    Object,
    Math,
    String,
    Number,
    URLSearchParams,
    performance: { now: () => ++clock },
    console: { log: (message) => logs.push(String(message)) },
    documentApi: options.documentApi || null
  };
  vm.runInNewContext(`
    ${functionSource(indexHtml, 'createStartupDiagnostics')}
    this.diagnostics = createStartupDiagnostics({
      performanceApi: performance,
      consoleApi: console,
      documentApi: documentApi,
      navigatorApi: null,
      enabled: ${options.enabled === true}
    });
  `, context);
  return { diagnostics: context.diagnostics, logs };
}

function createPanelDocument() {
  const listeners = {};
  const elements = {
    'startup-diagnostics': { hidden: true },
    'startup-diagnostics-stages': {
      children: [],
      replaceChildren() { this.children = []; },
      appendChild(child) { this.children.push(child); }
    },
    'startup-diagnostics-total': { textContent: '' },
    'startup-diagnostics-close': {
      addEventListener(type, handler) { listeners[`close:${type}`] = handler; }
    },
    'startup-diagnostics-copy': {
      addEventListener(type, handler) { listeners[`copy:${type}`] = handler; }
    },
    'startup-diagnostics-copy-status': { textContent: '' }
  };
  return {
    elements,
    listeners,
    getElementById: (id) => elements[id] || null,
    createElement: () => ({
      className: '', textContent: '', children: [],
      setAttribute() {},
      appendChild(child) { this.children.push(child); }
    })
  };
}

test('startup diagnostics panel exists only behind edit access plus debugStartup=1', () => {
  assert.match(indexHtml, /id="startup-diagnostics"[^>]*hidden/);
  for (const id of [
    'startup-diagnostics-stages', 'startup-diagnostics-total',
    'startup-diagnostics-copy', 'startup-diagnostics-close', 'startup-diagnostics-copy-status'
  ]) assert.match(indexHtml, new RegExp(`id="${id}"`));

  const enabledSource = functionSource(indexHtml, 'isStartupDiagnosticsEnabled');
  const context = { URLSearchParams };
  vm.runInNewContext(`${enabledSource}\nthis.enabled = isStartupDiagnosticsEnabled;`, context);
  assert.equal(context.enabled(true, '?debugStartup=1'), true);
  assert.equal(context.enabled(true, '?debugStartup=0'), false);
  assert.equal(context.enabled(true, ''), false);
  assert.equal(context.enabled(false, '?debugStartup=1'), false);
  assert.match(indexHtml, /isStartupDiagnosticsEnabled\(hasEditToken,\s*window\.location\.search\)/);

  const hiddenDocument = createPanelDocument();
  createDiagnosticsHarness({ enabled: false, documentApi: hiddenDocument });
  assert.equal(hiddenDocument.elements['startup-diagnostics'].hidden, true);
  const visibleDocument = createPanelDocument();
  createDiagnosticsHarness({ enabled: true, documentApi: visibleDocument });
  assert.equal(visibleDocument.elements['startup-diagnostics'].hidden, false);
});

test('client diagnostics uses performance.now and contains every fixed startup stage', () => {
  const source = functionSource(indexHtml, 'createStartupDiagnostics');
  assert.match(source, /performanceApi\.now\(\)/);
  for (const stage of [
    'app-start', 'map-ready', 'settings-load', 'pins-load', 'pins-render',
    'routes-load', 'tracks-load', 'presets-load', 'pin-add-ready'
  ]) assert.match(source, new RegExp(`['"]${stage}['"]`), stage);
  assert.match(source, /\[startup\] stage=/);
  assert.match(source, /result=/);
  assert.match(source, /durationMs=/);
});

test('pin-add-ready is recorded once at the first formal enable', () => {
  const refresh = functionBody(indexHtml, 'refreshPinAddButtonState');
  assert.match(refresh, /ready[\s\S]*startupDiagnostics\.recordPinAddReady\(\)/);

  const { diagnostics, logs } = createDiagnosticsHarness();
  diagnostics.recordPinAddReady();
  diagnostics.recordPinAddReady();
  const records = JSON.parse(JSON.stringify(diagnostics.getRecords()));
  const readyRecords = records.filter((record) => record.stage === 'pin-add-ready');
  assert.equal(readyRecords.length, 1);
  assert.equal(readyRecords[0].result, 'success');
  assert.equal(logs.filter((line) => line.includes('stage=pin-add-ready')).length, 1);
});

test('rejected startup work closes its stage as failure and preserves rejection', async () => {
  const { diagnostics } = createDiagnosticsHarness();
  const failure = new Error('secret raw exception');
  await assert.rejects(
    diagnostics.measurePromise('pins-load', () => Promise.reject(failure)),
    (error) => error === failure
  );
  const record = JSON.parse(JSON.stringify(diagnostics.getRecords()))
    .find((entry) => entry.stage === 'pins-load');
  assert.equal(record.result, 'failure');
  assert.equal(typeof record.durationMs, 'number');
});

test('copy output and client logger contain fixed stage result and duration fields only', () => {
  const source = functionSource(indexHtml, 'createStartupDiagnostics');
  for (const unsafe of [
    /error\.message/i, /JSON\.stringify\([^)]*response/i, /fileName/i,
    /coordinates/i, /base64/i, /editToken/i, /__EDIT_TOKEN__/i
  ]) assert.doesNotMatch(source, unsafe);

  const { diagnostics } = createDiagnosticsHarness();
  diagnostics.recordPinAddReady();
  const copied = diagnostics.copyText();
  assert.match(copied, /^\[startup\] stage=/m);
  assert.doesNotMatch(copied, /secret|token|base64|latitude|longitude/i);
});

test('startup failure paths never log or display raw exceptions IDs or filenames', () => {
  const initialize = functionBody(indexHtml, 'initializeApp');
  const mapLoad = functionBody(indexHtml, 'loadMapData');
  const tracksLoad = functionBody(indexHtml, 'loadTracks');
  for (const source of [initialize, mapLoad, tracksLoad]) {
    assert.doesNotMatch(source, /console\.(?:error|warn)\([^)]*,\s*(?:error|caught|reason|trackId|safeTrackId)/);
    assert.doesNotMatch(source, /error\.message|String\(error\)/);
  }
  assert.doesNotMatch(initialize, /message:[^\n]*(?:error|caught|reason)/);
});

test('startup GAS calls keep their original count order and parallel relationship', () => {
  assert.deepEqual(noArgGasCalls('loadAppSettings'), ['getAppSettings']);
  assert.deepEqual(noArgGasCalls('loadMapData'), ['getMapData']);
  assert.deepEqual(noArgGasCalls('loadRouteGroups'), ['getRouteGroups']);
  assert.deepEqual(noArgGasCalls('loadTracks'), ['getTracks']);

  const initialize = functionBody(indexHtml, 'initializeApp');
  const settingsAt = initialize.indexOf('await loadAppSettings()');
  const pinsAt = initialize.indexOf('await loadMapData()');
  const tracksAt = initialize.indexOf('await loadTracks()');
  assert.ok(settingsAt !== -1 && settingsAt < pinsAt && pinsAt < tracksAt);
  assert.doesNotMatch(initialize, /listInputPresets|loadInputPresetCatalog|ensureInputPresetsLoaded/);

  const mapLoad = functionBody(indexHtml, 'loadMapData');
  assert.match(mapLoad, /Promise\.allSettled\(\[[\s\S]*withGASNoArg\('getMapData'\)[\s\S]*loadRouteGroups\(\)[\s\S]*\]\)/);
});

test('client stages wrap existing startup paths including lazy presets without new GAS calls', () => {
  assert.match(functionBody(indexHtml, 'loadAppSettings'), /measurePromise\('settings-load'/);
  assert.match(functionBody(indexHtml, 'loadMapData'), /measurePromise\('pins-load'/);
  assert.match(functionBody(indexHtml, 'loadRouteGroups'), /measurePromise\('routes-load'/);
  assert.match(functionBody(indexHtml, 'loadTracks'), /measurePromise\('tracks-load'/);
  assert.match(functionBody(indexHtml, 'loadInputPresetCatalog'), /measurePromise\('presets-load'/);
  assert.match(functionBody(indexHtml, 'loadMapData'), /measureSync\('pins-render'/);
});

test('GAS startup logging uses fixed safe fields and covers requested internal stages', () => {
  const logger = functionSource(codeJs, 'logStartupGasStage_');
  assert.match(logger, /\[startup\] gas=/);
  assert.match(logger, /stage=/);
  assert.match(logger, /result=/);
  assert.match(logger, /durationMs=/);
  assert.doesNotMatch(logger, /error|JSON\.stringify|pinId|fileId|fileName|coordinate|token|base64/i);

  const mapData = functionSource(codeJs, 'getMapData');
  for (const stage of ['sheet-read', 'row-convert', 'drive-info', 'response-build', 'total']) {
    assert.match(mapData, new RegExp(`['"]${stage}['"]`), stage);
  }
  const tracks = functionSource(codeJs, 'getTracks');
  for (const stage of [
    'sheets-read', 'tracks-read', 'track-segments-read',
    'json-restore', 'summary-calculate', 'response-build', 'total'
  ]) assert.match(tracks, new RegExp(`['"]${stage}['"]`), stage);
  assert.match(functionSource(codeJs, 'getRouteGroups'), /logStartupGasStage_/);
  assert.match(functionSource(codeJs, 'getAppSettings'), /logStartupGasStage_/);
});

test('shared viewing implementation remains free of startup diagnostics', () => {
  const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
  assert.doesNotMatch(sharedHtml, /startup-diagnostics|debugStartup|pin-add-ready/);
});

test('README documents safe startup diagnosis without claiming a speed optimization', () => {
  assert.match(readme, /debugStartup=1/);
  assert.match(readme, /pin-add-ready/);
  assert.match(readme, /GAS.*呼び出し.*追加しません/);
  assert.match(readme, /高速化.*実装していません/);
});

const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const fixture = (name) => fs.readFileSync(path.join(__dirname, 'fixtures', name), 'utf8');
const codeJs = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
const TRACKS_HEADERS = [
  'trackId', 'name', 'description', 'color', 'sourceType', 'sourceName', 'activeRevision',
  'payloadHash', 'segmentCount', 'pointCount', 'distanceMeters', 'minElevation', 'maxElevation',
  'startTime', 'endTime', 'boundsJson', 'createdAt', 'updatedAt', 'orderIndex', 'visible',
  'lineStyle', 'lineWidth'
];
const TRACK_SEGMENTS_HEADERS = [
  'trackId', 'revisionId', 'segmentIndex', 'chunkIndex', 'pointsJson', 'pointCount', 'createdAt', 'updatedAt'
];

function plain(value) { return JSON.parse(JSON.stringify(value)); }

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function sourceFunction(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected ${name}`);
  const open = source.indexOf('{', start);
  let depth = 0;
  for (let cursor = open; cursor < source.length; cursor += 1) {
    if (source[cursor] === '{') depth += 1;
    if (source[cursor] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, cursor + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

class TextNode {
  constructor(value) { this.nodeType = 3; this.nodeValue = value; }
  get textContent() { return this.nodeValue; }
}

class ElementNode {
  constructor(tagName, namespaceURI, attributes) {
    this.nodeType = 1;
    this.tagName = tagName;
    this.localName = tagName.split(':').pop();
    this.namespaceURI = namespaceURI;
    this.childNodes = [];
    this.attributes = attributes;
  }
  get children() { return this.childNodes.filter((node) => node.nodeType === 1); }
  get textContent() { return this.childNodes.map((node) => node.textContent).join(''); }
  getAttribute(name) { return Object.prototype.hasOwnProperty.call(this.attributes, name) ? this.attributes[name] : null; }
  getElementsByTagName(name) {
    const result = [];
    const visit = (node) => {
      if (node.nodeType !== 1) return;
      if (node.tagName === name || node.localName === name) result.push(node);
      node.childNodes.forEach(visit);
    };
    this.childNodes.forEach(visit);
    return result;
  }
}

class DocumentNode {
  constructor(rootElement) { this.documentElement = rootElement; }
  getElementsByTagName(name) {
    if (!this.documentElement) return [];
    const own = this.documentElement.tagName === name || this.documentElement.localName === name
      ? [this.documentElement] : [];
    return own.concat(this.documentElement.getElementsByTagName(name));
  }
}

function decodeXml(value) {
  return value.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function parserErrorDocument() {
  return new DocumentNode(new ElementNode('parsererror', 'http://www.mozilla.org/newlayout/xml/parsererror.xml', {}));
}

function parseFixtureXml(source) {
  const tokens = source.match(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<[^>]*>|[^<]+/g) || [];
  if (tokens.join('') !== source) return parserErrorDocument();
  const stack = [];
  let rootElement = null;
  for (const token of tokens) {
    if (token.startsWith('<?') || token.startsWith('<!--')) continue;
    if (token.startsWith('<![CDATA[')) {
      if (!stack.length) return parserErrorDocument();
      stack.at(-1).element.childNodes.push(new TextNode(token.slice(9, -3)));
      continue;
    }
    if (token.startsWith('</')) {
      const name = token.slice(2, -1).trim();
      if (!stack.length || stack.at(-1).element.tagName !== name) return parserErrorDocument();
      stack.pop();
      continue;
    }
    if (token.startsWith('<')) {
      if (/^<!/i.test(token)) return parserErrorDocument();
      const match = token.match(/^<([^\s/>]+)([\s\S]*?)(\/?)>$/);
      if (!match) return parserErrorDocument();
      const attributes = {};
      const attributePattern = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
      let attributeMatch;
      while ((attributeMatch = attributePattern.exec(match[2]))) {
        attributes[attributeMatch[1]] = decodeXml(attributeMatch[2] === undefined ? attributeMatch[3] : attributeMatch[2]);
      }
      const inherited = stack.length ? stack.at(-1).namespaces : {};
      const namespaces = Object.assign({}, inherited);
      Object.keys(attributes).forEach((name) => {
        if (name === 'xmlns') namespaces[''] = attributes[name];
        else if (name.startsWith('xmlns:')) namespaces[name.slice(6)] = attributes[name];
      });
      const prefix = match[1].includes(':') ? match[1].split(':')[0] : '';
      const element = new ElementNode(match[1], namespaces[prefix] || '', attributes);
      if (stack.length) stack.at(-1).element.childNodes.push(element);
      else if (rootElement) return parserErrorDocument();
      else rootElement = element;
      if (!match[3]) stack.push({ element, namespaces });
      continue;
    }
    if (stack.length) stack.at(-1).element.childNodes.push(new TextNode(decodeXml(token)));
    else if (token.trim()) return parserErrorDocument();
  }
  return stack.length || !rootElement ? parserErrorDocument() : new DocumentNode(rootElement);
}

class DOMParser {
  parseFromString(source) { return parseFixtureXml(source); }
}

function createDocument() {
  const elements = new Map();
  function element(id) {
    if (elements.has(id)) return elements.get(id);
    const classes = new Set();
    const listeners = Object.create(null);
    const value = {
      id, value: '', checked: false, disabled: false, textContent: '', style: {}, dataset: {}, children: [],
      classList: {
        add(name) { classes.add(name); }, remove(name) { classes.delete(name); },
        contains(name) { return classes.has(name); },
        toggle(name, force) { if (force) classes.add(name); else classes.delete(name); }
      },
      addEventListener(name, handler) { listeners[name] = handler; },
      appendChild(child) { this.children.push(child); return child; },
      replaceChildren(...children) { this.children = children; },
      setAttribute(name, content) { this[name] = String(content); },
      click() { if (listeners.click) return listeners.click({ target: this }); }
    };
    elements.set(id, value);
    return value;
  }
  return {
    getElementById(id) { return element(id); },
    createElement(tag) { return element(`${tag}-${elements.size + 1}`); }
  };
}

function loadModules(documentApi) {
  const start = indexHtml.indexOf('const ImportJobCore = (function() {');
  const end = indexHtml.indexOf('\n    const state = {', start);
  const context = {
    Promise, Date, Math, JSON, Object, Array, Number, String, RegExp, Set, URL, Blob, DOMParser,
    document: documentApi,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935', label: '赤' }, { hex: '#2196f3', label: '青' }],
    PIN_ICONS: [{ id: 'default', label: '標準' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    crypto: { randomUUID: (() => { let id = 0; return () => `uuid-${++id}`; })() }
  };
  vm.createContext(context);
  vm.runInContext(`${indexHtml.slice(start, end)}\n` +
    'globalThis.__modules = { geometry: TrackGeometryCore, ui: TrackFileImportUI, '
    + 'geojson: GeoJsonTrackImportAdapter, gpx: GpxTrackImportAdapter };', context);
  return context.__modules;
}

function createState() {
  return {
    owner: '', loading: false, requestToken: 0, batch: null, draft: null,
    submittedPayload: null, submittedPayloads: null, saveIndex: 0,
    saving: false, saved: false, errorCode: '', retryable: null, sourceKind: '', sourceName: ''
  };
}

function createWorkflow(options = {}) {
  const documentApi = options.documentApi || createDocument();
  const modules = loadModules(documentApi);
  const state = createState();
  const calls = [];
  const savedTracks = [];
  const root = modules.ui.create({
    state, adapters: [modules.geojson, modules.gpx], documentApi,
    canUse: () => true, canStartImport: () => true, isSettingsOpen: () => true,
    closeSettings: options.closeSettings || (() => {}),
    callGAS: options.callGAS || (async (method, payload) => {
      calls.push({ method, payload: plain(payload) });
      return { ok: true, track: { ...plain(payload), orderIndex: 0 } };
    }),
    withEditToken: (payload) => ({ ...plain(payload), __editToken: 'token' }),
    getTrackCount: options.getTrackCount || (() => 0),
    normalizeSavedTrack: modules.geometry.normalizeTrack,
    onSaved: (track) => {
      savedTracks.push(plain(track));
      if (options.onSaved) return options.onSaved(track, modules);
      return undefined;
    },
    onBusy: options.onBusy || (() => {})
  });
  return {
    root, gpx: root.forKind('gpx'), geojson: root.forKind('geojson'), state,
    calls, savedTracks, documentApi, modules
  };
}

function gpxFile(overrides = {}) {
  return {
    name: 'walk.gpx', type: 'application/gpx+xml', size: 512,
    text: async () => fixture('gpx-mixed.gpx'), ...overrides
  };
}

function interruptedGpx() {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<gpx version="1.1" creator="test" xmlns="http://www.topografix.com/GPX/1/1">',
    '<metadata><name>縦走</name></metadata><trk><trkseg>',
    '<trkpt lat="35.0000" lon="139.0000"><time>2026-07-20T00:00:00Z</time></trkpt>',
    '<trkpt lat="35.0001" lon="139.0001"><time>2026-07-20T00:00:05Z</time></trkpt>',
    '<trkpt lat="35.1000" lon="139.1000"><time>2026-07-20T04:00:05Z</time></trkpt>',
    '<trkpt lat="35.1001" lon="139.1001"><time>2026-07-20T04:00:10Z</time></trkpt>',
    '</trkseg></trk></gpx>'
  ].join('');
}

function threeStageGpx() {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<gpx version="1.1" creator="test" xmlns="http://www.topografix.com/GPX/1/1">',
    '<metadata><name>三日行程</name></metadata><trk><trkseg>',
    '<trkpt lat="35.0000" lon="139.0000"><time>2026-07-20T00:00:00Z</time></trkpt>',
    '<trkpt lat="35.0001" lon="139.0001"><time>2026-07-20T00:00:05Z</time></trkpt>',
    '<trkpt lat="35.1000" lon="139.1000"><time>2026-07-20T04:00:05Z</time></trkpt>',
    '<trkpt lat="35.1001" lon="139.1001"><time>2026-07-20T04:00:10Z</time></trkpt>',
    '<trkpt lat="35.2000" lon="139.2000"><time>2026-07-20T08:00:10Z</time></trkpt>',
    '<trkpt lat="35.2001" lon="139.2001"><time>2026-07-20T08:00:15Z</time></trkpt>',
    '</trkseg></trk></gpx>'
  ].join('');
}

function createTrackServer() {
  const audit = { driveCalls: 0 };
  function makeSheet(rows) {
    const sheet = {
      rows: rows.map((row) => row.slice()), formulas: rows.map((row) => row.map(() => '')),
      maxRows: Math.max(1, rows.length),
      getLastRow() {
        for (let index = this.rows.length - 1; index >= 0; index -= 1) {
          if ((this.rows[index] || []).some((value) => value !== '' && value != null)) return index + 1;
        }
        return 0;
      },
      getLastColumn() { return this.rows.reduce((max, row) => Math.max(max, row.length), 0); },
      getMaxRows() { return this.maxRows; },
      insertRowsAfter(_position, count) { this.maxRows += count; },
      getDataRange() {
        return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
      },
      getRange(row, column, numRows = 1, numColumns = 1) {
        return {
          getValue: () => ((sheet.rows[row - 1] || [])[column - 1] ?? ''),
          getValues: () => Array.from({ length: numRows }, (_, rowOffset) =>
            Array.from({ length: numColumns }, (_, columnOffset) =>
              ((sheet.rows[row - 1 + rowOffset] || [])[column - 1 + columnOffset] ?? ''))),
          getFormulas: () => Array.from({ length: numRows }, (_, rowOffset) =>
            Array.from({ length: numColumns }, (_, columnOffset) =>
              ((sheet.formulas[row - 1 + rowOffset] || [])[column - 1 + columnOffset] ?? ''))),
          setValue(value) { return this.setValues([[value]]); },
          setValues(values) {
            values.forEach((valuesRow, rowOffset) => {
              while (sheet.rows.length <= row - 1 + rowOffset) sheet.rows.push([]);
              while (sheet.formulas.length <= row - 1 + rowOffset) sheet.formulas.push([]);
              valuesRow.forEach((value, columnOffset) => {
                sheet.rows[row - 1 + rowOffset][column - 1 + columnOffset] = value;
                sheet.formulas[row - 1 + rowOffset][column - 1 + columnOffset] =
                  typeof value === 'string' && value.startsWith('=') ? value : '';
              });
            });
            sheet.maxRows = Math.max(sheet.maxRows, row - 1 + values.length);
            return this;
          }
        };
      }
    };
    return sheet;
  }
  const originalRows = {
    routes: [['routeId'], ['route-a']], route_pins: [['routeId'], ['route-a']],
    route_cache: [['cacheKey'], ['cache-a']], map_info: [['ID'], ['pin-a']]
  };
  const sheets = {
    tracks: makeSheet([TRACKS_HEADERS]), track_segments: makeSheet([TRACK_SEGMENTS_HEADERS]),
    routes: makeSheet(originalRows.routes), route_pins: makeSheet(originalRows.route_pins),
    route_cache: makeSheet(originalRows.route_cache), map_info: makeSheet(originalRows.map_info)
  };
  const properties = new Map();
  const context = {
    Logger: { log() {} },
    CacheService: { getScriptCache: () => ({ get: (key) => key === 'EDIT_TOKEN_token' ? '1' : null }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty(key) { return properties.has(key) ? properties.get(key) : null; },
      setProperty(key, value) { properties.set(key, String(value)); },
      deleteProperty(key) { properties.delete(key); }
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: (name) => sheets[name] || null }), flush() {}
    },
    DriveApp: new Proxy({}, { get() { audit.driveCalls += 1; throw new Error('Drive must not be used'); } }),
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest(_algorithm, value) {
        return Array.from(crypto.createHash('sha256').update(String(value)).digest())
          .map((byte) => byte > 127 ? byte - 256 : byte);
      },
      getUuid() { throw new Error('Track storage must not generate identifiers'); }
    }
  };
  vm.runInNewContext(`${codeJs}\nglobalThis.__trackApi = { saveTrackBundle, getTracks };`, context);
  return { api: context.__trackApi, audit, sheets, originalRows };
}

test('production wires GeoJSON and GPX adapters into one track workflow and save API', () => {
  assert.match(indexHtml, /const GeoJsonTrackImportAdapter\s*=\s*createTrackFileImportAdapter\(\{/);
  assert.match(indexHtml, /const GpxTrackImportAdapter\s*=\s*createTrackFileImportAdapter\(\{/);
  assert.match(indexHtml, /adapters:\s*\[\s*GeoJsonTrackImportAdapter,\s*GpxTrackImportAdapter\s*\]/s);
  const start = indexHtml.indexOf('const TrackFileImportUI =');
  const end = indexHtml.indexOf('\n    const GeoJsonTrackImportUI =', start);
  assert.notEqual(start, -1, 'Expected TrackFileImportUI');
  assert.notEqual(end, -1, 'Expected compatibility alias boundary');
  const source = indexHtml.slice(start, end);
  assert.equal((source.match(/saveTrackBundle/g) || []).length, 1);
  assert.match(source, /adapter\.toSavePayloads\(importState\.batch\)/);
  assert.match(source, /owner !== kind|owner === kind/);
  assert.match(source, /sourceKind/);
  assert.match(source, /sourceName/);
});

test('GPX launch handlers reset through the shared controller and shared preview actions are format neutral', () => {
  assert.match(indexHtml, /gpx-track-import-button[\s\S]{0,260}isProductionImportBusy\(\)[\s\S]{0,180}gpx-track-file-input/);
  assert.match(indexHtml, /gpx-track-file-input[\s\S]{0,220}handleTrackFileInputChange\('gpx', event\)/);
  assert.match(sourceFunction(indexHtml, 'handleTrackFileInputChange'), /controller\.handleFileSelected\(event\)/);
  for (const method of ['save', 'retry', 'discard', 'close']) {
    assert.match(indexHtml, new RegExp(`trackImportController\\.${method}\\(\\)`));
  }
});

test('button rendering covers route loading saving and compact idle labels', () => {
  const start = indexHtml.indexOf('function renderTrackImportBusy(');
  assert.notEqual(start, -1, 'Expected renderTrackImportBusy');
  const source = indexHtml.slice(start, indexHtml.indexOf('\n    function isCsvImportBusy', start));
  assert.match(source, /GPX取込準備中\.\.\./);
  assert.match(source, /GPX取込中\.\.\./);
  assert.match(source, /'取込'/);
  assert.match(source, /GeoJSON取込準備中\.\.\./);
});

test('GPX validates extension MIME and the exact 5MB boundary before File.text', async () => {
  for (const accepted of [
    { name: 'a.gpx', type: '' }, { name: 'a.bin', type: 'application/gpx+xml' },
    { name: 'a.gpx', type: 'application/octet-stream' }, { name: 'a.gpx', type: 'binary/octet-stream' },
    { name: 'a.xml', type: 'application/xml' }, { name: 'a.html', type: 'application/xml' },
    { name: 'a.xml', type: 'text/xml' }
  ]) {
    const setup = createWorkflow();
    assert.ok(await setup.gpx.importFile(gpxFile({ ...accepted, size: 5 * 1024 * 1024 })));
    assert.equal(setup.gpx.discard(), true);
  }
  const setup = createWorkflow();
  let reads = 0;
  const input = setup.documentApi.getElementById('gpx-track-file-input');
  input.value = 'C:\\fakepath\\walk.gpx';
  input.files = [gpxFile({ size: 5 * 1024 * 1024 + 1, text: async () => { reads += 1; return ''; } })];
  assert.equal(await setup.gpx.handleFileSelected({ target: input }), null);
  assert.equal(input.value, '');
  assert.equal(reads, 0);
  assert.equal(setup.state.owner, '');
  assert.equal(setup.state.errorCode, 'GPX_FILE_TOO_LARGE');
  assert.equal(setup.state.retryable, false);
  assert.equal(setup.documentApi.getElementById('gpx-track-operation-error').textContent, 'GPXファイルは5MB以内にしてください。');
});

test('GPX rejects missing malformed unreadable and unsafe files with safe non-retryable messages', async () => {
  const throwing = (property) => {
    const candidate = { name: 'walk.gpx', type: 'application/gpx+xml', size: 1, text: async () => '' };
    Object.defineProperty(candidate, property, { get() { throw new Error(`${property} getter failed`); } });
    return candidate;
  };
  const cases = [
    [null, 'GPX_FILE_TYPE_INVALID'],
    [throwing('name'), 'GPX_FILE_TYPE_INVALID'],
    [throwing('type'), 'GPX_FILE_TYPE_INVALID'],
    [throwing('size'), 'GPX_FILE_TYPE_INVALID'],
    [throwing('text'), 'GPX_FILE_READ_FAILED'],
    [{ name: 'walk.txt', type: 'text/plain', size: 1, text: async () => '' }, 'GPX_FILE_TYPE_INVALID'],
    [gpxFile({ size: undefined }), 'GPX_FILE_TOO_LARGE'],
    [gpxFile({ size: null }), 'GPX_FILE_TOO_LARGE'],
    [gpxFile({ size: '' }), 'GPX_FILE_TOO_LARGE'],
    [gpxFile({ size: false }), 'GPX_FILE_TOO_LARGE'],
    [gpxFile({ size: '1' }), 'GPX_FILE_TOO_LARGE'],
    [gpxFile({ size: -1 }), 'GPX_FILE_TOO_LARGE'],
    [{ name: 'walk.gpx', type: '', size: 1 }, 'GPX_FILE_READ_FAILED']
  ];
  for (const [file, errorCode] of cases) {
    const setup = createWorkflow();
    assert.equal(await setup.gpx.importFile(file), null);
    assert.equal(setup.state.owner, '');
    assert.equal(setup.state.errorCode, errorCode);
    assert.equal(setup.state.retryable, false);
    assert.doesNotMatch(setup.documentApi.getElementById('gpx-track-operation-error').textContent, /<gpx|stack|parsererror|139/i);
  }
  const invalidXml = createWorkflow();
  assert.equal(await invalidXml.gpx.importFile(gpxFile({ text: async () => '<gpx><secret>139</gpx>' })), null);
  assert.equal(invalidXml.state.errorCode, 'GPX_INVALID_XML');
  assert.equal(invalidXml.state.retryable, false);
  assert.equal(invalidXml.documentApi.getElementById('gpx-track-operation-error').textContent,
    'GPXファイルのXML形式を確認してください。');
});

test('a new owner clears the previous format error before its delayed read resolves', async () => {
  const setup = createWorkflow();
  assert.equal(await setup.gpx.importFile(gpxFile({ size: null })), null);
  assert.equal(setup.state.errorCode, 'GPX_FILE_TOO_LARGE');
  setup.documentApi.getElementById('gpx-track-operation-status').textContent = 'old gpx status';
  setup.documentApi.getElementById('geojson-track-operation-status').textContent = 'old geojson status';
  setup.documentApi.getElementById('geojson-track-operation-error').textContent = 'old geojson error';
  setup.documentApi.getElementById('geojson-track-operation-error').style.display = 'block';
  setup.documentApi.getElementById('track-import-preview-error').textContent = 'old preview error';
  setup.documentApi.getElementById('track-import-preview-error').style.display = 'block';
  const gate = deferred();
  const pending = setup.geojson.importFile({
    name: 'next.geojson', type: 'application/geo+json', size: 100, text: () => gate.promise
  });
  await Promise.resolve();
  assert.equal(setup.state.owner, 'geojson');
  assert.equal(setup.state.loading, true);
  assert.equal(setup.state.errorCode, '');
  assert.equal(setup.state.retryable, null);
  assert.equal(setup.state.sourceKind, 'geojson');
  for (const id of [
    'gpx-track-operation-status', 'geojson-track-operation-status',
    'gpx-track-operation-error', 'geojson-track-operation-error', 'track-import-preview-error'
  ]) {
    assert.equal(setup.documentApi.getElementById(id).textContent, '', id);
  }
  assert.equal(setup.documentApi.getElementById('gpx-track-operation-error').style.display, 'none');
  assert.equal(setup.documentApi.getElementById('geojson-track-operation-error').style.display, 'none');
  assert.equal(setup.documentApi.getElementById('track-import-preview-error').style.display, 'none');
  assert.equal(setup.geojson.invalidate(), true);
  gate.resolve('{}');
  assert.equal(await pending, null);

  setup.documentApi.getElementById('geojson-track-operation-status').textContent = 'old geojson status';
  setup.documentApi.getElementById('geojson-track-operation-error').textContent = 'old geojson error';
  setup.documentApi.getElementById('geojson-track-operation-error').style.display = 'block';
  const reverseGate = deferred();
  const reversePending = setup.gpx.importFile(gpxFile({ text: () => reverseGate.promise }));
  await Promise.resolve();
  assert.equal(setup.documentApi.getElementById('geojson-track-operation-status').textContent, '');
  assert.equal(setup.documentApi.getElementById('geojson-track-operation-error').textContent, '');
  assert.equal(setup.documentApi.getElementById('geojson-track-operation-error').style.display, 'none');
  assert.equal(setup.gpx.invalidate(), true);
  reverseGate.resolve(fixture('gpx-mixed.gpx'));
  assert.equal(await reversePending, null);
});

test('closeSettings failure clears GPX ownership and permits the next file', async () => {
  let failClose = true;
  const setup = createWorkflow({
    closeSettings() {
      if (failClose) throw new Error('settings close failed');
    }
  });
  assert.equal(await setup.gpx.importFile(gpxFile()), null);
  assert.equal(setup.state.owner, '');
  assert.equal(setup.state.loading, false);
  assert.equal(setup.state.draft, null);
  assert.equal(setup.state.submittedPayload, null);
  assert.equal(setup.documentApi.getElementById('track-import-preview-overlay').classList.contains('open'), false);
  failClose = false;
  assert.ok(await setup.gpx.importFile(gpxFile({ name: 'again.gpx' })));
});

test('renderPreview failure keeps settings visible, rolls back the preview, and permits the next file', async () => {
  const documentApi = createDocument();
  const settings = documentApi.getElementById('settings-overlay');
  settings.classList.add('open');
  const preview = documentApi.getElementById('track-import-preview-overlay');
  const originalAdd = preview.classList.add;
  let failRender = true;
  preview.classList.add = function(name) {
    if (failRender) throw new Error('preview render failed');
    return originalAdd.call(this, name);
  };
  const setup = createWorkflow({
    documentApi,
    closeSettings() { settings.classList.remove('open'); }
  });

  assert.equal(await setup.gpx.importFile(gpxFile()), null);
  assert.equal(settings.classList.contains('open'), true);
  assert.equal(preview.classList.contains('open'), false);
  assert.equal(setup.state.owner, '');
  assert.equal(setup.state.draft, null);
  assert.match(documentApi.getElementById('gpx-track-operation-error').textContent, /GPX/);

  failRender = false;
  assert.ok(await setup.gpx.importFile(gpxFile({ name: 'again.gpx' })));
  assert.equal(settings.classList.contains('open'), false);
  assert.equal(preview.classList.contains('open'), true);
});

test('busy observers cannot break ownership initialization or cleanup', async () => {
  let setup;
  assert.doesNotThrow(() => {
    setup = createWorkflow({ onBusy() { throw new Error('observer failed'); } });
  });
  assert.ok(await setup.gpx.importFile(gpxFile()));
  assert.equal(setup.gpx.discard(), true);
  assert.equal(setup.state.owner, '');
});

test('GPX preview shows safe summary stats warnings and editable metadata without XML or coordinates', async () => {
  const setup = createWorkflow();
  const xml = fixture('gpx-mixed.gpx');
  const draft = await setup.gpx.importFile(gpxFile({ name: 'C:\\fake\\mixed.gpx', text: async () => xml }));
  assert.ok(draft);
  assert.equal(setup.state.owner, 'gpx');
  assert.equal(setup.state.sourceKind, 'gpx');
  assert.equal(setup.state.sourceName, 'mixed.gpx');
  assert.equal(setup.documentApi.getElementById('track-import-preview-title').textContent, 'ルート取込プレビュー');
  assert.equal(setup.documentApi.getElementById('track-import-preview-format').textContent, 'GPX');
  assert.equal(setup.documentApi.getElementById('track-import-preview-source').textContent, 'mixed.gpx');
  assert.match(setup.documentApi.getElementById('track-import-preview-stats').textContent, /trk要素 1/);
  assert.match(setup.documentApi.getElementById('track-import-preview-stats').textContent, /rte要素 2/);
  assert.match(setup.documentApi.getElementById('track-import-preview-warnings').textContent, /ウェイポイント1件/);
  const rendered = [
    setup.documentApi.getElementById('track-import-preview-summary').textContent,
    setup.documentApi.getElementById('track-import-preview-stats').textContent,
    setup.documentApi.getElementById('track-import-preview-warnings').textContent
  ].join('\n');
  assert.doesNotMatch(rendered, /30\.0|139|<gpx|trkpt|xmlns|2026-/i);
  assert.equal(Object.values(setup.state).includes(xml), false);
  assert.deepEqual(plain(setup.modules.gpx.renderWarnings({ warnings: [
    { code: 'UNKNOWN', count: 3 },
    { code: 'GPX_WAYPOINTS_IGNORED', count: 0 },
    { code: 'GPX_POINTS_WITHOUT_TIME', count: NaN }
  ] })), []);

  assert.equal(setup.gpx.discard(), true);
  const geoJson = JSON.stringify({
    type: 'FeatureCollection', features: [{ type: 'Feature', properties: {},
      geometry: { type: 'LineString', coordinates: [[139, 35], [139.1, 35.1]] } }]
  });
  assert.ok(await setup.geojson.importFile({
    name: 'next.geojson', type: 'application/geo+json', size: 200, text: async () => geoJson
  }));
  assert.equal(setup.documentApi.getElementById('track-import-preview-stats').textContent, '');
  assert.equal(setup.documentApi.getElementById('track-import-preview-stats').style.display, 'none');
  assert.equal(setup.documentApi.getElementById('track-import-preview-warnings').textContent, '');
  assert.equal(setup.documentApi.getElementById('track-import-preview-warnings').style.display, 'none');
});

test('GPX preview lists every route created by an interruption and shows aggregate transform stats', async () => {
  const setup = createWorkflow();
  const batch = await setup.gpx.importFile(gpxFile({ text: async () => interruptedGpx() }));

  assert.equal(batch.drafts.length, 2);
  assert.equal(setup.state.batch.drafts.length, 2);
  assert.equal(setup.state.draft.name, '縦走(1/2)');
  const parts = setup.documentApi.getElementById('track-import-preview-parts');
  assert.equal(parts.style.display, '');
  assert.match(parts.textContent, /縦走\(1\/2\)/);
  assert.match(parts.textContent, /縦走\(2\/2\)/);
  assert.match(setup.documentApi.getElementById('track-import-preview-stats').textContent,
    /記録中断 1件.*生成ルート 2件/s);
  assert.match(setup.documentApi.getElementById('track-import-preview-stats').textContent,
    /時刻付きpoint 4/);
  assert.equal(setup.documentApi.getElementById('track-import-preview-points').textContent, '4');
});

test('GPX common preview edits are applied to every generated route and refresh part names', async () => {
  const setup = createWorkflow();
  await setup.gpx.importFile(gpxFile({ text: async () => interruptedGpx() }));
  setup.documentApi.getElementById('track-import-preview-name').value = '編集後';
  setup.documentApi.getElementById('track-import-preview-description').value = '共通説明';
  setup.documentApi.getElementById('track-import-preview-line-style').value = 'dashed';
  setup.documentApi.getElementById('track-import-preview-visible').checked = false;

  const batch = setup.gpx.syncDraftFromForm();

  assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), ['編集後(1/2)', '編集後(2/2)']);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.description)), ['共通説明', '共通説明']);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.lineStyle)), ['dashed', 'dashed']);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.visible)), [false, false]);
  assert.match(setup.documentApi.getElementById('track-import-preview-parts').textContent, /編集後\(1\/2\)/);
  assert.match(setup.documentApi.getElementById('track-import-preview-parts').textContent, /編集後\(2\/2\)/);
});

test('GPX generated routes are saved sequentially in displayed order', async () => {
  const requests = [];
  let activeRequests = 0;
  let maxActiveRequests = 0;
  const setup = createWorkflow({
    async callGAS(method, payload) {
      assert.equal(method, 'saveTrackBundle');
      activeRequests += 1;
      maxActiveRequests = Math.max(maxActiveRequests, activeRequests);
      await Promise.resolve();
      const submitted = plain(payload);
      delete submitted.__editToken;
      requests.push(submitted);
      activeRequests -= 1;
      return { ok: true, track: { ...submitted, orderIndex: requests.length - 1 } };
    }
  });
  await setup.gpx.importFile(gpxFile({ text: async () => interruptedGpx() }));

  assert.equal(await setup.root.save(), true);
  assert.equal(maxActiveRequests, 1);
  assert.deepEqual(requests.map((payload) => payload.name), ['縦走(1/2)', '縦走(2/2)']);
  assert.deepEqual(setup.savedTracks.map((track) => track.name), ['縦走(1/2)', '縦走(2/2)']);
  assert.equal(setup.state.saved, true);
  assert.equal(setup.state.saveIndex, 2);
});

test('GPX retry resumes at the failed generated route without resending completed routes', async () => {
  const requests = [];
  let secondAttempt = 0;
  const setup = createWorkflow({
    async callGAS(method, payload) {
      assert.equal(method, 'saveTrackBundle');
      const submitted = plain(payload);
      delete submitted.__editToken;
      requests.push(submitted);
      if (submitted.name === '三日行程(2/3)' && secondAttempt++ === 0) {
        throw new Error('response lost');
      }
      return { ok: true, track: { ...submitted, orderIndex: requests.length - 1 } };
    }
  });
  await setup.gpx.importFile(gpxFile({ text: async () => threeStageGpx() }));

  assert.equal(await setup.root.save(), false);
  assert.deepEqual(requests.map((payload) => payload.name), ['三日行程(1/3)', '三日行程(2/3)']);
  assert.deepEqual(setup.savedTracks.map((track) => track.name), ['三日行程(1/3)']);
  assert.equal(setup.state.saveIndex, 1);
  assert.equal(setup.state.retryable, true);
  assert.match(setup.documentApi.getElementById('track-import-preview-error').textContent,
    /2\/3件目で停止.*1件は保存済み/);
  const firstFailedPayload = requests[1];

  assert.equal(await setup.root.retry(), true);
  assert.deepEqual(requests.map((payload) => payload.name), [
    '三日行程(1/3)', '三日行程(2/3)', '三日行程(2/3)', '三日行程(3/3)'
  ]);
  assert.deepEqual(requests[2], firstFailedPayload);
  assert.deepEqual(setup.savedTracks.map((track) => track.name),
    ['三日行程(1/3)', '三日行程(2/3)', '三日行程(3/3)']);
  assert.equal(setup.state.saveIndex, 3);
  assert.equal(setup.state.saved, true);
  assert.equal(setup.state.submittedPayloads.some((payload) => '__editToken' in payload), false);
});

test('GPX batch capacity is checked before any generated route is saved', async () => {
  const rejected = createWorkflow({ getTrackCount: () => 99 });
  await rejected.gpx.importFile(gpxFile({ text: async () => interruptedGpx() }));
  assert.equal(await rejected.root.save(), false);
  assert.equal(rejected.calls.length, 0);
  assert.equal(rejected.state.errorCode, 'TRACK_LIMIT_EXCEEDED');
  assert.equal(rejected.state.retryable, false);
  assert.equal(rejected.state.submittedPayloads, null);

  const accepted = createWorkflow({ getTrackCount: () => 98 });
  await accepted.gpx.importFile(gpxFile({ text: async () => interruptedGpx() }));
  assert.equal(await accepted.root.save(), true);
  assert.equal(accepted.calls.length, 2);
});

test('owner invalidation prevents an old GPX promise from changing a newer GeoJSON import', async () => {
  const gate = deferred();
  const setup = createWorkflow();
  const oldRead = setup.gpx.importFile(gpxFile({ text: () => gate.promise }));
  assert.equal(setup.state.owner, 'gpx');
  assert.equal(setup.geojson.invalidate(), false);
  assert.equal(setup.gpx.invalidate(), true);
  const geoJson = JSON.stringify({
    type: 'FeatureCollection', features: [{ type: 'Feature', properties: {},
      geometry: { type: 'LineString', coordinates: [[139, 35], [139.1, 35.1]] } }]
  });
  const imported = await setup.geojson.importFile({
    name: 'new.geojson', type: 'application/geo+json', size: 200, text: async () => geoJson
  });
  let directError = null;
  if (!imported) {
    try { setup.modules.geojson.buildDraft(geoJson, { sourceName: 'new.geojson' }); }
    catch (error) { directError = error; }
  }
  assert.ok(imported, JSON.stringify({
    errorCode: setup.state.errorCode,
    message: setup.documentApi.getElementById('geojson-track-operation-error').textContent,
    directError: directError && String(directError.stack || directError)
  }));
  gate.resolve(fixture('gpx-mixed.gpx'));
  assert.equal(await oldRead, null);
  assert.equal(setup.state.owner, 'geojson');
  assert.equal(setup.state.draft.sourceType, 'geojson');
  assert.equal(setup.gpx.discard(), false);
  assert.equal(setup.geojson.discard(), true);
});

test('GPX response loss and malformed success retry the identical fixed payload until valid deduplication', async () => {
  const requests = [];
  let attempt = 0;
  const setup = createWorkflow({
    async callGAS(method, payload) {
      assert.equal(method, 'saveTrackBundle');
      const submitted = plain(payload);
      delete submitted.__editToken;
      requests.push(submitted);
      attempt += 1;
      if (attempt === 1) throw new Error('response lost');
      if (attempt === 2) return { ok: true };
      if (attempt === 3) return { ok: true, track: submitted };
      if (attempt === 4) return { ok: true, track: { ...submitted, orderIndex: -1 } };
      if (attempt === 5) return { ok: true, track: { ...submitted, orderIndex: '0' } };
      return { ok: true, deduplicated: true, track: { ...submitted, orderIndex: 0 } };
    }
  });
  await setup.gpx.importFile(gpxFile());
  setup.documentApi.getElementById('track-import-preview-name').value = 'Edited GPX';
  assert.equal(await setup.root.save(), false);
  assert.equal(setup.state.retryable, true);
  assert.equal(await setup.root.retry(), false);
  assert.equal(setup.state.retryable, true);
  assert.equal(await setup.root.retry(), false);
  assert.equal(setup.state.retryable, true);
  assert.equal(await setup.root.retry(), false);
  assert.equal(setup.state.retryable, true);
  assert.equal(await setup.root.retry(), false);
  assert.equal(setup.state.retryable, true);
  assert.equal(await setup.root.retry(), true);
  assert.equal(requests.length, 6);
  requests.slice(1).forEach((request) => assert.deepEqual(request, requests[0]));
  assert.equal(requests[0].name, 'Edited GPX');
  for (const key of ['orderIndex', 'summary', 'stats', 'warnings', '__editToken']) {
    assert.equal(Object.prototype.hasOwnProperty.call(requests[0], key), false, key);
  }
  assert.equal(setup.state.submittedPayload.__editToken, undefined);
  assert.equal(setup.savedTracks.length, 1);
  assert.equal(setup.root.close(), true);
  assert.equal(setup.state.owner, '');
  assert.equal(setup.state.draft, null);
  assert.equal(setup.state.submittedPayload, null);
  assert.equal(setup.state.sourceKind, '');
  assert.equal(setup.state.sourceName, '');
});

test('GPX Core common preview saveTrackBundle response loss retry and production upsert converge once', async () => {
  const server = createTrackServer();
  const requests = [];
  const clientState = { tracks: [], trackVisibilityOverrides: Object.create(null) };
  let layersRendered = 0;
  let panelRendered = 0;
  let loseFirstResponse = true;
  let modules;
  const upsertContext = {
    state: clientState,
    normalizeTrackForState(track) { return modules.geometry.normalizeTrack(track); },
    renderTrackLayers() { layersRendered += 1; },
    renderTrackPanel() { panelRendered += 1; }
  };
  vm.runInNewContext(`${sourceFunction(indexHtml, 'upsertTrack')}\nglobalThis.__upsertTrack = upsertTrack;`, upsertContext);
  const setup = createWorkflow({
    callGAS(method, payload) {
      assert.equal(method, 'saveTrackBundle');
      const submitted = plain(payload);
      delete submitted.__editToken;
      requests.push(submitted);
      const result = plain(server.api.saveTrackBundle(payload));
      if (loseFirstResponse) {
        loseFirstResponse = false;
        throw new Error('response lost after commit');
      }
      return result;
    },
    onSaved(track, loadedModules) {
      modules = loadedModules;
      return upsertContext.__upsertTrack(track);
    }
  });
  modules = setup.modules;
  const draft = await setup.gpx.importFile(gpxFile({ text: async () => fixture('gpx-1.1-multisegment.gpx') }));
  const expectedSegments = plain(draft.segments);
  setup.documentApi.getElementById('track-import-preview-name').value = '保存済みGPX';
  assert.equal(await setup.root.save(), false);
  assert.equal(await setup.root.retry(), true);
  assert.equal(requests.length, 2);
  assert.deepEqual(requests[1], requests[0]);
  assert.deepEqual(requests[0].segments, expectedSegments);
  assert.equal(Object.prototype.hasOwnProperty.call(requests[0], 'orderIndex'), false);
  assert.equal(clientState.tracks.length, 1);
  assert.equal(clientState.tracks[0].trackId, requests[0].trackId);
  assert.equal(clientState.tracks[0].revisionId, requests[0].revisionId);
  assert.deepEqual(plain(clientState.tracks[0].segments), expectedSegments);
  assert.equal(clientState.tracks[0].orderIndex, 0);
  assert.equal(layersRendered, 1);
  assert.equal(panelRendered, 1);
  assert.equal(server.sheets.tracks.getLastRow(), 2);
  assert.equal(server.sheets.track_segments.getLastRow(), 3);
  assert.equal(server.audit.driveCalls, 0);
  for (const name of ['routes', 'route_pins', 'route_cache', 'map_info']) {
    assert.deepEqual(server.sheets[name].rows, server.originalRows[name], name);
  }
  const stored = plain(server.api.getTracks());
  assert.equal(stored.tracks.length, 1);
  assert.equal(stored.tracks[0].trackId, requests[0].trackId);
  assert.equal(stored.tracks[0].revisionId, requests[0].revisionId);
  assert.equal(setup.root.close(), true);
});

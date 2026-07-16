const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(__dirname, '..', 'shared.html'), 'utf8');
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

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((done, fail) => { resolve = done; reject = fail; });
  return { promise, resolve, reject };
}

function createDocument() {
  const elements = new Map();
  function element(id) {
    if (elements.has(id)) return elements.get(id);
    const classes = new Set();
    const listeners = Object.create(null);
    const value = {
      id,
      value: '',
      checked: false,
      disabled: false,
      textContent: '',
      style: {},
      dataset: {},
      children: [],
      classList: {
        add(name) { classes.add(name); },
        remove(name) { classes.delete(name); },
        contains(name) { return classes.has(name); },
        toggle(name, force) { if (force) classes.add(name); else classes.delete(name); }
      },
      addEventListener(name, handler) { listeners[name] = handler; },
      appendChild(child) { this.children.push(child); return child; },
      replaceChildren(...children) { this.children = children; },
      setAttribute(name, content) { this[name] = String(content); },
      removeAttribute(name) { delete this[name]; },
      click() { if (listeners.click) return listeners.click({ target: this }); }
    };
    elements.set(id, value);
    return value;
  }
  return {
    getElementById(id) { return element(id); },
    createElement(tag) { return element(`${tag}-${elements.size + 1}`); },
    elements
  };
}

function loadModules(documentApi = createDocument()) {
  const start = indexHtml.indexOf('const ImportJobCore = (function() {');
  const end = indexHtml.indexOf('\n    const state = {', start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);
  const context = {
    console,
    Promise,
    Date,
    Math,
    Set,
    URL,
    Blob,
    document: documentApi,
    SAFE_COLOR_RE: /^#[0-9a-fA-F]{6}$/,
    PIN_COLORS: [{ hex: '#e53935', label: '赤' }, { hex: '#2196f3', label: '青' }],
    PIN_ICONS: [{ id: 'default', label: '標準' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    crypto: { randomUUID: (() => { let id = 0; return () => `uuid-${++id}`; })() }
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__trackModules = {'
      + 'core: GeoJsonTrackInterchangeCore, geometry: TrackGeometryCore, '
      + 'ui: typeof GeoJsonTrackImportUI === "undefined" ? null : GeoJsonTrackImportUI'
      + '};',
    context
  );
  assert.ok(context.__trackModules.ui, 'Expected GeoJsonTrackImportUI');
  return { ...context.__trackModules, documentApi };
}

function createState() {
  return {
    loading: false,
    requestToken: 0,
    draft: null,
    submittedPayload: null,
    saving: false,
    saved: false,
    errorCode: '',
    retryable: null
  };
}

function createTrackServer() {
  const audit = { driveCalls: 0, lockHeld: false };
  function makeSheet(rows) {
    const sheet = {
      rows: rows.map((row) => row.slice()),
      formulas: rows.map((row) => row.map(() => '')),
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
    routes: [['routeId'], ['route-a']],
    route_pins: [['routeId'], ['route-a']],
    route_cache: [['cacheKey'], ['cache-a']],
    map_info: [['ID'], ['pin-a']]
  };
  const sheets = {
    tracks: makeSheet([TRACKS_HEADERS]),
    track_segments: makeSheet([TRACK_SEGMENTS_HEADERS]),
    routes: makeSheet(originalRows.routes),
    route_pins: makeSheet(originalRows.route_pins),
    route_cache: makeSheet(originalRows.route_cache),
    map_info: makeSheet(originalRows.map_info)
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
    LockService: { getScriptLock: () => ({
      tryLock() { audit.lockHeld = true; return true; },
      releaseLock() { audit.lockHeld = false; }
    }) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getSheetByName: (name) => sheets[name] || null }),
      flush() {}
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

function geoJson() {
  return JSON.stringify({
    type: 'FeatureCollection',
    name: 'Morning walk',
    features: [{
      type: 'Feature',
      geometry: { type: 'LineString', coordinates: [[139, 35, 10], [140, 36, 20]] },
      properties: { coordTimes: ['2026-07-11T01:00:00Z', '2026-07-11T01:01:00Z'] }
    }]
  });
}

function file(overrides = {}) {
  return {
    name: 'walk.geojson',
    type: 'application/geo+json',
    size: 512,
    text: async () => geoJson(),
    ...overrides
  };
}

function createController(options = {}) {
  const documentApi = options.documentApi || createDocument();
  const modules = loadModules(documentApi);
  const state = options.state || createState();
  const calls = [];
  const preserved = [];
  const savedTracks = [];
  const controller = modules.ui.create({
    state,
    trackCore: modules.core,
    documentApi,
    canUse: options.canUse || (() => true),
    canStartImport: options.canStartImport || (() => true),
    isSettingsOpen: options.isSettingsOpen || (() => true),
    closeSettings(settings) { preserved.push(settings); },
    callGAS: options.callGAS || (async (method, payload) => {
      calls.push({ method, payload: plain(payload) });
      return { ok: true, deduplicated: false, track: { ...plain(payload), orderIndex: 0 } };
    }),
    withEditToken(payload) { return { ...plain(payload), __editToken: 'token' }; },
    normalizeSavedTrack: options.normalizeSavedTrack || modules.geometry.normalizeTrack,
    onSaved: options.onSaved || ((track) => { savedTracks.push(plain(track)); }),
    onBusy: options.onBusy || (() => {})
  });
  return { controller, state, calls, preserved, savedTracks, documentApi, modules };
}

test('Data exposes separate Point and route GeoJSON inputs while shared view exposes neither', () => {
  for (const id of [
    'geojson-export-button', 'geojson-import-button', 'geojson-file-input',
    'geojson-track-import-button', 'geojson-track-file-input',
    'geojson-track-operation-status', 'geojson-track-operation-error'
  ]) {
    assert.match(indexHtml, new RegExp(`id="${id}"`), id);
  }
  assert.match(indexHtml, /id="geojson-interchange-section"[\s\S]*?GeoJSON Point[\s\S]*?id="geojson-import-button"[^>]*>[\s\S]*?<span class="action-label">取込<\/span><\/button>/);
  assert.match(indexHtml, /id="geojson-track-import-section"[\s\S]*?GeoJSONルート[\s\S]*?id="geojson-track-import-button"[^>]*>[\s\S]*?<span class="action-label">取込<\/span><\/button>/);
  assert.match(indexHtml, /id="geojson-track-file-input"[^>]*accept="\.geojson,\.json,application\/geo\+json,application\/json"/);
  assert.doesNotMatch(sharedHtml, /geojson-track-|GeoJSONトラックインポート/);
  assert.doesNotMatch(sharedHtml, /track-import-preview-overlay/);
});

test('route preview is dedicated, editable, summary-only, and text-safe', () => {
  for (const id of [
    'track-import-preview-overlay', 'track-import-preview-name', 'track-import-preview-description',
    'track-import-preview-color', 'track-import-preview-line-style',
    'track-import-preview-visible', 'track-import-preview-summary', 'track-import-preview-source',
    'track-import-preview-segments', 'track-import-preview-points', 'track-import-preview-distance',
    'track-import-preview-time', 'track-import-preview-elevation', 'track-import-preview-state',
    'track-import-preview-error', 'track-import-preview-save', 'track-import-preview-retry',
    'track-import-preview-discard', 'track-import-preview-close'
  ]) {
    assert.match(indexHtml, new RegExp(`id="${id}"`), id);
  }
  assert.doesNotMatch(indexHtml, /track-import-preview-(?:coordinates|geojson)/);
  assert.doesNotMatch(indexHtml, /id="track-import-preview-line-width"/);
  assert.match(indexHtml, /id="track-import-preview-color"[^>]*class="color-palette"/);
  assert.match(indexHtml, /track-import-preview-name[^>]*maxlength="100"/);
  assert.match(indexHtml, /track-import-preview-description[^>]*maxlength="400"/);
});

test('track preview color swatches are rebuilt from PIN_COLORS without duplicates', async () => {
  const setup = createController();
  await setup.controller.importFile(file());
  const palette = setup.documentApi.getElementById('track-import-preview-color');
  assert.deepEqual(palette.children.map((button) => [
    button.dataset.pinColor, button.title, button['aria-label'], button['aria-pressed']
  ]), [
    ['#e53935', '赤', '赤', 'true'],
    ['#2196f3', '青', '青', 'false']
  ]);
  palette.children[1].click();
  assert.equal(palette.value, '#2196f3');
  assert.equal(palette.children[1]['aria-pressed'], 'true');
  assert.equal(setup.controller.discard(), true);
  await setup.controller.importFile(file({ name: 'second.geojson' }));
  assert.equal(palette.children.length, 2);
});

test('file selection resets immediately and rejects oversized files before File.text', async () => {
  const exact = createController();
  assert.ok(await exact.controller.importFile(file({ size: 2 * 1024 * 1024 })));
  assert.equal(exact.controller.discard(), true);

  const { controller, state, documentApi } = createController();
  let reads = 0;
  const input = documentApi.getElementById('geojson-track-file-input');
  input.value = 'C:\\fakepath\\large.geojson';
  input.files = [file({ size: 2 * 1024 * 1024 + 1, text: async () => { reads += 1; return geoJson(); } })];
  const result = await controller.handleFileSelected({ target: input });
  assert.equal(input.value, '');
  assert.equal(result, null);
  assert.equal(reads, 0);
  assert.equal(state.loading, false);
  assert.equal(state.draft, null);
  assert.equal(state.errorCode, 'GEOJSON_FILE_TOO_LARGE');
  assert.equal(documentApi.getElementById('geojson-track-operation-error').textContent, 'GeoJSONファイルは2MB以内にしてください。');
});

test('one active read blocks another and settings close invalidates delayed results without retaining File or text', async () => {
  const gate = deferred();
  let settingsOpen = true;
  const setup = createController({ isSettingsOpen: () => settingsOpen });
  const firstFile = file({ text: () => gate.promise });
  const first = setup.controller.importFile(firstFile);
  assert.equal(setup.state.loading, true);
  assert.equal(await setup.controller.importFile(file()), null);
  settingsOpen = false;
  setup.controller.invalidate();
  gate.resolve(geoJson());
  assert.equal(await first, null);
  assert.equal(setup.state.loading, false);
  assert.equal(setup.state.draft, null);
  assert.equal(setup.state.submittedPayload, null);
  assert.equal(Object.values(setup.state).includes(firstFile), false);
  assert.equal(Object.values(setup.state).includes(geoJson()), false);
});

test('successful parsing opens preview through the sole preserve path and displays metadata without points', async () => {
  const setup = createController();
  const result = await setup.controller.importFile(file());
  assert.ok(result);
  assert.equal(setup.state.draft.trackId, 'uuid-1');
  assert.equal(setup.state.draft.revisionId, 'uuid-2');
  assert.deepEqual(plain(setup.preserved), [{ preserveTrackImport: true, restoreFocus: false }]);
  assert.equal(setup.documentApi.getElementById('track-import-preview-overlay').classList.contains('open'), true);
  assert.equal(setup.documentApi.getElementById('track-import-preview-name').value, 'Morning walk');
  assert.equal(setup.documentApi.getElementById('track-import-preview-source').textContent, 'walk.geojson');
  assert.match(setup.documentApi.getElementById('track-import-preview-summary').textContent, /1 segment/);
  assert.match(setup.documentApi.getElementById('track-import-preview-summary').textContent, /2 point/);
  assert.doesNotMatch(setup.documentApi.getElementById('track-import-preview-summary').textContent, /139|140/);
  assert.equal(setup.state.submittedPayload, null);
});

test('metadata remains editable until first save then the submitted payload and controls are fixed', async () => {
  const saveGate = deferred();
  const requests = [];
  const setup = createController({
    callGAS(method, payload) {
      requests.push({ method, payload: plain(payload) });
      return saveGate.promise;
    }
  });
  await setup.controller.importFile(file());
  setup.documentApi.getElementById('track-import-preview-name').value = 'Edited name';
  setup.documentApi.getElementById('track-import-preview-description').value = 'Edited description';
  setup.documentApi.getElementById('track-import-preview-color').value = '#2196f3';
  setup.documentApi.getElementById('track-import-preview-line-style').value = 'dashed';
  setup.documentApi.getElementById('track-import-preview-visible').checked = false;
  const saving = setup.controller.save();
  await Promise.resolve();
  assert.equal(setup.state.saving, true);
  assert.equal(setup.state.submittedPayload.name, 'Edited name');
  assert.equal(setup.state.submittedPayload.description, 'Edited description');
  assert.equal(setup.state.submittedPayload.color, '#2196f3');
  assert.equal(setup.state.submittedPayload.lineStyle, 'dashed');
  assert.equal(setup.state.submittedPayload.lineWidth, 4);
  assert.equal(setup.state.submittedPayload.visible, false);
  assert.equal(Object.prototype.hasOwnProperty.call(setup.state.submittedPayload, 'orderIndex'), false);
  for (const id of [
    'track-import-preview-name', 'track-import-preview-description', 'track-import-preview-color',
    'track-import-preview-line-style',
    'track-import-preview-visible', 'track-import-preview-save', 'track-import-preview-discard',
    'track-import-preview-close'
  ]) assert.equal(setup.documentApi.getElementById(id).disabled, true, id);
  assert.equal(setup.documentApi.getElementById('track-import-preview-color').children.every((button) => button.disabled), true);
  assert.equal(setup.controller.discard(), false);
  saveGate.resolve({ ok: true, track: { ...requests[0].payload, orderIndex: 0 } });
  assert.equal(await saving, true);
  assert.equal(setup.state.saved, true);
});

test('response loss retries the identical track payload and deduplicated success closes cleanly', async () => {
  const requests = [];
  const savedByRevision = new Map();
  const setup = createController({
    async callGAS(method, payload) {
      assert.equal(method, 'saveTrackBundle');
      const trackPayload = plain(payload);
      delete trackPayload.__editToken;
      requests.push(trackPayload);
      const key = `${trackPayload.trackId}:${trackPayload.revisionId}`;
      const existing = savedByRevision.get(key);
      if (!existing) {
        savedByRevision.set(key, trackPayload);
        throw new Error('response lost');
      }
      return { ok: true, deduplicated: true, track: { ...existing, orderIndex: 0 } };
    }
  });
  await setup.controller.importFile(file());
  assert.equal(await setup.controller.save(), false);
  assert.equal(setup.state.retryable, true);
  assert.equal(setup.state.saved, false);
  assert.equal(setup.documentApi.getElementById('track-import-preview-retry').style.display, '');
  assert.equal(await setup.controller.retry(), true);
  assert.equal(requests.length, 2);
  assert.deepEqual(requests[1], requests[0]);
  assert.equal(setup.state.saved, true);
  assert.equal(setup.savedTracks.length, 1);
  assert.equal(setup.savedTracks[0].trackId, requests[0].trackId);
  assert.equal(setup.controller.close(), true);
  assert.equal(setup.state.draft, null);
  assert.equal(setup.state.submittedPayload, null);
  assert.equal(setup.state.saved, false);
  assert.equal(setup.documentApi.getElementById('track-import-preview-overlay').classList.contains('open'), false);
});

test('GeoJSON UI and saveTrackBundle converge after response loss and append multiple tracks in import order', async () => {
  const server = createTrackServer();
  const requests = [];
  const clientTracks = [];
  let loseFirstResponse = true;
  const setup = createController({
    callGAS(method, payload) {
      assert.equal(method, 'saveTrackBundle');
      const submitted = plain(payload);
      delete submitted.__editToken;
      requests.push(submitted);
      const result = plain(server.api.saveTrackBundle(payload));
      if (loseFirstResponse) {
        loseFirstResponse = false;
        throw new Error('simulated response loss');
      }
      return result;
    },
    onSaved(track) {
      const index = clientTracks.findIndex((existing) => existing.trackId === track.trackId);
      if (index === -1) clientTracks.push(plain(track));
      else clientTracks[index] = plain(track);
      clientTracks.sort((left, right) => left.orderIndex - right.orderIndex
        || left.trackId.localeCompare(right.trackId));
    }
  });

  await setup.controller.importFile(file());
  const firstTrackId = setup.state.draft.trackId;
  const firstRevisionId = setup.state.draft.revisionId;
  assert.equal(await setup.controller.save(), false);
  const fixedPayload = plain(setup.state.submittedPayload);
  assert.equal(Object.prototype.hasOwnProperty.call(fixedPayload, 'orderIndex'), false);
  assert.equal(await setup.controller.retry(), true);
  assert.deepEqual(requests[1], requests[0]);
  assert.deepEqual(plain(setup.state.submittedPayload), fixedPayload);
  assert.equal(setup.controller.close(), true);

  await setup.controller.importFile(file({
    name: 'evening.geojson',
    text: async () => geoJson().replace('Morning walk', 'Evening walk')
  }));
  const secondTrackId = setup.state.draft.trackId;
  const secondRevisionId = setup.state.draft.revisionId;
  assert.notEqual(secondTrackId, firstTrackId);
  assert.notEqual(secondRevisionId, firstRevisionId);
  assert.equal(await setup.controller.save(), true);

  const stored = plain(server.api.getTracks());
  assert.equal(stored.ok, true);
  assert.deepEqual(stored.tracks.map((track) => [track.trackId, track.orderIndex]), [
    [firstTrackId, 0], [secondTrackId, 1]
  ]);
  assert.deepEqual(clientTracks.map((track) => track.trackId), [firstTrackId, secondTrackId]);
  assert.equal(server.sheets.tracks.rows.filter((row) => row[0]).length, 3);
  assert.equal(server.sheets.track_segments.rows.filter((row) => row[0]).length, 3);
  assert.equal(server.sheets.track_segments.rows.filter((row) => row[0] === firstTrackId).length, 1);
  assert.equal(server.sheets.track_segments.rows.filter((row) => row[0] === secondTrackId).length, 1);
  assert.equal(server.audit.driveCalls, 0);
  ['routes', 'route_pins', 'route_cache', 'map_info'].forEach((name) => {
    assert.deepEqual(server.sheets[name].rows, server.originalRows[name]);
  });
});

test('malformed successful responses keep the fixed payload retryable until a valid deduplicated track arrives', async () => {
  const malformedResults = [
    { ok: true },
    { ok: true, track: null },
    { ok: true, track: { id: 'x' } },
    { ok: true, track: { trackId: 'x', segments: 'invalid' } }
  ];
  for (const malformed of malformedResults) {
    const requests = [];
    let attempt = 0;
    const setup = createController({
      async callGAS(_method, payload) {
        const request = plain(payload);
        delete request.__editToken;
        requests.push(request);
        attempt += 1;
        return attempt === 1
          ? malformed
          : { ok: true, deduplicated: true, track: { ...request, orderIndex: 0 } };
      }
    });
    await setup.controller.importFile(file());
    assert.equal(await setup.controller.save(), false);
    assert.equal(setup.state.errorCode, 'TRACK_STORAGE_FAILED');
    assert.equal(setup.state.retryable, true);
    assert.equal(setup.state.saved, false);
    assert.ok(setup.state.submittedPayload);
    assert.equal(setup.documentApi.getElementById('track-import-preview-retry').style.display, '');
    assert.equal(setup.documentApi.getElementById('track-import-preview-name').disabled, true);
    assert.equal(await setup.controller.retry(), true);
    assert.deepEqual(requests[1], requests[0]);
    assert.equal(setup.state.saved, true);
  }

  const inherited = createController({
    async callGAS(_method, payload) {
      const prototype = plain(payload);
      delete prototype.__editToken;
      return { ok: true, track: Object.create(prototype) };
    }
  });
  await inherited.controller.importFile(file());
  assert.equal(await inherited.controller.save(), false);
  assert.equal(inherited.state.errorCode, 'TRACK_STORAGE_FAILED');
  assert.equal(inherited.state.retryable, true);
});

test('successful response requires an own track and exact canonical submitted content', async () => {
  const mismatches = [
    (request) => Object.assign(Object.create({ track: { ...request, orderIndex: 0 } }), { ok: true }),
    (request) => ({ ok: true, track: { ...request, name: 'Different name', orderIndex: 0 } }),
    (request) => ({ ok: true, track: { ...request, description: 'Different description', orderIndex: 0 } }),
    (request) => ({ ok: true, track: { ...request, color: '#2196f3', orderIndex: 0 } }),
    (request) => ({ ok: true, track: { ...request, sourceName: 'different.geojson', orderIndex: 0 } }),
    (request) => ({ ok: true, track: { ...request, visible: !request.visible, orderIndex: 0 } }),
    (request) => ({ ok: true, track: { ...request, lineStyle: 'dashed', orderIndex: 0 } }),
    (request) => ({ ok: true, track: { ...request, lineWidth: request.lineWidth + 1, orderIndex: 0 } }),
    (request) => {
      const changed = plain(request);
      changed.segments[0].points[0].lat += 0.01;
      return { ok: true, track: { ...changed, orderIndex: 0 } };
    }
  ];

  for (const mismatch of mismatches) {
    const requests = [];
    let attempt = 0;
    const setup = createController({
      async callGAS(_method, payload) {
        const request = plain(payload);
        delete request.__editToken;
        requests.push(request);
        attempt += 1;
        return attempt === 1
          ? mismatch(request)
          : { ok: true, deduplicated: true, track: { ...request, orderIndex: 0 } };
      }
    });
    await setup.controller.importFile(file());
    assert.equal(await setup.controller.save(), false);
    assert.equal(setup.state.saved, false);
    assert.equal(setup.state.errorCode, 'TRACK_STORAGE_FAILED');
    assert.equal(setup.state.retryable, true);
    assert.equal(await setup.controller.retry(), true);
    assert.deepEqual(requests[1], requests[0]);
  }
});

test('onSaved sync and async failures never turn an acknowledged save into failure', async () => {
  for (const onSaved of [
    () => { throw new Error('callback failed'); },
    () => Promise.reject(new Error('callback rejected'))
  ]) {
    const setup = createController({ onSaved });
    await setup.controller.importFile(file());
    assert.equal(await setup.controller.save(), true);
    await new Promise((resolve) => setImmediate(resolve));
    assert.equal(setup.state.saved, true);
    assert.equal(setup.state.errorCode, '');
  }
});

test('server errors use safe messages and retryability without exposing internals', async () => {
  const cases = [
    ['INVALID_TRACK_PAYLOAD', true, 'ルートの入力内容を確認してください。取込を破棄してファイルを選び直してください。'],
    ['TRACK_SHEETS_MISSING', false, '初期設定が必要です。スプレッドシートの設定メニューから初期設定を実行してください。'],
    ['TRACK_STORAGE_BUSY', true, 'ルート保存が混み合っています。時間をおいて再試行してください。'],
    ['TRACK_STORAGE_CORRUPTED', true, 'ルート保存領域が不整合です。管理者へ連絡してください。'],
    ['TRACK_REVISION_PAYLOAD_CONFLICT', false, 'このルートの保存内容が競合しました。取込を破棄してファイルを選び直してください。'],
    ['TRACK_LIMIT_EXCEEDED', false, '保存できるルート数の上限を超えています。'],
    ['TRACK_SEGMENT_LIMIT_EXCEEDED', false, 'ルートのsegment数が上限を超えています。'],
    ['TRACK_POINT_LIMIT_EXCEEDED', false, 'ルートのpoint数が上限を超えています。'],
    ['TRACK_PAYLOAD_INVALID', false, 'ルートの入力内容を確認してください。'],
    ['TRACK_STORAGE_FAILED', true, 'ルートを保存できませんでした。時間をおいて再試行してください。']
  ];
  for (const [errorCode, serverRetryable, message] of cases) {
    const setup = createController({
      callGAS: async () => ({
        ok: false, errorCode, retryable: serverRetryable,
        error: 'pointsJson internal_property stack trace secret'
      })
    });
    await setup.controller.importFile(file());
    assert.equal(await setup.controller.save(), false);
    assert.equal(setup.state.errorCode, errorCode);
    const retryable = errorCode === 'INVALID_TRACK_PAYLOAD' || errorCode === 'TRACK_STORAGE_CORRUPTED'
      ? false : serverRetryable;
    assert.equal(setup.state.retryable, retryable);
    assert.equal(setup.documentApi.getElementById('track-import-preview-error').textContent, message);
    assert.doesNotMatch(setup.documentApi.getElementById('track-import-preview-error').textContent, /pointsJson|stack|internal|secret/i);
    assert.equal(setup.documentApi.getElementById('track-import-preview-retry').disabled, !retryable);
  }
});

test('discard and preview initialization failure release all ownership and permit a later file', async () => {
  const setup = createController();
  await setup.controller.importFile(file());
  assert.equal(setup.controller.discard(), true);
  assert.equal(setup.state.draft, null);
  assert.equal(setup.state.submittedPayload, null);
  assert.equal(await setup.controller.importFile(file({ name: 'again.geojson' })) !== null, true);

  const brokenDocument = createDocument();
  const overlay = brokenDocument.getElementById('track-import-preview-overlay');
  overlay.classList.add = () => { throw new Error('DOM failed'); };
  const broken = createController({ documentApi: brokenDocument });
  assert.equal(await broken.controller.importFile(file()), null);
  assert.equal(broken.state.draft, null);
  assert.equal(broken.state.submittedPayload, null);
  assert.equal(broken.state.loading, false);
  assert.equal(broken.state.saving, false);
});

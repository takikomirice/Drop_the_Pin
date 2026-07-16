const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(__dirname, '..', 'shared.html'), 'utf8');

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

function loadModule() {
  const start = indexHtml.indexOf('const GeoJsonInterchangeUI = (function() {');
  const end = indexHtml.indexOf('\n    const state = {', start);
  assert.notEqual(start, -1, 'Expected GeoJsonInterchangeUI');
  assert.notEqual(end, -1, 'Expected state after GeoJsonInterchangeUI');
  const context = { Date, Promise, Blob };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\nglobalThis.__geoJsonUi = GeoJsonInterchangeUI;`,
    context
  );
  return context.__geoJsonUi;
}

function harness(overrides = {}) {
  const moduleApi = loadModule();
  const audit = {
    serialized: [], blobs: [], createdUrls: [], revokedUrls: [], anchors: [], removedAnchors: [],
    status: [], errors: [], previews: [], busy: []
  };
  const state = { loading: false, requestToken: 0, job: null, controller: null, registering: false };
  const pins = overrides.pins || [{ id: 'one' }, { id: 'two' }];
  const documentApi = {
    body: {
      appendChild(anchor) { audit.anchors.push(anchor); },
      removeChild(anchor) { audit.removedAnchors.push(anchor); }
    },
    createElement(tag) {
      assert.equal(tag, 'a');
      return {
        style: {}, href: '', download: '', clicked: false,
        click() {
          this.clicked = true;
          if (overrides.anchorClickError) throw new Error('download click failed');
        }
      };
    }
  };
  class BlobCtor {
    constructor(parts, options) {
      this.parts = parts;
      this.type = options.type;
      audit.blobs.push(this);
    }
  }
  let editable = overrides.editable !== false;
  let settingsOpen = overrides.settingsOpen !== false;
  const controller = moduleApi.create({
    state,
    canUse() { return editable; },
    canExport: overrides.canExport || (() => true),
    canStartImport: overrides.canStartImport || (() => true),
    isSettingsOpen() { return settingsOpen; },
    getPins() { return pins; },
    geoJsonCore: overrides.geoJsonCore || {
      serializePins(value) { audit.serialized.push(value); return '{"type":"FeatureCollection"}'; },
      buildImportJob(text) { return { id: `job:${text}`, sourceType: 'geojson', items: [{ id: 'item' }] }; }
    },
    documentApi,
    urlApi: {
      createObjectURL(blob) { audit.createdUrls.push(blob); return 'blob:geojson-download'; },
      revokeObjectURL(url) { audit.revokedUrls.push(url); }
    },
    BlobCtor,
    now() { return new Date('2026-07-11T01:02:03Z'); },
    onStatus(value) { audit.status.push(value); },
    onError(value) { audit.errors.push(value); },
    onBusy(value) { audit.busy.push(value); },
    onPreview(job) {
      audit.previews.push(job);
      if (typeof overrides.onPreview === 'function') return overrides.onPreview(job);
    }
  });
  return {
    moduleApi, controller, state, pins, audit,
    setEditable(value) { editable = value; },
    setSettingsOpen(value) { settingsOpen = value; }
  };
}

test('settings contains edit-only GeoJSON controls and shared view exposes no GeoJSON feature', () => {
  const ids = [
    'geojson-export-button', 'geojson-import-button', 'geojson-file-input',
    'geojson-operation-status', 'geojson-operation-error'
  ];
  ids.forEach((id) => {
    assert.match(indexHtml, new RegExp(`id=["']${id}["']`));
    assert.doesNotMatch(sharedHtml, new RegExp(`id=["']${id}["']`));
  });
  assert.match(indexHtml, /geojson-file-input[^>]+accept=["'][^"']*\.geojson[^"']*\.json[^"']*application\/geo\+json[^"']*application\/json/i);
  assert.doesNotMatch(sharedHtml, /GeoJsonPinInterchangeCore|GeoJsonInterchangeUI|saveImportPinItem/);
});

test('export snapshots every pin into a GeoJSON Blob, names it, clicks once, and revokes the URL', () => {
  const { controller, pins, audit } = harness();
  const before = JSON.stringify(pins);
  assert.equal(controller.exportPins(), true);
  assert.equal(audit.serialized.length, 1);
  assert.notEqual(audit.serialized[0], pins);
  assert.deepEqual(JSON.parse(JSON.stringify(audit.serialized[0])), pins);
  assert.equal(audit.blobs.length, 1);
  assert.equal(audit.blobs[0].type, 'application/geo+json;charset=utf-8');
  assert.deepEqual(JSON.parse(JSON.stringify(audit.blobs[0].parts)), ['{"type":"FeatureCollection"}']);
  assert.equal(audit.anchors[0].download, 'drop-the-pin-20260711-100203.geojson');
  assert.equal(audit.anchors[0].clicked, true);
  assert.deepEqual(audit.revokedUrls, ['blob:geojson-download']);
  assert.deepEqual(audit.removedAnchors, [audit.anchors[0]]);
  assert.equal(JSON.stringify(pins), before);
});

test('export cleanup runs exactly once on failure while zero pins still produce a Blob', () => {
  const failure = harness({ anchorClickError: true });
  assert.equal(failure.controller.exportPins(), false);
  assert.deepEqual(failure.audit.removedAnchors, [failure.audit.anchors[0]]);
  assert.deepEqual(failure.audit.revokedUrls, ['blob:geojson-download']);
  assert.match(failure.audit.errors.at(-1), /GeoJSON.*書き出せません/);

  const denied = harness({ editable: false });
  assert.equal(denied.controller.exportPins(), false);
  assert.equal(denied.audit.blobs.length, 0);
  assert.match(denied.audit.errors.at(-1), /編集モード/);

  const empty = harness({ pins: [] });
  assert.equal(empty.controller.exportPins(), true);
  assert.equal(empty.audit.blobs.length, 1);
  assert.deepEqual(empty.audit.serialized, [[]]);
  assert.equal(empty.audit.status.at(-1), 'GeoJSONを0件で書き出しました。');
  assert.deepEqual(empty.audit.revokedUrls, ['blob:geojson-download']);
  assert.deepEqual(empty.audit.removedAnchors, [empty.audit.anchors[0]]);

  const busy = harness({ canExport: () => false });
  assert.equal(busy.controller.exportPins(), false);
  assert.equal(busy.audit.blobs.length, 0);
  assert.match(busy.audit.errors.at(-1), /インポート|処理|完了/);
});

test('empty FeatureCollection completes without Preview Flow save or retained ownership', async () => {
  let flowCreates = 0;
  let saveCalls = 0;
  const current = harness({
    geoJsonCore: {
      serializePins() {
        return JSON.stringify({ type: 'FeatureCollection', dropThePinSchemaVersion: 1, features: [] });
      },
      buildImportJob(text) {
        if (text === 'next') return { id: 'next-job', sourceType: 'geojson', items: [{ id: 'item' }] };
        return { sourceType: 'geojson', items: [], warnings: [], empty: true };
      }
    },
    onPreview() {
      flowCreates += 1;
      saveCalls += 1;
    }
  });
  const firstInput = {
    value: 'C:\\fakepath\\empty.geojson',
    files: [{ name: 'empty.geojson', type: '', size: 100, text: () => Promise.resolve('empty') }]
  };
  const result = await current.controller.handleFileSelected({ target: firstInput });
  assert.equal(result.empty, true);
  assert.equal(firstInput.value, '');
  assert.equal(flowCreates, 0);
  assert.equal(saveCalls, 0);
  assert.deepEqual(current.state, {
    loading: false, requestToken: 1, job: null, controller: null, registering: false
  });
  assert.equal(current.audit.status.at(-1), '登録対象のピンは0件です');
  assert.equal(current.audit.errors.at(-1), '');

  await current.controller.importFile({
    name: 'next.geojson', type: '', size: 10, text: () => Promise.resolve('next')
  });
  assert.equal(flowCreates, 1);
  assert.equal(saveCalls, 1);
});

test('file selection resets immediately and accepts supported extensions and MIME types', async () => {
  const { controller, audit } = harness();
  const files = [
    { name: 'pins.geojson', type: '', value: 'one' },
    { name: 'pins.json', type: '', value: 'two' },
    { name: 'pins.txt', type: 'application/geo+json', value: 'three' },
    { name: 'pins.data', type: 'application/json', value: 'four' }
  ];
  for (const source of files) {
    const input = {
      value: `C:\\fakepath\\${source.name}`,
      files: [{
        name: source.name, type: source.type, size: 10,
        text: () => Promise.resolve(source.value)
      }]
    };
    const pending = controller.handleFileSelected({ target: input });
    assert.equal(input.value, '');
    await pending;
    controller.invalidate();
  }
  assert.equal(audit.previews.length, 4);
});

test('file validation rejects unsupported type and over 2MB before File.text', async () => {
  const { controller, audit } = harness();
  let reads = 0;
  await controller.importFile({
    name: 'pins.txt', type: 'text/plain', size: 10,
    text() { reads += 1; return Promise.resolve('x'); }
  });
  await controller.importFile({
    name: 'pins.geojson', type: 'application/geo+json', size: (2 * 1024 * 1024) + 1,
    text() { reads += 1; return Promise.resolve('x'); }
  });
  assert.equal(reads, 0);
  assert.equal(audit.previews.length, 0);
  const errors = audit.errors.filter(Boolean);
  assert.match(errors[0], /GeoJSON|JSONファイル/);
  assert.match(errors[1], /2MB/);
});

test('active File.text blocks another import and invalidation ignores a delayed result', async () => {
  const { controller, state, audit, setSettingsOpen } = harness();
  const first = deferred();
  let secondReads = 0;
  const firstRun = controller.importFile({
    name: 'first.geojson', type: '', size: 10, text: () => first.promise
  });
  await controller.importFile({
    name: 'second.geojson', type: '', size: 10,
    text() { secondReads += 1; return Promise.resolve('new'); }
  });
  assert.equal(secondReads, 0);
  assert.match(audit.errors.at(-1), /進行中|閉じ|完了/);
  first.resolve('old');
  await firstRun;
  assert.equal(audit.previews.length, 1);
  assert.equal(state.job.id, 'job:old');

  controller.invalidate();
  const delayed = deferred();
  const delayedRun = controller.importFile({
    name: 'delayed.json', type: '', size: 10, text: () => delayed.promise
  });
  setSettingsOpen(false);
  controller.invalidate();
  delayed.resolve('late secret contents');
  await delayedRun;
  assert.equal(audit.previews.length, 1);
  assert.equal(state.job, null);
  assert.equal(state.loading, false);
  assert.equal(Object.values(state).includes('late secret contents'), false);
});

test('preview handoff and parse failures are sanitized and release all GeoJSON ownership', async () => {
  const handoff = harness({ onPreview() { throw new Error('secret preview failure'); } });
  await handoff.controller.importFile({
    name: 'pins.geojson', type: '', size: 10, text: () => Promise.resolve('parsed')
  });
  assert.equal(handoff.state.job, null);
  assert.equal(handoff.state.controller, null);
  assert.equal(handoff.state.loading, false);
  assert.equal(handoff.state.registering, false);
  assert.equal(handoff.audit.errors.at(-1).includes('secret'), false);

  const secret = '<script>secret geojson</script>';
  const parse = harness({
    geoJsonCore: {
      serializePins() { return ''; },
      buildImportJob() {
        const error = new Error(secret);
        error.code = 'GEOJSON_INVALID_JSON';
        throw error;
      }
    }
  });
  await parse.controller.importFile({
    name: 'secret.json', type: '', size: 10, text: () => Promise.resolve(secret)
  });
  assert.equal(parse.state.job, null);
  assert.equal(Object.values(parse.state).includes(secret), false);
  assert.equal(parse.audit.errors.at(-1).includes(secret), false);
  assert.match(parse.audit.errors.at(-1), /GeoJSON.*内容|JSON構文/);
});

test('whole-file GeoJSON codes map to specific safe operation errors', async () => {
  const cases = [
    ['IMPORT_ITEM_LIMIT_EXCEEDED', /20件/],
    ['GEOJSON_FEATURES_REQUIRED', /1件以上/],
    ['GEOJSON_SCHEMA_VERSION_UNSUPPORTED', /形式|バージョン/],
    ['GEOJSON_UNSUPPORTED_CRS', /CRS/],
    ['GEOJSON_FILE_READ_FAILED', /読み込/]
  ];
  for (const [code, expected] of cases) {
    const current = harness({
      geoJsonCore: {
        serializePins() { return ''; },
        buildImportJob() { const error = new Error('secret'); error.code = code; throw error; }
      }
    });
    await current.controller.importFile({
      name: 'pins.geojson', type: '', size: 10, text: () => Promise.resolve('{}')
    });
    assert.match(current.audit.errors.at(-1), expected, code);
  }
});

test('GeoJSON state keeps only the parsed job and refuses conflicting production work before reading', async () => {
  const owned = harness();
  await owned.controller.importFile({
    name: 'first.geojson', type: '', size: 10, text: () => Promise.resolve('first')
  });
  assert.ok(owned.state.job);
  await owned.controller.importFile({
    name: 'second.geojson', type: '', size: 10, text: () => Promise.resolve('second')
  });
  assert.equal(owned.audit.previews.length, 1);
  assert.match(owned.audit.errors.at(-1), /進行中|閉じ|完了/);

  let reads = 0;
  const blocked = harness({ canStartImport: () => false });
  await blocked.controller.importFile({
    name: 'blocked.geojson', type: '', size: 10,
    text() { reads += 1; return Promise.resolve('secret'); }
  });
  assert.equal(reads, 0);
  assert.equal(blocked.audit.previews.length, 0);
  assert.match(blocked.audit.errors.at(-1), /別のインポート|進行中/);
});

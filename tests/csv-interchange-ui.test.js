const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(__dirname, '..', 'shared.html'), 'utf8');
const readme = fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8');

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
  const start = indexHtml.indexOf('const CsvInterchangeUI = (function() {');
  const end = indexHtml.indexOf('\n    const state = {', start);
  assert.notEqual(start, -1, 'Expected CsvInterchangeUI');
  assert.notEqual(end, -1, 'Expected state after CsvInterchangeUI');
  const context = { Date, Promise, Blob, setTimeout };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\nglobalThis.__csvUi = CsvInterchangeUI;`,
    context
  );
  return context.__csvUi;
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

function harness(overrides = {}) {
  const moduleApi = loadModule();
  const audit = {
    serialized: [], serializedReferences: [], blobs: [], createdUrls: [], revokedUrls: [], anchors: [],
    removedAnchors: [], status: [], errors: [], previews: [], busy: []
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
    csvCore: overrides.csvCore || {
      serializePins(value) { audit.serialized.push(value); return '\uFEFFcsv'; },
      serializeReference() { audit.serializedReferences.push(true); return '\uFEFFreference'; },
      buildImportJob(text) { return { id: `job:${text}`, items: [{ id: 'item' }] }; }
    },
    documentApi,
    urlApi: {
      createObjectURL(blob) { audit.createdUrls.push(blob); return 'blob:csv-download'; },
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

test('settings contains edit-only CSV controls and shared view exposes none', () => {
  const ids = [
    'csv-export-button', 'csv-import-button', 'csv-reference-button', 'csv-file-input',
    'csv-operation-status', 'csv-operation-error'
  ];
  ids.forEach((id) => {
    assert.match(indexHtml, new RegExp(`id=["']${id}["']`));
    assert.doesNotMatch(sharedHtml, new RegExp(`id=["']${id}["']`));
  });
  assert.match(indexHtml, /csv-file-input[^>]+accept=["'][^"']*\.csv[^"']*text\/csv/i);
  const referenceButton = indexHtml.match(/<button\b[^>]*id=["']csv-reference-button["'][^>]*>[\s\S]*?<\/button>/i);
  assert.ok(referenceButton);
  assert.match(referenceButton[0], /title=["']色・アイコン対応表をダウンロード["']/);
  assert.match(referenceButton[0], /aria-label=["']色・アイコン対応表をダウンロード["']/);
  assert.match(referenceButton[0], /class=["'][^"']*ghost-btn[^"']*["']/);
});

test('export snapshots every pin into a CSV Blob, names it, clicks once, and revokes the URL', () => {
  const { controller, pins, audit } = harness();
  const before = JSON.stringify(pins);
  assert.equal(controller.exportPins(), true);
  assert.equal(audit.serialized.length, 1);
  assert.notEqual(audit.serialized[0], pins);
  assert.deepEqual(JSON.parse(JSON.stringify(audit.serialized[0])), pins);
  assert.equal(audit.blobs.length, 1);
  assert.equal(audit.blobs[0].type, 'text/csv;charset=utf-8');
  assert.deepEqual(JSON.parse(JSON.stringify(audit.blobs[0].parts)), ['\uFEFFcsv']);
  assert.equal(audit.anchors[0].download, 'drop-the-pin-20260711-100203.csv');
  assert.equal(audit.anchors[0].clicked, true);
  assert.deepEqual(audit.revokedUrls, ['blob:csv-download']);
  assert.deepEqual(audit.removedAnchors, [audit.anchors[0]]);
  assert.equal(JSON.stringify(pins), before);
});

test('export failure removes its anchor and revokes its Object URL exactly once', () => {
  const { controller, audit } = harness({ anchorClickError: true });
  assert.equal(controller.exportPins(), false);
  assert.equal(audit.anchors.length, 1);
  assert.deepEqual(audit.removedAnchors, [audit.anchors[0]]);
  assert.deepEqual(audit.revokedUrls, ['blob:csv-download']);
  assert.match(audit.errors.at(-1), /書き出せません/);
});

test('reference download uses the fixed name, blocks a concurrent press, shows busy, and cleans up', async () => {
  const { controller, state, audit } = harness();
  const first = controller.downloadReference();
  assert.equal(state.referenceDownloading, true);
  const duplicate = controller.downloadReference();
  assert.equal(await duplicate, false);
  assert.equal(await first, true);
  assert.equal(state.referenceDownloading, false);
  assert.equal(audit.serializedReferences.length, 1);
  assert.equal(audit.blobs.length, 1);
  assert.equal(audit.blobs[0].type, 'text/csv;charset=utf-8');
  assert.deepEqual(JSON.parse(JSON.stringify(audit.blobs[0].parts)), ['\uFEFFreference']);
  assert.equal(audit.anchors[0].download, 'pin-csv-reference.csv');
  assert.equal(audit.anchors[0].clicked, true);
  assert.deepEqual(audit.revokedUrls, ['blob:csv-download']);
  assert.deepEqual(audit.removedAnchors, [audit.anchors[0]]);
  assert.equal(audit.status.at(-1), '色・アイコン対応表CSVを書き出しました。');
  assert.equal(audit.busy.length >= 2, true);
});

test('reference download failure releases busy state and cleans up exactly once', async () => {
  const { controller, state, audit } = harness({ anchorClickError: true });
  assert.equal(await controller.downloadReference(), false);
  assert.equal(state.referenceDownloading, false);
  assert.deepEqual(audit.removedAnchors, [audit.anchors[0]]);
  assert.deepEqual(audit.revokedUrls, ['blob:csv-download']);
  assert.match(audit.errors.at(-1), /対応表CSVを書き出せません/);
});

test('production busy renderer disables and labels the reference download button', () => {
  const body = sourceFunction(indexHtml, 'renderCsvInterchangeBusy');
  assert.match(body, /csv-reference-button/);
  assert.match(body, /referenceDownloading/);
  assert.match(body, /対応表作成中\.\.\./);
  assert.match(indexHtml, /csv-reference-button[^\n]+addEventListener\('click'/);
});

test('export is gated and zero pins produce a header-only Blob with normal cleanup', () => {
  const denied = harness({ editable: false });
  assert.equal(denied.controller.exportPins(), false);
  assert.equal(denied.audit.blobs.length, 0);
  assert.match(denied.audit.errors.at(-1), /編集モード/);

  const empty = harness({ pins: [] });
  assert.equal(empty.controller.exportPins(), true);
  assert.equal(empty.audit.blobs.length, 1);
  assert.deepEqual(empty.audit.serialized, [[]]);
  assert.equal(empty.audit.status.at(-1), 'CSVを0件で書き出しました。');
  assert.deepEqual(empty.audit.revokedUrls, ['blob:csv-download']);
  assert.deepEqual(empty.audit.removedAnchors, [empty.audit.anchors[0]]);

  const busy = harness({ canExport: () => false });
  assert.equal(busy.controller.exportPins(), false);
  assert.equal(busy.audit.blobs.length, 0);
  assert.match(busy.audit.errors.at(-1), /インポート|完了|処理/);
});

test('header-only CSV completes empty without Preview Flow save or retained ownership', async () => {
  let flowCreates = 0;
  let saveCalls = 0;
  const current = harness({
    csvCore: {
      serializePins() {
        return '\uFEFFschemaVersion,sourceId,title,description,lat,lng,color,icon,status,tags,eventAt,links';
      },
      buildImportJob(text) {
        if (text === 'next') return { id: 'next-job', items: [{ id: 'item' }] };
        return { sourceType: 'csv', items: [], warnings: [], empty: true };
      }
    },
    onPreview() {
      flowCreates += 1;
      saveCalls += 1;
    }
  });
  const firstInput = {
    value: 'C:\\fakepath\\empty.csv',
    files: [{ name: 'empty.csv', type: 'text/csv', size: 100, text: () => Promise.resolve('empty') }]
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
    name: 'next.csv', type: 'text/csv', size: 10, text: () => Promise.resolve('next')
  });
  assert.equal(flowCreates, 1);
  assert.equal(saveCalls, 1);
});

test('file selection resets the input immediately and accepts csv extension or text/csv MIME', async () => {
  const { controller, audit } = harness();
  const input = {
    value: 'C:\\fakepath\\pins.csv',
    files: [{ name: 'pins.csv', type: '', size: 10, text: () => Promise.resolve('first') }]
  };
  const pending = controller.handleFileSelected({ target: input });
  assert.equal(input.value, '');
  await pending;
  assert.equal(audit.previews.length, 1);

  controller.invalidate();
  await controller.importFile({
    name: 'pins.txt', type: 'text/csv', size: 10, text: () => Promise.resolve('second')
  });
  assert.equal(audit.previews.length, 2);
});

test('file validation rejects type and 2MB overflow before reading', async () => {
  const { controller, audit } = harness();
  let reads = 0;
  await controller.importFile({
    name: 'pins.txt', type: 'text/plain', size: 10,
    text() { reads += 1; return Promise.resolve('x'); }
  });
  await controller.importFile({
    name: 'pins.csv', type: 'text/csv', size: (2 * 1024 * 1024) + 1,
    text() { reads += 1; return Promise.resolve('x'); }
  });
  assert.equal(reads, 0);
  assert.equal(audit.previews.length, 0);
  const errors = audit.errors.filter(Boolean);
  assert.match(errors[0], /CSVファイル/);
  assert.match(errors[1], /2MB/);
});

test('an active File.text blocks a newer read and invalidation ignores delayed results after settings closes', async () => {
  const { controller, state, audit, setSettingsOpen } = harness();
  const first = deferred();
  let secondReads = 0;
  const firstRun = controller.importFile({
    name: 'first.csv', type: 'text/csv', size: 10, text: () => first.promise
  });
  const secondRun = controller.importFile({
    name: 'second.csv', type: 'text/csv', size: 10,
    text() { secondReads += 1; return Promise.resolve('new'); }
  });
  await secondRun;
  assert.equal(secondReads, 0);
  assert.match(audit.errors.at(-1), /進行中|完了|閉じ/);
  first.resolve('old');
  await firstRun;
  assert.equal(audit.previews.length, 1);
  assert.equal(state.job.id, 'job:old');

  controller.invalidate();
  const delayed = deferred();
  const delayedRun = controller.importFile({
    name: 'delayed.csv', type: 'text/csv', size: 10, text: () => delayed.promise
  });
  setSettingsOpen(false);
  controller.invalidate();
  delayed.resolve('late');
  await delayedRun;
  assert.equal(audit.previews.length, 1);
  assert.equal(state.job, null);
  assert.equal(state.loading, false);
});

test('preview handoff exceptions clear parsed ownership without retaining the CSV job', async () => {
  const created = [];
  const { controller, state, audit } = harness({
    onPreview(job) {
      created.push(job);
      throw new Error('preview initialization failed with secret details');
    }
  });
  await controller.importFile({
    name: 'pins.csv', type: 'text/csv', size: 10, text: () => Promise.resolve('parsed')
  });
  assert.equal(created.length, 1);
  assert.equal(state.job, null);
  assert.equal(state.controller, null);
  assert.equal(state.loading, false);
  assert.equal(state.registering, false);
  assert.equal(audit.errors.at(-1).includes('secret details'), false);
});

test('read and parse failures are sanitized without retaining File or CSV text', async () => {
  const secret = '<script>secret csv</script>';
  const { controller, state, audit } = harness({
    csvCore: {
      serializePins() { return ''; },
      buildImportJob() { const error = new Error(secret); error.code = 'CSV_ROW_INVALID'; throw error; }
    }
  });
  await controller.importFile({
    name: 'secret.csv', type: 'text/csv', size: 10, text: () => Promise.resolve(secret)
  });
  assert.equal(state.job, null);
  assert.equal(Object.values(state).includes(secret), false);
  assert.equal(audit.errors.at(-1).includes(secret), false);
  assert.match(audit.errors.at(-1), /CSVの内容/);
});

test('production CSV preview uses the pin processor and production flow with bounded settings', () => {
  assert.match(indexHtml, /function openCsvImportPreview\(job\)/);
  assert.match(indexHtml, /title:\s*'CSV取込を確認'/);
  assert.match(indexHtml, /sourceLabel:\s*'CSV '\s*\+\s*job\.items\.length\s*\+\s*'件'/);
  assert.match(indexHtml, /closePolicy:\s*'discard-only'/);
  assert.match(indexHtml, /timeFieldLabel:\s*'イベント時刻'/);
  assert.match(indexHtml, /completionItemLabel:\s*'ピン'/);
  assert.match(indexHtml, /ImportPinItemProcessor\.create\([\s\S]*?callGAS:\s*withGAS/);
  assert.match(indexHtml, /ImportFlowController\.create\([\s\S]*?concurrency:\s*2/);
  assert.match(indexHtml, /validateJob:\s*ImportPinDraftValidator\.validateJob/);
  assert.match(indexHtml, /onSaved:\s*upsertImportedPin/);
  assert.doesNotMatch(indexHtml, /function openCsvImportPreview[\s\S]{0,1800}hidePrimaryAction/);
});

test('CSV preview handoff atomically clears ownership for every initialization failure', () => {
  const scenarios = ['processor', 'controller', 'close', 'open'];
  scenarios.forEach((failure) => {
    const job = { id: 'csv-job', items: [{ id: 'csv-item' }] };
    const state = {
      csvInterchange: { loading: true, requestToken: 1, job, controller: null, registering: true }
    };
    const controller = {
      open() {
        if (failure === 'open') throw new Error('preview init failed');
        return job;
      }
    };
    const context = {
      state,
      isMultiPhotoImportBusyState: () => false,
      isGeoJsonImportBusy: () => false,
      ImportPinItemProcessor: {
        create() {
          if (failure === 'processor') throw new Error('processor create failed');
          return { processItem() {} };
        }
      },
      ImportPinDraftValidator: { validateJob() {} },
      ImportFlowController: {
        create() {
          if (failure === 'controller') throw new Error('controller create failed');
          return controller;
        }
      },
      closeOverlay() {
        if (failure === 'close') throw new Error('settings close failed');
      },
      withGAS() {}, withEditToken(value) { return value; }, upsertImportedPin() {},
      renderCsvInterchangeBusy() {}, reloadInputPresets() {}, ensureInputPresetsLoaded() {},
      getEnabledInputPresets() { return []; }
    };
    vm.runInNewContext(
      `${sourceFunction(indexHtml, 'resetInterchangePreviewOwnership')}\n`
        + `${sourceFunction(indexHtml, 'openCsvImportPreview')}\nthis.openCsv = openCsvImportPreview;`,
      context
    );
    assert.throws(() => context.openCsv(job));
    assert.equal(state.csvInterchange.job, null, failure);
    assert.equal(state.csvInterchange.controller, null, failure);
    assert.equal(state.csvInterchange.registering, false, failure);
    assert.equal(state.csvInterchange.loading, false, failure);
    assert.equal(state.csvInterchange.requestToken, 2, failure);
  });
});

test('CSV state keeps parsed jobs but refuses another import while a preview owns one', async () => {
  const first = harness();
  await first.controller.importFile({
    name: 'first.csv', type: 'text/csv', size: 10, text: () => Promise.resolve('first')
  });
  assert.ok(first.state.job);
  await first.controller.importFile({
    name: 'second.csv', type: 'text/csv', size: 10, text: () => Promise.resolve('second')
  });
  assert.equal(first.audit.previews.length, 1);
  assert.match(first.audit.errors.at(-1), /進行中|完了|閉じ/);
});

test('CSV import refuses a conflicting production workflow before File.text', async () => {
  let reads = 0;
  const blocked = harness({ canStartImport: () => false });
  await blocked.controller.importFile({
    name: 'blocked.csv', type: 'text/csv', size: 10,
    text() { reads += 1; return Promise.resolve('secret csv'); }
  });
  assert.equal(reads, 0);
  assert.equal(blocked.audit.previews.length, 0);
  assert.match(blocked.audit.errors.at(-1), /別のインポート|進行中/);
});

test('README summarizes CSV limits, export exclusions, and manual verification status', () => {
  assert.match(readme, /CSVピン[^\n]*2MB[^\n]*20データ行/);
  assert.match(readme, /CSV／GeoJSON Pointのエクスポートには写真[^\n]*Drive ID[^\n]*ルート[^\n]*含めません/);
  assert.match(readme, /CSV／GeoJSON Pointの1件／20件取込[^\n]*上限超過[^\n]*不正項目削除[^\n]*エクスポート/);
  assert.match(readme, /Nodeテスト[^\n]*Chromiumテスト[^\n]*Apps Script[^\n]*実Drive[^\n]*未実施/);
});

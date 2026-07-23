const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function functionSource(source, name) {
  let start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  if (source.slice(Math.max(0, start - 6), start) === 'async ') start -= 6;
  const bodyStart = source.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < source.length; index += 1) {
    const character = source[index];
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
    if (character === '}') {
      depth -= 1;
      if (depth === 0) return source.slice(start, index + 1);
    }
  }
  assert.fail(`Could not parse ${name}`);
}

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((resolvePromise, rejectPromise) => {
    resolve = resolvePromise;
    reject = rejectPromise;
  });
  return { promise, resolve, reject };
}

async function flush() {
  await Promise.resolve();
  await Promise.resolve();
}

test('production editor Promise rejects exactly once on native destroy and restores editor mode', async () => {
  const calls = [];
  let destroyHandler = null;
  const editor = {
    init() { calls.push(['init']); },
    loadFile() { return Promise.resolve({ mode: 'editing' }); },
    getState() { return { totalDuration: 30 }; },
    setResultHandler(handler) { calls.push(['result', typeof handler]); },
    setDestroyHandler(handler) { destroyHandler = handler; calls.push(['destroy', typeof handler]); },
    setImportMode(value) { calls.push(['mode', value]); }
  };
  const context = {
    window: { HemisphereAudioEditor: editor },
    state: { audioImport: { editorReject: null } },
    Promise,
    Error,
    Number,
    Object
  };
  vm.runInNewContext(`${functionSource(indexHtml, 'openAudioImportEditor')}; this.open = openAudioImportEditor;`, context);

  const pending = context.open({ name: 'voice.mp3' });
  await flush();
  assert.equal(typeof destroyHandler, 'function', 'the production bridge must observe native editor destruction');
  const nativeDestroy = destroyHandler;
  const rejected = assert.rejects(pending, (error) => error && error.code === 'AUDIO_IMPORT_CANCELLED');
  nativeDestroy();
  nativeDestroy();
  await rejected;
  assert.deepEqual(calls.filter((entry) => entry[0] === 'mode'), [['mode', true], ['mode', false]]);
  assert.equal(calls.at(-1)[0], 'destroy');
  assert.equal(calls.at(-1)[1], 'object', 'cleanup clears the destroy observer with null');
});

test('post-result native destroy keeps ownership until async validation rejects the stale result', async () => {
  const headerGate = deferred();
  const headerStarted = deferred();
  const audit = { modes: [], editorReleases: 0, finishes: 0 };
  let resultHandler = null;
  let destroyHandler = null;
  class DelayedHeaderBlob extends Blob {
    slice() {
      return {
        arrayBuffer() {
          headerStarted.resolve();
          return headerGate.promise;
        }
      };
    }
  }
  const editor = {
    init() {},
    loadFile() { return Promise.resolve({ mode: 'editing' }); },
    getState() { return { totalDuration: 45 }; },
    setResultHandler(handler) { resultHandler = handler; },
    setDestroyHandler(handler) { destroyHandler = handler; },
    setImportMode(value) { audit.modes.push(value); },
    nativeDestroy() {
      audit.editorReleases += 1;
      const handler = destroyHandler;
      destroyHandler = null;
      if (handler) handler();
    }
  };
  const context = {
    console,
    Blob,
    Uint8Array,
    ArrayBuffer,
    Promise,
    Error,
    Number,
    Object,
    window: { HemisphereAudioEditor: editor },
    state: { audioImport: { editorReject: null, editorDisarm: null, processor: null } }
  };
  const processorStart = indexHtml.indexOf('const ImportAudioItemProcessor =');
  const processorEnd = indexHtml.indexOf('const AudioPinImportWorkflow =', processorStart);
  vm.runInNewContext([
    indexHtml.slice(processorStart, processorEnd),
    functionSource(indexHtml, 'openAudioImportEditor'),
    'this.processorApi = ImportAudioItemProcessor;',
    'this.openEditor = openAudioImportEditor;'
  ].join('\n'), context);
  const processor = context.processorApi.create({ openAudioEditor: context.openEditor });
  context.state.audioImport.processor = processor;
  const file = { name: 'voice.mp3', type: 'audio/mpeg', size: 1024 };
  const processing = processor.processLocalFile(file).then(
    (value) => { audit.finishes += 1; return { value }; },
    (error) => ({ error })
  );
  await flush();
  const blob = new DelayedHeaderBlob(
    [new Uint8Array([0x49, 0x44, 0x33, 4, 0, 0, 0, 0, 0, 0])],
    { type: 'audio/mpeg' }
  );
  const publishResult = resultHandler;
  publishResult({
    blob, fileName: 'voice-trimmed.mp3', mimeType: 'audio/mpeg', sizeBytes: blob.size,
    durationSeconds: 30, selectionStart: 0, selectionEnd: 30,
    sampleRate: 48000, bitrate: 192000, numberOfChannels: 1
  });
  await headerStarted.promise;

  assert.deepEqual(audit.modes, [true], 'workflow-owned replacement controls remain guarded after result publication');
  assert.equal(typeof destroyHandler, 'function', 'native destroy ownership remains armed through async validation');
  editor.nativeDestroy();
  headerGate.resolve(Uint8Array.from([0x49, 0x44, 0x33, 4]).buffer);
  const outcome = await processing;

  assert.equal(outcome.error && outcome.error.code, 'AUDIO_IMPORT_CANCELLED');
  assert.equal(outcome.value, undefined);
  assert.equal(audit.finishes, 0);
  assert.equal(audit.editorReleases, 1);
  assert.equal(processor.getResult(), null);
});

test('successful finish disarms editor cancellation immediately before programmatic destroy', async () => {
  const order = [];
  const context = {
    window: { HemisphereAudioEditor: { destroy() { order.push('destroy'); } } },
    state: {
      audioImport: {
        editorDisarm() { order.push('disarm'); },
        sourceFileName: '',
        operation: { id: 'operation' }
      }
    },
    audioPinImportWorkflow: { start() { order.push('start'); } },
    closeOverlay() { order.push('close'); }
  };
  vm.runInNewContext(`${functionSource(indexHtml, 'finishAudioImportProcessing')}; this.finish = finishAudioImportProcessing;`, context);

  await context.finish({ blob: {} }, 'voice.mp3');
  assert.deepEqual(order, ['disarm', 'destroy', 'start', 'close']);
});

test('dismissing a Drive audio editor closes the picker before returning to the source overlay', () => {
  const order = [];
  const context = {
    state: {
      audioImport: {
        operation: { sourceKind: 'drive' },
        editorReject: null
      }
    },
    window: {},
    closeOverlay(id, options) {
      order.push(['close', id, options && options.restoreFocus]);
    },
    cancelPendingAudioImport(reason) {
      order.push(['cancel', reason]);
    },
    openAudioImportReturnTarget() {
      order.push(['return']);
    }
  };
  vm.runInNewContext(`${functionSource(indexHtml, 'dismissOverlayById')}; this.dismiss = dismissOverlayById;`, context);

  assert.equal(context.dismiss('audio-editor-overlay'), true);
  assert.deepEqual(order, [
    ['close', 'audio-drive-import-overlay', false],
    ['cancel', 'editor-closed'],
    ['return']
  ]);
});

test('Drive list request can be cancelled and its late success cannot repopulate the picker', async () => {
  const gate = deferred();
  const audit = { renders: 0, opens: [], errors: [] };
  const state = {
    audioImport: {
      requestToken: 0, operation: null, shell: null,
      driveItems: [], driveSelectedId: '', driveLoading: false
    }
  };
  const context = {
    state,
    canStartAddAction() { return true; },
    closeOverlay() {},
    openOverlay(id) { audit.opens.push(id); },
    createAudioImportOperation() {
      state.audioImport.requestToken += 1;
      state.audioImport.operation = { id: 'operation' };
      state.audioImport.shell = {};
      return state.audioImport.operation;
    },
    setAudioDriveImportError(value) { audit.errors.push(value); },
    renderAudioDriveImportItems() { audit.renders += 1; },
    withGAS() { return gate.promise; },
    withEditToken(value) { return value; },
    cancelPendingAudioImport() {
      state.audioImport.requestToken += 1;
      state.audioImport.operation = null;
      state.audioImport.shell = null;
      state.audioImport.driveItems = [];
      state.audioImport.driveLoading = false;
    },
    openAudioImportReturnTarget() {
      audit.opens.push('add-menu-overlay');
    }
  };
  vm.runInNewContext([
    functionSource(indexHtml, 'handleDriveAudioImportButtonClick'),
    functionSource(indexHtml, 'cancelDriveAudioImport'),
    'this.startList = handleDriveAudioImportButtonClick;',
    'this.cancelList = cancelDriveAudioImport;'
  ].join('\n'), context);

  const listing = context.startList();
  await flush();
  assert.equal(state.audioImport.driveLoading, true);
  assert.equal(context.cancelList(), true, 'cancel remains available while the read is pending');
  gate.resolve({ ok: true, items: [{ id: 'late-audio', name: 'late.wav' }] });
  await listing;

  assert.deepEqual(state.audioImport.driveItems, []);
  assert.equal(state.audioImport.operation, null);
  assert.equal(audit.errors.length, 1, 'only the initial error clear is allowed; no late settlement writes');
  assert.deepEqual(audit.opens, ['audio-drive-import-overlay', 'add-menu-overlay']);
});

test('Drive read loading ends before editor wait and cancellation ignores late editor settlement', async () => {
  const readGate = deferred();
  const editorGate = deferred();
  const audit = { finishes: 0, failures: 0, renders: 0 };
  const state = {
    audioImport: {
      requestToken: 7,
      driveSelectedId: 'audio_AAAAAAAAAAA',
      driveItems: [{ id: 'audio_AAAAAAAAAAA', name: 'voice.wav' }],
      driveLoading: false,
      operation: { id: 'operation' },
      shell: {
        setSourceDriveFileId() { return { id: 'operation-with-source' }; }
      },
      processor: null
    }
  };
  const context = {
    state,
    setAudioDriveImportError() {},
    renderAudioDriveImportItems() { audit.renders += 1; },
    withGAS() { return readGate.promise; },
    withEditToken(value) { return value; },
    createProductionAudioProcessor() {
      return { processDriveResponse() { return editorGate.promise; } };
    },
    async finishAudioImportProcessing() { audit.finishes += 1; },
    handleAudioImportFailure() { audit.failures += 1; },
    closeOverlay() {},
    openOverlay() {},
    cancelPendingAudioImport() {
      state.audioImport.requestToken += 1;
      state.audioImport.operation = null;
      state.audioImport.shell = null;
      state.audioImport.driveLoading = false;
    },
    openAudioImportReturnTarget() {}
  };
  vm.runInNewContext([
    functionSource(indexHtml, 'confirmDriveAudioImport'),
    functionSource(indexHtml, 'cancelDriveAudioImport'),
    'this.confirmRead = confirmDriveAudioImport;',
    'this.cancelRead = cancelDriveAudioImport;'
  ].join('\n'), context);

  const confirmation = context.confirmRead();
  assert.equal(state.audioImport.driveLoading, true);
  readGate.resolve({ ok: true, file: {} });
  await flush();
  assert.equal(state.audioImport.driveLoading, false, 'Drive loading covers list/read only, not the editor Promise');
  assert.equal(context.cancelRead(), true);
  editorGate.resolve({ blob: {} });
  assert.equal(await confirmation, false);
  assert.equal(audit.finishes, 0);
  assert.equal(audit.failures, 0);
  assert.equal(state.audioImport.operation, null);
});

test('Drive radio change preserves the focused node instead of rebuilding the list', () => {
  const listeners = new Map();
  const appended = [];
  let replacements = 0;
  const confirm = { disabled: true };
  const status = { textContent: '' };
  const container = {
    replaceChildren() { replacements += 1; appended.length = 0; },
    appendChild(value) { appended.push(value); }
  };
  function element(tag) {
    return {
      tag,
      style: {},
      appendChild() {},
      addEventListener(type, handler) { listeners.set(this, { type, handler }); }
    };
  }
  const state = {
    audioImport: {
      driveItems: [{ id: 'one', name: 'One' }, { id: 'two', name: 'Two' }],
      driveSelectedId: '', driveLoading: false
    }
  };
  const context = {
    state,
    document: {
      getElementById(id) {
        if (id === 'audio-drive-import-list') return container;
        if (id === 'audio-drive-import-confirm') return confirm;
        return status;
      },
      createElement: element
    }
  };
  vm.runInNewContext(`${functionSource(indexHtml, 'renderAudioDriveImportItems')}; this.render = renderAudioDriveImportItems;`, context);
  context.render();
  const firstInput = Array.from(listeners.keys())[0];
  listeners.get(firstInput).handler();

  assert.equal(replacements, 1, 'selection updates state and controls without replacing focused radios');
  assert.equal(state.audioImport.driveSelectedId, 'one');
  assert.equal(confirm.disabled, false);
});

test('audio Preview hides the photo-only EXIF datetime action', () => {
  assert.match(
    indexHtml,
    /#upload-overlay\.audio-import-preview[^{}]*#upload-event-at-exif[^{}]*\{[^}]*display:\s*none\s*!important/,
    'audio Preview must not expose photo EXIF UI'
  );
});

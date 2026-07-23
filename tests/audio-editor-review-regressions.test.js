const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const rootPath = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(rootPath, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(rootPath, 'shared.html'), 'utf8');
const EDITOR_START = '<template data-dtp-audio-editor-boundary="start"></template>';
const EDITOR_END = '<template data-dtp-audio-editor-boundary="end"></template>';

const audioEditorConditionBlock = /<\?\s*if\s*\(editToken\)\s*\{\s*\?>\s*<template data-dtp-audio-editor-boundary="start"><\/template>[\s\S]*?<template data-dtp-audio-editor-boundary="end"><\/template>\s*<\?\s*\}\s*\?>/;
const editorStartIndex = indexHtml.indexOf(EDITOR_START);
const editorEndIndex = indexHtml.indexOf(EDITOR_END);
assert.ok(editorStartIndex >= 0 && editorEndIndex > editorStartIndex, 'inline editor boundaries');
const editorHtml = indexHtml.slice(editorStartIndex + EDITOR_START.length, editorEndIndex).replace(/^\r?\n/, '');

function evaluateIndexTemplate(hasEditToken) {
  assert.match(indexHtml, audioEditorConditionBlock);
  return indexHtml.replace(audioEditorConditionBlock, hasEditToken ? editorHtml : '');
}

function functionSource(source, name) {
  const start = source.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
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

function editorScript() {
  const start = editorHtml.lastIndexOf('<script>');
  const end = editorHtml.lastIndexOf('</script>');
  assert.ok(start >= 0 && end > start);
  return editorHtml.slice(start + '<script>'.length, end);
}

function classList(initial = []) {
  const values = new Set(initial);
  return {
    add(...items) { items.forEach((item) => values.add(item)); },
    remove(...items) { items.forEach((item) => values.delete(item)); },
    contains(item) { return values.has(item); },
    toggle(item, force) {
      const enabled = force === undefined ? !values.has(item) : !!force;
      if (enabled) values.add(item);
      else values.delete(item);
      return enabled;
    }
  };
}

function createLifecycleHarness() {
  const audit = { added: 0, removed: 0, loads: 0, opens: 0, closes: 0, stack: [], documentEscapes: 0 };
  const nodes = new Map();
  function node(key) {
    if (nodes.has(key)) return nodes.get(key);
    const eventListeners = new Map();
    const item = {
      key,
      hidden: false,
      disabled: false,
      open: false,
      value: '',
      checked: false,
      files: [],
      classList: classList(),
      style: {},
      dataset: {},
      addEventListener(type, handler) {
        audit.added += 1;
        if (!eventListeners.has(type)) eventListeners.set(type, []);
        eventListeners.get(type).push(handler);
      },
      removeEventListener(type, handler) {
        audit.removed += 1;
        if (!eventListeners.has(type)) return;
        eventListeners.set(type, eventListeners.get(type).filter((candidate) => candidate !== handler));
      },
      dispatchEvent(event) {
        const nextEvent = event;
        nextEvent.target = nextEvent.target || this;
        nextEvent.currentTarget = this;
        nextEvent.defaultPrevented = false;
        nextEvent.propagationStopped = false;
        nextEvent.preventDefault = nextEvent.preventDefault || function preventDefault() { this.defaultPrevented = true; };
        nextEvent.stopPropagation = nextEvent.stopPropagation || function stopPropagation() { this.propagationStopped = true; };
        (eventListeners.get(nextEvent.type) || []).slice().forEach((handler) => handler(nextEvent));
        if (nextEvent.bubbles !== false && !nextEvent.propagationStopped) documentApi.dispatchEvent(nextEvent);
        return !nextEvent.defaultPrevented;
      },
      setAttribute() {},
      removeAttribute() {},
      querySelector(selector) { return node(`${key} ${selector}`); },
      querySelectorAll() { return []; },
      close() { this.open = false; },
      showModal() { this.open = true; },
      pause() {},
      load() {},
      focus() {
        documentApi.activeElement = this;
        this.dispatchEvent({ type: 'focus', bubbles: false });
      },
      blur() {
        this.dispatchEvent({ type: 'blur', bubbles: false });
        if (documentApi.activeElement === this) documentApi.activeElement = node('document-body');
      },
      click() {}
    };
    nodes.set(key, item);
    return item;
  }
  const editorRoot = node('editor-root');
  const overlay = node('audio-editor-overlay');
  const help = node('help-dialog');
  editorRoot.querySelector = (selector) => node(selector);
  editorRoot.querySelectorAll = () => [];
  const documentListeners = new Map();
  const documentApi = {
    activeElement: node('opener'),
    querySelector(selector) {
      if (selector === '[data-hemisphere-audio-editor]') return editorRoot;
      if (selector === '[data-hae-help-dialog]') return help;
      if (selector === '[data-hae-help-close]') return node('help-close');
      if (selector === '#hae-audio-file') return node('file-input');
      return node(selector);
    },
    getElementById(id) { return id === 'audio-editor-overlay' ? overlay : null; },
    addEventListener(type, handler) {
      if (!documentListeners.has(type)) documentListeners.set(type, []);
      documentListeners.get(type).push(handler);
    },
    dispatchEvent(event) {
      (documentListeners.get(event.type) || []).slice().forEach((handler) => handler(event));
    }
  };
  let rejectFirstLoad;
  const firstLoad = new Promise((_resolve, reject) => { rejectFirstLoad = reject; });
  const context = {
    console,
    document: documentApi,
    URL,
    URLSearchParams,
    Blob,
    setTimeout,
    clearTimeout,
    location: { search: '' },
    requestAnimationFrame() { return 1; },
    cancelAnimationFrame() {},
    addEventListener() { audit.added += 1; },
    removeEventListener() { audit.removed += 1; },
    loadAudioVendorBundle() {
      audit.loads += 1;
      return audit.loads === 1 ? firstLoad : Promise.resolve();
    },
    openOverlay(id) {
      audit.opens += 1;
      audit.stack = audit.stack.filter((entry) => entry !== id);
      audit.stack.push(id);
      overlay.classList.add('open');
    },
    closeOverlay(id) {
      audit.closes += 1;
      audit.stack = audit.stack.filter((entry) => entry !== id);
      overlay.classList.remove('open');
    }
  };
  context.window = context;
  context.globalThis = context;
  vm.runInNewContext(editorScript(), context);
  return { api: context.HemisphereAudioEditor, audit, overlay, rejectFirstLoad, document: documentApi, node };
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

function createAudioBuffer(duration = 1) {
  const length = 480;
  const samples = [new Float32Array(length)];
  return {
    duration,
    numberOfChannels: 1,
    sampleRate: 48000,
    length,
    getChannelData(channel) { return samples[channel]; }
  };
}

function createEncodeHarness() {
  const verification = deferred();
  const verificationStarted = deferred();
  const audit = { createdUrls: [], revokedUrls: [], decodeCalls: 0, outputCancels: 0 };
  const inputBuffer = createAudioBuffer();
  class FakeAudioContext {
    decodeAudioData() {
      audit.decodeCalls += 1;
      if (audit.decodeCalls === 2) {
        verificationStarted.resolve();
        return verification.promise;
      }
      return Promise.resolve(inputBuffer);
    }
    createBuffer(channels, length, sampleRate) {
      const data = Array.from({ length: channels }, () => new Float32Array(length));
      return {
        duration: length / sampleRate,
        numberOfChannels: channels,
        length,
        sampleRate,
        getChannelData(channel) { return data[channel]; }
      };
    }
    close() { return Promise.resolve(); }
  }
  class BufferTarget { constructor() { this.buffer = null; } }
  class Output {
    constructor(options) { this.target = options.target; }
    addAudioTrack() {}
    start() { return Promise.resolve(); }
    finalize() {
      this.target.buffer = Uint8Array.from([0xff, 0xfb, 0x90, 0x64]).buffer;
      return Promise.resolve();
    }
    cancel() { audit.outputCancels += 1; return Promise.resolve(); }
  }
  class AudioBufferSource {
    add() { return Promise.resolve(); }
    close() {}
  }
  class Mp3OutputFormat {}
  const context = {
    console,
    Blob,
    URL: {
      createObjectURL() {
        const value = `blob:review-${audit.createdUrls.length + 1}`;
        audit.createdUrls.push(value);
        return value;
      },
      revokeObjectURL(value) { audit.revokedUrls.push(value); }
    },
    URLSearchParams,
    AudioContext: FakeAudioContext,
    Mediabunny: {
      canEncodeAudio() { return Promise.resolve(true); },
      BufferTarget,
      Output,
      Mp3OutputFormat,
      AudioBufferSource
    },
    MediabunnyMp3Encoder: { registerMp3Encoder() {} },
    loadAudioVendorBundle() { return Promise.resolve(); },
    requestAnimationFrame() { return 1; },
    cancelAnimationFrame() {},
    location: { search: '' }
  };
  context.window = context;
  context.globalThis = context;
  vm.runInNewContext(editorScript(), context);
  return { api: context.HemisphereAudioEditor, audit, verification, verificationStarted };
}

function mp3File(name) {
  const bytes = Uint8Array.from([0xff, 0xfb, 0x90, 0x64]);
  return {
    name,
    type: 'audio/mpeg',
    size: bytes.byteLength,
    arrayBuffer() { return Promise.resolve(bytes.buffer.slice(0)); }
  };
}

test('editor fragment is a bounded sheet overlay managed by the shared lifecycle', () => {
  assert.match(editorHtml, /<div\s+id="audio-editor-overlay"\s+class="sheet-overlay"[^>]*role="dialog"[^>]*aria-modal="true"[^>]*aria-labelledby="hae-editor-title"/);
  assert.match(editorHtml, /id="audio-editor-overlay"[\s\S]*?<div\s+class="sheet-body[^"\n]*dialog-scroll-frame/);
  assert.match(editorHtml, /dialog-scroll-content[\s\S]*?data-hemisphere-audio-editor/);
  assert.match(editorHtml, /id="hae-editor-title"/);
  assert.match(editorHtml, /data-hae-overlay-close[^>]*data-overlay-initial-focus|data-overlay-initial-focus[^>]*data-hae-overlay-close/);
  assert.match(editorHtml, /\.audio-editor-sheet-header svg\s*\{[^}]*width:22px[^}]*stroke:currentColor/);
  assert.match(editorHtml, /data-hemisphere-audio-editor[\s\S]*?<dialog[\s\S]*?data-hae-help-dialog[\s\S]*?<\/dialog>[\s\S]*?<\/div>/);

  assert.match(functionSource(editorHtml, 'init'), /openOverlay\(['"]audio-editor-overlay['"]\)/);
  assert.match(functionSource(editorHtml, 'destroy'), /closeOverlay\(['"]audio-editor-overlay['"]\)/);
  assert.match(functionSource(editorHtml, 'bindInteractions'), /elements\.cancel[^\n]*destroy/);
  assert.match(functionSource(indexHtml, 'dismissOverlayById'), /audio-editor-overlay[\s\S]*HemisphereAudioEditor\.destroy\(\)/);
  assert.match(indexHtml, /MAIN_DISMISSIBLE_OVERLAY_IDS[\s\S]{0,500}['"]audio-editor-overlay['"]/);
  assert.match(indexHtml, /BACKDROP_DISMISS_OVERLAY_IDS[\s\S]{0,500}['"]audio-editor-overlay['"]/);
  assert.ok(
    indexHtml.indexOf(EDITOR_START) < indexHtml.indexOf("['help-overlay'].concat(MAIN_DISMISSIBLE_OVERLAY_IDS)"),
    'the fragment DOM must exist before shared backdrop listeners are registered'
  );
});

test('tokenless evaluated index skips the absent audio overlay and continues startup', () => {
  const tokenlessIndex = evaluateIndexTemplate(false);
  assert.equal(tokenlessIndex.includes('id="audio-editor-overlay"'), false);
  assert.match(tokenlessIndex, /id="app-shell"/);
  assert.equal(sharedHtml.includes('id="audio-editor-overlay"'), false);

  const registrationStart = indexHtml.indexOf("['help-overlay'].concat(MAIN_DISMISSIBLE_OVERLAY_IDS).forEach");
  const registrationEnd = indexHtml.indexOf('function getPinById', registrationStart);
  assert.ok(registrationStart >= 0 && registrationEnd > registrationStart);
  const registrationSource = indexHtml.slice(registrationStart, registrationEnd);
  const audit = { listeners: 0, initialized: 0 };
  const helpOverlay = { addEventListener() { audit.listeners += 1; } };
  const context = {
    MAIN_DISMISSIBLE_OVERLAY_IDS: ['audio-editor-overlay'],
    document: {
      getElementById(id) { return id === 'help-overlay' ? helpOverlay : null; }
    },
    initializeApp() { audit.initialized += 1; }
  };

  assert.doesNotThrow(() => vm.runInNewContext([
    functionSource(indexHtml, 'setupOverlayBackdropDismissal'),
    registrationSource,
    'initializeApp();'
  ].join('\n'), context));
  assert.equal(audit.listeners, 4, 'the existing help backdrop remains bound');
  assert.equal(audit.initialized, 1, 'startup continues after optional overlay registration');
});

test('native audio help dialog owns Tab and Escape while open', () => {
  const helpDialog = { open: true };
  const audit = { trap: 0, escape: 0, enter: 0, shortcut: 0 };
  const context = {
    document: { querySelector() { return helpDialog; } },
    handleAppDialogEnter() { audit.enter += 1; return false; },
    trapOverlayFocus() { audit.trap += 1; return true; },
    dispatchEscape() { audit.escape += 1; return true; },
    handleDuplicateShortcut() { audit.shortcut += 1; }
  };
  vm.runInNewContext([
    functionSource(indexHtml, 'isAudioEditorHelpDialogOpen'),
    functionSource(indexHtml, 'handleGlobalKeydown'),
    'this.handle = handleGlobalKeydown;'
  ].join('\n'), context);
  const tab = { key: 'Tab', preventDefault() {}, stopPropagation() {} };
  const escape = { key: 'Escape', preventDefault() {}, stopPropagation() {} };
  context.handle(tab);
  context.handle(escape);
  assert.deepEqual(audit, { trap: 0, escape: 0, enter: 0, shortcut: 0 });

  helpDialog.open = false;
  context.handle(tab);
  context.handle(escape);
  assert.equal(audit.trap, 1);
  assert.equal(audit.escape, 1);
});

test('audio overlay owns pin copy duplicate and paste shortcuts without changing typing guards', () => {
  const audioOverlay = { classList: classList(['open']) };
  const control = { tagName: 'BUTTON', isContentEditable: false };
  const waveform = { tagName: 'DIV', isContentEditable: false };
  const input = { tagName: 'INPUT', isContentEditable: false };
  const audit = { duplicates: [], hints: [] };
  const state = {
    activePinId: 'pin-1',
    copiedPinSourceId: null,
    selectedPinIds: new Set()
  };
  const context = {
    state,
    Set,
    document: {
      activeElement: control,
      getElementById(id) { return id === 'audio-editor-overlay' ? audioOverlay : null; }
    },
    canEdit() { return true; },
    getPinById(id) { return id === 'pin-1' ? { id: 'pin-1', title: 'Pin 1' } : null; },
    duplicatePinFromSource(...args) { audit.duplicates.push(args); },
    beginDuplicatePlacement(...args) { audit.duplicates.push(args); },
    showHint(message) { audit.hints.push(message); }
  };
  ['isTypingTarget', 'isShortcutOverlayOpen', 'hasMultipleSelectedPins', 'handleDuplicateShortcut'].forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(indexHtml, name)})`, context);
  });
  function shortcut(key, target, options = {}) {
    const event = {
      key,
      target,
      ctrlKey: true,
      metaKey: false,
      altKey: false,
      shiftKey: !!options.shiftKey,
      preventCount: 0,
      preventDefault() { this.preventCount += 1; }
    };
    context.document.activeElement = target;
    context.handleDuplicateShortcut(event);
    return event;
  }

  shortcut('c', control);
  shortcut('d', waveform);
  state.copiedPinSourceId = 'pin-1';
  shortcut('v', control, { shiftKey: true });
  assert.equal(state.copiedPinSourceId, 'pin-1');
  assert.deepEqual(audit.duplicates, []);
  assert.deepEqual(audit.hints, []);

  audioOverlay.classList.remove('open');
  state.copiedPinSourceId = null;
  assert.equal(shortcut('c', control).preventCount, 1);
  assert.equal(state.copiedPinSourceId, 'pin-1');
  assert.equal(shortcut('d', waveform).preventCount, 1);
  assert.equal(shortcut('v', control, { shiftKey: true }).preventCount, 1);
  assert.equal(audit.duplicates.length, 2, 'closed editor preserves duplicate and paste behavior');

  state.copiedPinSourceId = null;
  assert.equal(shortcut('c', input).preventCount, 0);
  assert.equal(state.copiedPinSourceId, null, 'typing targets keep native shortcuts');
});

test('a failed per-open vendor preparation retries without duplicate bindings or overlay records', async () => {
  const harness = createLifecycleHarness();
  harness.api.init();
  const bindingsAfterFirstOpen = harness.audit.added;
  assert.equal(harness.audit.loads, 1);
  assert.equal(harness.audit.opens, 1);
  assert.deepEqual(harness.audit.stack, ['audio-editor-overlay']);

  harness.rejectFirstLoad(new Error('first open failed'));
  await Promise.resolve();
  harness.api.init();
  await Promise.resolve();

  assert.equal(harness.audit.loads, 2, 'each open attempt prepares the vendor again after failure');
  assert.equal(harness.audit.added, bindingsAfterFirstOpen, 'DOM listeners are bound only once');
  assert.equal(harness.audit.opens, 1, 'an already-open overlay is not pushed again');
  assert.deepEqual(harness.audit.stack, ['audio-editor-overlay']);

  assert.doesNotThrow(() => harness.api.destroy());
  assert.doesNotThrow(() => harness.api.destroy());
  assert.equal(harness.audit.closes, 1, 'idempotent destroy closes an open overlay once');
  assert.equal(harness.audit.stack.length, 0);
});

for (const selector of ['[data-hae-cancel]', '[data-hae-overlay-close]']) {
  test(`${selector} and recursive cleanup notify editor destruction exactly once`, () => {
    const harness = createLifecycleHarness();
    let notifications = 0;
    harness.api.setDestroyHandler(() => {
      notifications += 1;
      harness.api.destroy('recursive-cleanup');
      throw new Error('observer failure must be isolated');
    });
    harness.api.init();
    harness.node(selector).dispatchEvent({ type: 'click', bubbles: false });
    harness.api.destroy('duplicate-cleanup');

    assert.equal(notifications, 1);
    assert.equal(harness.audit.closes, 1);
    assert.deepEqual(harness.audit.stack, []);
  });
}

test('workflow-owned editor mode hides replacement controls and destroy restores standalone defaults', () => {
  const harness = createLifecycleHarness();
  harness.api.init();
  const chooseAnother = harness.node('[data-hae-choose-another]');
  const errorChoose = harness.node('[data-hae-error-choose]');

  harness.api.setImportMode(true);
  assert.equal(chooseAnother.hidden, true);
  assert.equal(chooseAnother.disabled, true);
  assert.equal(errorChoose.hidden, true);
  assert.equal(errorChoose.disabled, true);

  harness.api.destroy();
  assert.equal(chooseAnother.hidden, false);
  assert.equal(chooseAnother.disabled, false);
  assert.equal(errorChoose.hidden, false);
  assert.equal(errorChoose.disabled, false);
});

test('Escape rolls back every time input without bubbling into editor dismissal', () => {
  const harness = createLifecycleHarness();
  harness.api.init();
  harness.document.addEventListener('keydown', (event) => {
    if (event.key !== 'Escape') return;
    harness.audit.documentEscapes += 1;
    harness.api.destroy();
  });

  ['[data-hae-start-time]', '[data-hae-end-time]', '[data-hae-duration]'].forEach((selector, index) => {
    const input = harness.node(selector);
    const original = `original-${index}`;
    input.value = original;
    input.focus();
    input.value = `changed-${index}`;
    const event = { type: 'keydown', key: 'Escape', bubbles: true };
    input.dispatchEvent(event);

    assert.equal(event.defaultPrevented, true, selector);
    assert.equal(event.propagationStopped, true, selector);
    assert.equal(input.value, original, selector);
    assert.equal(harness.overlay.classList.contains('open'), true, selector);
    assert.deepEqual(harness.audit.stack, ['audio-editor-overlay'], selector);
  });
  assert.equal(harness.audit.documentEscapes, 0);
  assert.equal(harness.audit.closes, 0);
});

for (const scenario of [
  {
    label: 'cancel',
    invalidate(harness) { harness.api.destroy(); }
  },
  {
    label: 'file replacement',
    async invalidate(harness) { await harness.api.loadFile(mp3File('replacement.mp3')); }
  },
  {
    label: 'selection change',
    invalidate(harness) { harness.api.setSelection(0, 0.5); }
  },
  {
    label: 'destroy',
    invalidate(harness) { harness.api.destroy(); }
  }
]) {
  test(`final MP3 verification cannot publish a stale result after ${scenario.label}`, async () => {
    const harness = createEncodeHarness();
    await harness.api.loadFile(mp3File('source.mp3'));
    const encoding = harness.api.encodeSelection();
    await harness.verificationStarted.promise;

    await scenario.invalidate(harness);
    harness.verification.resolve({ duration: 1 });
    assert.equal(await encoding, null);
    assert.equal(harness.api.getResult(), null);
    assert.notEqual(harness.api.getState().mode, 'ready');
    assert.deepEqual(harness.audit.createdUrls, [], 'no stale Object URL is created');
    assert.deepEqual(harness.audit.revokedUrls, [], 'there is no leaked URL requiring cleanup');
  });
}

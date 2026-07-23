const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const PLAYER_START = '<template data-dtp-audio-player-boundary="start"></template>';
const PLAYER_END = '<template data-dtp-audio-player-boundary="end"></template>';

function playerHtml() {
  const startIndex = indexHtml.indexOf(PLAYER_START);
  const endIndex = indexHtml.indexOf(PLAYER_END);
  assert.ok(startIndex >= 0 && endIndex > startIndex, 'inline player boundaries');
  return indexHtml.slice(startIndex + PLAYER_START.length, endIndex).replace(/^\r?\n/, '');
}

function scriptSource(html) {
  const matches = Array.from(html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi));
  assert.equal(matches.length, 1, 'audio player fragment must have one script');
  return matches[0][1];
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

const base64BySize = new Map();
function response(byteLength = 1024 * 1024) {
  if (!base64BySize.has(byteLength)) {
    base64BySize.set(byteLength, Buffer.alloc(byteLength).toString('base64'));
  }
  return { ok: true, mimeType: 'audio/mpeg', byteLength, base64: base64BySize.get(byteLength) };
}

function createAudioElement() {
  const listeners = new Map();
  const audio = {
    src: '', currentTime: 0, defaultPlaybackRate: 0.5, playbackRate: 1.5,
    playCalls: 0, pauseCalls: 0, loadCalls: 0,
    play() { audio.playCalls += 1; return Promise.resolve(); },
    pause() { audio.pauseCalls += 1; },
    load() { audio.loadCalls += 1; },
    removeAttribute(name) { if (name === 'src') audio.src = ''; },
    addEventListener(type, handler) {
      if (!listeners.has(type)) listeners.set(type, new Set());
      listeners.get(type).add(handler);
    },
    removeEventListener(type, handler) {
      if (listeners.has(type)) listeners.get(type).delete(handler);
    },
    emit(type) {
      Array.from(listeners.get(type) || []).forEach((handler) => handler());
    }
  };
  return audio;
}

function createElement() {
  const attributes = new Map();
  return {
    hidden: false, disabled: false, textContent: '',
    setAttribute(name, value) { attributes.set(name, String(value)); },
    getAttribute(name) { return attributes.get(name) || null; }
  };
}

function loadPlayer() {
  const created = [];
  const revoked = [];
  const lifecycleListeners = new Map();
  let nextUrl = 0;
  class FakeBlob {
    constructor(parts, options) {
      this.parts = parts;
      this.type = options && options.type;
      this.size = parts.reduce((total, part) => total + Number(part && part.byteLength || 0), 0);
    }
  }
  const window = {
    addEventListener(type, handler) {
      if (!lifecycleListeners.has(type)) lifecycleListeners.set(type, new Set());
      lifecycleListeners.get(type).add(handler);
    },
    removeEventListener(type, handler) {
      if (lifecycleListeners.has(type)) lifecycleListeners.get(type).delete(handler);
    }
  };
  const context = {
    window, Blob: FakeBlob, Uint8Array, Promise, Map, Set, Object, Number, String, Error,
    atob(value) { return Buffer.from(String(value), 'base64').toString('binary'); },
    URL: {
      createObjectURL(blob) {
        const url = `blob:audio-${++nextUrl}`;
        created.push({ url, blob });
        return url;
      },
      revokeObjectURL(url) { revoked.push(String(url)); }
    }
  };
  vm.createContext(context);
  vm.runInContext(scriptSource(playerHtml()), context);
  assert.ok(window.AudioPinPlayer);
  return {
    api: window.AudioPinPlayer,
    created,
    revoked,
    emitPagehide() {
      Array.from(lifecycleListeners.get('pagehide') || []).forEach((handler) => handler());
    }
  };
}

test('player fragment exposes the native accessible playback UI without a download action', () => {
  const html = playerHtml();
  assert.equal((html.match(/<audio\s+id="pin-audio-runtime"/g) || []).length, 1);
  assert.match(html, /<audio\s+id="pin-audio-runtime"[^>]*\bcontrols\b/);
  assert.match(html, /<audio\s+id="pin-audio-runtime"[^>]*\bcontrolsList="nodownload"/);
  assert.match(html, /<audio\s+id="pin-audio-runtime"[^>]*\bpreload="metadata"/);
  assert.doesNotMatch(html, /id="pin-audio-player-toggle"/);
  assert.match(html, /aria-live="polite"/);
  for (const text of ['音声を準備中…', '音声を再生できませんでした', '再試行']) {
    assert.ok(html.includes(text), `missing player copy: ${text}`);
  }
});

test('shared renderer projects hidden loading ready and retry states to the native audio element', () => {
  const { api } = loadPlayer();
  const elements = {
    container: createElement(), statusElement: createElement(),
    audioElement: createElement(), retryButton: createElement()
  };
  const render = api.createRenderer(elements);

  render({ status: 'loading' });
  assert.equal(elements.container.hidden, false);
  assert.equal(elements.statusElement.textContent, '音声を準備中…');
  assert.equal(elements.audioElement.hidden, true);

  render({ status: 'ready' });
  assert.equal(elements.statusElement.textContent, '');
  assert.equal(elements.audioElement.hidden, false);
  assert.equal(elements.retryButton.hidden, true);

  render({ status: 'error' });
  assert.equal(elements.statusElement.textContent, '音声を再生できませんでした');
  assert.equal(elements.audioElement.hidden, true);
  assert.equal(elements.retryButton.hidden, false);
  assert.equal(elements.retryButton.textContent, '再試行');

  render({ status: 'hidden' });
  assert.equal(elements.container.hidden, true);
});

test('player coalesces same-pin fetches and rejects switched closed and retried stale responses', async () => {
  const { api, created } = loadPlayer();
  const audio = createAudioElement();
  const views = [];
  const gates = new Map();
  const fetches = [];
  function fetchAudio(pinId) {
    fetches.push(pinId);
    const gate = deferred();
    if (!gates.has(pinId)) gates.set(pinId, []);
    gates.get(pinId).push(gate);
    return gate.promise;
  }
  const controller = api.create({
    audioElement: audio,
    fetchAudio,
    renderState(view) { views.push(JSON.parse(JSON.stringify(view))); }
  });
  assert.deepEqual(fetches, [], 'construction must not fetch audio');

  const firstA = controller.open('pin-a');
  const secondA = controller.open('pin-a');
  assert.equal(secondA, firstA, 'same-pin in-flight work must share one Promise');
  assert.deepEqual(fetches, ['pin-a']);

  const firstB = controller.open('pin-b');
  gates.get('pin-a')[0].resolve(response());
  assert.equal(await firstA, false);
  assert.equal(created.length, 0, 'a switched stale response must not allocate an Object URL');
  gates.get('pin-b')[0].resolve(response());
  assert.equal(await firstB, true);
  assert.equal(views.at(-1).status, 'ready');
  assert.equal(views.at(-1).pinId, 'pin-b');

  const closed = controller.open('pin-closed');
  controller.close();
  gates.get('pin-closed')[0].resolve(response());
  assert.equal(await closed, false);
  assert.equal(views.at(-1).status, 'hidden');

  const oldRetry = controller.open('pin-retry');
  const freshRetry = controller.retry();
  assert.deepEqual(fetches.slice(-2), ['pin-retry', 'pin-retry']);
  gates.get('pin-retry')[0].resolve(response());
  assert.equal(await oldRetry, false);
  gates.get('pin-retry')[1].resolve(response());
  assert.equal(await freshRetry, true);
  assert.equal(views.at(-1).pinId, 'pin-retry');
});

test('Object URL cache is true three-entry LRU and destroy revokes every retained URL', async () => {
  const { api, revoked } = loadPlayer();
  const fetches = [];
  const controller = api.create({
    audioElement: createAudioElement(),
    fetchAudio(pinId) { fetches.push(pinId); return Promise.resolve(response(4 * 1024 * 1024)); },
    renderState() {}
  });

  await controller.open('a');
  await controller.open('b');
  await controller.open('c');
  await controller.open('a');
  await controller.open('d');

  assert.deepEqual(fetches, ['a', 'b', 'c', 'd'], 'cache hit must avoid a second fetch');
  assert.deepEqual(revoked, ['blob:audio-2'], 'touching a makes b the oldest entry');
  controller.destroy();
  assert.deepEqual(new Set(revoked), new Set([
    'blob:audio-1', 'blob:audio-2', 'blob:audio-3', 'blob:audio-4'
  ]));
});

test('invalidating a changed pin revokes its cache and forces a fresh fetch', async () => {
  const { api, revoked } = loadPlayer();
  let fetchCount = 0;
  const controller = api.create({
    audioElement: createAudioElement(),
    fetchAudio() { fetchCount += 1; return Promise.resolve(response()); },
    renderState() {}
  });

  await controller.open('changed-pin');
  controller.close();
  assert.equal(controller.invalidate('changed-pin'), true);
  assert.deepEqual(revoked, ['blob:audio-1']);
  await controller.open('changed-pin');
  assert.equal(fetchCount, 2);
});

test('a mismatched decoded length fails closed without allocating an Object URL', async () => {
  const { api, created } = loadPlayer();
  const views = [];
  const controller = api.create({
    audioElement: createAudioElement(),
    fetchAudio() {
      return Promise.resolve({
        ok: true,
        mimeType: 'audio/mpeg',
        byteLength: 1024 * 1024,
        base64: Buffer.alloc(1024 * 1024 - 1).toString('base64')
      });
    },
    renderState(view) { views.push(JSON.parse(JSON.stringify(view))); }
  });

  assert.equal(await controller.open('corrupt-pin'), false);
  assert.equal(created.length, 0);
  assert.equal(views.at(-1).status, 'error');
});

test('player accepts a 1 KiB short clip response and rejects anything smaller', async () => {
  const { api, created } = loadPlayer();
  const controller = api.create({
    audioElement: createAudioElement(),
    fetchAudio(pinId) {
      return Promise.resolve(response(pinId === 'short-valid' ? 1024 : 1023));
    },
    renderState() {}
  });

  assert.equal(await controller.open('short-valid'), true);
  assert.equal(created.length, 1);
  assert.equal(await controller.open('too-small'), false);
  assert.equal(created.length, 1);
});

test('player stops native playback on pin switch close and pagehide', async () => {
  const loaded = loadPlayer();
  const audio = createAudioElement();
  const controller = loaded.api.create({
    audioElement: audio,
    fetchAudio() { return Promise.resolve(response()); },
    renderState() {}
  });

  await controller.open('a');
  await audio.play();
  audio.currentTime = 8;
  await controller.open('b');
  assert.equal(audio.currentTime, 0);

  audio.currentTime = 4;
  controller.close();
  assert.equal(audio.currentTime, 0);

  await controller.open('c');
  loaded.emitPagehide();
  assert.equal(audio.src, '');
  assert.equal(audio.pauseCalls >= 3, true);
  assert.equal(loaded.revoked.length, 3);
  assert.equal(await controller.open('after-destroy'), false);
});

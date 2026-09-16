const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

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
    if (character === '}' && --depth === 0) return source.slice(start, index + 1);
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

const base64BySize = new Map();
function response(byteLength = 4, mimeType = 'image/jpeg') {
  if (!base64BySize.has(byteLength)) {
    base64BySize.set(byteLength, Buffer.alloc(byteLength, 7).toString('base64'));
  }
  return {
    ok: true,
    mimeType,
    byteLength,
    base64: base64BySize.get(byteLength)
  };
}

function loadFactory() {
  const created = [];
  const revoked = [];
  let nextUrl = 0;
  class FakeBlob {
    constructor(parts, options) {
      this.parts = parts;
      this.type = options && options.type;
      this.size = parts.reduce(
        (total, part) => total + Number(part && part.byteLength || 0),
        0
      );
    }
  }
  const context = {
    Blob: FakeBlob,
    Uint8Array,
    Promise,
    Object,
    Number,
    String,
    Error,
    atob(value) {
      return Buffer.from(String(value), 'base64').toString('binary');
    },
    URL: {
      createObjectURL(blob) {
        const url = `blob:photo-${++nextUrl}`;
        created.push({ url, blob });
        return url;
      },
      revokeObjectURL(url) {
        revoked.push(String(url));
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(
    `${functionSource(indexHtml, 'createPinPhotoLoader')}
globalThis.__createPinPhotoLoader = createPinPhotoLoader;`,
    context
  );
  if (indexHtml.includes('function createPinPhotoCache(')) {
    vm.runInContext(`${functionSource(indexHtml, 'createPinPhotoCache')}
globalThis.__createPinPhotoCache = createPinPhotoCache;`, context);
  }
  return {
    create: context.__createPinPhotoLoader,
    createCache: context.__createPinPhotoCache,
    created,
    revoked
  };
}

function lifecycleTarget() {
  const listeners = new Map();
  return {
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
}

test('photo loader is lazy and coalesces the same in-flight pin request', async () => {
  const loaded = loadFactory();
  const gate = deferred();
  const fetches = [];
  const views = [];
  const controller = loaded.create({
    fetchPhoto(pinId) {
      fetches.push(pinId);
      return gate.promise;
    },
    renderState(view) {
      views.push({ ...view });
    },
    lifecycleTarget: lifecycleTarget()
  });

  assert.deepEqual(fetches, []);
  const first = controller.open('pin-a', 'A');
  const second = controller.open('pin-a', 'A');
  assert.equal(second, first);
  assert.deepEqual(fetches, ['pin-a']);
  assert.equal(views.at(-1).status, 'loading');

  gate.resolve(response());
  assert.equal(await first, true);
  assert.equal(views.at(-1).status, 'ready');
  assert.equal(views.at(-1).objectUrl, 'blob:photo-1');
  assert.equal(loaded.created[0].blob.type, 'image/jpeg');
});

test('photo cache avoids refetching across closed and destroyed views without retaining URLs', async () => {
  const loaded = loadFactory();
  assert.equal(typeof loaded.createCache, 'function');
  const cache = loaded.createCache();
  let fetches = 0;
  const config = {
    cache,
    fetchPhoto() { fetches += 1; return Promise.resolve(response()); },
    renderState() {},
    lifecycleTarget: lifecycleTarget()
  };
  const first = loaded.create(config);
  await first.open('a', 'A');
  first.close();
  assert.deepEqual(loaded.revoked, ['blob:photo-1']);
  await first.open('a', 'A');
  first.destroy();
  const second = loaded.create(config);
  assert.equal(await second.open('a', 'A'), true);
  assert.equal(fetches, 1);
  assert.equal(loaded.created.length, 3);
  assert.deepEqual(loaded.revoked, ['blob:photo-1', 'blob:photo-2']);
});

test('photo cache expires from fetch time even when repeatedly revisited', async () => {
  const loaded = loadFactory();
  assert.equal(typeof loaded.createCache, 'function');
  let time = 0;
  let fetches = 0;
  const cache = loaded.createCache({ now: () => time });
  const controller = loaded.create({ cache,
    fetchPhoto() { fetches += 1; return Promise.resolve(response()); },
    renderState() {}, lifecycleTarget: lifecycleTarget()
  });
  await controller.open('a', 'A');
  controller.close();
  time = 59999;
  await controller.open('a', 'A');
  assert.equal(fetches, 1);
  controller.close();
  time = 60000;
  await controller.open('a', 'A');
  assert.equal(fetches, 2);
});

test('photo cache evicts least recently viewed photos at both entry and byte limits', () => {
  const loaded = loadFactory();
  assert.equal(typeof loaded.createCache, 'function');
  const cache = loaded.createCache();
  const small = { size: 1 };
  for (const key of ['a', 'b', 'c']) cache.put(key, small);
  assert.equal(cache.get('a'), small);
  cache.put('d', small);
  assert.equal(cache.get('b'), null);
  assert.equal(cache.get('a'), small);
  cache.clear();
  const large = { size: 13 * 1024 * 1024 };
  cache.put('a', large);
  cache.put('b', large);
  assert.equal(cache.get('a'), null);
  assert.equal(cache.get('b'), large);
  cache.put('too-large', { size: 25 * 1024 * 1024 });
  assert.equal(cache.get('too-large'), null);
  cache.clear();
  assert.equal(cache.get('b'), null);
});

test('photo revisions retry and explicit invalidation bypass cached data', async () => {
  const loaded = loadFactory();
  assert.equal(typeof loaded.createCache, 'function');
  const cache = loaded.createCache();
  let revision = 'v1';
  let fetches = 0;
  const controller = loaded.create({ cache, getCacheKey: id => id + revision,
    fetchPhoto() { fetches += 1; return Promise.resolve(response()); },
    renderState() {}, lifecycleTarget: lifecycleTarget()
  });
  await controller.open('a', 'A');
  revision = 'v2';
  await controller.open('a', 'A');
  assert.equal(fetches, 2);
  await controller.retry();
  assert.equal(fetches, 3);
  controller.invalidate('a');
  await controller.open('a', 'A');
  assert.equal(fetches, 4);
});

test('invalid failed and stale photo responses never enter the cache', async () => {
  const loaded = loadFactory();
  assert.equal(typeof loaded.createCache, 'function');
  const cache = loaded.createCache();
  const gate = deferred();
  const replies = [response(4, 'text/plain'), Promise.reject(new Error('offline')), gate.promise, response()];
  replies[1].catch(() => {});
  let fetches = 0;
  const controller = loaded.create({ cache,
    fetchPhoto() { return replies[fetches++]; },
    renderState() {}, lifecycleTarget: lifecycleTarget()
  });
  assert.equal(await controller.open('a', 'A'), false);
  assert.equal(await controller.open('a', 'A'), false);
  const pending = controller.open('a', 'A');
  controller.close();
  gate.resolve(response());
  assert.equal(await pending, false);
  assert.equal(await controller.open('a', 'A'), true);
  assert.equal(fetches, 4);
});

test('a photo changed during fetch exposes retry and does not cache the previous revision', async () => {
  const loaded = loadFactory();
  const cache = loaded.createCache();
  const gate = deferred();
  let revision = 'v1';
  let fetches = 0;
  const views = [];
  const controller = loaded.create({ cache, getCacheKey: id => id + revision,
    fetchPhoto() { fetches += 1; return fetches === 1 ? gate.promise : response(); },
    renderState(view) { views.push(view); }, lifecycleTarget: lifecycleTarget()
  });
  const pending = controller.open('a', 'A');
  revision = 'v2';
  gate.resolve(response());
  assert.equal(await pending, false);
  assert.equal(views.at(-1).status, 'error');
  assert.equal(cache.get('av1'), null);
  assert.equal(await controller.retry(), true);
  assert.equal(fetches, 2);
});

test('photo loader rejects switched closed and retried stale responses', async () => {
  const loaded = loadFactory();
  const gates = new Map();
  const views = [];
  function fetchPhoto(pinId) {
    const gate = deferred();
    if (!gates.has(pinId)) gates.set(pinId, []);
    gates.get(pinId).push(gate);
    return gate.promise;
  }
  const controller = loaded.create({
    fetchPhoto,
    renderState(view) {
      views.push({ ...view });
    },
    lifecycleTarget: lifecycleTarget()
  });

  const firstA = controller.open('pin-a', 'A');
  const firstB = controller.open('pin-b', 'B');
  gates.get('pin-a')[0].resolve(response());
  assert.equal(await firstA, false);
  assert.equal(loaded.created.length, 0);
  gates.get('pin-b')[0].resolve(response());
  assert.equal(await firstB, true);
  assert.equal(views.at(-1).pinId, 'pin-b');

  const closed = controller.open('pin-closed', 'Closed');
  controller.close();
  gates.get('pin-closed')[0].resolve(response());
  assert.equal(await closed, false);
  assert.equal(views.at(-1).status, 'hidden');

  const oldRetry = controller.open('pin-retry', 'Retry');
  const freshRetry = controller.retry();
  gates.get('pin-retry')[0].resolve(response());
  assert.equal(await oldRetry, false);
  gates.get('pin-retry')[1].resolve(response());
  assert.equal(await freshRetry, true);
  assert.equal(views.at(-1).pinId, 'pin-retry');
});

test('photo loader validates MIME Base64 and decoded length before allocating a URL', async () => {
  const invalidResponses = [
    response(4, 'text/plain'),
    { ...response(4), byteLength: 0 },
    { ...response(4), byteLength: 15 * 1024 * 1024 + 1 },
    { ...response(4), base64: 'not canonical!' },
    { ...response(4), byteLength: 5 }
  ];

  for (const invalidResponse of invalidResponses) {
    const loaded = loadFactory();
    const views = [];
    const controller = loaded.create({
      fetchPhoto() {
        return Promise.resolve(invalidResponse);
      },
      renderState(view) {
        views.push({ ...view });
      },
      lifecycleTarget: lifecycleTarget()
    });
    assert.equal(await controller.open('invalid-pin', 'Invalid'), false);
    assert.equal(loaded.created.length, 0);
    assert.equal(views.at(-1).status, 'error');
  }
});

test('photo loader accepts the supported image MIME types', async () => {
  for (const mimeType of ['image/jpeg', 'image/png', 'image/gif', 'image/webp']) {
    const loaded = loadFactory();
    const controller = loaded.create({
      fetchPhoto() {
        return Promise.resolve(response(4, mimeType));
      },
      renderState() {},
      lifecycleTarget: lifecycleTarget()
    });
    assert.equal(await controller.open('supported-pin', 'Supported'), true);
    assert.equal(loaded.created[0].blob.type, mimeType);
  }
});

test('photo loader retains one URL and revokes it on switch close invalidate and pagehide', async () => {
  const loaded = loadFactory();
  const lifecycle = lifecycleTarget();
  const controller = loaded.create({
    fetchPhoto() {
      return Promise.resolve(response());
    },
    renderState() {},
    lifecycleTarget: lifecycle
  });

  await controller.open('a', 'A');
  await controller.open('b', 'B');
  assert.deepEqual(loaded.revoked, ['blob:photo-1']);

  controller.close();
  assert.deepEqual(loaded.revoked, ['blob:photo-1', 'blob:photo-2']);

  await controller.open('c', 'C');
  assert.equal(controller.invalidate('c'), true);
  assert.deepEqual(loaded.revoked, ['blob:photo-1', 'blob:photo-2', 'blob:photo-3']);

  await controller.open('d', 'D');
  lifecycle.emit('pagehide');
  assert.deepEqual(loaded.revoked, [
    'blob:photo-1', 'blob:photo-2', 'blob:photo-3', 'blob:photo-4'
  ]);
  assert.equal(await controller.open('after-destroy', 'After'), false);
});

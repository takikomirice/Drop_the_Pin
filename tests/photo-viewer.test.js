const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const indexCss = indexHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const sharedCss = sharedHtml.match(/<style>([\s\S]*?)<\/style>/)[1];
const indexMainScriptStart = indexHtml.lastIndexOf('<script>', indexHtml.indexOf('const appStartupStartedAt'));
const sharedMainScriptStart = sharedHtml.lastIndexOf('<script>', sharedHtml.indexOf('const SHARED_DEFAULT_COLOR'));
const indexBody = indexHtml.slice(indexHtml.indexOf('<body'), indexMainScriptStart)
  .replace(/<script(?:\s[^>]*)?>[\s\S]*?<\/script>/gi, '');
const sharedBody = sharedHtml.slice(sharedHtml.indexOf('<body'), sharedMainScriptStart)
  .replace(/<script(?:\s[^>]*)?>[\s\S]*?<\/script>/gi, '');

function functionSource(source, name) {
  const mainMarker = source === indexHtml ? 'const appStartupStartedAt' : 'const SHARED_DEFAULT_COLOR';
  const searchStart = source.lastIndexOf('<script>', source.indexOf(mainMarker));
  const start = source.indexOf(`function ${name}(`, searchStart);
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
  assert.fail(`Could not parse function ${name}`);
}

function openingTag(body, id) {
  const match = body.match(new RegExp(`<[^>]+id=["']${id}["'][^>]*>`));
  assert.ok(match, `Expected #${id}`);
  return match[0];
}

function elementBlock(body, id, endNeedle) {
  const start = body.indexOf(`id="${id}"`);
  assert.notEqual(start, -1, `Expected #${id}`);
  const tagStart = body.lastIndexOf('<', start);
  const end = body.indexOf(endNeedle, start);
  assert.notEqual(end, -1, `Expected ${endNeedle} after #${id}`);
  return body.slice(tagStart, end + endNeedle.length);
}

function cssRule(source, selector) {
  const escaped = selector.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const match = source.match(new RegExp(`${escaped}\\s*\\{([^}]*)\\}`));
  assert.ok(match, `Expected CSS rule ${selector}`);
  return match[1];
}

function containedSize(sourceWidth, sourceHeight, frameWidth, frameHeight) {
  const scale = Math.min(frameWidth / sourceWidth, frameHeight / sourceHeight);
  return { width: sourceWidth * scale, height: sourceHeight * scale };
}

test('index and shared each expose exactly one labelled protected photo viewer', () => {
  assert.equal((indexBody.match(/id="photo-viewer-overlay"/g) || []).length, 1);
  assert.equal((sharedBody.match(/id="shared-photo-viewer-overlay"/g) || []).length, 1);

  const indexOverlay = openingTag(indexBody, 'photo-viewer-overlay');
  assert.match(indexOverlay, /class="[^"]*sheet-overlay[^"]*photo-viewer-overlay/);
  assert.match(indexOverlay, /role="dialog"/);
  assert.match(indexOverlay, /aria-modal="true"/);
  assert.match(indexOverlay, /aria-labelledby="photo-viewer-title"/);
  const indexViewer = elementBlock(indexBody, 'photo-viewer-overlay', '</div>');
  assert.match(indexViewer, /id="photo-viewer-title"/);
  assert.match(indexViewer, /id="photo-viewer-status"[^>]*role="status"[^>]*aria-live="polite"/);
  assert.match(indexViewer, /id="photo-viewer-image" class="[^"]*photo-fit-contain[^"]*protected-photo[^"]*" draggable="false"/);
  assert.match(indexViewer, /id="photo-viewer-close"[^>]*data-overlay-initial-focus/);

  const sharedOverlay = openingTag(sharedBody, 'shared-photo-viewer-overlay');
  assert.match(sharedOverlay, /class="[^"]*sheet-overlay[^"]*photo-viewer-overlay/);
  assert.match(sharedOverlay, /role="dialog"/);
  assert.match(sharedOverlay, /aria-modal="true"/);
  assert.match(sharedOverlay, /aria-labelledby="shared-photo-viewer-title"/);
  const sharedViewer = sharedBody.slice(sharedBody.indexOf('<div id="shared-photo-viewer-overlay"'));
  assert.match(sharedViewer, /id="shared-photo-viewer-title"/);
  assert.match(sharedViewer, /id="shared-photo-viewer-status"[^>]*role="status"[^>]*aria-live="polite"/);
  assert.match(sharedViewer, /id="shared-photo-viewer-image" class="[^"]*photo-fit-contain[^"]*protected-photo[^"]*" draggable="false"/);
  assert.match(sharedViewer, /id="shared-photo-viewer-close"[^>]*data-shared-initial-focus/);
});

test('photo surfaces carry explicit fit classes and native viewer triggers', () => {
  const contracts = [
    [indexBody, 'upload-preview-trigger', 'photo-viewer-overlay', 'upload-preview', 'photo-fit-contain'],
    [indexBody, 'import-preview-image-trigger', 'photo-viewer-overlay', 'import-preview-image', 'photo-fit-contain'],
    [indexBody, 'pin-detail-image-trigger', 'photo-viewer-overlay', 'pin-detail-image', 'photo-fit-cover'],
    [sharedBody, 'shared-detail-image-trigger', 'shared-photo-viewer-overlay', 'shared-detail-image', 'photo-fit-cover']
  ];
  contracts.forEach(([body, triggerId, viewerId, imageId, fitClass]) => {
    const trigger = openingTag(body, triggerId);
    assert.match(trigger, /^<button\b/);
    assert.match(trigger, /type="button"/);
    assert.match(trigger, /class="[^"]*photo-viewer-trigger/);
    assert.match(trigger, new RegExp(`aria-controls="${viewerId}"`));
    assert.match(trigger, /hidden/);
    assert.match(trigger, /disabled/);
    const image = openingTag(body, imageId);
    assert.match(image, new RegExp(`class="[^"]*${fitClass}`));
  });

  assert.match(cssRule(indexCss, '.photo-fit-cover'), /object-fit:\s*cover/);
  assert.match(cssRule(indexCss, '.photo-fit-contain'), /object-fit:\s*contain/);
  assert.match(cssRule(sharedCss, '.photo-fit-cover'), /object-fit:\s*cover/);
  assert.match(cssRule(sharedCss, '.photo-fit-contain'), /object-fit:\s*contain/);
  assert.match(cssRule(indexCss, '.list-thumb img'), /object-fit:\s*cover/);
  assert.doesNotMatch(functionSource(indexHtml, 'renderListThumb'), /<img/);
  assert.doesNotMatch(functionSource(sharedHtml, 'sharedPinListThumbMarkup'), /<img/);
});

test('viewer layout contains every aspect ratio within its frame and preserves safe areas', () => {
  [indexCss, sharedCss].forEach((css) => {
    const image = cssRule(css, '.photo-viewer-image');
    ['width', 'height', 'max-width', 'max-height'].forEach((property) => {
      assert.match(image, new RegExp(`${property}:\\s*100%`));
    });
    assert.match(image, /object-fit:\s*contain/);
    assert.match(image, /touch-action:\s*pinch-zoom/);
    const overlay = cssRule(css, '.photo-viewer-overlay');
    assert.match(overlay, /safe-area-inset-top/);
    assert.match(overlay, /safe-area-inset-right/);
    assert.match(overlay, /safe-area-inset-bottom/);
    assert.match(overlay, /safe-area-inset-left/);
    assert.match(overlay, /overflow:\s*hidden/);
  });

  [
    [300, 900], [1200, 400], [640, 640], [32, 24]
  ].forEach(([width, height]) => {
    const fitted = containedSize(width, height, 390, 700);
    assert.ok(fitted.width <= 390 + Number.EPSILON);
    assert.ok(fitted.height <= 700 + Number.EPSILON);
    assert.ok(Math.abs((fitted.width / fitted.height) - (width / height)) < 1e-12);
  });
});

test('index viewer is routed through the existing stack, Escape dispatcher, and safe backdrop gesture', () => {
  const inventory = indexHtml.match(/const MAIN_DISMISSIBLE_OVERLAY_IDS = Object\.freeze\(\[([\s\S]*?)\]\);/)[1];
  const backdrop = indexHtml.slice(
    indexHtml.indexOf('const BACKDROP_DISMISS_OVERLAY_IDS = ['),
    indexHtml.indexOf('];', indexHtml.indexOf('const BACKDROP_DISMISS_OVERLAY_IDS = ['))
  );
  assert.match(inventory, /'photo-viewer-overlay'/);
  assert.match(backdrop, /'photo-viewer-overlay'/);
  assert.match(functionSource(indexHtml, 'dismissOverlayById'), /photo-viewer-overlay[\s\S]*closePhotoViewer/);
  assert.match(functionSource(indexHtml, 'closePhotoViewer'), /closeOverlay\('photo-viewer-overlay'/);
  assert.match(indexHtml, /\['help-overlay'\]\.concat\(MAIN_DISMISSIBLE_OVERLAY_IDS\)\.forEach/);
  assert.match(indexHtml, /const overlay = document\.getElementById\(id\);\s*if \(!overlay\) return;\s*setupOverlayBackdropDismissal\(overlay/);
});

test('shared viewer is the top shared surface and closes before the detail', () => {
  const inventory = sharedHtml.match(/const SHARED_DISMISSIBLE_SURFACE_IDS = Object\.freeze\(\[([\s\S]*?)\]\);/)[1];
  assert.match(inventory, /'shared-photo-viewer-overlay'/);
  assert.match(functionSource(sharedHtml, 'dispatchSharedEscape'), /shared-photo-viewer-overlay[\s\S]*closeSharedPhotoViewer/);
  assert.match(functionSource(sharedHtml, 'closeSharedPhotoViewer'), /closeSharedSurface\('shared-photo-viewer-overlay'/);
  assert.match(sharedHtml, /setupSharedOverlayBackdropDismissal\(document\.getElementById\('shared-photo-viewer-overlay'\), closeSharedPhotoViewer\)/);
  assert.match(functionSource(sharedHtml, 'closeSharedDetail'), /closeSharedPhotoViewer[\s\S]*closeSharedSurface\('shared-detail-overlay'/);
});

function fakeClassList(initial = []) {
  const values = new Set(initial);
  return {
    add(...items) { items.forEach((item) => values.add(item)); },
    remove(...items) { items.forEach((item) => values.delete(item)); },
    contains(item) { return values.has(item); }
  };
}

function fakeElement(id, tagName = 'div') {
  const attributes = Object.create(null);
  const listeners = Object.create(null);
  return {
    id,
    tagName: tagName.toUpperCase(),
    attributes,
    listeners,
    classList: fakeClassList(),
    style: {},
    hidden: false,
    disabled: false,
    src: '',
    alt: '',
    textContent: '',
    image: null,
    setAttribute(name, value) {
      attributes[name] = String(value);
      if (name === 'src') this.src = String(value);
    },
    getAttribute(name) {
      return Object.prototype.hasOwnProperty.call(attributes, name) ? attributes[name] : null;
    },
    removeAttribute(name) {
      delete attributes[name];
      if (name === 'src') this.src = '';
    },
    addEventListener(name, listener) {
      if (!listeners[name]) listeners[name] = [];
      listeners[name].push(listener);
    },
    removeEventListener(name, listener) {
      if (!listeners[name]) return;
      listeners[name] = listeners[name].filter((candidate) => candidate !== listener);
    },
    dispatch(name) {
      (listeners[name] || []).slice().forEach((listener) => listener({ target: this, currentTarget: this }));
    },
    querySelector(selector) { return selector === 'img' ? this.image : null; }
  };
}

function createViewerHarness(source, shared) {
  const prefix = shared ? 'Shared' : '';
  const lowerPrefix = shared ? 'shared' : '';
  const ids = shared ? {
    overlay: 'shared-photo-viewer-overlay', title: 'shared-photo-viewer-title',
    image: 'shared-photo-viewer-image', status: 'shared-photo-viewer-status'
  } : {
    overlay: 'photo-viewer-overlay', title: 'photo-viewer-title',
    image: 'photo-viewer-image', status: 'photo-viewer-status'
  };
  const elements = {
    [ids.overlay]: fakeElement(ids.overlay),
    [ids.title]: fakeElement(ids.title),
    [ids.image]: fakeElement(ids.image, 'img'),
    [ids.status]: fakeElement(ids.status)
  };
  const calls = { open: 0, close: 0, create: 0, revoke: 0 };
  const stateName = shared ? 'sharedPhotoViewerState' : 'photoViewerState';
  const context = {
    document: { getElementById(id) { return elements[id] || null; } },
    URL: {
      createObjectURL() { calls.create += 1; return 'blob:unexpected'; },
      revokeObjectURL() { calls.revoke += 1; }
    },
    [stateName]: {
      sourceTrigger: null, sourceUrl: '', requestToken: 0,
      loadHandler: null, errorHandler: null
    }
  };
  context[shared ? 'openSharedSurface' : 'openOverlay'] = function(id) {
    calls.open += 1;
    elements[id].classList.add('open');
  };
  context[shared ? 'closeSharedSurface' : 'closeOverlay'] = function(id) {
    calls.close += 1;
    elements[id].classList.remove('open');
  };
  const names = [
    `get${prefix}PhotoViewerSourceUrl`, `set${prefix}PhotoViewerStatus`,
    `remove${prefix}PhotoViewerImageListeners`, `update${prefix}PhotoViewerTrigger`,
    `open${prefix}PhotoViewer`, `open${prefix}PhotoViewerFromTrigger`,
    `close${prefix}PhotoViewer`, `close${prefix}PhotoViewerForTrigger`,
    `bind${prefix}PhotoViewerTrigger`
  ];
  names.forEach((name) => {
    context[name] = vm.runInNewContext(`(${functionSource(source, name)})`, context);
  });
  return {
    context, elements, calls, state: context[stateName], ids,
    names: {
      update: `update${prefix}PhotoViewerTrigger`, openFrom: `open${prefix}PhotoViewerFromTrigger`,
      close: `close${prefix}PhotoViewer`, bind: `bind${prefix}PhotoViewerTrigger`
    },
    trigger(id) {
      const trigger = fakeElement(id, 'button');
      const image = fakeElement(`${id}-image`, 'img');
      trigger.image = image;
      return { trigger, image };
    },
    lowerPrefix
  };
}

[['index', indexHtml, false], ['shared', sharedHtml, true]].forEach(([label, source, shared]) => {
  test(`${label} viewer ignores stale image events, avoids duplicate opens, and never owns URLs`, () => {
    const harness = createViewerHarness(source, shared);
    const first = harness.trigger(`${label}-first`);
    const second = harness.trigger(`${label}-second`);
    harness.context[harness.names.update](first.trigger, first.image, {
      sourceUrl: 'blob:first', title: '最初', alt: '最初の写真'
    });
    harness.context[harness.names.bind](first.trigger);
    first.trigger.dispatch('click');
    assert.equal(harness.calls.open, 1);
    assert.equal(harness.elements[harness.ids.status].textContent, '写真を読み込み中...');
    const staleLoad = harness.state.loadHandler;

    harness.context[harness.names.update](second.trigger, second.image, {
      sourceUrl: 'https://example.test/second.jpg', title: '次', alt: '次の写真'
    });
    harness.context[harness.names.openFrom](second.trigger);
    assert.equal(harness.calls.open, 1, 'an already-open viewer must keep one stack record');
    staleLoad();
    assert.equal(harness.elements[harness.ids.status].textContent, '写真を読み込み中...');
    harness.state.errorHandler();
    assert.equal(harness.elements[harness.ids.status].textContent, '写真を表示できませんでした');

    harness.context[harness.names.openFrom](second.trigger);
    harness.state.loadHandler();
    assert.equal(harness.elements[harness.ids.status].textContent, '');
    harness.context[harness.names.close]();
    assert.equal(harness.elements[harness.ids.image].src, '');
    assert.equal(harness.state.sourceTrigger, null);
    assert.equal(harness.state.sourceUrl, '');
    assert.equal(harness.calls.create, 0);
    assert.equal(harness.calls.revoke, 0);
  });
});

test('surface cleanup closes viewers before owners revoke or replace photo URLs', () => {
  const clearUpload = functionSource(indexHtml, 'clearUploadPhotoState');
  assert.ok(clearUpload.indexOf('closePhotoViewerForTrigger') < clearUpload.indexOf('URL.revokeObjectURL'));
  assert.match(functionSource(indexHtml, 'handleUploadPhotoSelected'), /clearUploadPhotoState\(\)[\s\S]*URL\.createObjectURL/);

  const renderEditor = functionSource(indexHtml, 'renderEditor');
  assert.match(renderEditor, /updatePhotoViewerTrigger[\s\S]*previewUrl/);
  const deleteItem = functionSource(indexHtml, 'handleDelete');
  assert.ok(deleteItem.indexOf('clearImportPreviewPhoto') < deleteItem.indexOf('removeDraftItem'));
  const closeImport = functionSource(indexHtml, 'close');
  assert.ok(closeImport.indexOf('clearImportPreviewPhoto') < closeImport.indexOf('releaseJobResources'));
  assert.match(functionSource(indexHtml, 'clearImportPreviewPhoto'), /closePhotoViewerForTrigger[\s\S]*sourceUrl:\s*''/);

  assert.match(functionSource(indexHtml, 'openPinDetail'), /updatePhotoViewerTrigger/);
  assert.match(functionSource(indexHtml, 'closePinDetail'), /closePhotoViewerForTrigger[\s\S]*closeOverlay/);
  assert.match(functionSource(sharedHtml, 'openSharedDetail'), /updateSharedPhotoViewerTrigger/);
  assert.match(functionSource(sharedHtml, 'closeSharedDetail'), /closeSharedPhotoViewerForTrigger[\s\S]*closeSharedSurface/);
});

test('viewer controller state retains no binary, Drive, credential, or base64 payload fields', () => {
  const indexState = indexHtml.match(/const photoViewerState = \{([\s\S]*?)\n\s*\};/);
  const sharedState = sharedHtml.match(/const sharedPhotoViewerState = \{([\s\S]*?)\n\s*\};/);
  assert.ok(indexState);
  assert.ok(sharedState);
  [indexState[1], sharedState[1]].forEach((stateSource) => {
    assert.doesNotMatch(stateSource, /File|Blob|base64|Drive|fileId|credential|authToken|payload/i);
    assert.match(stateSource, /sourceTrigger/);
    assert.match(stateSource, /sourceUrl/);
    assert.match(stateSource, /requestToken/);
  });
  [
    functionSource(indexHtml, 'openPhotoViewer'), functionSource(indexHtml, 'closePhotoViewer'),
    functionSource(sharedHtml, 'openSharedPhotoViewer'), functionSource(sharedHtml, 'closeSharedPhotoViewer')
  ].forEach((source) => {
    assert.doesNotMatch(source, /createObjectURL|revokeObjectURL|File|Blob|base64|Drive|fileId|editToken/i);
  });
});

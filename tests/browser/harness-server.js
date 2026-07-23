const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');

const host = '127.0.0.1';
const port = Number(process.env.DTP_BROWSER_PORT || 4173);
const projectRoot = path.resolve(__dirname, '..', '..');

function readProjectFile(name) {
  return fs.readFileSync(path.join(projectRoot, name), 'utf8');
}

const indexSource = readProjectFile('index.html');
const sharedSource = readProjectFile('shared.html');

const audioEditorStartMarker = '<template data-dtp-audio-editor-boundary="start"></template>';
const audioEditorEndMarker = '<template data-dtp-audio-editor-boundary="end"></template>';
const audioPlayerStartMarker = '<template data-dtp-audio-player-boundary="start"></template>';
const audioPlayerEndMarker = '<template data-dtp-audio-player-boundary="end"></template>';
const audioVendorStartMarker = 'AUDIO_VENDOR_BUNDLE_START';
const audioVendorEndMarker = 'AUDIO_VENDOR_BUNDLE_END';

function findUniqueText(source, text, name) {
  const first = String(source).indexOf(text);
  if (first < 0) throw new Error(`${name} is missing.`);
  if (String(source).indexOf(text, first + text.length) >= 0) {
    throw new Error(`${name} is duplicated.`);
  }
  return first;
}

function extractUniqueDelimitedRegion(source, startText, endText, name) {
  const start = findUniqueText(source, startText, `${name} start boundary`);
  const end = findUniqueText(source, endText, `${name} end boundary`);
  if (end < start + startText.length) {
    throw new Error(`${name} boundaries are reversed.`);
  }
  return {
    content: source.slice(start + startText.length, end),
    start: start,
    end: end,
    endExclusive: end + endText.length
  };
}

function replaceUniqueText(source, target, replacement, name) {
  const start = findUniqueText(source, target, name);
  return source.slice(0, start) + replacement + source.slice(start + target.length);
}

function lineBreakAfter(source, index, name) {
  if (source.startsWith('\r\n', index)) return '\r\n';
  if (source.startsWith('\n', index)) return '\n';
  throw new Error(`${name} must be followed by a line break.`);
}

function lineBreakBefore(source, index, name) {
  if (source.slice(index - 2, index) === '\r\n') return '\r\n';
  if (source.slice(index - 1, index) === '\n') return '\n';
  throw new Error(`${name} must be preceded by a line break.`);
}

const audioEditorRegion = extractUniqueDelimitedRegion(
  indexSource,
  audioEditorStartMarker,
  audioEditorEndMarker,
  'Index audio editor'
);
const indexAudioPlayerRegion = extractUniqueDelimitedRegion(
  indexSource,
  audioPlayerStartMarker,
  audioPlayerEndMarker,
  'Index audio player'
);
const sharedAudioPlayerRegion = extractUniqueDelimitedRegion(
  sharedSource,
  audioPlayerStartMarker,
  audioPlayerEndMarker,
  'Shared audio player'
);
if (sharedAudioPlayerRegion.content !== indexAudioPlayerRegion.content) {
  throw new Error('Index and shared audio player content must be byte-identical.');
}
const audioEditorContent = audioEditorRegion.content;
const audioPlayerContent = indexAudioPlayerRegion.content;

const audioVendorRegion = extractUniqueDelimitedRegion(
  indexSource,
  audioVendorStartMarker,
  audioVendorEndMarker,
  'Index audio vendor bundle'
);
if (audioVendorRegion.start !== 0) {
  throw new Error('Index audio vendor bundle must be the document prefix.');
}
const audioVendorMatch = /^(?:\r\n|\n)<script>(?:\r\n|\n)([\s\S]*?)(?:\r\n|\n)<\/script>(?:\r\n|\n)$/.exec(audioVendorRegion.content);
if (!audioVendorMatch) throw new Error('Audio vendor bundle must use the exact plain script wrapper.');
const audioVendorSource = audioVendorMatch[1];

function withoutAudioVendorPrefix(source) {
  const region = extractUniqueDelimitedRegion(
    source,
    audioVendorStartMarker,
    audioVendorEndMarker,
    'Index audio vendor bundle'
  );
  if (region.start !== 0) {
    throw new Error('Index audio vendor bundle must be the document prefix.');
  }
  const lineBreak = lineBreakAfter(source, region.endExclusive, 'Index audio vendor end boundary');
  const documentStart = region.endExclusive + lineBreak.length;
  if (!source.startsWith('<!DOCTYPE html>', documentStart)) {
    throw new Error('Index document must begin at the doctype after the audio vendor prefix.');
  }
  return source.slice(documentStart);
}

function evaluateEditTokenWrapper(source) {
  const region = extractUniqueDelimitedRegion(
    source,
    audioEditorStartMarker,
    audioEditorEndMarker,
    'Index audio editor'
  );
  const openingDirective = '  <? if (editToken) { ?>';
  const closingDirective = '  <? } ?>';
  const openingLineBreak = lineBreakBefore(source, region.start, 'Index audio editor start boundary');
  const openingEnd = region.start - openingLineBreak.length;
  const openingStart = openingEnd - openingDirective.length;
  const closingLineBreak = lineBreakAfter(source, region.endExclusive, 'Index audio editor end boundary');
  const closingStart = region.endExclusive + closingLineBreak.length;
  const uniqueOpeningStart = findUniqueText(source, openingDirective, 'Index edit-token opening directive');
  const uniqueClosingStart = findUniqueText(source, closingDirective, 'Index edit-token closing directive');
  if (uniqueOpeningStart !== openingStart || uniqueClosingStart !== closingStart) {
    throw new Error('Index edit-token wrapper is not exactly adjacent to the audio editor boundaries.');
  }
  return source.slice(0, openingStart)
    + source.slice(region.start, region.endExclusive)
    + source.slice(closingStart + closingDirective.length);
}

function extractBetween(source, startText, endText) {
  const start = source.indexOf(startText);
  const end = source.indexOf(endText, start + startText.length);
  if (start < 0 || end < 0) {
    throw new Error(`Production source region was not found: ${startText}`);
  }
  return source.slice(start, end);
}

function exposeConstIife(source, name, nextName) {
  const region = extractBetween(
    source,
    `const ${name} = (function()`,
    `const ${nextName}`
  );
  return region.replace(`const ${name} =`, `window.${name} =`);
}

function extractFunction(source, name) {
  const marker = `function ${name}(`;
  const start = source.indexOf(marker);
  if (start < 0) throw new Error(`Production function was not found: ${name}`);
  const bodyStart = source.indexOf('{', start + marker.length);
  let depth = 0;
  let quote = '';
  let escaped = false;
  let lineComment = false;
  let blockComment = false;
  for (let index = bodyStart; index < source.length; index += 1) {
    const character = source[index];
    const next = source[index + 1];
    if (lineComment) {
      if (character === '\n') lineComment = false;
      continue;
    }
    if (blockComment) {
      if (character === '*' && next === '/') {
        blockComment = false;
        index += 1;
      }
      continue;
    }
    if (quote) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === quote) quote = '';
      continue;
    }
    if (character === '/' && next === '/') {
      lineComment = true;
      index += 1;
      continue;
    }
    if (character === '/' && next === '*') {
      blockComment = true;
      index += 1;
      continue;
    }
    if (character === '\'' || character === '"' || character === '`') {
      quote = character;
      continue;
    }
    if (character === '{') depth += 1;
    if (character === '}') {
      depth -= 1;
      if (depth === 0) return source.slice(start, index + 1);
    }
  }
  throw new Error(`Production function was incomplete: ${name}`);
}

const workflowModules = [
  exposeConstIife(indexSource, 'MediaImportShell', 'ImportAudioItemProcessor'),
  exposeConstIife(indexSource, 'ImportAudioItemProcessor', 'AudioPinImportWorkflow'),
  exposeConstIife(indexSource, 'AudioPinImportWorkflow', 'DrivePhotoImportSourceCore')
].join('\n');
const editPlayerFactory = extractFunction(indexSource, 'createPinAudioPlayerController');
const sharedPlayerFactory = extractFunction(sharedSource, 'createSharedPinAudioPlayerController');

const baseStyles = `
  :root {
    --border: #d5dee5;
    --color-surface-muted: #f3f6f8;
    --text-sub: #607080;
    --dialog-viewport-height: 100dvh;
    --app-modal-viewport-inset: 28px;
    font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
  }
  * { box-sizing: border-box; }
  [hidden] { display: none !important; }
  html, body { width: 100%; min-height: 100%; margin: 0; overflow-x: hidden; }
  body { padding: 16px; color: #24313d; background: #eef3f6; }
  button, input { font: inherit; }
  button, .ghost-btn { min-width: 44px; min-height: 44px; }
  .ghost-btn { padding: 8px 12px; border: 1px solid #9aacb7; border-radius: 8px; background: #fff; }
  .sheet-overlay {
    position: fixed;
    z-index: 20;
    inset: 0;
    display: none;
    align-items: center;
    justify-content: center;
    padding: 14px;
    overflow: hidden;
    background: #24313d78;
  }
  .sheet-overlay.open { display: flex; }
  .sheet-body {
    display: flex;
    flex-direction: column;
    min-width: 0;
    max-width: 100%;
    overflow: hidden;
    border-radius: 14px;
    background: #fff;
  }
  .dialog-scroll-content { min-height: 0; overflow: auto; }
  .surface {
    width: min(100%, 640px);
    margin: 12px auto;
    padding: 16px;
    overflow-x: hidden;
    border: 1px solid var(--border);
    border-radius: 14px;
    background: #fff;
  }
  .button-row { display: flex; flex-wrap: wrap; gap: 8px; }
  .detail-surface[hidden], .source-chooser[hidden], .workflow-preview[hidden] { display: none !important; }
  .detail-surface, .source-chooser, .workflow-preview {
    width: min(100%, 620px);
    margin: 12px auto;
    padding: 14px;
    overflow-x: hidden;
    border: 1px solid var(--border);
    border-radius: 12px;
    background: #fff;
  }
  .field { display: grid; gap: 5px; margin: 10px 0; }
  .field input { width: 100%; min-width: 0; min-height: 44px; }
  .status { min-height: 1.5em; white-space: pre-wrap; }
`;

function commonInstrumentationScript() {
  return `<script>
    window.__runtime = {
      pageErrors: [],
      unhandledRejections: [],
      mainReady: true,
      geocoderCalls: 0
    };
    window.addEventListener('error', function(event) {
      window.__runtime.pageErrors.push(String(event.error && event.error.message || event.message || 'error'));
    });
    window.addEventListener('unhandledrejection', function(event) {
      window.__runtime.unhandledRejections.push(String(event.reason && event.reason.message || event.reason || 'rejection'));
    });
  </script>`;
}

function gasMockScript() {
  return `<script>
    (function() {
      const queues = new Map();
      const deferred = new Map();
      const calls = [];

      function audioResponse(seed) {
        const byte = String.fromCharCode(Number(seed || 0) & 255);
        const binary = byte.repeat(8192).repeat(128);
        return {
          ok: true,
          mimeType: 'audio/mpeg',
          byteLength: binary.length,
          base64: btoa(binary)
        };
      }

      function enqueue(method, item) {
        const list = queues.get(method) || [];
        list.push(item || {});
        queues.set(method, list);
      }

      function deliver(item, success, failure) {
        if (item && item.drop === true) return;
        if (item && Object.prototype.hasOwnProperty.call(item, 'failure')) {
          failure(item.failure instanceof Error ? item.failure : new Error(String(item.failure)));
          return;
        }
        const response = item && Object.prototype.hasOwnProperty.call(item, 'response')
          ? item.response
          : (item && Object.prototype.hasOwnProperty.call(item, 'audioSeed')
            ? audioResponse(item.audioSeed) : { ok: true });
        success(response);
      }

      function invoke(method, payload, success, failure) {
        calls.push({ method: String(method), payload: payload });
        const list = queues.get(method) || [];
        const item = list.length ? list.shift() : { failure: 'No mocked GAS response.' };
        queues.set(method, list);
        if (item && item.defer) {
          deferred.set(String(item.defer), { item: item, success: success, failure: failure });
          return;
        }
        if (item && Number(item.delayMs) > 0) {
          setTimeout(function() { deliver(item, success, failure); }, Number(item.delayMs));
          return;
        }
        queueMicrotask(function() { deliver(item, success, failure); });
      }

      function settle(key, override, reject) {
        const pending = deferred.get(String(key));
        if (!pending) throw new Error('Deferred GAS response was not found: ' + key);
        deferred.delete(String(key));
        if (reject) {
          pending.failure(override instanceof Error ? override : new Error(String(override)));
          return;
        }
        const item = Object.assign({}, pending.item);
        if (override !== undefined) item.response = override;
        deliver(item, pending.success, pending.failure);
      }

      window.__gasMock = {
        calls: calls,
        enqueue: enqueue,
        resolve: function(key, response) { settle(key, response, false); },
        reject: function(key, error) { settle(key, error, true); },
        audioResponse: audioResponse,
        clear: function() { queues.clear(); deferred.clear(); calls.splice(0); }
      };

      function createRunner() {
        let success = function() {};
        let failure = function() {};
        const target = {
          withSuccessHandler: function(handler) { success = handler; return proxy; },
          withFailureHandler: function(handler) { failure = handler; return proxy; }
        };
        const proxy = new Proxy(target, {
          get: function(object, property) {
            if (property in object) return object[property];
            return function(payload) { invoke(String(property), payload, success, failure); };
          }
        });
        return proxy;
      }

      window.google = { script: {} };
      Object.defineProperty(window.google.script, 'run', { get: createRunner });
      window.withGAS = function(name, payload) {
        return new Promise(function(resolve, reject) {
          google.script.run.withSuccessHandler(resolve).withFailureHandler(reject)[name](payload);
        });
      };
    })();
  </script>`;
}

function productionBrowserStubsScript() {
  return `<script>
    (function() {
      function chainObject(overrides) {
        const target = Object.assign({
          setView: function() { return proxy; },
          addTo: function() { return proxy; },
          remove: function() { return proxy; },
          removeFrom: function() { return proxy; },
          on: function() { return proxy; },
          off: function() { return proxy; },
          once: function() { return proxy; },
          bindPopup: function() { return proxy; },
          openPopup: function() { return proxy; },
          closePopup: function() { return proxy; },
          clearLayers: function() { return proxy; },
          addLayer: function() { return proxy; },
          removeLayer: function() { return proxy; },
          invalidateSize: function() { return proxy; },
          flyTo: function() { return proxy; },
          fitBounds: function() { return proxy; },
          createPane: function() { return { style: {} }; },
          getPane: function() { return { style: {} }; },
          getZoom: function() { return 5; },
          getCenter: function() { return { lat: 36.5, lng: 138 }; },
          getBounds: function() { return { contains: function() { return true; } }; },
          getElement: function() { return null; },
          latLngToContainerPoint: function() { return { x: 0, y: 0 }; },
          containerPointToLatLng: function() { return { lat: 36.5, lng: 138 }; }
        }, overrides || {});
        const proxy = new Proxy(target, {
          get: function(object, property) {
            if (property in object) return object[property];
            return function() { return proxy; };
          }
        });
        return proxy;
      }
      window.L = {
        map: function() { return chainObject(); },
        tileLayer: function() { return chainObject(); },
        marker: function() { return chainObject(); },
        polyline: function() { return chainObject(); },
        featureGroup: function() { return chainObject(); },
        divIcon: function(options) { return options || {}; },
        point: function(x, y) { return { x: Number(x || 0), y: Number(y || 0) }; },
        latLngBounds: function() { return chainObject({ isValid: function() { return true; } }); },
        control: { zoom: function() { return chainObject(); } }
      };
      window.Sortable = { create: function() { return { destroy: function() {}, option: function() {} }; } };
      window.EXIF = { getData: function(_file, callback) { callback.call(_file); }, getTag: function() { return null; } };
      window.ExifReader = { load: function() { return {}; } };
      window.heicTo = function() { return Promise.reject(new Error('HEIC is disabled in the browser harness.')); };

      const media = HTMLMediaElement.prototype;
      media.play = function() { return Promise.resolve(); };
      media.pause = function() {};
      media.load = function() {};
    })();
  </script>`;
}

function withoutRemoteAssets(source) {
  return String(source)
    .replace(/\s*<link[^>]+href=["']https?:[^>]+>\s*/gi, '\n')
    .replace(/\s*<script[^>]+src=["']https?:[^>]*><\/script>\s*/gi, '\n');
}

function ensureNoTemplateDirectives(source, name) {
  if (String(source).includes('<?')) {
    throw new Error(`${name} contains an unprocessed Apps Script template directive.`);
  }
  return source;
}

function productionEditPage() {
  let html = withoutAudioVendorPrefix(indexSource);
  html = replaceUniqueText(
    html,
    "<?!= JSON.stringify(editToken || '') ?>",
    JSON.stringify('edit-token-browser-test'),
    'Index edit-token JSON expression'
  );
  html = evaluateEditTokenWrapper(html);
  html = withoutRemoteAssets(html);
  html = html.replace('<head>', '<head>' + commonInstrumentationScript() + gasMockScript() + productionBrowserStubsScript());
  html = html.replace(
    '    initializeApp();',
    `    state.initializing = false;
    state.narrowView = false;
    window.__productionEdit = Object.freeze({
      state: state,
      openPinDetail: openPinDetail,
      closePinDetail: closePinDetail,
      openPinAudioSource: openPinAudioSource,
      removeAudioFromPinDetail: removeAudioFromPinDetail,
      handleLocalAudioImportSelected: handleLocalAudioImportSelected,
      handleDriveAudioImportButtonClick: handleDriveAudioImportButtonClick,
      saveAudioImportedPin: saveAudioImportedPin,
      pinAudioPlayer: pinAudioPlayer,
      initializeApp: initializeApp,
      warmAudioVendorBundle: warmAudioVendorBundle,
      loadAudioVendorBundle: loadAudioVendorBundle,
      map: map
    });
    window.__productionInitializationSuppressed = true;`
  );
  return ensureNoTemplateDirectives(html, 'index.html');
}

function productionSharedPage() {
  let html = sharedSource;
  html = replaceUniqueText(
    html,
    "<?!= JSON.stringify(execUrl || '') ?>",
    JSON.stringify(''),
    'Shared exec URL JSON expression'
  );
  html = replaceUniqueText(
    html,
    "<?!= JSON.stringify(token || '') ?>",
    JSON.stringify('share-token-browser-test'),
    'Shared token JSON expression'
  );
  html = withoutRemoteAssets(html);
  html = html.replace('<head>', '<head>' + commonInstrumentationScript() + gasMockScript() + productionBrowserStubsScript());
  html = html.replace(
    '    initializeSharedView();',
    `    window.__productionShared = Object.freeze({
      state: state,
      openSharedDetail: openSharedDetail,
      closeSharedDetail: closeSharedDetail,
      renderSharedPins: renderSharedPins,
      initializeSharedView: initializeSharedView,
      sharedPinAudioPlayer: sharedPinAudioPlayer
    });
    window.__productionInitializationSuppressed = true;`
  );
  return ensureNoTemplateDirectives(html, 'shared.html');
}

function page(title, body) {
  return `<!doctype html>
  <html lang="ja">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width,initial-scale=1">
      <title>${title}</title>
      <style>${baseStyles}</style>
    </head>
    <body>${commonInstrumentationScript()}${body}</body>
  </html>`;
}

function editorPage() {
  const catchPrefix = '    } catch (error) {';
  const generationGuard = '      if (generation === encodeGeneration && inputGeneration === loadGeneration)';
  const catchLineBreak = audioEditorContent.includes(catchPrefix + '\r\n' + generationGuard) ? '\r\n' : '\n';
  const instrumentedEditorContent = replaceUniqueText(
    audioEditorContent,
    catchPrefix + catchLineBreak + generationGuard,
    catchPrefix + catchLineBreak
      + "      global.__lastEncodeError = String(error && error.stack || error && error.message || error || 'unknown');"
      + catchLineBreak + generationGuard,
    'Audio editor encode error handler'
  );
  const bridge = `<script>
    (function() {
      const returnFocus = new Map();
      let vendorPromise = null;
      const params = new URLSearchParams(location.search);
      let failuresRemaining = Number(params.get('vendorFailures') || 0);
      let deferredRelease = null;
      window.__vendor = { calls: 0, loaded: false, failures: 0 };

      window.openOverlay = function(id) {
        const overlay = document.getElementById(id);
        if (!overlay) return false;
        returnFocus.set(id, document.activeElement);
        overlay.classList.add('open');
        overlay.removeAttribute('aria-hidden');
        const initial = overlay.querySelector('[data-overlay-initial-focus], button, input');
        if (initial) queueMicrotask(function() { initial.focus(); });
        return true;
      };
      window.closeOverlay = function(id) {
        const overlay = document.getElementById(id);
        if (!overlay) return false;
        overlay.classList.remove('open');
        overlay.setAttribute('aria-hidden', 'true');
        const trigger = returnFocus.get(id);
        returnFocus.delete(id);
        if (trigger && typeof trigger.focus === 'function') queueMicrotask(function() { trigger.focus(); });
        return true;
      };

      function evaluateVendor() {
        return fetch('/audio-vendor.js').then(function(response) {
          if (!response.ok) throw new Error('Vendor fetch failed.');
          return response.text();
        }).then(function(source) {
          (0, eval)(source);
          window.__vendor.loaded = !!(window.Mediabunny && window.MediabunnyMp3Encoder);
          return window.__vendor.loaded;
        });
      }

      window.loadAudioVendorBundle = function() {
        window.__vendor.calls += 1;
        if (window.Mediabunny && window.MediabunnyMp3Encoder) {
          window.__vendor.loaded = true;
          return Promise.resolve(true);
        }
        if (vendorPromise) return vendorPromise;
        if (failuresRemaining > 0) {
          failuresRemaining -= 1;
          window.__vendor.failures += 1;
          return Promise.reject(new Error('Configured vendor failure.'));
        }
        if (params.get('vendorDeferred') === '1' && !deferredRelease) {
          vendorPromise = new Promise(function(resolve) { deferredRelease = resolve; })
            .then(evaluateVendor)
            .finally(function() { vendorPromise = null; });
          return vendorPromise;
        }
        vendorPromise = evaluateVendor().finally(function() { vendorPromise = null; });
        return vendorPromise;
      };
      window.__releaseVendor = function() {
        if (!deferredRelease) return false;
        const release = deferredRelease;
        deferredRelease = null;
        release();
        return true;
      };
    })();
  </script>`;
  const controls = `
    <main class="surface" aria-label="edit harness">
      <div class="button-row">
        <button id="editor-trigger" type="button">音声を編集</button>
        <button id="map-action" type="button">地図を動かす</button>
      </div>
      <output id="map-action-count">0</output>
    </main>`;
  const launch = `<script>
    ${workflowModules}
    document.getElementById('editor-trigger').addEventListener('click', function() {
      window.HemisphereAudioEditor.init();
    });
    document.getElementById('map-action').addEventListener('click', function() {
      const output = document.getElementById('map-action-count');
      output.value = String(Number(output.value || output.textContent || 0) + 1);
      output.textContent = output.value;
    });
    window.__saveRealEditorResult = async function() {
      const editorResult = window.HemisphereAudioEditor.getResult();
      if (!editorResult || !(editorResult.blob instanceof Blob)) {
        throw new Error('A real editor result is required.');
      }
      const operation = MediaImportShell.create({
        createId: (function() {
          let counter = 0;
          return function(prefix) { counter += 1; return prefix + '-real-' + counter; };
        })()
      }).begin({
        mediaKind: 'audio', sourceKind: 'local', operationMode: 'create-pin', selectionLimit: 1
      });
      const sourceFile = new File([new Uint8Array([82, 73, 70, 70])], 'real-source.wav', { type: 'audio/wav' });
      const processor = ImportAudioItemProcessor.create({
        openAudioEditor: function() {
          return Promise.resolve(Object.assign({}, editorResult, { sourceDurationSeconds: 46 }));
        }
      });
      const validated = await processor.processLocalFile(sourceFile, operation);
      let captured = null;
      const workflow = AudioPinImportWorkflow.create({
        callGAS: async function(name, payload) {
          const decodedBytes = atob(payload.audioBase64).length;
          captured = {
            method: name,
            mimeType: payload.audioMimeType,
            decodedBytes: decodedBytes,
            blobBytes: validated.blob.size,
            base64Length: payload.audioBase64.length,
            sourceKind: payload.sourceKind,
            sourceFileName: payload.sourceFileName,
            pin: Object.assign({}, payload.pin)
          };
          return {
            ok: true,
            pin: {
              id: 'real-mp3-pin', title: payload.pin.title, description: '', eventAt: '',
              lat: payload.pin.lat, lng: payload.pin.lng, color: '', icon: '', status: '',
              tags: [], links: [], updatedAt: '2026-07-22T00:00:00.000Z', hasAudio: true
            }
          };
        }
      });
      const draft = workflow.start({
        operation: operation,
        sourceFileName: sourceFile.name,
        editorResult: validated
      });
      workflow.setLocationChoice({ kind: 'map', lat: 35, lng: 135 });
      await workflow.save(draft);
      processor.release();
      return captured;
    };
  </script>`;
  return page('Audio editor harness', bridge + controls + instrumentedEditorContent + launch);
}

function playerPage(shared) {
  const factory = shared ? sharedPlayerFactory : editPlayerFactory;
  const apiName = shared ? 'sharedPinAudioPlayer' : 'pinAudioPlayer';
  const stateSetup = shared
    ? `window.state = { token: 'share-token-browser-test' };`
    : `window.withEditToken = function(payload) { return Object.assign({ editToken: 'edit-token-browser-test' }, payload); };`;
  const body = `
    ${gasMockScript()}
    <main class="surface">
      <div class="button-row" id="pin-openers">
        <button type="button" data-open-pin="pin-a">Aを開く</button>
        <button type="button" data-open-pin="pin-b">Bを開く</button>
        <button type="button" data-open-pin="pin-c">Cを開く</button>
        <button type="button" data-open-pin="pin-d">Dを開く</button>
        <button type="button" data-open-pin="pin-e">Eを開く</button>
      </div>
      <section id="pin-detail" class="detail-surface" aria-label="ピン詳細" hidden>
        <h1 id="detail-title">ピン詳細</h1>
        ${audioPlayerContent}
        <button id="detail-close" type="button">詳細を閉じる</button>
      </section>
    </main>
    <script>
      (function() {
        ${stateSetup}
        const nativeCreateObjectURL = URL.createObjectURL.bind(URL);
        const nativeRevokeObjectURL = URL.revokeObjectURL.bind(URL);
        window.__media = { created: [], revoked: [], plays: 0, pauses: 0, loads: 0 };
        URL.createObjectURL = function(blob) {
          const url = nativeCreateObjectURL(blob);
          window.__media.created.push({ url: url, size: blob.size, type: blob.type });
          return url;
        };
        URL.revokeObjectURL = function(url) {
          window.__media.revoked.push(String(url));
          return nativeRevokeObjectURL(url);
        };
        const audioPrototype = HTMLMediaElement.prototype;
        audioPrototype.play = function() { window.__media.plays += 1; return Promise.resolve(); };
        audioPrototype.pause = function() { window.__media.pauses += 1; };
        audioPrototype.load = function() { window.__media.loads += 1; };

        ${factory}
        const ${apiName} = ${shared ? 'createSharedPinAudioPlayerController' : 'createPinAudioPlayerController'}();
        window.__player = ${apiName};
        window.__currentPin = '';
        let detailTrigger = null;
        document.querySelectorAll('[data-open-pin]').forEach(function(button) {
          button.addEventListener('click', function() {
            detailTrigger = button;
            window.__currentPin = String(button.dataset.openPin || '');
            document.getElementById('pin-detail').hidden = false;
            ${apiName}.open(window.__currentPin);
          });
        });
        document.getElementById('detail-close').addEventListener('click', function() {
          ${apiName}.close();
          document.getElementById('pin-detail').hidden = true;
          if (detailTrigger) detailTrigger.focus();
        });
        window.__dispatchEnded = function() {
          document.getElementById('pin-audio-runtime').dispatchEvent(new Event('ended'));
        };
      })();
    </script>`;
  return page(shared ? 'Shared audio harness' : 'Edit audio player harness', body);
}

function workflowPage() {
  const body = `
    ${gasMockScript()}
    <main class="surface">
      <div class="button-row">
        <button id="workflow-trigger" type="button">音声を追加</button>
        <button id="existing-player" type="button">既存音声を再生</button>
      </div>
      <section id="source-chooser" class="source-chooser" aria-label="音声の取込元" hidden>
        <h1>音声の取込元</h1>
        <input id="workflow-local-file" type="file" accept=".m4a,.mp3,.wav,audio/*">
        <div class="button-row">
          <button type="button" data-source="local">端末から</button>
          <button type="button" data-source="drive">Driveから</button>
        </div>
      </section>
      <section id="workflow-preview" class="workflow-preview" aria-label="音声ピンの確認" hidden>
        <h1 id="preview-heading">音声ピンを確認</h1>
        <label class="field">タイトル<input id="workflow-title"></label>
        <div class="button-row" id="location-choices">
          <button id="choose-map" type="button">地図上の位置</button>
          <button id="choose-unplaced" type="button">未配置</button>
        </div>
        <div class="button-row">
          <button id="workflow-save" type="button">保存</button>
          <button id="workflow-cancel" type="button">キャンセル</button>
        </div>
        <output id="workflow-status" class="status" role="status"></output>
      </section>
    </main>
    <script>
      ${workflowModules}
      (function() {
        let idCounter = 0;
        let activeShell = null;
        let activeProcessor = null;
        let activeDraft = null;
        let requestedMode = 'create-pin';
        let requestedSource = 'local';
        const state = {
          pins: {
            existing: {
              id: 'existing', title: '既存ピン', description: '説明', eventAt: '',
              lat: 35.1, lng: 135.2, color: '#176a8d', icon: 'pin', status: 'active',
              tags: ['音声'], links: [], updatedAt: '2026-07-22T00:00:00.000Z', hasAudio: true
            }
          },
          previewOptions: null,
          saves: [],
          removals: [],
          editorSources: []
        };
        window.__workflowState = state;

        function fakeEditorResult(details) {
          state.editorSources.push({
            sourceKind: String(details && details.sourceKind || ''),
            sourceDriveFileId: String(details && details.sourceDriveFileId || '')
          });
          const bytes = new Uint8Array(1024 * 1024);
          bytes.set([0xff, 0xfb, 0x90, 0x64]);
          const blob = new Blob([bytes], { type: 'audio/mpeg' });
          return {
            blob: blob,
            mimeType: 'audio/mpeg',
            sizeBytes: blob.size,
            sourceDurationSeconds: 10,
            selectionStart: 0,
            selectionEnd: 10,
            durationSeconds: 10,
            numberOfChannels: 1,
            sampleRate: 48000,
            bitrate: 192000
          };
        }

        function createProcessor() {
          return ImportAudioItemProcessor.create({
            openAudioEditor: function(_file, details) { return Promise.resolve(fakeEditorResult(details)); },
            environment: {
              atob: function(value) { return window.atob(value); },
              Uint8Array: Uint8Array,
              Blob: Blob,
              File: File
            }
          });
        }

        const workflow = AudioPinImportWorkflow.create({
          callGAS: withGAS,
          withEditToken: function(payload) {
            return Object.assign({ editToken: 'edit-token-browser-test' }, payload);
          },
          openPreview: function(draft, options) {
            activeDraft = draft;
            state.previewOptions = options;
            document.getElementById('workflow-title').value = String(draft.title || '');
            document.getElementById('location-choices').hidden = !!options.readOnlyFields;
            document.getElementById('workflow-preview').hidden = false;
            document.getElementById('source-chooser').hidden = true;
          },
          onSaved: function(pin) {
            state.pins[String(pin.id)] = Object.assign({}, pin);
            state.saves.push(Object.assign({}, pin));
            if (String(pin.id) === 'existing') {
              document.getElementById('existing-player').hidden = !pin.hasAudio;
            }
          }
        });

        function operationFor(sourceKind, operationMode) {
          const existing = operationMode !== 'create-pin';
          activeShell = MediaImportShell.create({
            createId: function(prefix) { idCounter += 1; return prefix + '-' + idCounter; }
          });
          let operation = activeShell.begin({
            mediaKind: 'audio',
            sourceKind: sourceKind,
            operationMode: operationMode,
            selectionLimit: 1,
            targetPinId: existing ? 'existing' : '',
            expectedUpdatedAt: existing ? state.pins.existing.updatedAt : ''
          });
          if (sourceKind === 'drive') {
            operation = activeShell.setSourceDriveFileId('driveaudio123');
          }
          return operation;
        }

        async function beginLocal(operationMode) {
          requestedMode = operationMode || requestedMode;
          requestedSource = 'local';
          const input = document.getElementById('workflow-local-file');
          const file = input.files && input.files[0];
          if (!file) throw new Error('Local test file is required.');
          const operation = operationFor('local', requestedMode);
          activeProcessor = createProcessor();
          const result = await activeProcessor.processLocalFile(file, operation);
          const target = requestedMode === 'create-pin' ? null : state.pins.existing;
          return workflow.start({
            operation: operation,
            targetPin: target,
            sourceFileName: file.name,
            editorResult: result
          });
        }

        async function beginDrive(operationMode) {
          requestedMode = operationMode || requestedMode;
          requestedSource = 'drive';
          const operation = operationFor('drive', requestedMode);
          activeProcessor = createProcessor();
          const raw = 'RIFFbrowser-drive-audio';
          const response = {
            ok: true,
            file: {
              id: 'driveaudio123', name: 'Driveの録音.wav', mimeType: 'audio/wav',
              sizeBytes: raw.length, base64: btoa(raw), modifiedAt: '2026-07-22T00:00:00.000Z'
            }
          };
          const result = await activeProcessor.processDriveResponse(response, {
            requestFileId: 'driveaudio123', operation: operation
          });
          const target = requestedMode === 'create-pin' ? null : state.pins.existing;
          return workflow.start({
            operation: operation,
            targetPin: target,
            sourceFileName: response.file.name,
            editorResult: result
          });
        }

        function chooseLocation(kind) {
          return workflow.setLocationChoice(kind === 'unplaced'
            ? { kind: 'unplaced' }
            : { kind: 'map', lat: 34.987, lng: 135.765 });
        }

        async function save() {
          if (!activeDraft) throw new Error('No active draft.');
          activeDraft.title = document.getElementById('workflow-title').value;
          try {
            const result = await workflow.save(activeDraft);
            document.getElementById('workflow-status').textContent = '保存しました';
            return result;
          } catch (error) {
            document.getElementById('workflow-status').textContent = String(error.code || error.message);
            throw error;
          }
        }

        async function removeExisting() {
          try {
            const response = await withGAS('removePinAudio', {
              editToken: 'edit-token-browser-test', pinId: 'existing',
              expectedUpdatedAt: state.pins.existing.updatedAt
            });
            if (!response || response.ok !== true || !response.pin) throw new Error('Remove failed.');
            state.pins.existing = Object.assign({}, response.pin);
            state.removals.push(response.pin.id);
            document.getElementById('existing-player').hidden = !response.pin.hasAudio;
            return response;
          } catch (error) {
            document.getElementById('workflow-status').textContent = String(error.message || error);
            throw error;
          }
        }

        window.__workflow = {
          beginLocal: beginLocal,
          beginDrive: beginDrive,
          chooseLocation: chooseLocation,
          save: save,
          retry: function() { return workflow.retry(); },
          cancel: function() { return workflow.cancel(); },
          removeExisting: removeExisting,
          setMode: function(mode) { requestedMode = String(mode); },
          getRequestedSource: function() { return requestedSource; }
        };

        document.getElementById('workflow-trigger').addEventListener('click', function() {
          document.getElementById('source-chooser').hidden = false;
        });
        document.querySelector('[data-source="local"]').addEventListener('click', function() {
          beginLocal(requestedMode).catch(function(error) {
            document.getElementById('workflow-status').textContent = String(error.code || error.message);
          });
        });
        document.querySelector('[data-source="drive"]').addEventListener('click', function() {
          beginDrive(requestedMode).catch(function(error) {
            document.getElementById('workflow-status').textContent = String(error.code || error.message);
          });
        });
        document.getElementById('choose-map').addEventListener('click', function() { chooseLocation('map'); });
        document.getElementById('choose-unplaced').addEventListener('click', function() { chooseLocation('unplaced'); });
        document.getElementById('workflow-save').addEventListener('click', function() { save().catch(function() {}); });
        document.getElementById('workflow-cancel').addEventListener('click', function() {
          workflow.cancel();
          document.getElementById('workflow-preview').hidden = true;
          document.getElementById('workflow-trigger').focus();
        });
      })();
    </script>`;
  return page('Audio workflow harness', body);
}

const server = http.createServer((request, response) => {
  const requestUrl = new URL(request.url, `http://${host}:${port}`);
  if (requestUrl.pathname === '/__health') {
    response.writeHead(200, { 'content-type': 'text/plain; charset=utf-8', 'cache-control': 'no-store' });
    response.end('ok');
    return;
  }
  if (requestUrl.pathname === '/audio-vendor.js') {
    response.writeHead(200, { 'content-type': 'text/javascript; charset=utf-8', 'cache-control': 'no-store' });
    response.end(audioVendorSource);
    return;
  }
  let html;
  if (requestUrl.pathname === '/editor') html = editorPage();
  else if (requestUrl.pathname === '/player') html = playerPage(false);
  else if (requestUrl.pathname === '/shared') html = playerPage(true);
  else if (requestUrl.pathname === '/workflow') html = workflowPage();
  else if (requestUrl.pathname === '/production-edit') html = productionEditPage();
  else if (requestUrl.pathname === '/production-shared') html = productionSharedPage();
  else html = page('Not found', '<h1>Not found</h1>');
  response.writeHead(requestUrl.pathname === '/' ? 404 : 200, {
    'content-type': 'text/html; charset=utf-8',
    'cache-control': 'no-store'
  });
  response.end(html);
});

server.listen(port, host);

function stop() {
  server.close(function() { process.exit(0); });
  setTimeout(function() { process.exit(0); }, 1000).unref();
}

process.once('SIGINT', stop);
process.once('SIGTERM', stop);
process.once('SIGHUP', stop);

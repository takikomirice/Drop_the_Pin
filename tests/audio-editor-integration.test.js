const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const EDITOR_START = '<template data-dtp-audio-editor-boundary="start"></template>';
const EDITOR_END = '<template data-dtp-audio-editor-boundary="end"></template>';
const VENDOR_START = 'AUDIO_VENDOR_BUNDLE_START';
const VENDOR_END = 'AUDIO_VENDOR_BUNDLE_END';
const fixtureVendorSource = 'globalThis.Mediabunny={};globalThis.MediabunnyMp3Encoder={};';

function editorHtml() {
  assert.equal(countOccurrences(indexHtml, EDITOR_START), 1, 'editor start boundary');
  assert.equal(countOccurrences(indexHtml, EDITOR_END), 1, 'editor end boundary');
  const contentStart = indexHtml.indexOf(EDITOR_START) + EDITOR_START.length;
  const contentEnd = indexHtml.indexOf(EDITOR_END);
  assert.ok(contentEnd > contentStart, 'editor boundaries must be ordered');
  return indexHtml.slice(contentStart, contentEnd).replace(/^\r?\n/, '');
}

function countOccurrences(text, needle) {
  return text.split(needle).length - 1;
}

function rawIndexWithVendor(source = fixtureVendorSource) {
  return `${VENDOR_START}\n<script>\n${source}\n</script>\n${VENDOR_END}\n<!DOCTYPE html>\n<html></html>\n`;
}

function createServerHarness(options = {}) {
  const audit = { fileReads: [], tokenReads: 0 };
  const context = {
    console,
    CacheService: {
      getScriptCache() {
        return {
          get(key) {
            audit.tokenReads += 1;
            return key === 'EDIT_TOKEN_valid-token' ? '1' : null;
          }
        };
      }
    },
    HtmlService: {
      createHtmlOutputFromFile(name) {
        audit.fileReads.push(name);
        return {
          getContent: () => options.indexContent === undefined
            ? rawIndexWithVendor()
            : options.indexContent
        };
      }
    }
  };
  vm.runInNewContext(`${codeJs}\nglobalThis.__audioEditorIntegrationApi = {\n  locateAudioVendorBundleInIndex_: typeof locateAudioVendorBundleInIndex_ === 'function' ? locateAudioVendorBundleInIndex_ : null,\n  stripAudioVendorBundleFromIndex_: typeof stripAudioVendorBundleFromIndex_ === 'function' ? stripAudioVendorBundleFromIndex_ : null,\n  getAudioVendorBundle: typeof getAudioVendorBundle === 'function' ? getAudioVendorBundle : null\n};`, context);
  return { api: context.__audioEditorIntegrationApi, audit };
}

test('editor block is inlined only inside the edit-token index condition', () => {
  assert.match(
    indexHtml,
    /<\?\s*if\s*\(editToken\)\s*\{\s*\?>[\s\S]*?data-dtp-audio-editor-boundary="start"[\s\S]*?data-dtp-audio-editor-boundary="end"[\s\S]*?<\?\s*\}\s*\?>/
  );
  assert.equal(countOccurrences(indexHtml, EDITOR_START), 1);
  assert.equal(countOccurrences(indexHtml, EDITOR_END), 1);
  assert.equal(sharedHtml.includes('data-dtp-audio-editor-boundary'), false);
  assert.equal(sharedHtml.includes(VENDOR_START), false);
});

test('server removes the obsolete fragment include API', () => {
  assert.doesNotMatch(codeJs, /HTML_INCLUDE_ALLOWLIST_/);
});

test('audio vendor authentication happens before raw index is read', () => {
  const harness = createServerHarness({ indexContent: 'malformed private source' });
  assert.equal(typeof harness.api.getAudioVendorBundle, 'function');
  assert.match(codeJs, /function\s+getAudioVendorBundle\s*\(payload\)\s*\{\s*assertEditToken_\(payload\);/);
  assert.throws(() => harness.api.getAudioVendorBundle({}), /編集/);
  assert.throws(
    () => harness.api.getAudioVendorBundle({ __editToken: 'invalid-token' }),
    /編集/
  );
  assert.deepEqual(harness.audit.fileReads, []);
  assert.equal(harness.audit.tokenReads, 1);
});

test('authenticated audio vendor API reads index once and returns only source', () => {
  const harness = createServerHarness();
  assert.equal(typeof harness.api.getAudioVendorBundle, 'function');
  const result = harness.api.getAudioVendorBundle({ __editToken: 'valid-token' });

  assert.equal(result.version, '1.50.8');
  assert.equal(result.source, fixtureVendorSource);
  assert.equal(result.source.includes('AUDIO_VENDOR_BUNDLE_START'), false);
  assert.equal(result.source.includes('AUDIO_VENDOR_BUNDLE_END'), false);
  assert.equal(result.source.includes('<script'), false);
  assert.equal(result.source.includes('</script>'), false);
  assert.equal(countOccurrences(result.source, 'globalThis.Mediabunny='), 1);
  assert.equal(countOccurrences(result.source, 'globalThis.MediabunnyMp3Encoder='), 1);
  assert.deepEqual(harness.audit.fileReads, ['index']);
});

test('audio vendor extraction fails closed for malformed full-index regions', () => {
  const fixtures = [
    `unexpected prefix\n${rawIndexWithVendor()}`,
    rawIndexWithVendor().replace('<script>', '<script type="text/javascript">'),
    rawIndexWithVendor().replace(VENDOR_START, `${VENDOR_START}\n${VENDOR_START}`),
    rawIndexWithVendor().replace(VENDOR_END, `${VENDOR_END}\n${VENDOR_END}`),
    `${VENDOR_END}\n<script>safe</script>\n${VENDOR_START}\n<!DOCTYPE html>`,
    rawIndexWithVendor(''),
    rawIndexWithVendor('globalThis.Mediabunny={};console.log("</script>");globalThis.MediabunnyMp3Encoder={};'),
    rawIndexWithVendor(`${fixtureVendorSource}globalThis.Mediabunny={};`),
    `${VENDOR_START}\n<script>\n${fixtureVendorSource}\n</script>\n${VENDOR_END}\nnot a document`
  ];

  fixtures.forEach((indexContent) => {
    const harness = createServerHarness({ indexContent });
    assert.equal(typeof harness.api.getAudioVendorBundle, 'function');
    assert.throws(
      () => harness.api.getAudioVendorBundle({ __editToken: 'valid-token' }),
      /vendor|bundle|source/i
    );
  });
});

test('editor fragment preserves the input, duration, selection, output and channel contracts', () => {
  const html = editorHtml();
  assert.equal(html.includes('<main'), false, 'included fragment must not add a second main landmark');
  for (const format of ['M4A', 'MP3', 'WAV']) assert.match(html, new RegExp(format));
  for (const contract of [
    /WARN_SIZE\s*=\s*50\s*\*\s*1024\s*\*\s*1024/,
    /MAX_SIZE\s*=\s*200\s*\*\s*1024\s*\*\s*1024/,
    /WARN_DURATION\s*=\s*300/,
    /MAX_DURATION\s*=\s*600/,
    /MIN_SELECTION\s*=\s*0\.5/,
    /INITIAL_SELECTION_DURATION\s*=\s*30/,
    /MAX_SELECTION_DURATION\s*=\s*120/,
    /MAX_OUTPUT_SIZE\s*=\s*4\s*\*\s*1024\s*\*\s*1024/,
    /sampleRate:\s*48000/,
    /bitrate:\s*192000/,
    /bitrateMode:\s*['"]constant['"]/,
    /numberOfChannels\s*>\s*2/,
    /3チャンネル以上/,
    /モノラルまたはステレオ/
  ]) assert.match(html, contract);
});

test('editor hands a valid MP3 to its result handler without a download action', () => {
  const html = editorHtml();
  assert.match(html, /この音声を使用/);
  assert.match(html, /function setResultHandler\s*\(/);
  assert.match(html, /resultHandler\s*\(copyResultMetadata\(resultMetadata\)\)/);
  assert.match(html, /resultBlob\s+instanceof\s+Blob|resultMetadata\.blob\s+instanceof\s+Blob/);
  assert.match(html, /audio\/mpeg/);
  assert.match(html, /data-hae-result-audio[^>]*controlsList="nodownload"/);
  assert.doesNotMatch(
    html,
    /controlsList="(?:[^"]*\s)?(?:nofullscreen|noplaybackrate|noremoteplayback)/
  );
  assert.equal(html.includes('data-hae-download'), false);
  assert.equal(html.includes('downloadResult'), false);
  assert.doesNotMatch(html, /\.download\s*=/);
});

test('editor invalidates stale work and releases cancellable and object URL resources', () => {
  const html = editorHtml();
  for (const contract of [
    /function cancelActiveOutput\s*\(/,
    /Promise\.resolve\(output\.cancel\(\)\)\.catch/,
    /function invalidateResult\s*\(\)\s*\{[\s\S]*?encodeGeneration\s*\+=\s*1[\s\S]*?cancelActiveOutput\(\)[\s\S]*?clearResultInternal\(\)/,
    /revokeObjectURL\(resultObjectUrl\)/,
    /function resetForLoad\s*\([^)]*\)\s*\{[\s\S]*?loadGeneration\s*\+=\s*1[\s\S]*?invalidateResult\(\)/,
    /function setSelection\s*\([^)]*\)\s*\{[\s\S]*?invalidateResult\(\)/,
    /generation\s*!==\s*encodeGeneration\s*\|\|\s*inputGeneration\s*!==\s*loadGeneration/,
    /function destroy\s*\(\)\s*\{[\s\S]*?invalidateResult\(\)/
  ]) assert.match(html, contract);
});

test('client vendor loader is singleton, authenticated, validated, retryable and nonblocking', () => {
  assert.match(indexHtml, /let audioVendorBundlePromise\s*=\s*null/);
  assert.match(indexHtml, /function loadAudioVendorBundle\s*\(/);
  assert.match(indexHtml, /if\s*\(audioVendorBundlePromise\)\s*return audioVendorBundlePromise/);
  assert.match(indexHtml, /getAudioVendorBundle\s*\(\s*\{\s*__editToken:\s*window\.__EDIT_TOKEN__/);
  assert.match(indexHtml, /document\.createElement\(['"]script['"]\)/);
  assert.match(indexHtml, /script\.(?:text|textContent)\s*=\s*(?:bundle\.)?source/);
  assert.match(indexHtml, /window\.Mediabunny/);
  assert.match(indexHtml, /window\.MediabunnyMp3Encoder/);
  for (const api of ['canEncodeAudio', 'BufferTarget', 'Output', 'Mp3OutputFormat', 'AudioBufferSource', 'registerMp3Encoder']) {
    assert.match(indexHtml, new RegExp(api));
  }
  assert.match(indexHtml, /\.catch\(function\s*\([^)]*\)\s*\{[\s\S]*?audioVendorBundlePromise\s*=\s*null/);

  const ready = indexHtml.indexOf('state.initializing = false;');
  const warmup = indexHtml.indexOf('warmAudioVendorBundle();', ready);
  assert.ok(ready >= 0 && warmup > ready, 'vendor warmup must start after the main UI becomes ready');
  assert.doesNotMatch(indexHtml.slice(ready, warmup + 40), /await\s+(?:warmAudioVendorBundle|loadAudioVendorBundle)/);
  assert.match(editorHtml(), /function init\s*\(\)\s*\{[\s\S]*?loadAudioVendorBundle\(\)\.catch/);
});

test('client-facing audio editor copy does not expose private implementation details', () => {
  const html = editorHtml();
  assert.doesNotMatch(html, /Drive\s*ID|folder\s*ID|__editToken|AUDIO_VENDOR_BUNDLE|globalThis\.Mediabunny/);
  assert.doesNotMatch(html, /error\.message|String\(error\)/);
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const PLAYER_START = '<template data-dtp-audio-player-boundary="start"></template>';
const PLAYER_END = '<template data-dtp-audio-player-boundary="end"></template>';
const EDITOR_START = '<template data-dtp-audio-editor-boundary="start"></template>';
const EDITOR_END = '<template data-dtp-audio-editor-boundary="end"></template>';
const VENDOR_START = 'AUDIO_VENDOR_BUNDLE_START';
const VENDOR_END = 'AUDIO_VENDOR_BUNDLE_END';

function countOccurrences(source, marker) {
  return source.split(marker).length - 1;
}

function extractMarkedBlock(source, startMarker, endMarker) {
  if (countOccurrences(source, startMarker) !== 1 || countOccurrences(source, endMarker) !== 1) {
    throw new Error('Marker boundaries must each appear exactly once.');
  }
  const startIndex = source.indexOf(startMarker);
  const endIndex = source.indexOf(endMarker);
  if (endIndex <= startIndex) throw new Error('Marker boundaries are reversed.');
  return source.slice(startIndex, endIndex + endMarker.length);
}

test('the Apps Script project has exactly the two root HTML files', () => {
  const htmlFiles = fs.readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === '.html')
    .map((entry) => entry.name)
    .sort();
  assert.deepEqual(htmlFiles, ['index.html', 'shared.html']);
});

test('the Apps Script project has exactly one deployable root JavaScript file', () => {
  const claspIgnoredPaths = new Set(
    fs.readFileSync(path.join(root, '.claspignore'), 'utf8')
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean)
  );
  const jsFiles = fs.readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isFile() && path.extname(entry.name).toLowerCase() === '.js')
    .filter((entry) => !claspIgnoredPaths.has(entry.name))
    .map((entry) => entry.name)
    .sort();
  assert.deepEqual(jsFiles, ['Code.js']);
});

test('index and shared contain one byte-identical inline player block', () => {
  const indexPlayer = extractMarkedBlock(indexHtml, PLAYER_START, PLAYER_END);
  const sharedPlayer = extractMarkedBlock(sharedHtml, PLAYER_START, PLAYER_END);
  assert.equal(indexPlayer, sharedPlayer);
});

test('each page contains exactly one pin audio runtime', () => {
  assert.equal(countOccurrences(indexHtml, 'id="pin-audio-runtime"'), 1);
  assert.equal(countOccurrences(sharedHtml, 'id="pin-audio-runtime"'), 1);
});

test('the editor block appears once only in index and remains edit-token gated', () => {
  extractMarkedBlock(indexHtml, EDITOR_START, EDITOR_END);
  assert.equal(countOccurrences(sharedHtml, EDITOR_START), 0);
  assert.equal(countOccurrences(sharedHtml, EDITOR_END), 0);
  const conditionStart = '<? if (editToken) { ?>';
  const conditionEnd = '<? } ?>';
  const conditionStartIndex = indexHtml.lastIndexOf(conditionStart, indexHtml.indexOf(EDITOR_START));
  const conditionEndIndex = indexHtml.indexOf(conditionEnd, indexHtml.indexOf(EDITOR_END) + EDITOR_END.length);
  assert.ok(conditionStartIndex >= 0 && conditionEndIndex > conditionStartIndex);
  assert.equal(
    indexHtml.slice(conditionStartIndex + conditionStart.length, indexHtml.indexOf(EDITOR_START)).trim(),
    ''
  );
  assert.equal(
    indexHtml.slice(indexHtml.indexOf(EDITOR_END) + EDITOR_END.length, conditionEndIndex).trim(),
    ''
  );
});

test('vendor sentinels occur once at the raw index prefix and nowhere in shared', () => {
  assert.equal(countOccurrences(indexHtml, VENDOR_START), 1);
  assert.equal(countOccurrences(indexHtml, VENDOR_END), 1);
  assert.equal(countOccurrences(sharedHtml, VENDOR_START), 0);
  assert.equal(countOccurrences(sharedHtml, VENDOR_END), 0);
  assert.equal(countOccurrences(sharedHtml, EDITOR_START), 0);
  assert.equal(countOccurrences(sharedHtml, EDITOR_END), 0);

  const startIndex = indexHtml.indexOf(VENDOR_START);
  const endIndex = indexHtml.indexOf(VENDOR_END);
  const doctypeIndex = indexHtml.indexOf('<!DOCTYPE html>');
  assert.equal(indexHtml.slice(0, startIndex).trim(), '');
  assert.ok(startIndex >= 0 && endIndex > startIndex && doctypeIndex > endIndex);
});

test('marker extraction rejects missing, duplicate, and reversed boundaries', () => {
  assert.throws(() => extractMarkedBlock('start only', 'start', 'end'), /exactly once/);
  assert.throws(() => extractMarkedBlock('start start end', 'start', 'end'), /exactly once/);
  assert.throws(() => extractMarkedBlock('start end end', 'start', 'end'), /exactly once/);
  assert.throws(() => extractMarkedBlock('end then start', 'start', 'end'), /reversed/);
});

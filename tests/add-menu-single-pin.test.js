const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function tagFor(id) {
  const match = indexHtml.match(new RegExp(`<[^>]+id=["']${id}["'][^>]*>`, 'i'));
  assert.ok(match, `${id} must exist`);
  return match[0];
}

test('add menu preserves photo entries and adds one local and one Drive single-audio path', () => {
  const menu = indexHtml.match(/<div class="add-menu-actions"[^>]*>([\s\S]*?)<\/div>\s*<div id="multi-photo-conflict-message"/);
  assert.ok(menu);
  const actions = menu[1].match(/<button\b[^>]*data-add-action=/g) || [];
  assert.equal(actions.length, 5);
  for (const label of [
    'ピンを追加', '端末から写真を取込', 'Driveから写真を取込',
    '端末から音声を取込', 'Driveから音声を取込'
  ]) assert.ok(menu[1].includes(label), label);
});

test('local audio input is single-select and limited to M4A, MP3, and WAV', () => {
  const input = tagFor('audio-import-file-input');
  assert.doesNotMatch(input, /\bmultiple\b/i);
  assert.match(input, /accept=["']\.m4a,\.mp3,\.wav,audio\/mp4,audio\/mpeg,audio\/wav["']/i);
  assert.match(indexHtml, /getElementById\('audio-import-file-input'\)\.addEventListener\('change',\s*handleLocalAudioImportSelected\)/);
});

test('audio Drive path lists audio only, reads one selected item, and photo remains capped at twenty', () => {
  assert.match(indexHtml, /listDriveMediaInbox[\s\S]{0,500}mediaKind:\s*'audio'/);
  assert.match(indexHtml, /readDriveAudioImportFile/);
  assert.match(indexHtml, /selectionLimit:\s*1/);
  assert.match(indexHtml, /getSelectionLimit:[\s\S]{0,180}ImportJobCore\.MAX_ITEMS/);
  assert.match(indexHtml, /const MAX_ITEMS\s*=\s*20|MAX_ITEMS:\s*20/);
  assert.match(tagFor('multi-photo-file-input'), /\bmultiple\b/i);
});

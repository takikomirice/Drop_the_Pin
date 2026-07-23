const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const {
  makeHarness, audioPayload, validMp3Bytes, pinRow, MAP_HEADERS, RECEIPT_HEADERS
} = require('./audio-storage-harness');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const PLAYER_START = '<template data-dtp-audio-player-boundary="start"></template>';
const PLAYER_END = '<template data-dtp-audio-player-boundary="end"></template>';

function plain(value) { return JSON.parse(JSON.stringify(value)); }
function playerHtml(source) {
  assert.equal(count(source, PLAYER_START), 1, 'player start boundary');
  assert.equal(count(source, PLAYER_END), 1, 'player end boundary');
  const contentStart = source.indexOf(PLAYER_START) + PLAYER_START.length;
  const contentEnd = source.indexOf(PLAYER_END);
  assert.ok(contentEnd > contentStart, 'player boundaries must be ordered');
  return source.slice(contentStart, contentEnd).replace(/^\r?\n/, '');
}
function count(text, needle) { return text.split(needle).length - 1; }
function sheetsWithPin(row) {
  return {
    map_info: [MAP_HEADERS, row],
    import_receipts: [RECEIPT_HEADERS],
    config: [
      ['設定項目', '値', '説明'],
      ['IMAGE_DRIVE_URL', 'https://drive.google.com/drive/folders/root_media_1234567890', ''],
      ['EDIT_KEY', 'key', '']
    ]
  };
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
    if (character === '}' && --depth === 0) return source.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

test('common player block is byte-identical and appears once in edit and shared detail pages', () => {
  const indexPlayer = playerHtml(indexHtml);
  const sharedPlayer = playerHtml(sharedHtml);
  assert.equal(indexPlayer, sharedPlayer);
  assert.equal(count(indexHtml, '<audio id="pin-audio-runtime"'), 1);
  assert.equal(count(sharedHtml, '<audio id="pin-audio-runtime"'), 1);
  assert.match(indexPlayer, /<audio\s+id="pin-audio-runtime"[^>]*\bcontrols\b/);
  assert.match(indexPlayer, /<audio\s+id="pin-audio-runtime"[^>]*\bcontrolsList="nodownload"/);
  assert.equal(indexPlayer.includes('pin-audio-player-toggle'), false);
  assert.equal(indexPlayer.includes('data-dtp-audio-editor-boundary'), false);
});

test('authenticated edit playback returns only bounded MP3 bytes and no Drive metadata', () => {
  const harness = makeHarness();
  const saved = plain(harness.api.saveImportAudioItem(audioPayload({
    audioBase64: Buffer.from(validMp3Bytes(1024)).toString('base64')
  })));
  assert.equal(saved.ok, true);
  assert.equal(typeof harness.api.getPinAudioData, 'function');

  const result = plain(harness.api.getPinAudioData({
    __editToken: 'valid-token', pinId: saved.pin.id
  }));
  assert.deepEqual(Object.keys(result).sort(), ['base64', 'byteLength', 'mimeType', 'ok']);
  assert.equal(result.ok, true);
  assert.equal(result.mimeType, 'audio/mpeg');
  assert.equal(result.byteLength, 1024);
  assert.equal(Buffer.from(result.base64, 'base64').length, result.byteLength);

  const privateValues = [
    harness.receipt()[RECEIPT_HEADERS.indexOf('fileId')],
    harness.audioFolder.id,
    'voice.wav'
  ];
  privateValues.forEach((value) => assert.equal(JSON.stringify(result).includes(value), false));
  for (const field of ['audioId', 'fileId', 'name', 'folder', 'url', 'sourceDriveFileId']) {
    assert.equal(Object.hasOwn(result, field), false);
  }
});

test('edit playback authenticates and rejects invalid pins or malformed managed audio', () => {
  const denied = makeHarness({ validToken: false });
  assert.equal(typeof denied.api.getPinAudioData, 'function');
  assert.throws(
    () => denied.api.getPinAudioData({ __editToken: 'invalid', pinId: 'pin-existing-0001' }),
    /編集/
  );

  const invalidId = makeHarness();
  assert.throws(
    () => invalidId.api.getPinAudioData({ __editToken: 'valid-token', pinId: '=formula' }),
    /audio|pin|invalid|unavailable/i
  );

  const shortId = 'managed_short_audio_1';
  const short = makeHarness({ sheets: sheetsWithPin(pinRow({ audioId: shortId })) });
  short.addFile(shortId, 'short.mp3', 'audio/mpeg', [0x49, 0x44, 0x33], [short.audioFolder.id]);
  assert.throws(
    () => short.api.getPinAudioData({ __editToken: 'valid-token', pinId: 'pin-existing-0001' }),
    /audio|unavailable|invalid/i
  );

  const wrongMimeId = 'managed_wrong_mime_1';
  const wrongMime = makeHarness({ sheets: sheetsWithPin(pinRow({ audioId: wrongMimeId })) });
  wrongMime.addFile(
    wrongMimeId, 'wrong.wav', 'audio/wav', new Uint8Array(1024 * 1024),
    [wrongMime.audioFolder.id]
  );
  assert.throws(
    () => wrongMime.api.getPinAudioData({ __editToken: 'valid-token', pinId: 'pin-existing-0001' }),
    /audio|unavailable|invalid/i
  );
});

test('edit detail alone starts playback fetch and always closes or destroys the controller', () => {
  assert.match(indexHtml, /AudioPinPlayer\.create\s*\(/);
  assert.match(indexHtml, /getPinAudioData[\s\S]*?withEditToken\s*\(\s*\{\s*pinId:/);
  const openSource = functionSource(indexHtml, 'openPinDetail');
  const closeSource = functionSource(indexHtml, 'closePinDetail');
  assert.match(
    openSource,
    /hasEditToken\s*&&\s*pin\.hasAudio[\s\S]*?pinAudioPlayer\.open\s*\(\s*pin\.id\s*\)/
  );
  assert.doesNotMatch(openSource, /canEdit\(\)\s*&&\s*pin\.hasAudio/);
  assert.ok(
    openSource.indexOf("openOverlay('pin-detail-overlay')") < openSource.indexOf('pinAudioPlayer.open'),
    'audio fetch starts only after detail is opened'
  );
  assert.match(openSource, /pinAudioPlayer\.close\s*\(\)/);
  assert.match(closeSource, /pinAudioPlayer\.close\s*\(\)/);
  assert.match(indexHtml, /onSaved:\s*function\s*\(pin\)[\s\S]*?pinAudioPlayer\.invalidate\s*\(\s*pin\.id\s*\)[\s\S]*?upsertImportedPin\s*\(\s*pin\s*\)/);
  assert.match(indexHtml, /removeAudioFromPinDetail[\s\S]*?pinAudioPlayer\.invalidate\s*\(\s*snapshot\.pinId\s*\)/);
  assert.match(indexHtml, /pagehide[\s\S]*?pinAudioPlayer\.destroy\s*\(\)/);
  assert.equal(sharedHtml.includes("withGAS('getPinAudioData'"), false, 'shared playback must not use edit audio authorization');
  assert.match(sharedHtml, /withGAS\('getSharedPinAudioData'/);
});

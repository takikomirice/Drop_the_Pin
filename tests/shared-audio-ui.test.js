const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');
const PLAYER_START = '<template data-dtp-audio-player-boundary="start"></template>';
const PLAYER_END = '<template data-dtp-audio-player-boundary="end"></template>';

function functionSource(name) {
  const start = sharedHtml.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const bodyStart = sharedHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = bodyStart; index < sharedHtml.length; index += 1) {
    const character = sharedHtml[index];
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
    if (character === '}' && --depth === 0) return sharedHtml.slice(start, index + 1);
  }
  assert.fail(`Could not parse ${name}`);
}

test('shared detail wires the common minimal player with share token only', () => {
  assert.equal(sharedHtml.split(PLAYER_START).length - 1, 1);
  assert.equal(sharedHtml.split(PLAYER_END).length - 1, 1);
  assert.equal((sharedHtml.match(/id="pin-audio-runtime"/g) || []).length, 1);
  assert.match(sharedHtml, /AudioPinPlayer\.create\s*\(/);
  assert.match(
    sharedHtml,
    /getSharedPinAudioData[\s\S]*?shareToken:\s*state\.token[\s\S]*?pinId:/
  );
  assert.equal(sharedHtml.includes("withGAS('getPinAudioData'"), false);
  assert.equal(sharedHtml.includes('withEditToken'), false);

  const openSource = functionSource('openSharedDetail');
  const closeSource = functionSource('closeSharedDetail');
  assert.match(openSource, /pin\.hasAudio[\s\S]*?sharedPinAudioPlayer\.open\s*\(\s*pin\.id\s*\)/);
  assert.ok(
    openSource.indexOf("openSharedSurface('shared-detail-overlay')")
      < openSource.indexOf('sharedPinAudioPlayer.open'),
    'audio fetch must start after the shared detail is opened'
  );
  assert.match(openSource, /sharedPinAudioPlayer\.close\s*\(\)/);
  assert.match(closeSource, /sharedPinAudioPlayer\.close\s*\(\)/);
  assert.match(sharedHtml, /pagehide[\s\S]*?sharedPinAudioPlayer\.destroy\s*\(\)/);
});

test('shared page never includes editor or vendor code and initial load never fetches audio', () => {
  assert.equal(sharedHtml.includes('data-dtp-audio-editor-boundary'), false);
  assert.equal(sharedHtml.includes('AUDIO_VENDOR_BUNDLE_START'), false);
  assert.equal(sharedHtml.includes('AUDIO_VENDOR_BUNDLE_END'), false);
  assert.equal(sharedHtml.includes('getAudioVendorBundle'), false);
  assert.equal(/\bMediabunny\b|\bMediaBunny\b/.test(sharedHtml), false);
  assert.equal(sharedHtml.includes('HemisphereAudioEditor'), false);
  assert.equal(functionSource('initializeSharedView').includes('getSharedPinAudioData'), false);
  assert.equal(functionSource('renderSharedPins').includes('getSharedPinAudioData'), false);
  assert.equal(functionSource('renderSharedMap').includes('getSharedPinAudioData'), false);
});

test('shared player uses native controls with only download hidden and keeps retry wiring', () => {
  assert.match(sharedHtml, /pin-audio-runtime[^>]*controls[^>]*controlsList="nodownload"/);
  assert.equal(sharedHtml.includes('pin-audio-player-toggle'), false);
  assert.match(sharedHtml, /pin-audio-player-retry/);
  assert.match(sharedHtml, /pin-audio-player-retry[\s\S]*?addEventListener\s*\(\s*['"]click['"]/);
  const playerMarkup = sharedHtml.slice(
    sharedHtml.indexOf('<section id="pin-audio-player"'),
    sharedHtml.indexOf('</section>', sharedHtml.indexOf('<section id="pin-audio-player"'))
  );
  for (const forbidden of ['nofullscreen', 'noplaybackrate', 'noremoteplayback']) {
    assert.equal(playerMarkup.includes(forbidden), false);
  }
});

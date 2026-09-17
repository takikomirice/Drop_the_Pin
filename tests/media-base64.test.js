const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const vm = require('node:vm');
const code = fs.readFileSync(require('node:path').join(__dirname, '../Code.js'), 'utf8');

test('media encoding preserves signed bytes, padding and chunk boundaries without GAS service calls', () => {
  const context = vm.createContext({ Utilities: { base64Encode() { throw new Error('service conversion used'); } } });
  vm.runInContext(code, context);
  for (const size of [1, 2, 3, 256, 24575, 24576, 24577, 1400000]) {
    const bytes = Array.from({ length: size }, (_, i) => (i % 256) - 128);
    const result = context.pinPhotoDataFromBlob_({ getBytes: () => bytes, getContentType: () => 'image/jpeg' });
    assert.equal(result.base64, Buffer.from(bytes).toString('base64'));
    assert.equal(result.byteLength, size);
  }
  const bytes = Array.from({ length: 4096 }, (_, i) => (i % 256) - 128);
  const audio = context.pinAudioDataFromBlob_({ getBytes: () => bytes, getContentType: () => 'audio/mpeg' });
  assert.equal(audio.base64, Buffer.from(bytes).toString('base64'));
});

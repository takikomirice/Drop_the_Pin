const assert = require('node:assert/strict');
const test = require('node:test');
const {
  makeHarness, pinRow, MAP_HEADERS, RECEIPT_HEADERS
} = require('./audio-storage-harness');

const MAX_PHOTO_BYTES = 15 * 1024 * 1024;

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

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

function harnessWithPhoto(options = {}) {
  const fileId = options.fileId || 'managed_photo_0001';
  const row = pinRow({
    fileId,
    imageUrl: `https://drive.google.com/thumbnail?id=${fileId}`
  });
  const harness = makeHarness({
    validToken: options.validToken,
    sheets: sheetsWithPin(row)
  });
  const file = harness.addFile(
    fileId,
    options.name || 'photo.jpg',
    options.mimeType || 'image/jpeg',
    options.bytes || [0xff, 0xd8, 0xff, 0xd9],
    [harness.photosFolder.id],
    { trashed: options.trashed === true }
  );
  if (options.sizeOverride != null) file.getSize = () => options.sizeOverride;
  return { harness, file, fileId };
}

test('authenticated photo read resolves the server-side pin file and returns only bounded bytes', () => {
  const { harness, fileId } = harnessWithPhoto();
  const attackerId = 'attacker_controlled_photo';
  harness.addFile(
    attackerId, 'attacker.png', 'image/png', [1, 2, 3, 4],
    [harness.photosFolder.id]
  );

  assert.equal(typeof harness.api.getPinPhotoData, 'function');
  const result = plain(harness.api.getPinPhotoData({
    __editToken: 'valid-token',
    pinId: 'pin-existing-0001',
    fileId: attackerId
  }));

  assert.deepEqual(Object.keys(result).sort(), ['base64', 'byteLength', 'mimeType', 'ok']);
  assert.equal(result.ok, true);
  assert.equal(result.mimeType, 'image/jpeg');
  assert.equal(result.byteLength, 4);
  assert.deepEqual(Array.from(Buffer.from(result.base64, 'base64')), [0xff, 0xd8, 0xff, 0xd9]);

  const serialized = JSON.stringify(result);
  for (const privateValue of [
    fileId, attackerId, 'photo.jpg', 'https://drive.google.com', 'valid-token'
  ]) {
    assert.equal(serialized.includes(privateValue), false);
  }
  for (const field of ['fileId', 'name', 'folder', 'url', 'imageUrl']) {
    assert.equal(Object.hasOwn(result, field), false);
  }
});

test('photo read requires edit authentication and rejects invalid or missing pins safely', () => {
  const denied = harnessWithPhoto({ validToken: false }).harness;
  assert.throws(
    () => denied.api.getPinPhotoData({
      __editToken: 'invalid',
      pinId: 'pin-existing-0001'
    }),
    /編集/
  );

  const valid = harnessWithPhoto().harness;
  for (const pinId of ['=formula', 'pin-not-found']) {
    assert.throws(
      () => valid.api.getPinPhotoData({ __editToken: 'valid-token', pinId }),
      /photo|pin|unavailable/i
    );
  }

  const noPhoto = makeHarness({
    sheets: sheetsWithPin(pinRow({ fileId: '', imageUrl: '' }))
  });
  assert.throws(
    () => noPhoto.api.getPinPhotoData({
      __editToken: 'valid-token',
      pinId: 'pin-existing-0001'
    }),
    /photo|unavailable/i
  );
});

test('photo read accepts supported image MIME types', () => {
  for (const mimeType of ['image/jpeg', 'image/png', 'image/gif', 'image/webp']) {
    const { harness } = harnessWithPhoto({ mimeType });
    const result = plain(harness.api.getPinPhotoData({
      __editToken: 'valid-token',
      pinId: 'pin-existing-0001'
    }));
    assert.equal(result.mimeType, mimeType);
    assert.equal(result.byteLength, 4);
  }
});

test('photo read rejects missing trashed non-image empty and oversized Drive files', () => {
  const missing = makeHarness({
    sheets: sheetsWithPin(pinRow({
      fileId: 'missing_photo_0001',
      imageUrl: 'https://drive.google.com/thumbnail?id=missing_photo_0001'
    }))
  });
  assert.throws(
    () => missing.api.getPinPhotoData({
      __editToken: 'valid-token',
      pinId: 'pin-existing-0001'
    }),
    /photo|unavailable/i
  );

  const cases = [
    harnessWithPhoto({ trashed: true }).harness,
    harnessWithPhoto({ mimeType: 'text/plain' }).harness,
    harnessWithPhoto({ bytes: [] }).harness,
    harnessWithPhoto({ sizeOverride: MAX_PHOTO_BYTES + 1 }).harness
  ];
  assert.equal(typeof cases[0].api.getPinPhotoData, 'function');
  for (const harness of cases) {
    assert.throws(
      () => harness.api.getPinPhotoData({
        __editToken: 'valid-token',
        pinId: 'pin-existing-0001'
      }),
      /photo|unavailable/i
    );
  }
});

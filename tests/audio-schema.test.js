const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const codeJs = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
const context = { console };

vm.runInNewContext(`${codeJs}
globalThis.__audioSchemaApi = {
  PinData,
  MAP_INFO_HEADERS,
  IMPORT_RECEIPT_HEADERS,
  toClientPin_: typeof toClientPin_ === 'function' ? toClientPin_ : null,
  toSharedPin_,
  importReceiptFromRow_
};`, context);

const api = context.__audioSchemaApi;
const plain = (value) => JSON.parse(JSON.stringify(value));

const LEGACY_MAP_INFO_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明', '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧', '状態', 'タグ',
  'イベント時刻', '更新時刻', 'アイコン'
];

const LEGACY_IMPORT_RECEIPT_HEADERS = [
  'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
  'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
  'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId'
];

function pinRow(audioId) {
  return [
    '2026/07/22 10:00:00', 'Audio pin', '', 35, 139, '#e53935', '', '', 'pin-audio',
    '', '未対応', '', '', '', 'default', audioId
  ];
}

test('map_info appends only the audio ID header', () => {
  assert.deepEqual(plain(api.MAP_INFO_HEADERS), LEGACY_MAP_INFO_HEADERS.concat(['音声ID']));
});

test('server pins retain audioId while client pins expose only hasAudio', () => {
  assert.equal(typeof api.toClientPin_, 'function');

  const oldRowPin = api.PinData.rowToPin(pinRow().slice(0, 15));
  const oldRowClient = api.toClientPin_(oldRowPin);
  assert.equal(oldRowPin.audioId, '');
  assert.equal(oldRowClient.hasAudio, false);
  assert.equal(Object.hasOwn(oldRowClient, 'audioId'), false);

  const newRowPin = api.PinData.rowToPin(pinRow('audio-file-id'));
  const newRowClient = api.toClientPin_(newRowPin);
  assert.equal(newRowPin.audioId, 'audio-file-id');
  assert.equal(newRowClient.hasAudio, true);
  assert.equal(Object.hasOwn(newRowClient, 'audioId'), false);
});

test('shared pins expose hasAudio without exposing audioId', () => {
  const shared = api.toSharedPin_(api.PinData.rowToPin(pinRow('private-audio-id')), []);
  assert.equal(shared.hasAudio, true);
  assert.equal(Object.hasOwn(shared, 'audioId'), false);
});

test('import receipts append media metadata and normalize legacy rows as photos', () => {
  assert.deepEqual(plain(api.IMPORT_RECEIPT_HEADERS.slice(-4)), [
    'mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId'
  ]);
  assert.deepEqual(
    plain(api.IMPORT_RECEIPT_HEADERS.slice(0, LEGACY_IMPORT_RECEIPT_HEADERS.length)),
    LEGACY_IMPORT_RECEIPT_HEADERS
  );

  const legacyRow = LEGACY_IMPORT_RECEIPT_HEADERS.map((header) => ({
    idempotencyKey: 'legacy-key',
    state: 'completed',
    pinId: 'legacy-pin',
    imageUrl: 'https://example.com/photo.jpg'
  })[header] || '');
  const readLegacyReceipt = api.importReceiptFromRow_(legacyRow, 2);
  assert.equal(readLegacyReceipt.mediaKind, 'photo');
  assert.equal(readLegacyReceipt.operationMode, '');
  assert.equal(readLegacyReceipt.targetPinId, '');
  assert.equal(readLegacyReceipt.cleanupFileId, '');
});

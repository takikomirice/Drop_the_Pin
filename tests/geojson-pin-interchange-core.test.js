const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const importUrlVectors = require('./fixtures/import-url-vectors');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadCore(extra = {}) {
  const start = indexHtml.indexOf('const ImportJobCore = (function() {');
  const end = indexHtml.indexOf('\n    const ImportQueueRunner = (function() {', start);
  assert.notEqual(start, -1, 'Expected ImportJobCore');
  assert.notEqual(end, -1, 'Expected interchange cores before ImportQueueRunner');
  const context = {
    PIN_COLORS: [{ hex: '#e53935' }, { hex: '#2196f3' }],
    PIN_ICONS: [{ id: 'default' }, { id: 'photo' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    crypto: { randomUUID: () => 'uuid' },
    Date,
    URL,
    ...extra
  };
  vm.createContext(context);
  vm.runInContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__jobCore = ImportJobCore;\n'
      + 'globalThis.__geoJsonCore = typeof GeoJsonPinInterchangeCore === "undefined" ? null : GeoJsonPinInterchangeCore;',
    context
  );
  assert.ok(context.__geoJsonCore, 'Expected GeoJsonPinInterchangeCore');
  return { geo: context.__geoJsonCore, jobs: context.__jobCore };
}

function collection(features, extras = {}) {
  return JSON.stringify({ type: 'FeatureCollection', features, ...extras });
}

function feature(properties = {}, geometry = null, extras = {}) {
  return { type: 'Feature', geometry, properties, ...extras };
}

function point(lng, lat, altitude) {
  const coordinates = altitude === undefined ? [lng, lat] : [lng, lat, altitude];
  return { type: 'Point', coordinates };
}

function sequentialIds() {
  let value = 0;
  return () => `generated-${++value}`;
}

function errorCode(fn) {
  assert.throws(fn, (error) => typeof error.code === 'string');
  try {
    fn();
  } catch (error) {
    return error.code;
  }
  return '';
}

test('GeoJSON parser accepts BOM, Japanese, emoji, escaped content, schema 1, and WGS84 CRS names', () => {
  const { geo } = loadCore();
  const crsValues = [
    'CRS84',
    'urn:ogc:def:crs:OGC:1.3:CRS84',
    'EPSG:4326',
    'urn:ogc:def:crs:EPSG::4326'
  ];
  crsValues.forEach((name, index) => {
    const text = `${index === 0 ? '\uFEFF' : ''}${collection([
      feature({ title: '東京🙂 "駅"\n改行' }, point(139.7671, 35.6812))
    ], {
      dropThePinSchemaVersion: index % 2 ? '1' : 1,
      crs: { type: 'name', properties: { name } }
    })}`;
    const parsed = geo.parse(text);
    assert.equal(parsed.type, 'FeatureCollection');
    assert.equal(parsed.features[0].properties.title, '東京🙂 "駅"\n改行');
  });
  assert.equal(geo.parse(collection([feature({ title: 'CRS省略' })])).features.length, 1);
});

test('GeoJSON CRS accepts documented representations and rejects link or conflicting names', () => {
  const { geo } = loadCore();
  const accepted = [
    ' crs84 ',
    { name: ' EPSG:4326 ' },
    { type: 'name', properties: { name: 'urn:ogc:def:crs:EPSG::4326' } },
    { type: 'name', name: 'CRS84', properties: { name: 'crs84' } },
    null
  ];
  accepted.forEach((crs) => {
    assert.equal(geo.parse(collection([feature({ title: 'accepted' })], { crs })).features.length, 1);
  });

  const rejected = [
    {},
    4326,
    { type: 'link', name: 'CRS84' },
    { type: 'link', properties: { name: 'EPSG:4326' } },
    { type: 'name', name: 'CRS84', properties: { name: 'EPSG:4326' } },
    { type: 'name', name: 4326 },
    { type: 'name', properties: { name: 4326 } },
    { type: 'name', properties: 'CRS84' }
  ];
  rejected.forEach((crs) => {
    assert.equal(
      errorCode(() => geo.parse(collection([feature({ title: 'rejected' })], { crs }))),
      'GEOJSON_UNSUPPORTED_CRS'
    );
  });
});

test('GeoJSON parser rejects unsafe or invalid whole-file input with stable codes', () => {
  const { geo } = loadCore();
  const cases = [
    ['', 'GEOJSON_EMPTY'],
    ['\uFEFF', 'GEOJSON_EMPTY'],
    ['{"type":', 'GEOJSON_INVALID_JSON'],
    ['{"type":"FeatureCollection","features":[]}\0', 'GEOJSON_NUL_CHARACTER'],
    [collection([feature({ title: 'escaped\0nul' })]), 'GEOJSON_NUL_CHARACTER'],
    ['null', 'GEOJSON_TOP_LEVEL_INVALID'],
    ['[]', 'GEOJSON_TOP_LEVEL_INVALID'],
    ['1', 'GEOJSON_TOP_LEVEL_INVALID'],
    [JSON.stringify({ type: 'Feature', features: [] }), 'GEOJSON_FEATURE_COLLECTION_REQUIRED'],
    [JSON.stringify({ type: 'FeatureCollection' }), 'GEOJSON_FEATURES_INVALID'],
    [JSON.stringify({ type: 'FeatureCollection', features: {} }), 'GEOJSON_FEATURES_INVALID'],
    [collection([]), 'GEOJSON_FEATURES_REQUIRED'],
    [collection([feature({ title: 'x' })], { dropThePinSchemaVersion: 2 }), 'GEOJSON_SCHEMA_VERSION_UNSUPPORTED'],
    [collection([feature({ title: 'x' })], { crs: { type: 'name', properties: { name: 'EPSG:3857' } } }), 'GEOJSON_UNSUPPORTED_CRS']
  ];
  cases.forEach(([input, code]) => assert.equal(errorCode(() => geo.parse(input)), code));
});

test('GeoJSON import accepts one through twenty Features and rejects all twenty-one', () => {
  const { geo, jobs } = loadCore();
  const generateId = sequentialIds();
  const one = geo.buildImportJob(collection([feature({ title: '1' })]), { generateId });
  assert.equal(one.items.length, 1);
  assert.equal(one.sourceType, 'geojson');
  const twenty = geo.buildImportJob(collection(Array.from({ length: jobs.MAX_ITEMS }, (_, index) => (
    feature({ title: String(index + 1) })
  ))), { generateId: sequentialIds() });
  assert.equal(twenty.items.length, 20);
  assert.equal(
    errorCode(() => geo.buildImportJob(collection(Array.from({ length: 21 }, (_, index) => (
      feature({ title: String(index + 1) })
    ))))),
    'IMPORT_ITEM_LIMIT_EXCEEDED'
  );
});

test('Point coordinates use longitude then latitude, preserve boundaries and negatives, and ignore altitude immutably', () => {
  const { geo } = loadCore();
  const coordinates = [-180, -90, 1234];
  const source = feature({ title: '境界' }, { type: 'Point', coordinates });
  const job = geo.buildImportJob(collection([
    source,
    feature({ title: '反対境界' }, point(180, 90))
  ]), { generateId: sequentialIds() });
  assert.deepEqual(plain(job.items.map((item) => [item.lat, item.lng])), [[-90, -180], [90, 180]]);
  assert.deepEqual(coordinates, [-180, -90, 1234]);
  assert.deepEqual(source.geometry.coordinates, coordinates);
});

test('Point permits only finite numeric altitude and later ordinates', () => {
  const { geo } = loadCore();
  const job = geo.buildImportJob(collection([
    feature({ title: 'numeric altitude' }, point(139, 35, 10)),
    feature({ title: 'string altitude' }, point(139, 35, '10')),
    feature({ title: 'null altitude' }, point(139, 35, null)),
    feature({ title: 'object altitude' }, point(139, 35, { value: 10 })),
    feature({ title: 'fourth ordinate' }, { type: 'Point', coordinates: [139, 35, 10, 'bad'] })
  ]), { generateId: sequentialIds() });
  assert.equal(job.items[0].uploadStatus, 'queued');
  assert.equal(job.items[0].lat, 35);
  assert.equal(job.items[0].lng, 139);
  assert.deepEqual(
    plain(job.items.slice(1).map((item) => item.errorCode)),
    Array(4).fill('GEOJSON_FEATURE_COORDINATES_INVALID')
  );
});

test('geometry null creates an unplaced item while invalid and unsupported geometries remain failed items', () => {
  const { geo } = loadCore();
  const geometries = [
    null,
    {},
    { type: 'Point', coordinates: '139,35' },
    { type: 'Point', coordinates: [139] },
    { type: 'Point', coordinates: ['139', 35] },
    { type: 'Point', coordinates: [181, 35] },
    { type: 'Point', coordinates: [139, -91] },
    { type: 'LineString', coordinates: [[139, 35], [140, 36]] },
    { type: 'Polygon', coordinates: [] },
    { type: 'GeometryCollection', geometries: [] }
  ];
  const job = geo.buildImportJob(collection(geometries.map((geometry, index) => (
    feature({ title: `geometry-${index}` }, geometry)
  ))), { generateId: sequentialIds() });
  assert.equal(job.items[0].uploadStatus, 'queued');
  assert.equal(job.items[0].lat, null);
  assert.equal(job.items[0].lng, null);
  assert.deepEqual(plain(job.items.slice(1, 7).map((item) => item.errorCode)), [
    'GEOJSON_FEATURE_GEOMETRY_UNSUPPORTED',
    ...Array(5).fill('GEOJSON_FEATURE_COORDINATES_INVALID')
  ]);
  assert.deepEqual(plain(job.items.slice(7).map((item) => item.errorCode)), Array(3).fill('GEOJSON_FEATURE_GEOMETRY_UNSUPPORTED'));
});

test('Feature structure failures are isolated, ordered, safe, and non-retryable', () => {
  const { geo } = loadCore();
  const values = [
    null,
    'text',
    [],
    { type: 'Other', geometry: null, properties: {} },
    { type: 'Feature', properties: {} },
    { type: 'Feature', geometry: null },
    { type: 'Feature', geometry: null, properties: [] },
    feature({ title: '正常' })
  ];
  const job = geo.buildImportJob(collection(values), { generateId: sequentialIds() });
  assert.equal(job.items.length, values.length);
  assert.deepEqual(plain(job.items.map((item) => item.uploadStatus)), [
    'failed', 'failed', 'failed', 'failed', 'failed', 'failed', 'failed', 'queued'
  ]);
  assert.equal(job.items[3].errorCode, 'GEOJSON_FEATURE_TYPE_INVALID');
  assert.equal(job.items[6].errorCode, 'GEOJSON_FEATURE_PROPERTIES_INVALID');
  job.items.slice(0, 7).forEach((item, index) => {
    assert.equal(item.retryable, false);
    assert.equal(item.attempts, 0);
    assert.match(item.error, new RegExp(`^GeoJSON ${index + 1}件目:`));
    assert.doesNotMatch(item.error, /stack|\[object Object\]/i);
  });
});

test('properties map title/name and sourceId/Feature id by precedence without reusing identity', () => {
  const { geo } = loadCore();
  const job = geo.buildImportJob(collection([
    feature({ title: '正式', name: '別名', sourceId: 'property-source' }, null, { id: 'feature-source' }),
    feature({ name: 'fallback' }, null, { id: 'feature-id' }),
    feature({ title: 'same', sourceId: 'duplicate' }),
    feature({ title: 'same', sourceId: 'duplicate' })
  ]), { generateId: sequentialIds(), jobId: 'explicit-job' });
  assert.equal(job.id, 'explicit-job');
  assert.deepEqual(plain(job.items.map((item) => item.title)), ['正式', 'fallback', 'same', 'same']);
  assert.deepEqual(plain(job.items.map((item) => item.sourceRef)), [
    'GeoJSON 1件目 / property-source', 'GeoJSON 2件目 / feature-id',
    'GeoJSON 3件目 / duplicate', 'GeoJSON 4件目 / duplicate'
  ]);
  assert.equal(new Set(job.items.map((item) => item.id)).size, 4);
  assert.equal(job.items.some((item) => ['property-source', 'feature-source', 'feature-id', 'duplicate'].includes(item.id)), false);
});

test('properties validate required lengths, catalogs, status, and eventAt with distinct codes', () => {
  const { geo } = loadCore();
  const invalid = [
    feature({}),
    feature({ title: 'x'.repeat(81) }),
    feature({ title: 'x', description: 'd'.repeat(401) }),
    feature({ title: 'x', color: '#ffffff' }),
    feature({ title: 'x', icon: 'unknown' }),
    feature({ title: 'x', status: 'unknown' }),
    feature({ title: 'x', eventAt: '2025-02-29T10:00' })
  ];
  const job = geo.buildImportJob(collection(invalid.concat([
    feature({ title: 'defaults' }),
    feature({ title: 'valid', description: '説明', color: '#2196f3', icon: 'photo', status: '', eventAt: '2000-02-29T23:59:59' })
  ])), { generateId: sequentialIds() });
  assert.deepEqual(plain(job.items.slice(0, 7).map((item) => item.errorCode)), [
    'GEOJSON_FEATURE_TITLE_REQUIRED',
    'GEOJSON_FEATURE_TITLE_TOO_LONG',
    'GEOJSON_FEATURE_DESCRIPTION_TOO_LONG',
    'GEOJSON_FEATURE_COLOR_INVALID',
    'GEOJSON_FEATURE_ICON_INVALID',
    'GEOJSON_FEATURE_STATUS_INVALID',
    'GEOJSON_FEATURE_EVENT_AT_INVALID'
  ]);
  assert.deepEqual(plain(job.items[7]), {
    id: 'generated-8', sourceType: 'geojson', sourceRef: 'GeoJSON 8件目',
    title: 'defaults', description: '', lat: null, lng: null, capturedAt: '', tags: [], links: [],
    color: '#e53935', icon: 'default', status: '', metadataStatus: 'not-applicable',
    conversionStatus: 'not-applicable', uploadStatus: 'queued', error: null, errorCode: null,
    retryable: null, attempts: 0, runtime: { originalFile: null, uploadFile: null, previewUrl: '' }
  });
  assert.equal(job.items[8].capturedAt, '2000-02-29T23:59:59');
});

test('tags accept arrays and GIS strings while preserving first spelling and enforcing five values', () => {
  const { geo } = loadCore();
  const job = geo.buildImportJob(collection([
    feature({ title: 'array', tags: ['植物', '#観察', '植物', 'PLANT', 'plant'] }),
    feature({ title: 'pipe', tags: '植物|#観察|植物' }),
    feature({ title: 'comma', tags: '植物,観察' }),
    feature({ title: 'hash-space', tags: '#植物 #観察' }),
    feature({ title: 'non-string', tags: ['ok', 1] }),
    feature({ title: 'object', tags: { value: 'x' } }),
    feature({ title: 'six', tags: '1|2|3|4|5|6' })
  ]), { generateId: sequentialIds() });
  assert.deepEqual(plain(job.items.slice(0, 4).map((item) => item.tags)), [
    ['植物', '観察', 'PLANT'], ['植物', '観察'], ['植物', '観察'], ['植物', '観察']
  ]);
  assert.deepEqual(plain(job.items.slice(4).map((item) => item.errorCode)), Array(3).fill('GEOJSON_FEATURE_TAGS_INVALID'));
});

test('links accept arrays and newline or pipe strings, deduplicate, and fail the Feature on one invalid URL', () => {
  const { geo } = loadCore();
  const validUrls = importUrlVectors.allowed.slice(0, 3);
  const invalidUrl = importUrlVectors.rejected[0];
  const job = geo.buildImportJob(collection([
    feature({ title: 'array', links: [validUrls[0], validUrls[1], validUrls[0]] }),
    feature({ title: 'string', links: `${validUrls[1]}\n${validUrls[2]}|${validUrls[1]}` }),
    feature({ title: 'invalid', links: [validUrls[0], invalidUrl] }),
    feature({ title: 'non-string', links: [validUrls[0], 1] }),
    feature({ title: 'object', links: { url: validUrls[0] } })
  ]), { generateId: sequentialIds() });
  assert.deepEqual(plain(job.items[0].links), validUrls.slice(0, 2));
  assert.deepEqual(plain(job.items[1].links), validUrls.slice(1, 3));
  assert.deepEqual(plain(job.items.slice(2).map((item) => item.errorCode)), Array(3).fill('GEOJSON_FEATURE_LINKS_INVALID'));
});

test('explicit null tags, links, and eventAt are invalid while omitted fields keep empty defaults', () => {
  const { geo } = loadCore();
  const job = geo.buildImportJob(collection([
    feature({ title: 'null tags', tags: null }),
    feature({ title: 'null links', links: null }),
    feature({ title: 'null eventAt', eventAt: null }),
    feature({ title: 'omitted' })
  ]), { generateId: sequentialIds() });
  assert.deepEqual(plain(job.items.slice(0, 3).map((item) => item.errorCode)), [
    'GEOJSON_FEATURE_TAGS_INVALID',
    'GEOJSON_FEATURE_LINKS_INVALID',
    'GEOJSON_FEATURE_EVENT_AT_INVALID'
  ]);
  assert.deepEqual(plain({
    tags: job.items[3].tags,
    links: job.items[3].links,
    eventAt: job.items[3].capturedAt
  }), { tags: [], links: [], eventAt: '' });
});

test('unknown and prototype-related fields are ignored and never copied into ImportItems', () => {
  const { geo } = loadCore();
  const input = '{"type":"FeatureCollection","unknown":{"secret":true},"features":['
    + '{"type":"Feature","id":"trace","geometry":null,"properties":'
    + '{"title":"safe","unknown":"secret","__proto__":{"title":"polluted"},'
    + '"constructor":{"title":"polluted"},"prototype":{"title":"polluted"}}}'
    + ']}';
  const item = geo.buildImportJob(input, { generateId: sequentialIds() }).items[0];
  assert.equal(item.title, 'safe');
  assert.equal(item.sourceRef, 'GeoJSON 1件目 / trace');
  assert.equal(Object.prototype.hasOwnProperty.call(item, 'unknown'), false);
  assert.equal(Object.prototype.hasOwnProperty.call(item, '__proto__'), false);
  assert.equal({}.title, undefined);
});

test('serialization emits immutable canonical GeoJSON with Point lng/lat order and null unplaced geometry', () => {
  const { geo } = loadCore();
  const pins = [
    {
      id: 'pin-1', title: '東京🙂', description: '改行\n"引用"', lat: 35.6812, lng: 139.7671,
      color: '#2196f3', icon: 'photo', status: '', tags: ['観察'],
      eventAt: '2026-07-11T10:30:45', links: ['https://example.com'],
      imageUrl: 'secret', fileId: 'file', folderUrl: 'folder', routeIds: ['route'],
      hasAudio: true, audioId: 'managed_audio_secret', audioBase64: 'SUQz-secret',
      sourceDriveFileId: 'source_audio_secret'
    },
    { id: 'pin-2', title: '未配置', description: '', lat: null, lng: null, tags: [], links: [] }
  ];
  const before = JSON.stringify(pins);
  const text = geo.serializePins(pins);
  assert.equal(text.startsWith('\uFEFF'), false);
  assert.match(text, /\n  "dropThePinSchemaVersion": 1,/);
  const exported = JSON.parse(text);
  assert.equal(exported.type, 'FeatureCollection');
  assert.equal(exported.dropThePinSchemaVersion, 1);
  assert.deepEqual(exported.features[0].geometry.coordinates, [139.7671, 35.6812]);
  assert.equal(exported.features[1].geometry, null);
  assert.deepEqual(Object.keys(exported.features[0].properties), [
    'sourceId', 'title', 'description', 'color', 'icon', 'status', 'tags', 'eventAt', 'links'
  ]);
  assert.equal(text.includes('secret'), false);
  assert.equal(text.includes('fileId'), false);
  assert.equal(text.includes('folderUrl'), false);
  assert.equal(text.includes('routeIds'), false);
  assert.equal(text.includes('managed_audio_secret'), false);
  assert.equal(text.includes('SUQz-secret'), false);
  assert.equal(text.includes('source_audio_secret'), false);
  assert.equal(JSON.stringify(pins), before);
});

test('empty pin serialization is a canonical FeatureCollection that imports as an empty Point exchange', () => {
  const { geo } = loadCore();
  const output = geo.serializePins([]);
  assert.deepEqual(JSON.parse(output), {
    type: 'FeatureCollection', dropThePinSchemaVersion: 1, features: []
  });
  assert.deepEqual(plain(geo.buildImportJob(output)), {
    sourceType: 'geojson', items: [], warnings: [], empty: true
  });
  assert.throws(
    () => geo.parse(output),
    (error) => error.code === 'GEOJSON_FEATURES_REQUIRED'
  );
});

test('serialized placed and unplaced pins round trip all exchange fields without source identity reuse', () => {
  const { geo } = loadCore();
  const pins = [
    {
      id: 'original-1', title: '日本語🙂', description: 'line 1\n"line 2"',
      lat: -33.5, lng: -70.6, color: '#2196f3', icon: 'photo', status: '',
      tags: ['植物', '観察'], eventAt: '2026-07-11T10:30:45',
      links: ['https://example.com/a', 'https://example.com/b']
    },
    {
      id: 'original-2', title: '未配置', description: '', lat: null, lng: null,
      color: '#e53935', icon: 'default', status: '完了', tags: [], eventAt: '', links: []
    }
  ];
  const job = geo.buildImportJob(geo.serializePins(pins), { generateId: sequentialIds() });
  assert.deepEqual(plain(job.items.map((item) => ({
    title: item.title, description: item.description, lat: item.lat, lng: item.lng,
    color: item.color, icon: item.icon, status: item.status, tags: item.tags,
    eventAt: item.capturedAt, links: item.links
  }))), pins.map((pin) => ({
    title: pin.title, description: pin.description, lat: pin.lat, lng: pin.lng,
    color: pin.color, icon: pin.icon, status: pin.status, tags: pin.tags,
    eventAt: pin.eventAt, links: pin.links
  })));
  assert.deepEqual(plain(job.items.map((item) => item.id)), ['generated-1', 'generated-2']);
  assert.equal(job.id, 'generated-3');
  assert.equal(job.items.some((item) => (
    Object.hasOwn(item, 'audioId') || Object.hasOwn(item, 'hasAudio')
      || Object.hasOwn(item, 'audioBase64') || Object.hasOwn(item, 'sourceDriveFileId')
  )), false);
});

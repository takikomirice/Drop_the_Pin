const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadCore(extra = {}) {
  const interchangeStart = indexHtml.indexOf('const ImportJobCore = (function() {');
  const interchangeEnd = indexHtml.indexOf('\n    const ImportQueueRunner = (function() {', interchangeStart);
  const trackStart = indexHtml.indexOf('const TrackGeometryCore = (function(');
  const trackEnd = indexHtml.indexOf('\n    const state = {', trackStart);
  assert.notEqual(interchangeStart, -1, 'Expected interchange cores');
  assert.notEqual(interchangeEnd, -1, 'Expected interchange core boundary');
  assert.notEqual(trackStart, -1, 'Expected TrackGeometryCore');
  assert.notEqual(trackEnd, -1, 'Expected track core boundary');
  const context = {
    PIN_COLORS: [{ hex: '#e53935' }, { hex: '#2196f3' }],
    PIN_ICONS: [{ id: 'default' }],
    PIN_STATUSES: ['未対応', '対応中', '完了', '保留'],
    crypto: { randomUUID: (() => { let id = 0; return () => `uuid-${++id}`; })() },
    Date,
    Math,
    URL,
    ...extra
  };
  vm.createContext(context);
  if (typeof extra.setupSource === 'string') vm.runInContext(extra.setupSource, context);
  vm.runInContext(
    `${indexHtml.slice(interchangeStart, interchangeEnd)}\n`
      + `${indexHtml.slice(trackStart, trackEnd)}\n`
      + 'globalThis.__trackGeoJson = typeof GeoJsonTrackInterchangeCore === "undefined" '
      + '? null : GeoJsonTrackInterchangeCore;\n'
      + 'globalThis.__trackGeometry = TrackGeometryCore;',
    context
  );
  assert.ok(context.__trackGeoJson, 'Expected GeoJsonTrackInterchangeCore');
  return { geo: context.__trackGeoJson, geometry: context.__trackGeometry };
}

function collection(features, extras = {}) {
  return JSON.stringify({ type: 'FeatureCollection', features, ...extras });
}

function feature(geometry, properties = {}, extras = {}) {
  return { type: 'Feature', geometry, properties, ...extras };
}

function line(coordinates) {
  return { type: 'LineString', coordinates };
}

function multiLine(coordinates) {
  return { type: 'MultiLineString', coordinates };
}

function draft(input, options = {}) {
  return loadCore().geo.buildDraft(input, { sourceName: 'walk.geojson', ...options });
}

function batch(input, options = {}) {
  return loadCore().geo.buildDraftBatch(input, { sourceName: 'walk.geojson', ...options });
}

function errorCode(run) {
  try {
    run();
  } catch (error) {
    return error && error.code;
  }
  return '';
}

test('LineString and MultiLineString become ordered disconnected segments in one track', () => {
  const coordinates = [[139, 35, 10], [139.1, 35.1, null, 'ignored']];
  const result = draft(collection([
    feature(line(coordinates), { name: 'first' }),
    feature(multiLine([
      [[140, 36], [140.1, 36.1, 20]],
      [[141, 37]]
    ])),
    feature(line([[142, 38]]))
  ]));

  assert.equal(result.trackId, 'uuid-1');
  assert.equal(result.id, 'uuid-1');
  assert.equal(result.revisionId, 'uuid-2');
  assert.equal(result.sourceType, 'geojson');
  assert.equal(result.summary.segmentCount, 4);
  assert.equal(result.summary.pointCount, 6);
  assert.deepEqual(plain(result.segments.map((segment) => segment.index)), [0, 1, 2, 3]);
  assert.deepEqual(plain(result.segments.map((segment) => segment.points.map((point) => [
    point.lat, point.lng, point.elevation, point.time
  ]))), [
    [[35, 139, 10, ''], [35.1, 139.1, null, '']],
    [[36, 140, null, ''], [36.1, 140.1, 20, '']],
    [[37, 141, null, '']],
    [[38, 142, null, '']]
  ]);
  assert.deepEqual(coordinates, [[139, 35, 10], [139.1, 35.1, null, 'ignored']]);
  const disconnectedDistance = loadCore().geometry.haversineMeters(
    { lat: 35, lng: 139 }, { lat: 35.1, lng: 139.1 }
  ) + loadCore().geometry.haversineMeters(
    { lat: 36, lng: 140 }, { lat: 36.1, lng: 140.1 }
  );
  assert.ok(Math.abs(result.summary.distanceMeters - disconnectedDistance) < 0.001);
});

test('coordinates require finite numeric lng lat and optional finite numeric elevation', () => {
  const invalid = [
    '139,35', [139], ['139', 35], [139, '35'],
    [181, 35], [-181, 35], [139, 91], [139, -91], [139, 35, '10']
  ];
  invalid.forEach((coordinate) => {
    assert.equal(
      errorCode(() => draft(collection([feature(line([coordinate]))]))),
      'GEOJSON_TRACK_COORDINATES_INVALID'
    );
  });
  [
    '{"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"LineString","coordinates":[[1e400,35]]},"properties":{}}]}',
    '{"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"LineString","coordinates":[[139,1e400]]},"properties":{}}]}',
    '{"type":"FeatureCollection","features":[{"type":"Feature","geometry":{"type":"LineString","coordinates":[[139,35,1e400]]},"properties":{}}]}'
  ].forEach((input) => {
    assert.equal(errorCode(() => draft(input)), 'GEOJSON_TRACK_COORDINATES_INVALID');
  });
  const accepted = draft(collection([feature(line([
    [-180, -90], [180, 90, null], [139, 35, 0], [140, 36, 12, 'ignored']
  ]))]));
  assert.deepEqual(plain(accepted.segments[0].points), [
    { lat: -90, lng: -180, elevation: null, time: '' },
    { lat: 90, lng: 180, elevation: null, time: '' },
    { lat: 35, lng: 139, elevation: 0, time: '' },
    { lat: 36, lng: 140, elevation: 12, time: '' }
  ]);
});

test('empty linear geometries and every unsupported or mixed geometry reject the whole file', () => {
  const rejected = [
    line([]), multiLine([]), multiLine([[]]),
    { type: 'Point', coordinates: [139, 35] }, null,
    { type: 'MultiPoint', coordinates: [] }, { type: 'Polygon', coordinates: [] },
    { type: 'MultiPolygon', coordinates: [] }, { type: 'GeometryCollection', geometries: [] },
    { type: 'Unknown', coordinates: [] }
  ];
  rejected.forEach((geometry, index) => {
    const code = errorCode(() => draft(collection([feature(geometry)])));
    if (index === 3 || index === 4) {
      assert.equal(code, 'GEOJSON_TRACK_USE_PIN_IMPORT');
    } else if (index < 3) {
      assert.equal(code, 'GEOJSON_TRACK_SEGMENT_EMPTY');
    } else {
      assert.equal(code, 'GEOJSON_TRACK_GEOMETRY_UNSUPPORTED');
    }
  });
  assert.equal(errorCode(() => draft(collection([
    feature(line([[139, 35]])),
    feature({ type: 'Point', coordinates: [140, 36] })
  ]))), 'GEOJSON_TRACK_USE_PIN_IMPORT');
});

test('properties null is empty track metadata while missing and non-object properties stay invalid', () => {
  const lineResult = loadCore().geo.buildDraft(collection([
    feature(line([[139, 35], [140, 36]]), null)
  ], { name: 'Collection fallback' }), { sourceName: 'line.geojson' });
  assert.equal(lineResult.name, 'Collection fallback');
  assert.equal(lineResult.summary.segmentCount, 1);

  const multiResult = loadCore().geo.buildDraft(collection([
    feature(multiLine([[[139, 35]], [[140, 36]]]), null)
  ]), { sourceName: 'multi.geojson' });
  assert.equal(multiResult.name, 'multi');
  assert.equal(multiResult.summary.segmentCount, 2);

  assert.equal(errorCode(() => draft(collection([
    feature({ type: 'Point', coordinates: [139, 35] }, null)
  ]))), 'GEOJSON_TRACK_USE_PIN_IMPORT');
  assert.equal(errorCode(() => draft(collection([
    feature(null, null)
  ]))), 'GEOJSON_TRACK_USE_PIN_IMPORT');

  const missingProperties = { type: 'Feature', geometry: line([[139, 35]]) };
  assert.equal(errorCode(() => draft(collection([missingProperties]))), 'GEOJSON_TRACK_FEATURE_INVALID');
  [[], 'text', 1, true].forEach((properties) => {
    assert.equal(
      errorCode(() => draft(collection([feature(line([[139, 35]]), properties)]))),
      'GEOJSON_TRACK_FEATURE_INVALID'
    );
  });
});

test('Feature and geometry structural fields must be own properties even in a polluted realm', () => {
  const pollutedFeature = loadCore({ setupSource: 'Object.prototype.type = "Feature";' }).geo;
  assert.equal(errorCode(() => pollutedFeature.buildDraft(collection([{
    geometry: line([[139, 35]]), properties: {}
  }]), { sourceName: 'feature.geojson' })), 'GEOJSON_TRACK_FEATURE_INVALID');

  const pollutedGeometry = loadCore({
    setupSource: 'Object.prototype.type = "LineString"; Object.prototype.coordinates = [[139,35]];'
  }).geo;
  assert.equal(errorCode(() => pollutedGeometry.buildDraft(collection([{
    type: 'Feature', geometry: { coordinates: [[139, 35]] }, properties: {}
  }]), { sourceName: 'geometry.geojson' })), 'GEOJSON_TRACK_GEOMETRY_UNSUPPORTED');
  assert.equal(errorCode(() => pollutedGeometry.buildDraft(collection([{
    type: 'Feature', geometry: { type: 'LineString' }, properties: {}
  }]), { sourceName: 'coordinates.geojson' })), 'GEOJSON_TRACK_COORDINATES_INVALID');
});

test('missing and scalar coordinates are invalid while empty linear arrays stay empty-segment errors', () => {
  [undefined, 'bad', {}, 1, true, null].forEach((coordinates) => {
    const geometry = { type: 'LineString' };
    if (coordinates !== undefined) geometry.coordinates = coordinates;
    assert.equal(
      errorCode(() => draft(collection([feature(geometry)]))),
      'GEOJSON_TRACK_COORDINATES_INVALID'
    );
  });
  assert.equal(errorCode(() => draft(collection([feature(
    multiLine([null]), { coordTimes: [[]] }
  )]))), 'GEOJSON_TRACK_COORDINATES_INVALID');
  assert.equal(errorCode(() => draft(collection([feature(line([]))]))), 'GEOJSON_TRACK_SEGMENT_EMPTY');
});

test('coordTimes and times follow coordinate shape, normalize RFC3339, and preserve point order', () => {
  const result = draft(collection([
    feature(line([[139, 35], [140, 36]]), {
      coordTimes: ['2026-07-11T02:00:00+09:00', null]
    }),
    feature(multiLine([[[141, 37]], [[142, 38], [143, 39]]]), {
      times: [[''], ['2026-07-11T01:00:00Z', '2026-07-11T00:00:00Z']]
    })
  ]));
  assert.deepEqual(plain(result.segments.map((segment) => segment.points.map((point) => point.time))), [
    ['2026-07-10T17:00:00.000Z', ''],
    [''],
    ['2026-07-11T01:00:00.000Z', '2026-07-11T00:00:00.000Z']
  ]);
  assert.deepEqual(plain(result.segments[2].points.map((point) => point.lng)), [142, 143]);
  assert.equal(result.summary.startTime, '2026-07-10T17:00:00.000Z');
  assert.equal(result.summary.endTime, '2026-07-11T00:00:00.000Z');
});

test('matching dual time extensions are accepted while conflicts and invalid shapes reject', () => {
  const same = ['2026-07-11T01:00:00Z', ''];
  assert.equal(draft(collection([feature(line([[139, 35], [140, 36]]), {
    coordTimes: same, times: same.slice()
  })])).summary.pointCount, 2);

  const invalidCases = [
    [{
      coordTimes: ['2026-07-11T01:00:00Z', ''],
      times: ['2026-07-11T02:00:00Z', '']
    }, 'GEOJSON_TRACK_TIME_CONFLICT'],
    [{ coordTimes: ['2026-07-11T01:00:00Z'] }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ coordTimes: null }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ times: null }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ coordTimes: null, times: null }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ coordTimes: {}, times: {} }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ coordTimes: '2026-07-11T01:00:00Z' }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ coordTimes: ['2026-07-11T01:00:00', ''] }, 'GEOJSON_TRACK_TIME_INVALID'],
    [{ coordTimes: [123, ''] }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ coordTimes: [true, ''] }, 'GEOJSON_TRACK_TIME_SHAPE_INVALID'],
    [{ coordTimes: ['not-a-date', ''] }, 'GEOJSON_TRACK_TIME_INVALID']
  ];
  invalidCases.forEach(([properties, code]) => {
    assert.equal(errorCode(() => draft(collection([
      feature(line([[139, 35], [140, 36]]), properties)
    ]))), code);
  });

  const multiShapeCases = [
    ['flat', 'flat'],
    [['2026-07-11T01:00:00Z'], []],
    [['2026-07-11T01:00:00Z'], ['']],
    [null, ['']]
  ];
  multiShapeCases.forEach((coordTimes) => {
    assert.equal(errorCode(() => draft(collection([
      feature(multiLine([[[139, 35]], [[140, 36], [141, 37]]]), { coordTimes })
    ]))), 'GEOJSON_TRACK_TIME_SHAPE_INVALID');
  });
});

test('metadata uses explicit own-property precedence and safe defaults', () => {
  const source = JSON.parse(collection([
    feature(line([[139, 35]]), {
      title: 'Feature title', name: 'Feature name', description: 'Feature description',
      color: '#2196F3', lineStyle: 'dashed', lineWidth: 7, visible: false,
      ignored: 'not copied', sourceId: 'not-an-id'
    }, { id: 'feature-id' })
  ], {
    name: 'Collection name', description: 'Collection description', color: '#E53935',
    lineStyle: 'dotted', lineWidth: 3, visible: true,
    ignored: 'not copied', sourceId: 'not-an-id'
  }));
  const result = draft(JSON.stringify(source));
  assert.equal(result.name, 'Collection name');
  assert.equal(result.description, 'Collection description');
  assert.equal(result.color, '#e53935');
  assert.equal(result.lineStyle, 'dotted');
  assert.equal(result.lineWidth, 4);
  assert.equal(result.visible, true);
  assert.equal(result.sourceName, 'walk.geojson');
  assert.equal(result.trackId, 'uuid-1');
  assert.equal(result.revisionId, 'uuid-2');
  for (const key of ['ignored', 'sourceId', 'properties', 'features', 'geometry']) {
    assert.equal(Object.prototype.hasOwnProperty.call(result, key), false, key);
  }

  const featureFallback = draft(collection([
    feature(line([[139, 35]]), {
      title: 'Feature title', name: 'Feature name', description: 'Feature description',
      color: '#2196f3', lineStyle: 'dashed', lineWidth: 7, visible: false
    })
  ]));
  assert.equal(featureFallback.name, 'Feature title');
  assert.equal(featureFallback.description, 'Feature description');
  assert.equal(featureFallback.color, '#2196f3');
  assert.equal(featureFallback.lineStyle, 'dashed');
  assert.equal(featureFallback.lineWidth, 4);
  assert.equal(featureFallback.visible, false);

  const filenameFallback = loadCore().geo.buildDraft(collection([
    feature(line([[139, 35]]), { name: '' })
  ]), { sourceName: 'folder\\fallback.json' });
  assert.equal(filenameFallback.name, 'fallback');
  assert.equal(filenameFallback.sourceName, 'fallback.json');
  assert.equal(draft(collection([feature(line([[139, 35]]))]), { sourceName: '' }).name, 'GeoJSONトラック');
});

test('parse itself rejects invalid whitelisted metadata before identity generation', () => {
  const { geo } = loadCore();
  const invalidCollections = [
    { name: 'x'.repeat(101) },
    { description: 'x'.repeat(401) },
    { color: '#ffffff' },
    { lineStyle: 'double' },
    { visible: 'true' }
  ];
  invalidCollections.forEach((extras) => {
    assert.equal(errorCode(() => geo.parse(
      collection([feature(line([[139, 35]]))], extras),
      { sourceName: 'safe.geojson' }
    )), 'GEOJSON_TRACK_METADATA_INVALID');
  });
  for (const lineWidth of [0, 4.5, 10, 'wide', null]) {
    const parsed = geo.parse(
      collection([feature(line([[139, 35]]))], { lineWidth }),
      { sourceName: 'safe.geojson' }
    );
    assert.equal(parsed.lineWidth, 4);
  }
});

test('higher-priority metadata ignores malformed lower-priority candidates', () => {
  const { geo } = loadCore();
  const result = geo.buildDraft(collection([
    feature(line([[139, 35]]), {
      title: { ignored: true }, name: ['ignored'], description: { ignored: true }
    })
  ], {
    name: 'Top-level name', description: 'Top-level description'
  }), { sourceName: 'safe.geojson' });
  assert.equal(result.name, 'Top-level name');
  assert.equal(result.description, 'Top-level description');
});

test('prototype values and dangerous keys never become track metadata', () => {
  const inherited = Object.create({
    name: 'inherited collection', description: 'inherited description', color: '#2196f3',
    lineStyle: 'dashed', lineWidth: 9, visible: false
  });
  inherited.type = 'FeatureCollection';
  inherited.features = [feature(line([[139, 35]]), Object.assign(Object.create({
    title: 'inherited title', name: 'inherited feature', color: '#2196f3'
  }), { description: 'own description' }))];
  const { geo } = loadCore();
  const parsed = geo.parse(JSON.stringify(inherited), { sourceName: 'safe.geojson' });
  assert.equal(parsed.name, 'safe');
  assert.equal(parsed.description, 'own description');
  assert.equal(parsed.color, '#e53935');

  const polluted = collection([feature(line([[139, 35]]), JSON.parse(
    '{"__proto__":{"title":"polluted"},"constructor":{"name":"bad"},"prototype":{"color":"#2196f3"}}'
  ))]);
  const result = geo.buildDraft(polluted, { sourceName: 'safe.geojson' });
  assert.equal(result.name, 'safe');
  assert.equal(result.color, '#e53935');
  assert.equal({}.title, undefined);
});

test('feature segment and point limits reject atomically without truncation', () => {
  const { geo, geometry } = loadCore();
  assert.equal(geometry.MAX_SEGMENTS, 200);
  assert.equal(geometry.MAX_POINTS, 20000);
  const segmentsAtLimit = Array.from({ length: geometry.MAX_SEGMENTS }, (_, index) => [[index / 10, 0]]);
  assert.equal(geo.buildDraft(collection([feature(multiLine(segmentsAtLimit))]), {
    sourceName: 'limit.geojson'
  }).summary.segmentCount, 200);
  assert.equal(errorCode(() => geo.buildDraft(collection([
    feature(multiLine([...segmentsAtLimit, [[21, 0]]]))
  ]), { sourceName: 'overflow.geojson' })), 'TRACK_SEGMENT_LIMIT_EXCEEDED');

  const pointsAtLimit = Array.from({ length: geometry.MAX_POINTS }, (_, index) => [
    (index % 36000) / 100 - 180, 0
  ]);
  assert.equal(geo.buildDraft(collection([feature(line(pointsAtLimit))]), {
    sourceName: 'limit.geojson'
  }).summary.pointCount, 20000);
  assert.equal(errorCode(() => geo.buildDraft(collection([
    feature(line([...pointsAtLimit, [0, 0]]))
  ]), { sourceName: 'overflow.geojson' })), 'TRACK_POINT_LIMIT_EXCEEDED');
});

test('GeoJSON batch removes only exact timed duplicates within one source segment', () => {
  const duplicate = [139, 35, 10];
  const time = '2026-01-01T00:00:00Z';
  const reduced = batch(collection([
    feature(line(Array.from({ length: 20001 }, () => duplicate.slice())), {
      coordTimes: Array.from({ length: 20001 }, () => time)
    })
  ]));
  assert.equal(reduced.stats.sourcePointCount, 20001);
  assert.equal(reduced.stats.duplicatePointCount, 20000);
  assert.equal(reduced.stats.pointCount, 1);
  assert.equal(reduced.drafts.length, 1);

  const exactness = batch(collection([
    feature(multiLine([
      [[139, 35, 10], [139, 35, 10], [139, 35, 11], [139.1, 35]],
      [[139, 35, 10]]
    ]), {
      coordTimes: [
        [time, '', time, time],
        [time]
      ]
    })
  ]));
  assert.equal(exactness.stats.sourcePointCount, 5);
  assert.equal(exactness.stats.duplicatePointCount, 0);
  assert.equal(exactness.stats.pointCount, 5);
});

test('GeoJSON batch splits four-hour interruptions but keeps a midnight crossing together', () => {
  const start = Date.parse('2026-01-01T23:59:59Z');
  const result = batch(collection([
    feature(line([[139, 35], [139.1, 35.1], [139.2, 35.2]]), {
      coordTimes: [
        new Date(start).toISOString(),
        new Date(start + 2000).toISOString(),
        new Date(start + 4 * 60 * 60 * 1000 + 2000).toISOString()
      ]
    })
  ]));
  assert.equal(result.drafts.length, 2);
  assert.equal(result.stats.interruptionCount, 1);
  assert.deepEqual(plain(result.drafts.map((value) => value.name)), [
    'walk(1/2)', 'walk(2/2)'
  ]);
  assert.deepEqual(plain(result.drafts.map((value) => value.summary.pointCount)), [2, 1]);
});

test('GeoJSON batch uses the smallest time interval and partitions irreducible timed data', () => {
  const start = Date.parse('2026-01-01T00:00:00Z');
  const seconds = Array.from({ length: 20001 }, (_, index) => index);
  const compressed = batch(collection([
    feature(line(seconds.map((index) => [139 + index / 1000000, 35])), {
      coordTimes: seconds.map((index) => new Date(start + index * 1000).toISOString())
    })
  ]));
  assert.equal(compressed.drafts.length, 1);
  assert.equal(compressed.stats.pointCount, 10001);
  assert.equal(compressed.stats.timeCompressedPointCount, 10000);
  assert.equal(compressed.stats.shapeCompressedPointCount, 0);
  assert.deepEqual(plain(compressed.stats.compressionIntervals), [2]);
  assert.deepEqual(
    plain(compressed.drafts[0].segments[0].points.slice(0, 3).map((value) => value.time)),
    [
      '2026-01-01T00:00:00.000Z',
      '2026-01-01T00:00:02.000Z',
      '2026-01-01T00:00:04.000Z'
    ]
  );
  assert.equal(
    compressed.drafts[0].segments[0].points.at(-1).time,
    '2026-01-01T05:33:20.000Z'
  );

  const everyTenSeconds = batch(collection([
    feature(line(seconds.map((index) => [139 + index / 1000000, 35])), {
      coordTimes: seconds.map((index) => new Date(start + index * 10000).toISOString())
    })
  ]));
  assert.deepEqual(
    plain(everyTenSeconds.drafts.map((value) => value.summary.pointCount)),
    [20000, 1]
  );
  assert.equal(everyTenSeconds.stats.compressedPointCount, 0);
  assert.deepEqual(plain(everyTenSeconds.stats.compressionIntervals), []);
});

test('GeoJSON batch shape reduction preserves protected untimed geometry and is inactive at the limit', () => {
  const coordinates = Array.from({ length: 20001 }, (_, index) => (
    [139 + index / 1000000, 35, index === 5000 ? -10 : (index === 15000 ? 50 : null)]
  ));
  const times = coordinates.map(() => '');
  times[10000] = '2026-01-01T00:00:00Z';
  const result = batch(collection([
    feature(line(coordinates), { coordTimes: times })
  ]));
  assert.equal(result.drafts.length, 1);
  assert.equal(result.stats.pointCount, 20000);
  assert.equal(result.stats.shapeCompressedPointCount, 1);
  assert.equal(result.stats.timeCompressedPointCount, 0);
  const retained = result.drafts[0].segments[0].points;
  assert.ok(retained.some((value) => value.elevation === -10));
  assert.ok(retained.some((value) => value.elevation === 50));
  assert.ok(retained.some((value) => value.time === '2026-01-01T00:00:00.000Z'));

  const atLimit = coordinates.slice(0, 20000);
  const unchanged = batch(collection([feature(line(atLimit))]));
  assert.deepEqual(
    plain(unchanged.drafts[0].segments[0].points.map((value) => [
      value.lng, value.lat, value.elevation, value.time
    ])),
    atLimit.map((value) => [value[0], value[1], value[2], ''])
  );
  assert.equal(unchanged.stats.compressedPointCount, 0);
});

test('GeoJSON batch enforces independent source segment and generated-track limits before IDs', () => {
  let generated = 0;
  const generateId = () => `generated-${++generated}`;
  const sourceLimitCoordinates = Array.from({ length: 100001 }, () => [139, 35]);
  assert.equal(errorCode(() => batch(collection([
    feature(line(sourceLimitCoordinates))
  ]), { generateId })), 'GEOJSON_SOURCE_POINT_LIMIT_EXCEEDED');
  assert.equal(generated, 0);

  const acceptedCoordinates = sourceLimitCoordinates.slice(0, 100000);
  const accepted = batch(collection([feature(line(acceptedCoordinates))]));
  assert.equal(accepted.stats.sourcePointCount, 100000);
  assert.equal(accepted.stats.pointCount, 20000);

  const segments = Array.from({ length: 200 }, (_, index) => [[139, 35 + index / 1000]]);
  assert.equal(batch(collection([feature(multiLine(segments))])).drafts.length, 1);
  assert.equal(errorCode(() => batch(collection([
    feature(multiLine([...segments, [[140, 36]]]))
  ]))), 'TRACK_SEGMENT_LIMIT_EXCEEDED');

  const start = Date.parse('2026-01-01T00:00:00Z');
  function interrupted(count) {
    return collection([feature(
      line(Array.from({ length: count }, (_, index) => [139, 35 + index / 1000])),
      { coordTimes: Array.from({ length: count }, (_, index) => (
        new Date(start + index * 4 * 60 * 60 * 1000).toISOString()
      )) }
    )]);
  }
  assert.equal(batch(interrupted(20)).drafts.length, 20);
  assert.equal(
    errorCode(() => batch(interrupted(21))),
    'GEOJSON_GENERATED_TRACK_LIMIT_EXCEEDED'
  );
});

test('GeoJSON batch edits regenerate bounded suffixes and save independent payloads', () => {
  const start = Date.parse('2026-01-01T00:00:00Z');
  const core = loadCore().geo;
  const created = core.buildDraftBatch(collection([
    feature(line([[139, 35], [140, 36]]), {
      coordTimes: [
        new Date(start).toISOString(),
        new Date(start + 4 * 60 * 60 * 1000).toISOString()
      ]
    })
  ]), { sourceName: 'walk.geojson' });
  const updated = core.updateDraftBatch(created, {
    name: 'x'.repeat(100),
    description: 'shared',
    color: '#2196f3',
    visible: false,
    lineStyle: 'dashed'
  });
  assert.deepEqual(plain(updated.drafts.map((value) => value.name)), [
    `${'x'.repeat(95)}(1/2)`,
    `${'x'.repeat(95)}(2/2)`
  ]);
  const payloads = core.toSavePayloads(updated);
  assert.equal(payloads.length, 2);
  assert.notEqual(payloads[0], payloads[1]);
  assert.notEqual(payloads[0].segments, payloads[1].segments);
  assert.deepEqual(plain(payloads.map((value) => value.description)), ['shared', 'shared']);
  assert.equal(updated.summary.pointCount, 2);
});

test('draft edits preserve identity and geometry while save payload is deeply whitelisted', () => {
  const { geo } = loadCore();
  const original = geo.buildDraft(collection([
    feature(line([[139, 35, 10], [140, 36]]))
  ]), { sourceName: 'walk.geojson' });
  const updated = geo.updateDraft(original, {
    name: 'Edited', description: 'Edited description', color: '#2196f3',
    lineStyle: 'dashed', lineWidth: 6, visible: false,
    trackId: 'attacker', revisionId: 'attacker', segments: [], arbitrary: 'ignored'
  });
  assert.equal(updated.trackId, original.trackId);
  assert.equal(updated.revisionId, original.revisionId);
  assert.deepEqual(plain(updated.segments), plain(original.segments));
  assert.notEqual(updated.segments, original.segments);
  assert.equal(updated.name, 'Edited');
  assert.equal(updated.visible, false);
  assert.equal(updated.lineWidth, 4);

  const payload = geo.toSavePayload({
    ...updated,
    file: { name: 'secret' }, originalGeoJson: 'secret', properties: { secret: true },
    summary: { fake: true }, editToken: 'secret'
  });
  assert.deepEqual(Object.keys(payload).sort(), [
    'color', 'description', 'lineStyle', 'lineWidth', 'name', 'revisionId',
    'segments', 'sourceName', 'sourceType', 'trackId', 'visible'
  ]);
  assert.equal(Object.prototype.hasOwnProperty.call(payload, 'orderIndex'), false);
  assert.equal(payload.trackId, original.trackId);
  assert.equal(payload.revisionId, original.revisionId);
  assert.equal(payload.lineWidth, 4);
  assert.notEqual(payload.segments, updated.segments);
});

test('each new selection creates new identifiers while a retry projection keeps them stable', () => {
  const { geo } = loadCore({
    crypto: { randomUUID: (() => { let id = 0; return () => `selection-${++id}`; })() }
  });
  const input = collection([feature(line([[139, 35]]), { sourceId: 'source-id' }, { id: 'feature-id' })]);
  const first = geo.buildDraft(input, { sourceName: 'one.geojson' });
  const second = geo.buildDraft(input, { sourceName: 'one.geojson' });
  assert.deepEqual([first.trackId, first.revisionId], ['selection-1', 'selection-2']);
  assert.deepEqual([second.trackId, second.revisionId], ['selection-3', 'selection-4']);
  assert.equal(first.trackId === 'feature-id' || first.trackId === 'source-id', false);
  assert.deepEqual(
    [geo.toSavePayload(first).trackId, geo.toSavePayload(first).revisionId],
    [first.trackId, first.revisionId]
  );
});

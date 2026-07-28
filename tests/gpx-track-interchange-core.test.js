const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const fixture = (name) => fs.readFileSync(path.join(__dirname, 'fixtures', name), 'utf8');

class TextNode {
  constructor(value) { this.nodeType = 3; this.nodeValue = value; }
  get textContent() { return this.nodeValue; }
}

class ElementNode {
  constructor(tagName, namespaceURI, attributes) {
    this.nodeType = 1;
    this.tagName = tagName;
    this.localName = tagName.split(':').pop();
    this.namespaceURI = namespaceURI;
    this.childNodes = [];
    this.attributes = attributes;
  }
  get children() { return this.childNodes.filter((node) => node.nodeType === 1); }
  get textContent() { return this.childNodes.map((node) => node.textContent).join(''); }
  getAttribute(name) { return Object.prototype.hasOwnProperty.call(this.attributes, name) ? this.attributes[name] : null; }
  getElementsByTagName(name) {
    const result = [];
    const visit = (node) => {
      if (node.nodeType !== 1) return;
      if (node.tagName === name || node.localName === name) result.push(node);
      node.childNodes.forEach(visit);
    };
    this.childNodes.forEach(visit);
    return result;
  }
}

class DocumentNode {
  constructor(rootElement) { this.documentElement = rootElement; }
  getElementsByTagName(name) {
    if (!this.documentElement) return [];
    const own = this.documentElement.tagName === name || this.documentElement.localName === name
      ? [this.documentElement] : [];
    return own.concat(this.documentElement.getElementsByTagName(name));
  }
}

function decodeXml(value) {
  return value.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'").replace(/&amp;/g, '&');
}

function parserErrorDocument() {
  return new DocumentNode(new ElementNode('parsererror', 'http://www.mozilla.org/newlayout/xml/parsererror.xml', {}));
}

function parseFixtureXml(source) {
  const tokens = source.match(/<!--[\s\S]*?-->|<!\[CDATA\[[\s\S]*?\]\]>|<\?[\s\S]*?\?>|<[^>]*>|[^<]+/g) || [];
  if (tokens.join('') !== source) return parserErrorDocument();
  const stack = [];
  let rootElement = null;
  for (const token of tokens) {
    if (token.startsWith('<?') || token.startsWith('<!--')) continue;
    if (token.startsWith('<![CDATA[')) {
      if (!stack.length) return parserErrorDocument();
      stack.at(-1).element.childNodes.push(new TextNode(token.slice(9, -3)));
      continue;
    }
    if (token.startsWith('</')) {
      const name = token.slice(2, -1).trim();
      if (!stack.length || stack.at(-1).element.tagName !== name) return parserErrorDocument();
      stack.pop();
      continue;
    }
    if (token.startsWith('<')) {
      if (/^<!/i.test(token)) return parserErrorDocument();
      const match = token.match(/^<([^\s/>]+)([\s\S]*?)(\/?)>$/);
      if (!match) return parserErrorDocument();
      const attributes = {};
      const attributeSource = match[2];
      const attributePattern = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
      let attributeMatch;
      while ((attributeMatch = attributePattern.exec(attributeSource))) {
        attributes[attributeMatch[1]] = decodeXml(attributeMatch[2] === undefined ? attributeMatch[3] : attributeMatch[2]);
      }
      const inherited = stack.length ? stack.at(-1).namespaces : {};
      const namespaces = Object.assign({}, inherited);
      Object.keys(attributes).forEach((name) => {
        if (name === 'xmlns') namespaces[''] = attributes[name];
        else if (name.startsWith('xmlns:')) namespaces[name.slice(6)] = attributes[name];
      });
      const prefix = match[1].includes(':') ? match[1].split(':')[0] : '';
      const element = new ElementNode(match[1], namespaces[prefix] || '', attributes);
      if (stack.length) stack.at(-1).element.childNodes.push(element);
      else if (rootElement) return parserErrorDocument();
      else rootElement = element;
      if (!match[3]) stack.push({ element, namespaces });
      continue;
    }
    if (stack.length) stack.at(-1).element.childNodes.push(new TextNode(decodeXml(token)));
    else if (token.trim()) return parserErrorDocument();
  }
  return stack.length || !rootElement ? parserErrorDocument() : new DocumentNode(rootElement);
}

function loadCore(extra = {}) {
  const start = indexHtml.indexOf('const TrackGeometryCore =');
  const end = indexHtml.indexOf('\n    const GeoJsonTrackImportUI =', start);
  assert.notEqual(start, -1, 'Expected TrackGeometryCore');
  assert.notEqual(end, -1, 'Expected track core boundary');
  const context = {
    PIN_COLORS: [{ hex: '#e53935' }, { hex: '#2196f3' }],
    Date,
    Math,
    crypto: { randomUUID: (() => { let id = 0; return () => `uuid-${++id}`; })() },
    ...extra
  };
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\n`
      + 'globalThis.__gpx = typeof GpxTrackInterchangeCore === "undefined" ? null : GpxTrackInterchangeCore;',
    context
  );
  assert.ok(context.__gpx, 'Expected GpxTrackInterchangeCore');
  return context.__gpx;
}

function plain(value) { return JSON.parse(JSON.stringify(value)); }
function options(overrides = {}) { return { sourceName: 'C:\\fake\\walk.gpx', parseXml: parseFixtureXml, ...overrides }; }
function parse(xml, overrides = {}) { return loadCore().parse(xml, options(overrides)); }
function build(xml, overrides = {}) { return loadCore().buildDraft(xml, options(overrides)); }
function error(run) { try { run(); } catch (caught) { return caught; } return null; }
function code(run) { return error(run)?.code || ''; }
function trkptXml(lat, lon, time, elevation = 10) {
  return `<trkpt lat="${lat}" lon="${lon}"><ele>${elevation}</ele><time>${time}</time></trkpt>`;
}

test('GPX 1.0 track fixture parses direct metadata and points without IDs', () => {
  const result = parse(fixture('gpx-1.0-track.gpx'));
  assert.equal(result.name, '架空の散歩');
  assert.equal(result.description, 'GPX 1.0 fixture');
  assert.equal(result.sourceType, 'gpx');
  assert.equal(result.sourceName, 'walk.gpx');
  assert.equal(result.trackId, undefined);
  assert.deepEqual(plain(result.segments), [{ index: 0, points: [
    { lat: 35, lng: 139, elevation: 10.5, time: '2026-07-11T01:02:03.000Z' },
    { lat: 35.001, lng: 139.001, elevation: null, time: '' }
  ] }]);
});

test('GPX 1.1 trkseg and route fixtures become disconnected ordered segments', () => {
  const track = parse(fixture('gpx-1.1-multisegment.gpx'));
  assert.deepEqual(plain(track.segments.map((segment) => segment.points.map((point) => point.lat))), [[36], [36.1, 36.2]]);
  const route = parse(fixture('gpx-route.gpx'));
  assert.equal(route.name, '架空の予定経路');
  assert.deepEqual(plain(route.segments[0].points[1]), {
    lat: 34.6, lng: 138.6, elevation: -5, time: '2026-07-11T03:00:00.000Z'
  });
  assert.equal(route.routeGroups, undefined);
  assert.equal(route.route_pins, undefined);
});

test('mixed prefixed GPX preserves root document order and ignores empty segments and waypoints', () => {
  const result = parse(fixture('gpx-mixed.gpx'), { sourceName: '' });
  assert.deepEqual(plain(result.segments.map((segment) => segment.points[0].lat)), [30, 32, 33]);
  assert.equal(result.segments[1].points[0].time, '');
  assert.equal(result.name, 'GPXトラック');
  assert.deepEqual(plain(result.warnings), [
    { code: 'GPX_WAYPOINTS_IGNORED', count: 1 },
    { code: 'GPX_EMPTY_SEGMENTS_IGNORED', count: 1 },
    { code: 'GPX_POINTS_WITHOUT_TIME', count: 3 },
    { code: 'GPX_POINTS_WITHOUT_ELEVATION', count: 3 }
  ]);
});

test('XML declaration comments CDATA namespaces and namespace-free GPX are accepted', () => {
  const values = [
    '<?xml version="1.0"?><gpx version="1.1"><metadata><name><![CDATA[ A < B ]]></name></metadata><trk><trkseg><trkpt lat=".5" lon="-.5"/></trkseg></trk></gpx>',
    '<x:gpx xmlns:x="http://www.topografix.com/GPX/1/0" version="1.0"><!--ok--><x:trk><x:trkseg><x:trkpt lat="+35" lon="-139"/></x:trkseg></x:trk></x:gpx>'
  ];
  assert.equal(parse(values[0]).name, 'A < B');
  assert.equal(parse(values[0]).segments[0].points[0].lng, -0.5);
  assert.equal(parse(values[1]).segments[0].points[0].lat, 35);
});

test('preflight rejects empty BOM-only NUL dangerous declarations and excessive text', () => {
  ['', '  \n', '\uFEFF', '\uFEFF \r\n'].forEach((xml) => assert.equal(code(() => parse(xml)), 'GPX_EMPTY'));
  assert.equal(code(() => parse('<gpx\0 version="1.1"/>')), 'GPX_NUL_CHARACTER');
  ['<!DOCTYPE gpx><gpx version="1.1"/>', '<!dOcTyPe gpx><gpx version="1.1"/>',
    '<!ENTITY x "y"><gpx version="1.1"/>', '<!eNtItY x "y"><gpx version="1.1"/>']
    .forEach((xml) => assert.equal(code(() => parse(xml)), 'GPX_DANGEROUS_XML'));
  const core = loadCore();
  assert.equal(core.MAX_GPX_TEXT_LENGTH, 5 * 1024 * 1024);
  const minimal = '<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>';
  const exact = minimal + ' '.repeat(core.MAX_GPX_TEXT_LENGTH - minimal.length);
  assert.equal(core.parse(exact, options()).stats.pointCount, 1);
  assert.equal(code(() => core.parse('x'.repeat(core.MAX_GPX_TEXT_LENGTH + 1), options())), 'GPX_TOO_LARGE');
});

test('parser anomalies and invalid GPX envelopes return safe errors', () => {
  const secret = '<gpx version="1.1"><secret>private-coordinate</secret>';
  const cases = [
    [() => { throw new Error(`parser leaked ${secret}`); }, secret],
    [() => parserErrorDocument(), secret],
    [() => new DocumentNode(null), secret],
    [() => parseFixtureXml('<html><body/></html>'), '<html><body/></html>'],
    [() => parseFixtureXml('<gpx version="1.1" xmlns="https://invalid.example/gpx"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>'), 'invalid.example'],
    [() => parseFixtureXml('<gpx xmlns="http://www.topografix.com/GPX/1/1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>'), 'missing-version-private'],
    [() => parseFixtureXml('<gpx version="1.2" xmlns="http://www.topografix.com/GPX/1/1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>'), '1.2']
  ];
  cases.forEach(([parseXml, marker]) => {
    const caught = error(() => loadCore().parse(secret, { sourceName: 'safe.gpx', parseXml }));
    assert.ok(caught?.code);
    assert.equal(caught.message.includes(marker), false);
    assert.equal(caught.message.includes('private-coordinate'), false);
  });
  assert.equal(code(() => parse('<gpx version="1.1"><trk>')), 'GPX_INVALID_XML');
  assert.equal(code(() => parse(
    '<gpx version="1.1" xmlns="http://www.topografix.com/GPX/1/0"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>'
  )), 'GPX_NAMESPACE_UNSUPPORTED');
  assert.equal(code(() => parse(
    '<gpx version="1.0" xmlns="http://www.topografix.com/GPX/1/1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>'
  )), 'GPX_NAMESPACE_UNSUPPORTED');
});

test('foreign namespace elements never become GPX tracks metadata or point fields', () => {
  const xml = [
    '<gpx version="1.1" xmlns="http://www.topografix.com/GPX/1/1" xmlns:e="urn:foreign">',
    '<e:metadata><e:name>Foreign metadata</e:name></e:metadata>',
    '<e:trk><e:trkseg><e:trkpt lat="9" lon="9"/></e:trkseg></e:trk>',
    '<rte><name>Valid route</name><rtept lat="1" lon="2"><e:ele>999</e:ele><e:time>2026-01-01T00:00:00Z</e:time></rtept></rte>',
    '</gpx>'
  ].join('');
  const result = parse(xml, { sourceName: 'safe.gpx' });
  assert.equal(result.name, 'Valid route');
  assert.equal(result.stats.trackElementCount, 0);
  assert.equal(result.stats.routeElementCount, 1);
  assert.deepEqual(plain(result.segments), [
    { index: 0, points: [{ lat: 1, lng: 2, elevation: null, time: '' }] }
  ]);
  assert.equal(code(() => parse(
    '<gpx version="1.1" xmlns="http://www.topografix.com/GPX/1/1" xmlns:e="urn:foreign"><e:trk><e:trkseg><e:trkpt lat="9" lon="9"/></e:trkseg></e:trk></gpx>'
  )), 'GPX_TRACK_REQUIRED');
});

test('namespaced parser errors and hostile DOM accessors cannot leak parser details', () => {
  const validRoot = parseFixtureXml('<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>').documentElement;
  const namespacedParserError = {
    documentElement: validRoot,
    getElementsByTagName: () => [],
    getElementsByTagNameNS: (_namespace, localName) => localName === 'parsererror' ? [{}] : []
  };
  assert.equal(code(() => loadCore().parse('<gpx version="1.1"/>', {
    sourceName: 'safe.gpx', parseXml: () => namespacedParserError
  })), 'GPX_INVALID_XML');

  const hostileRoot = {
    nodeType: 1,
    localName: 'gpx',
    namespaceURI: '',
    getAttribute: () => '1.1',
    get children() { throw new Error('private DOM contents'); }
  };
  const hostileError = error(() => loadCore().parse('<gpx version="1.1"/>', {
    sourceName: 'safe.gpx',
    parseXml: () => ({ documentElement: hostileRoot, getElementsByTagName: () => [] })
  }));
  assert.equal(hostileError?.code, 'GPX_INVALID_XML');
  assert.equal(hostileError?.message.includes('private DOM contents'), false);

  const codedHostileRoot = {
    nodeType: 1,
    get localName() {
      throw Object.assign(new Error('private-coordinate-139'), { code: 'GPX_POINT_TIME_INVALID' });
    }
  };
  const codedHostileError = error(() => loadCore().parse('<gpx version="1.1"/>', {
    sourceName: 'safe.gpx',
    parseXml: () => ({ documentElement: codedHostileRoot, getElementsByTagName: () => [] })
  }));
  assert.equal(codedHostileError?.code, 'GPX_INVALID_XML');
  assert.equal(codedHostileError?.message.includes('private-coordinate-139'), false);
});

test('strict decimal coordinates accept boundaries negative zero and reject malformed values', () => {
  const accepted = ['-90', '90', '-0', '+35.0', '.5', '-.5'];
  accepted.forEach((lat) => {
    const point = parse(`<gpx version="1.1"><trk><trkseg><trkpt lat="${lat}" lon="180"/></trkseg></trk></gpx>`).segments[0].points[0];
    assert.equal(Number.isFinite(point.lat), true);
  });
  const invalid = ['', ' ', '1e2', '1E2', '0x10', 'NaN', 'Infinity', '90.1'];
  invalid.forEach((lat) => assert.equal(code(() => parse(`<gpx version="1.1"><trk><trkseg><trkpt lat="${lat}" lon="139"/></trkseg></trk></gpx>`)), 'GPX_POINT_COORDINATES_INVALID'));
  ['-180.1', '180.1'].forEach((lon) => assert.equal(code(() => parse(`<gpx version="1.1"><rte><rtept lat="35" lon="${lon}"/></rte></gpx>`)), 'GPX_POINT_COORDINATES_INVALID'));
  assert.equal(code(() => parse('<gpx version="1.1"><trk><trkseg><trkpt lon="139"/></trkseg></trk></gpx>')), 'GPX_POINT_COORDINATES_INVALID');
  assert.equal(code(() => parse(fixture('gpx-invalid-point.gpx'))), 'GPX_POINT_COORDINATES_INVALID');
});

test('elevation is direct single strict decimal while extension values are ignored', () => {
  const result = parse('<gpx version="1.1"><trk><trkseg>'
    + '<trkpt lat="1" lon="2"><ele>-0</ele></trkpt>'
    + '<trkpt lat="3" lon="4"><extensions><ele>999</ele></extensions></trkpt>'
    + '</trkseg></trk></gpx>');
  assert.equal(Object.is(result.segments[0].points[0].elevation, -0), true);
  assert.equal(result.segments[0].points[1].elevation, null);
  ['', '1e2', 'Infinity'].forEach((ele) => assert.equal(code(() => parse(`<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"><ele>${ele}</ele></trkpt></trkseg></trk></gpx>`)), 'GPX_POINT_ELEVATION_INVALID'));
  assert.equal(code(() => parse('<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"><ele>1</ele><ele>2</ele></trkpt></trkseg></trk></gpx>')), 'GPX_POINT_ELEVATION_INVALID');
});

test('point time accepts the shared contract, stays ordered, and ignores extension time', () => {
  const result = parse('<gpx version="1.1"><trk><trkseg>'
    + '<trkpt lat="1" lon="2"><time>2026-07-11T02:00:00.123456789Z</time></trkpt>'
    + '<trkpt lat="3" lon="4"><time>2026-07-11T01:00:00+14:00</time></trkpt>'
    + '<trkpt lat="5" lon="6"><extensions><time>2026-01-01T00:00:00Z</time></extensions></trkpt>'
    + '</trkseg></trk></gpx>');
  assert.deepEqual(plain(result.segments[0].points.map((point) => [point.lat, point.time])), [
    [1, '2026-07-11T02:00:00.123Z'], [3, '2026-07-10T11:00:00.000Z'], [5, '']
  ]);
  ['2026-07-11T00:00:00', '2026-07-11T00:00:00+14:01', '2026-02-30T00:00:00Z', ''].forEach((time) => {
    assert.equal(code(() => parse(`<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"><time>${time}</time></trkpt></trkseg></trk></gpx>`)), 'GPX_POINT_TIME_INVALID');
  });
  assert.equal(code(() => parse('<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"><time>2026-01-01T00:00:00Z</time><time>2026-01-01T00:00:01Z</time></trkpt></trkseg></trk></gpx>')), 'GPX_POINT_TIME_INVALID');
});

test('empty trkseg trk and rte are warnings but no adopted segment rejects', () => {
  const result = parse('<gpx version="1.1"><trk/><trk><trkseg/><trkseg><trkpt lat="1" lon="2"/></trkseg></trk><rte/></gpx>');
  assert.equal(result.stats.ignoredEmptySegmentCount, 3);
  assert.deepEqual(plain(result.warnings.filter((warning) => warning.code === 'GPX_EMPTY_SEGMENTS_IGNORED')), [
    { code: 'GPX_EMPTY_SEGMENTS_IGNORED', count: 3 }
  ]);
  assert.equal(code(() => parse(fixture('gpx-waypoints-only.gpx'))), 'GPX_TRACK_REQUIRED');
  assert.equal(code(() => parse('<gpx version="1.1"><trk><trkseg/></trk></gpx>')), 'GPX_TRACK_REQUIRED');
});

test('segment and point limits reject the whole GPX without truncation', () => {
  const exactSegments = ('<trkseg>' + '<trkpt lat="1" lon="2"/>'.repeat(100) + '</trkseg>').repeat(200);
  const exact = parse(`<gpx version="1.1"><trk>${exactSegments}</trk></gpx>`);
  assert.equal(exact.stats.segmentCount, 200);
  assert.equal(exact.stats.pointCount, 20000);
  const segments = '<trkseg><trkpt lat="1" lon="2"/></trkseg>'.repeat(201);
  assert.equal(code(() => parse(`<gpx version="1.1"><trk>${segments}</trk></gpx>`)), 'TRACK_SEGMENT_LIMIT_EXCEEDED');
  const points = '<trkpt lat="1" lon="2"/>'.repeat(20001);
  assert.equal(code(() => parse(`<gpx version="1.1"><trk><trkseg>${points}</trkseg></trk></gpx>`)), 'TRACK_POINT_LIMIT_EXCEEDED');
});

test('GPX batch removes exact timed duplicates before applying the saved point limit', () => {
  const point = trkptXml(35, 139, '2026-01-01T00:00:00Z');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${point.repeat(20001)}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(batch.stats.sourcePointCount, 20001);
  assert.equal(batch.stats.duplicatePointCount, 20000);
  assert.equal(batch.drafts.length, 1);
  assert.equal(batch.drafts[0].summary.pointCount, 1);
});

test('GPX duplicate removal is exact timed and scoped to its source segment', () => {
  const same = trkptXml(35, 139, '2026-01-01T00:00:00Z');
  const noTime = '<trkpt lat="35" lon="139"><ele>10</ele></trkpt>';
  const differentElevation = trkptXml(35, 139, '2026-01-01T00:00:00Z', 11);
  const xml = '<gpx version="1.1"><trk>'
    + `<trkseg>${same}${noTime}${same}${noTime}${differentElevation}</trkseg>`
    + `<trkseg>${same}</trkseg></trk></gpx>`;
  const batch = loadCore().buildDraftBatch(xml, options());
  assert.equal(batch.stats.duplicatePointCount, 1);
  assert.deepEqual(plain(batch.drafts[0].segments.map((segment) => segment.points.length)), [4, 1]);
});

test('GPX batch splits a four-hour interruption but not a midnight crossing', () => {
  const continuous = [
    trkptXml(35, 139, '2026-01-01T23:59:59Z'),
    trkptXml(35.001, 139.001, '2026-01-02T00:00:01Z')
  ].join('');
  assert.equal(loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${continuous}</trkseg></trk></gpx>`,
    options()
  ).drafts.length, 1);

  const interrupted = [
    trkptXml(35, 139, '2026-01-01T18:00:00Z'),
    trkptXml(35.001, 139.001, '2026-01-01T21:59:59Z'),
    trkptXml(35.002, 139.002, '2026-01-02T02:00:00Z')
  ].join('');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${interrupted}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(batch.stats.interruptionCount, 1);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), ['walk(1/2)', 'walk(2/2)']);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.summary.pointCount)), [2, 1]);
});

test('GPX batch chooses the smallest whole-second interval that fits 20000 points', () => {
  const start = Date.parse('2026-01-01T00:00:00Z');
  const points = Array.from({ length: 20001 }, (_, index) =>
    trkptXml(35 + index / 1000000, 139, new Date(start + index * 1000).toISOString())
  ).join('');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${points}</trkseg></trk></gpx>`,
    options()
  );
  assert.deepEqual(plain(batch.stats.compressionIntervals), [2]);
  assert.equal(batch.drafts.length, 1);
  assert.equal(batch.drafts[0].summary.pointCount, 10001);
  assert.equal(batch.drafts[0].segments[0].points[0].time, '2026-01-01T00:00:00.000Z');
  assert.equal(batch.drafts[0].segments[0].points.at(-1).time,
    new Date(start + 20000 * 1000).toISOString());
});

test('GPX batch partitions a stage that still exceeds 20000 points without losing sampled points', () => {
  const start = Date.parse('2026-01-01T00:00:00Z');
  const points = Array.from({ length: 60001 }, (_, index) =>
    `<trkpt lat="35" lon="139"><time>${
      new Date(start + index * 2000).toISOString()
    }</time></trkpt>`
  ).join('');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${points}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(batch.drafts.length, 2);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.name)), ['walk(1/2)', 'walk(2/2)']);
  assert.deepEqual(plain(batch.drafts.map((draft) => draft.summary.pointCount)), [20000, 1]);
  assert.equal(batch.stats.compressedPointCount, 40000);
  assert.deepEqual(plain(batch.stats.compressionIntervals), [5]);
  const savedTimes = batch.drafts.flatMap((draft) =>
    draft.segments.flatMap((segment) => segment.points.map((point) => point.time)));
  assert.equal(savedTimes.length, 20001);
  assert.equal(new Set(savedTimes).size, 20001);
});

test('GPX batch uniformly reduces untimed overflow without claiming a time interval', () => {
  const points = Array.from({ length: 20001 }, (_, index) =>
    `<trkpt lat="${35 + index / 1000000}" lon="139"/>`
  ).join('');
  const batch = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${points}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(batch.drafts.length, 1);
  assert.equal(batch.drafts[0].summary.pointCount, 20000);
  assert.equal(batch.drafts[0].segments[0].points[0].lat, 35);
  assert.equal(batch.drafts[0].segments[0].points.at(-1).lat, 35.02);
  assert.equal(batch.stats.compressedPointCount, 1);
  assert.deepEqual(plain(batch.stats.compressionIntervals), []);
});

test('GPX batch rejects source points and generated tracks at their independent limits', () => {
  const root = new ElementNode('gpx', '', { version: '1.1' });
  const track = new ElementNode('trk', '', {});
  const segment = new ElementNode('trkseg', '', {});
  segment.childNodes = Array.from({ length: 100001 }, () =>
    new ElementNode('trkpt', '', { lat: '35', lon: '139' }));
  track.childNodes.push(segment);
  root.childNodes.push(track);
  assert.equal(code(() => loadCore().buildDraftBatch('<gpx version="1.1"/>', {
    sourceName: 'walk.gpx',
    parseXml: () => new DocumentNode(root)
  })), 'GPX_SOURCE_POINT_LIMIT_EXCEEDED');

  const start = Date.parse('2026-01-01T00:00:00Z');
  const interrupted = (count) => Array.from({ length: count }, (_, index) =>
    trkptXml(35, 139, new Date(start + index * 4 * 60 * 60 * 1000).toISOString())
  ).join('');
  const accepted = loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${interrupted(20)}</trkseg></trk></gpx>`,
    options()
  );
  assert.equal(accepted.drafts.length, 20);
  assert.equal(code(() => loadCore().buildDraftBatch(
    `<gpx version="1.1"><trk><trkseg>${interrupted(21)}</trkseg></trk></gpx>`,
    options()
  )), 'GPX_GENERATED_TRACK_LIMIT_EXCEEDED');
});

test('GPX batch common edits regenerate suffixes and save independent payloads', () => {
  const core = loadCore();
  const xml = '<gpx version="1.1"><trk><trkseg>'
    + trkptXml(35, 139, '2026-01-01T18:00:00Z')
    + trkptXml(35.1, 139.1, '2026-01-02T02:00:00Z')
    + '</trkseg></trk></gpx>';
  const batch = core.buildDraftBatch(xml, options());
  const updated = core.updateDraftBatch(batch, {
    name: '縦走',
    description: '2泊3日',
    color: '#2196f3',
    visible: false,
    lineStyle: 'dashed'
  });
  assert.deepEqual(plain(updated.drafts.map((draft) => draft.name)), ['縦走(1/2)', '縦走(2/2)']);
  assert.ok(updated.drafts.every((draft) =>
    draft.description === '2泊3日' && draft.color === '#2196f3'
    && draft.visible === false && draft.lineStyle === 'dashed'));
  const payloads = core.toSavePayloads(updated);
  assert.equal(payloads.length, 2);
  payloads[0].segments[0].points[0].lat = 0;
  assert.notEqual(payloads[1].segments[0].points[0].lat, 0);
  assert.notEqual(updated.drafts[0].segments[0].points[0].lat, 0);
});

test('metadata priority uses direct children and enforces safe length boundaries', () => {
  const xml = '<gpx version="1.1"><metadata><name> Meta </name><desc>  keep me  </desc><extensions><name>bad</name></extensions></metadata>'
    + '<name>Root</name><desc>Root desc</desc><extensions><name>bad root extension</name></extensions>'
    + '<rte><name>Route</name><desc>Route desc</desc><rtept lat="1" lon="2"/></rte></gpx>';
  const result = parse(xml, { sourceName: '/tmp/file.gpx' });
  assert.equal(result.name, 'Meta');
  assert.equal(result.description, '  keep me  ');
  assert.equal(parse('<gpx version="1.0"><name>Root</name><trk><name>Track</name><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>').name, 'Root');
  assert.equal(parse('<gpx version="1.1"><rte><name>Route</name><rtept lat="1" lon="2"/></rte></gpx>').name, 'Route');
  assert.equal(parse('<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>', { sourceName: 'folder\\fallback.gpx' }).name, 'fallback');
  assert.equal(parse('<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>', { sourceName: '' }).name, 'GPXトラック');
  assert.equal(code(() => parse(`<gpx version="1.1"><metadata><name>${'a'.repeat(101)}</name></metadata><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>`)), 'GPX_METADATA_INVALID');
  assert.equal(code(() => parse(`<gpx version="1.1"><metadata><desc>${'a'.repeat(401)}</desc></metadata><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>`)), 'GPX_METADATA_INVALID');
});

test('sourceName keeps only a safe basename and supports an empty value', () => {
  const xml = '<gpx version="1.1"><trk><trkseg><trkpt lat="1" lon="2"/></trkseg></trk></gpx>';
  assert.equal(parse(xml, { sourceName: '/private/path/track.gpx' }).sourceName, 'track.gpx');
  assert.equal(parse(xml, { sourceName: 'C:\\private\\track.gpx' }).sourceName, 'track.gpx');
  assert.equal(parse(xml, { sourceName: '' }).sourceName, '');
  assert.equal(code(() => parse(xml, { sourceName: `bad\nname.gpx` })), 'GPX_SOURCE_NAME_INVALID');
  assert.equal(code(() => parse(xml, { sourceName: `${'a'.repeat(197)}.gpx` })), 'GPX_SOURCE_NAME_INVALID');
});

test('stats and warnings contain only safe counts and are excluded from save payload', () => {
  const result = build(fixture('gpx-1.0-track.gpx'));
  assert.deepEqual(plain(result.stats), {
    trackElementCount: 1, routeElementCount: 0, segmentCount: 1, pointCount: 2,
    timedPointCount: 1, elevatedPointCount: 1, ignoredWaypointCount: 0, ignoredEmptySegmentCount: 0
  });
  assert.deepEqual(plain(result.warnings), [
    { code: 'GPX_POINTS_WITHOUT_TIME', count: 1 },
    { code: 'GPX_POINTS_WITHOUT_ELEVATION', count: 1 }
  ]);
  assert.equal(JSON.stringify(result.warnings).includes('35.'), false);
  const payload = loadCore().toSavePayload(result);
  assert.equal(payload.stats, undefined);
  assert.equal(payload.warnings, undefined);
});

test('buildDraft creates identity once and updateDraft changes only editable metadata', () => {
  const core = loadCore();
  const xml = fixture('gpx-route.gpx');
  const first = core.buildDraft(xml, options());
  const second = core.buildDraft(xml, options());
  assert.equal(first.trackId, 'uuid-1');
  assert.equal(first.id, first.trackId);
  assert.equal(first.revisionId, 'uuid-2');
  assert.notEqual(second.trackId, first.trackId);
  assert.ok(first.summary.distanceMeters > 0);
  const beforeSegments = JSON.stringify(first.segments);
  const updated = core.updateDraft(first, {
    name: 'Updated', description: 'D', color: '#2196f3', visible: false, lineStyle: 'dashed', lineWidth: 7,
    trackId: 'bad', revisionId: 'bad', sourceType: 'bad', sourceName: 'bad', segments: [], stats: {}, warnings: []
  });
  assert.equal(updated.name, 'Updated');
  assert.equal(updated.trackId, first.trackId);
  assert.equal(updated.revisionId, first.revisionId);
  assert.equal(updated.sourceType, 'gpx');
  assert.equal(updated.lineWidth, 4);
  assert.equal(JSON.stringify(updated.segments), beforeSegments);
  assert.deepEqual(plain(updated.stats), plain(first.stats));
  assert.deepEqual(plain(updated.warnings), plain(first.warnings));
});

test('identity generation starts only after the entire GPX has parsed successfully', () => {
  let calls = 0;
  const generateId = () => `generated-${++calls}`;
  assert.equal(code(() => loadCore().buildDraft(fixture('gpx-invalid-point.gpx'), options({ generateId }))),
    'GPX_POINT_COORDINATES_INVALID');
  assert.equal(calls, 0);
  const draft = loadCore().buildDraft(fixture('gpx-route.gpx'), options({ generateId }));
  assert.equal(calls, 2);
  assert.equal(draft.trackId, 'generated-1');
  assert.equal(draft.revisionId, 'generated-2');
});

test('toSavePayload is a deep whitelist without orderIndex or transient data', () => {
  const core = loadCore();
  const draft = core.buildDraft(fixture('gpx-route.gpx'), options());
  draft.orderIndex = 999;
  draft.File = { secret: true };
  draft.xml = '<private/>';
  draft.creator = 'drop';
  draft.extensions = { drop: true };
  draft.token = 'drop';
  const payload = core.toSavePayload(draft);
  assert.deepEqual(Object.keys(payload).sort(), [
    'color', 'description', 'lineStyle', 'lineWidth', 'name', 'revisionId', 'segments',
    'sourceName', 'sourceType', 'trackId', 'visible'
  ]);
  assert.equal(payload.trackId, draft.trackId);
  assert.equal(payload.revisionId, draft.revisionId);
  assert.equal(payload.lineWidth, 4);
  payload.segments[0].points[0].lat = 0;
  assert.notEqual(draft.segments[0].points[0].lat, 0);
});

test('default parser uses browser DOMParser and invalid parser options stay safe', () => {
  let calls = 0;
  class FakeDOMParser {
    parseFromString(text, type) {
      calls += 1;
      assert.equal(type, 'application/xml');
      return parseFixtureXml(text);
    }
  }
  const core = loadCore({ DOMParser: FakeDOMParser });
  const result = core.parse(fixture('gpx-route.gpx'), { sourceName: 'route.gpx' });
  assert.equal(result.stats.routeElementCount, 1);
  assert.equal(calls, 1);
  assert.equal(code(() => loadCore().parse(fixture('gpx-route.gpx'), { sourceName: 'route.gpx' })), 'GPX_INVALID_XML');
  assert.equal(code(() => loadCore().parse(fixture('gpx-route.gpx'), { sourceName: 'route.gpx', parseXml: true })), 'GPX_INVALID_XML');
});

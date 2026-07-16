const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');
const importUrlVectors = require('./fixtures/import-url-vectors');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const EXPECTED_COLORS = [
  ['#e53935', 'red', '赤'], ['#e91e63', 'pink', 'ピンク'],
  ['#9c27b0', 'purple', '紫'], ['#3f51b5', 'indigo', '藍'],
  ['#2196f3', 'blue', '青'], ['#00bcd4', 'cyan', '水色'],
  ['#009688', 'teal', 'ティール'], ['#4caf50', 'green', '緑'],
  ['#8bc34a', 'lime', '黄緑'], ['#ffeb3b', 'yellow', '黄'],
  ['#ff9800', 'orange', '橙'], ['#ff5722', 'deep-orange', '朱'],
  ['#795548', 'brown', '茶'], ['#607d8b', 'gray', 'グレー'],
  ['#212121', 'black', '黒']
];
const EXPECTED_ICONS = [
  ['default', '標準'], ['photo', '写真'], ['food', '食事'], ['hotel', '宿'],
  ['nature', '自然'], ['shop', '店'], ['transit', '交通'], ['warning', '注意']
];

function loadPinDefinitions() {
  const start = indexHtml.indexOf('const PIN_COLORS = [');
  const end = indexHtml.indexOf('const PIN_STATUSES =', start);
  assert.notEqual(start, -1, 'Expected PIN_COLORS');
  assert.notEqual(end, -1, 'Expected PIN_STATUSES after pin definitions');
  const context = {};
  vm.runInNewContext(
    `${indexHtml.slice(start, end)}\nthis.colors = PIN_COLORS; this.icons = PIN_ICONS;`,
    context
  );
  return { colors: plain(context.colors), icons: plain(context.icons) };
}

const { colors: PIN_COLORS, icons: PIN_ICONS } = loadPinDefinitions();

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function loadCore(extra = {}) {
  const start = indexHtml.indexOf('const ImportJobCore = (function() {');
  const end = indexHtml.indexOf('\n    const ImportQueueRunner = (function() {', start);
  assert.notEqual(start, -1, 'Expected ImportJobCore');
  assert.notEqual(end, -1, 'Expected import queue after CSV core');
  const context = {
    PIN_COLORS,
    PIN_ICONS,
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
      + 'globalThis.__csvCore = typeof CsvPinInterchangeCore === "undefined" ? null : CsvPinInterchangeCore;',
    context
  );
  return { csv: context.__csvCore, jobs: context.__jobCore };
}

function csvText(lines, bom = false, newline = '\n') {
  return (bom ? '\uFEFF' : '') + lines.join(newline);
}

test('CSV parser supports BOM, line endings, quotes, empty cells, blank rows, Japanese, and emoji', () => {
  const { csv } = loadCore();
  assert.ok(csv, 'Expected CsvPinInterchangeCore');
  const input = '\uFEFFa,b,c,d\r\n'
    + '"カンマ,あり","改行\r\nあり","引用符""あり",\r\n'
    + '\r\n'
    + '絵文字🙂,末尾,,\r\n';
  assert.deepEqual(plain(csv.parseCsv(input)), [
    ['a', 'b', 'c', 'd'],
    ['カンマ,あり', '改行\r\nあり', '引用符"あり', ''],
    [''],
    ['絵文字🙂', '末尾', '', '']
  ]);
  assert.deepEqual(plain(csv.parseCsv(csvText(['a,b', '1,2'], false, '\n'))), [
    ['a', 'b'], ['1', '2']
  ]);
});

test('CSV parser rejects unsafe or malformed whole-file input', () => {
  const { csv } = loadCore();
  for (const value of ['', '\uFEFF', 'a,"unterminated', 'a,"ok"x', 'a,\0b']) {
    assert.throws(() => csv.parseCsv(value), (error) => /^CSV_/.test(error.code));
  }
});

test('CSV stringifier is immutable, uses BOM and CRLF, and quotes cells safely', () => {
  const { csv } = loadCore();
  const rows = [['a', 'b,c', 'line\nbreak', 'quote"here', null, undefined, 3]];
  const before = JSON.stringify(rows);
  assert.equal(
    csv.stringifyCsv(rows),
    '\uFEFFa,"b,c","line\nbreak","quote""here",,,3'
  );
  assert.equal(JSON.stringify(rows), before);
});

test('pin serialization uses the fixed schema, preserves order, and round trips exchange data', () => {
  const { csv } = loadCore({ crypto: { randomUUID: () => 'roundtrip-item' } });
  const pins = [{
    id: 'pin-1', title: '東京,「観察」', description: '一行目\n二行目"引用"',
    lat: -35.5, lng: 139.75, color: '#2196F3', icon: 'photo', status: '',
    tags: ['植物', '観察'], eventAt: '2026-07-11T10:30:15',
    links: ['https://example.com', 'https://example.org']
  }, {
    id: 'pin-2', title: '未配置🙂', description: '', lat: null, lng: null,
    color: '#e53935', icon: 'default', status: '完了', tags: [], eventAt: '', links: []
  }];
  const output = csv.serializePins(pins);
  assert.equal(output.startsWith('\uFEFFschemaVersion,sourceId,title,description,lat,lng,color,icon,status,tags,eventAt,links\r\n'), true);
  const job = csv.buildImportJob(output, { jobId: 'job-1', generateId: (() => {
    let value = 0;
    return () => `item-${++value}`;
  })() });
  assert.deepEqual(plain(job.items.map((item) => ({
    sourceRef: item.sourceRef, title: item.title, description: item.description,
    lat: item.lat, lng: item.lng, color: item.color, icon: item.icon,
    status: item.status, tags: item.tags, capturedAt: item.capturedAt, links: item.links
  }))), [
    {
      sourceRef: 'CSV 2行目 / pin-1', title: '東京,「観察」', description: '一行目\n二行目"引用"',
      lat: -35.5, lng: 139.75, color: '#2196f3', icon: 'photo', status: '',
      tags: ['植物', '観察'], capturedAt: '2026-07-11T10:30:15',
      links: ['https://example.com', 'https://example.org']
    },
    {
      sourceRef: 'CSV 3行目 / pin-2', title: '未配置🙂', description: '', lat: null, lng: null,
      color: '#e53935', icon: 'default', status: '完了', tags: [], capturedAt: '', links: []
    }
  ]);
  assert.equal(job.sourceType, 'csv');
});

test('all standard colors export as fixed English names and re-import to the same lowercase hex', () => {
  const { csv } = loadCore();
  assert.deepEqual(PIN_COLORS.map((color) => [color.hex, color.csvName, color.label]), EXPECTED_COLORS);
  const pins = PIN_COLORS.map((color, index) => ({
    id: `pin-${index + 1}`, title: color.label, color: color.hex.toUpperCase(),
    icon: 'default', tags: [], links: []
  }));
  const output = csv.serializePins(pins);
  const rows = csv.parseCsv(output);
  assert.deepEqual(plain(rows.slice(1).map((row) => row[6])), PIN_COLORS.map((color) => color.csvName));
  assert.deepEqual(
    plain(csv.buildImportJob(output).items.map((item) => item.color)),
    PIN_COLORS.map((color) => color.hex)
  );
});

test('non-standard safe six-digit hex exports lowercase and round trips unchanged', () => {
  const { csv } = loadCore();
  const output = csv.serializePins([{
    id: 'custom', title: 'Custom', color: '#12AB9F', icon: 'default', tags: [], links: []
  }]);
  assert.equal(csv.parseCsv(output)[1][6], '#12ab9f');
  const item = csv.buildImportJob(output).items[0];
  assert.equal(item.uploadStatus, 'queued');
  assert.equal(item.color, '#12ab9f');
});

test('color English names, Japanese names, aliases, and hex normalize to lowercase six-digit hex', () => {
  const { csv } = loadCore();
  const values = [
    ['orange', '#ff9800'], ['ORANGE', '#ff9800'], ['橙', '#ff9800'],
    ['#FF9800', '#ff9800'], ['#12AB9F', '#12ab9f'],
    ['gray', '#607d8b'], ['GREY', '#607d8b'], ['グレー', '#607d8b']
  ];
  values.forEach(([value, expected]) => {
    const input = csv.stringifyCsv([
      ['schemaVersion', 'title', 'color'], ['1', value, value]
    ]);
    const item = csv.buildImportJob(input).items[0];
    assert.equal(item.uploadStatus, 'queued', value);
    assert.equal(item.color, expected, value);
  });
});

test('unknown colors plus three- and eight-digit hex fail with row-numbered errors', () => {
  const { csv } = loadCore();
  const input = csv.stringifyCsv([
    ['schemaVersion', 'title', 'color'],
    ['1', 'Unknown', 'sunset'],
    ['1', 'Short', '#fff'],
    ['1', 'Alpha', '#ff9800ff']
  ]);
  const items = csv.buildImportJob(input).items;
  items.forEach((item, index) => {
    assert.equal(item.uploadStatus, 'failed');
    assert.equal(item.errorCode, 'CSV_ROW_COLOR_INVALID');
    assert.match(item.error, new RegExp(`CSV ${index + 2}行目:`));
  });
});

test('icon English ids and Japanese names normalize case-insensitively to English ids', () => {
  const { csv } = loadCore();
  const values = [
    ['photo', 'photo'], ['PHOTO', 'photo'], ['写真', 'photo'],
    ['default', 'default'], ['標準', 'default'], ['注意', 'warning']
  ];
  values.forEach(([value, expected]) => {
    const input = csv.stringifyCsv([
      ['schemaVersion', 'title', 'icon'], ['1', value, value]
    ]);
    const item = csv.buildImportJob(input).items[0];
    assert.equal(item.uploadStatus, 'queued', value);
    assert.equal(item.icon, expected, value);
  });
});

test('unknown icon names fail with a row-numbered error', () => {
  const { csv } = loadCore();
  const item = csv.buildImportJob('schemaVersion,title,icon\n1,Pin,camera').items[0];
  assert.equal(item.uploadStatus, 'failed');
  assert.equal(item.errorCode, 'CSV_ROW_ICON_INVALID');
  assert.match(item.error, /CSV 2行目:/);
});

test('reference CSV has a BOM, 15 colors, 8 icons, and an empty fourth column on every row', () => {
  const { csv } = loadCore();
  assert.deepEqual(PIN_ICONS.map((icon) => [icon.id, icon.label]), EXPECTED_ICONS);
  const output = csv.serializeReference();
  assert.equal(output.charAt(0), '\uFEFF');
  const rows = csv.parseCsv(output);
  assert.equal(rows.length, 16);
  assert.deepEqual(plain(rows[0]), [
    'カラーコード', '色英名', '色和名', '', 'アイコン英名', 'アイコン和名'
  ]);
  assert.deepEqual(plain(rows[1]), ['#e53935', 'red', '赤', '', 'default', '標準']);
  assert.deepEqual(plain(rows[8].slice(4)), ['warning', '注意']);
  assert.deepEqual(plain(rows[9].slice(4)), ['', '']);
  assert.equal(rows.every((row) => row.length === 6 && row[3] === ''), true);
  assert.deepEqual(plain(rows.slice(1).map((row) => row.slice(0, 3))), PIN_COLORS.map((color) => [
    color.hex, color.csvName, color.label
  ]));
  assert.deepEqual(plain(rows.slice(1, 9).map((row) => row.slice(4))), PIN_ICONS.map((icon) => [
    icon.id, icon.label
  ]));
});

test('empty pin serialization is a header-only CSV that imports as an empty Point exchange', () => {
  const { csv } = loadCore();
  const output = csv.serializePins([]);
  assert.equal(
    output,
    '\uFEFFschemaVersion,sourceId,title,description,lat,lng,color,icon,status,tags,eventAt,links'
  );
  assert.deepEqual(plain(csv.buildImportJob(output)), {
    sourceType: 'csv', items: [], warnings: [], empty: true
  });
  assert.deepEqual(plain(csv.buildImportJob(output + '\r\n\r\n')), {
    sourceType: 'csv', items: [], warnings: [], empty: true
  });
});

test('formula injection protection covers user text and restores only its own apostrophe', () => {
  const { csv } = loadCore();
  const dangerous = ['=1+1', '+cmd', '-1+2', '@SUM(A1:A2)', '\tcalc', '\rcalc'];
  dangerous.forEach((title) => {
    const output = csv.serializePins([{
      id: title, title, description: title, lat: -1, lng: -2,
      color: '#e53935', icon: 'default', tags: [], links: []
    }]);
    const rows = csv.parseCsv(output);
    assert.equal(rows[1][1], `'${title}`);
    assert.equal(rows[1][2], `'${title}`);
    assert.equal(rows[1][3], `'${title}`);
    assert.equal(rows[1][4], '-1');
    const item = csv.buildImportJob(output).items[0];
    assert.equal(item.title, title);
    assert.equal(item.description, title);
  });
  const ordinary = csv.serializePins([{
    id: "'source", title: "'文章", description: "'通常", color: '#e53935',
    icon: 'default', tags: [], links: []
  }]);
  const ordinaryItem = csv.buildImportJob(ordinary).items[0];
  assert.equal(ordinaryItem.title, "'文章");
  assert.equal(ordinaryItem.description, "'通常");
});

test('headers are normalized, required and unique while unknown columns are ignored', () => {
  const { csv } = loadCore();
  const job = csv.buildImportJob(' TITLE , SCHEMAVERSION ,unknown\nA,1,ignored');
  assert.equal(job.items[0].title, 'A');
  assert.deepEqual(plain(job.warnings), ['不明な列を無視しました: unknown']);
  assert.throws(
    () => csv.buildImportJob('schemaVersion,sourceId\n1,x'),
    (error) => error.code === 'CSV_REQUIRED_HEADER_MISSING'
  );
  assert.throws(
    () => csv.buildImportJob('schemaVersion,title,TITLE\n1,A,B'),
    (error) => error.code === 'CSV_DUPLICATE_HEADER'
  );
});

test('buildImportJob accepts zero through MAX_ITEMS rows and rejects all 21 before adoption', () => {
  const { csv, jobs } = loadCore();
  assert.equal(jobs.MAX_ITEMS, 20);
  assert.equal(csv.buildImportJob('schemaVersion,title\n\n').empty, true);
  const twenty = ['schemaVersion,title'].concat(
    Array.from({ length: jobs.MAX_ITEMS }, (_, index) => `1,Pin ${index + 1}`)
  ).join('\n');
  assert.equal(csv.buildImportJob(twenty).items.length, jobs.MAX_ITEMS);
  const twentyOne = twenty + '\n1,Pin 21';
  assert.throws(
    () => csv.buildImportJob(twentyOne),
    (error) => error.code === 'IMPORT_ITEM_LIMIT_EXCEEDED'
  );
});

test('row conversion preserves order and generates unique ids independent of sourceId', () => {
  const { csv } = loadCore({ crypto: { randomUUID: () => 'same-id' } });
  const job = csv.buildImportJob([
    'schemaVersion,sourceId,title',
    '1,same,First',
    '1,same,Second'
  ].join('\n'), { jobId: 'same-id' });
  assert.deepEqual(plain(job.items.map((item) => item.title)), ['First', 'Second']);
  assert.equal(new Set(job.items.map((item) => item.id)).size, 2);
  assert.notEqual(job.items[0].id, 'same');
  assert.notEqual(job.items[1].id, 'same');
  assert.deepEqual(plain(job.items.map((item) => item.sourceRef)), [
    'CSV 2行目 / same', 'CSV 3行目 / same'
  ]);
});

test('invalid values remain isolated non-retryable failed rows with safe codes', () => {
  const { csv } = loadCore();
  const rows = csv.stringifyCsv([
    ['schemaVersion', 'title', 'description', 'lat', 'lng', 'color', 'icon', 'status', 'tags', 'eventAt', 'links'],
    ['1', 'Valid', '', '35', '139', '#E53935', 'default', '', '["植物","観察"]', '2026-07-11T10:30', '["https://example.com"]'],
    ['2', 'Bad version'],
    ['1', '', 'too short'],
    ['1', 'a'.repeat(81)],
    ['1', 'Description', 'd'.repeat(401)],
    ['1', 'Coordinates', '', '35', ''],
    ['1', 'Color', '', '', '', 'unknown-color'],
    ['1', 'Icon', '', '', '', '', 'unknown'],
    ['1', 'Status', '', '', '', '', '', 'unknown'],
    ['1', 'Tags', '', '', '', '', '', '', 'a|b|c|d|e|f'],
    ['1', 'Date', '', '', '', '', '', '', '', '2026-02-30T10:00'],
    ['1', 'URL', '', '', '', '', '', '', '', '', 'https://ok.example|javascript:bad']
  ]);
  const job = csv.buildImportJob(rows);
  assert.equal(job.items[0].uploadStatus, 'queued');
  assert.equal(job.items[0].color, '#e53935');
  assert.deepEqual(plain(job.items[0].tags), ['植物', '観察']);
  assert.deepEqual(plain(job.items[0].links), ['https://example.com']);
  job.items.slice(1).forEach((item, index) => {
    assert.equal(item.uploadStatus, 'failed', `row ${index + 3}`);
    assert.equal(item.retryable, false);
    assert.equal(item.attempts, 0);
    assert.match(item.errorCode, /^CSV_ROW_/);
    assert.match(item.error, new RegExp(`CSV ${index + 3}行目:`));
    assert.deepEqual(plain(item.runtime), {
      originalFile: null, uploadFile: null, previewUrl: ''
    });
  });
});

test('row validation accepts documented boundary values and defaults', () => {
  const { csv } = loadCore();
  const output = csv.stringifyCsv([
    ['schemaVersion', 'title', 'description', 'lat', 'lng', 'color', 'icon', 'status', 'tags', 'eventAt', 'links'],
    [
      '1', 't'.repeat(80), 'd'.repeat(400), '-90', '180', '', '', '',
      '["A","B","C","D","E"]', '2024-02-29T23:59:59',
      'https://one.example\nhttps://two.example'
    ],
    ['1', 'Unplaced', '', '', '', '#2196f3', 'photo', '保留', '', '', '']
  ]);
  const items = csv.buildImportJob(output).items;
  assert.equal(items[0].uploadStatus, 'queued');
  assert.equal(items[0].title.length, 80);
  assert.equal(items[0].description.length, 400);
  assert.equal(items[0].lat, -90);
  assert.equal(items[0].lng, 180);
  assert.equal(items[0].color, '#e53935');
  assert.equal(items[0].icon, 'default');
  assert.deepEqual(plain(items[0].tags), ['A', 'B', 'C', 'D', 'E']);
  assert.deepEqual(plain(items[0].links), ['https://one.example', 'https://two.example']);
  assert.equal(items[1].lat, null);
  assert.equal(items[1].lng, null);
  assert.equal(items[1].status, '保留');
});

test('tags and links support alternate input, preserve order, and deduplicate safely', () => {
  const { csv } = loadCore();
  const inputs = [
    ['#植物 #観察 #植物', ['植物', '観察']],
    ['植物,観察,植物', ['植物', '観察']],
    ['植物|観察|植物', ['植物', '観察']]
  ];
  inputs.forEach(([value, expected]) => {
    const item = csv.buildImportJob(`schemaVersion,title,tags\n1,Pin,"${value}"`).items[0];
    assert.deepEqual(plain(item.tags), expected);
  });
  const links = csv.buildImportJob(
    'schemaVersion,title,links\n1,Pin,"https://a.example\nhttps://b.example\nhttps://a.example"'
  ).items[0].links;
  assert.deepEqual(plain(links), ['https://a.example', 'https://b.example']);
});

test('CSV link validation uses the shared URL vectors with and without the URL class', () => {
  [URL, undefined].forEach((UrlCtor) => {
    const { csv } = loadCore({ URL: UrlCtor });
    importUrlVectors.allowed.forEach((link) => {
      const value = csv.stringifyCsv([['schemaVersion', 'title', 'links'], ['1', 'Pin', link]]);
      assert.equal(csv.buildImportJob(value).items[0].uploadStatus, 'queued', link);
    });
    importUrlVectors.rejected.forEach((link) => {
      const value = csv.stringifyCsv([['schemaVersion', 'title', 'links'], ['1', 'Pin', link]]);
      const item = csv.buildImportJob(value).items[0];
      assert.equal(item.uploadStatus, 'failed', link);
      assert.equal(item.errorCode, 'CSV_ROW_LINKS_INVALID', link);
    });
  });
});

test('CSV eventAt accepts real year and leap boundaries but rejects year zero', () => {
  const { csv } = loadCore();
  const values = [
    ['0001-01-01T00:00:00', 'queued'],
    ['2000-02-29T23:59:59', 'queued'],
    ['1900-02-29T00:00', 'failed'],
    ['0000-12-31T23:59', 'failed']
  ];
  values.forEach(([eventAt, status]) => {
    const input = csv.stringifyCsv([['schemaVersion', 'title', 'eventAt'], ['1', 'Pin', eventAt]]);
    assert.equal(csv.buildImportJob(input).items[0].uploadStatus, status, eventAt);
  });
});

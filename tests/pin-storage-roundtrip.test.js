const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
function harness(timeZone = 'Asia/Tokyo') {
  const rows = [];
  let timezoneReads = 0;
  const sheet = {
    getLastRow: () => rows.length,
    getDataRange: () => ({getValues: () => rows.map(row => row.slice())}),
    appendRow: row => rows.push(row.slice()),
    getRange: (r, c, n = 1, w = 1) => ({
      getValues: () => rows.slice(r - 1, r - 1 + n).map(row => row.slice(c - 1, c - 1 + w)),
      getFormulas: () => Array.from({length:n}, () => Array(w).fill('')),
      setValues: values => values.forEach((row, i) => row.forEach((v, j) => { rows[r - 1 + i][c - 1 + j] = v; }))
    })
  };
  const context = vm.createContext({
    Date, console,
    SpreadsheetApp: {getActiveSpreadsheet: () => ({
      getSpreadsheetTimeZone: () => {timezoneReads++; return timeZone;},
      getSheetByName: name => name === 'map_info' ? sheet : null
    }), flush() {}},
    LockService: {getScriptLock: () => ({tryLock:()=>true, releaseLock(){}})},
    CacheService: {getScriptCache: () => ({get: key => key === 'EDIT_TOKEN_test-token' ? '1' : null})},
    Session: {getScriptTimeZone: () => 'Etc/UTC'},
    Utilities: {
      getUuid: () => 'duplicate-local',
      formatDate: (date, zone, pattern) => {
        const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
          timeZone:zone, year:'numeric', month:'2-digit', day:'2-digit', hour:'2-digit', minute:'2-digit', second:'2-digit', hourCycle:'h23'
        }).formatToParts(date).map(part => [part.type, part.value]));
        return pattern.includes("'T'") ? `${p.year}-${p.month}-${p.day}T${p.hour}:${p.minute}:${p.second}` : `${p.year}/${p.month}/${p.day} ${p.hour}:${p.minute}:${p.second}`;
      }
    }
  });
  vm.runInContext(source + ';this.api={PinData,MAP_INFO_HEADERS,appendMapInfoRow_,updatePinDetails,duplicatePin,bulkUpdatePinMetadata,encodeSpreadsheetLiteral_};', context);
  const api = context.api;
  rows.push(Array.from(api.MAP_INFO_HEADERS));
  const row = Array(16).fill('');
  Object.assign(row, {0:'2026/09/09 14:00:00',1:'original',5:'#e53935',8:'pin-1',14:'default'});
  rows.push(row);
  return {api, rows, row, sheet, timezoneReads:()=>timezoneReads, edit:data=>api.updatePinDetails({__editToken:'test-token',id:'pin-1',title:'edited',...data})};
}

test('stored Date uses spreadsheet timezone across a day boundary and survives edit/read', () => {
  const h = harness();
  h.row[12] = new Date('2025-07-19T23:16:45Z');
  const pin = h.api.PinData.rowToPin(h.row);
  assert.equal(pin.eventAt, '2025-07-20T08:16:45');
  assert.equal(h.edit({eventAt:pin.eventAt}).ok, true);
  assert.equal(h.api.PinData.rowToPin(h.row).eventAt, pin.eventAt);
});

test('Date conversion honors spreadsheet DST instead of the script or machine timezone', () => {
  const h = harness('America/New_York');
  h.row[12] = new Date('2025-07-20T00:16:00Z');
  assert.equal(h.api.PinData.rowToPin(h.row).eventAt, '2025-07-19T20:16:00');
  h.row[12] = new Date('2025-01-20T00:16:00Z');
  assert.equal(h.api.PinData.rowToPin(h.row).eventAt, '2025-01-19T19:16:00');
  assert.equal(h.timezoneReads(), 1);
});

test('legacy text timestamps, blanks and invalid dates keep their validation behavior', () => {
  const h = harness();
  for (const [input, expected] of [['2024-02-29T12:30','2024-02-29T12:30'],['2025-02-29T12:30',''],['',''],[new Date(NaN),'']]) {
    h.row[12] = input;
    assert.equal(h.api.PinData.rowToPin(h.row).eventAt, expected);
  }
  assert.equal(h.timezoneReads(), 0);
});

test('editing formula-like text roundtrips as literal text in title, description and tags', () => {
  const h = harness();
  for (const text of ['=1+1','+example','-example','@example',"'example",'\u200Bdtp-sheet:v1:v:=example']) {
    assert.equal(h.edit({title:text,description:text,tags:[text]}).ok, true);
    const pin = h.api.PinData.rowToPin(h.row);
    assert.equal(pin.title, text);
    assert.equal(pin.description, text);
    assert.equal(pin.tags[0], text);
    for (const column of [1,2,11]) assert.equal(h.row[column].startsWith('='), false);
  }
});

test('duplicate and bulk tag editing retain literal encoding without leaking its marker', () => {
  const h = harness();
  for (const column of [1,2,11]) h.row[column] = h.api.encodeSpreadsheetLiteral_('=1+1');
  const duplicated = h.api.duplicatePin({__editToken:'test-token',sourcePinId:'pin-1',mode:'unplaced'});
  assert.equal(duplicated.ok, true);
  assert.equal(h.rows[2][1].startsWith('='), false);
  assert.equal(h.rows[2][2].startsWith('='), false);
  assert.equal(h.api.PinData.rowToPin(h.rows[2]).tags[0], '=1+1');
  const result = h.api.bulkUpdatePinMetadata({__editToken:'test-token',ids:['pin-1'],tagMode:'add',tags:['second']});
  assert.equal(result.ok, true);
  assert.deepEqual(Array.from(h.api.PinData.rowToPin(h.row).tags), ['=1+1','second']);
  assert.equal(h.row[11].startsWith('='), false);
});

test('duplicate response matches stored text even when user text resembles the codec marker', () => {
  const h = harness();
  const value = '\u200Bdtp-sheet:v1:v:=literal-user-text';
  for (const column of [1,2,11]) h.row[column] = h.api.encodeSpreadsheetLiteral_(value);
  const result = h.api.duplicatePin({__editToken:'test-token',sourcePinId:'pin-1',mode:'unplaced'});
  const loaded = h.api.PinData.rowToPin(h.rows[2]);
  assert.equal(result.pin.title, loaded.title);
  assert.equal(result.pin.description, value);
  assert.equal(result.pin.tags[0], value);
});

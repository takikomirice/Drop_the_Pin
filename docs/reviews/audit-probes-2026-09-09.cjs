// Historical audit probes for the 2026-09-09 findings, not a green regression suite.
// Assertions confirm the observed bugs; they must change when those bugs are fixed.
// Run from the repository root: node docs/reviews/audit-probes-2026-09-09.cjs
// Read-only audit probes: run production functions against local in-memory stubs.
// No Google API calls or real spreadsheet writes.
const fs = require('node:fs');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const code = fs.readFileSync('Code.js', 'utf8');
const context = vm.createContext({ console });
vm.runInContext(code, context);
const findings = vm.runInContext(`(() => {
  const result = {};
  const row = Array(MAP_INFO_HEADERS.length).fill('');
  row[0] = new Date('2026-09-09T05:02:27.000Z');
  row[1] = 'audit'; row[8] = 'audit-local-only';
  row[12] = new Date('2025-07-20T00:16:00.000Z');
  result.dateCell = { inputDate: row[12].toISOString(), outputEventAt: PinData.rowToPin(row).eventAt };
  row[12] = '2025-07-20T09:16';
  result.textCell = { outputEventAt: PinData.rowToPin(row).eventAt };
  row[12] = new Date('2025-07-20T00:16:00.000Z');
  let written;
  assertEditToken_ = () => {};
  withSpreadsheetMutationLock_ = fn => fn();
  findPinRowIndex_ = () => 1;
  getRenameFileWithTitle_ = () => false;
  currentUpdatedAt_ = () => '2026-09-09T06:00:00.000Z';
  openMapInfoSheet_ = () => ({ getRange: () => ({
    getValues: () => [row.slice()], getFormulas: () => [row.map(() => '')],
    setValues: values => { written = values[0].slice(); }
  }) });
  updatePinDetails({ id:'audit-local-only', title:'new title', eventAt:PinData.rowToPin(row).eventAt });
  result.editAfterReload = { eventAtWritten:written[12] };
  updatePinDetails({ id:'audit-local-only', title:'=1+1', description:'=2+2', tags:['=3+3'] });
  result.formulaWrite = { title:written[1], description:written[2], tags:written[11],
    safeAppendEncoding:encodeSpreadsheetLiteral_('=1+1') };
  let cacheRow;
  openRouteCacheSheet_ = () => ({getLastRow:()=>0,appendRow:r=>{cacheRow=r;}});
  const coords = Array.from({length:4000},(_,i)=>[35+i*0.000001,135+i*0.000001]);
  const cached = putRouteCache({cacheKey:'audit-key',routeId:'audit-route',coords});
  result.routeCache = {ok:cached.ok,points:coords.length,jsonCharacters:cacheRow[2].length};
  const chunks = chunkTrackSegments_({trackId:'audit-track',revisionId:'audit-revision',segments:[{
    points:Array.from({length:20000},(_,i)=>({lat:35+i*0.000001,lng:135+i*0.000001,elevation:100,time:'2025-07-20T00:16:00.000Z'}))
  }]});
  result.trackChunks = {chunks:chunks.length,points:chunks.reduce((n,c)=>n+c.pointCount,0),maxJsonCharacters:Math.max(...chunks.map(c=>c.pointsJson.length))};
  return result;
})()`, context);
assert.equal(findings.dateCell.outputEventAt, '');
assert.equal(findings.textCell.outputEventAt, '2025-07-20T09:16');
assert.equal(findings.editAfterReload.eventAtWritten, '');
assert.equal(findings.formulaWrite.title, '=1+1');
assert.ok(findings.routeCache.jsonCharacters > 50000);
assert.equal(findings.trackChunks.points, 20000);
assert.ok(findings.trackChunks.maxJsonCharacters <= 40000);
console.log(JSON.stringify(findings, null, 2));

// Run the real edit event handlers with delayed local promises.
(async () => {
  const html = fs.readFileSync('index.html', 'utf8');
  const start = html.indexOf('    function cancelPinEditor()');
  const end = html.indexOf('    function cancelDeleteConfirmation()', start);
  assert.ok(start > 0 && end > start);
  const handlers = {}, calls = [], resolves = [], closed = [];
  const pin = { id: 'local-pin-a', lat: 35, lng: 135 };
  const state = {editingPinId:pin.id, editColor:'#e53935', editIcon:'default'};
  const fields = {'edit-title':'local title','edit-desc':'','edit-event-at':'','edit-tags':'','edit-status':''};
  vm.runInNewContext(html.slice(start,end), {
    document:{getElementById:id=>({value:fields[id] || '',addEventListener:(_,fn)=>{handlers[id]=fn;}})},
    state, getPinById:()=>pin, normalizeTags:()=>[], collectLinks:()=>[],
    withEditToken:x=>x, withGAS:(name,payload)=>{calls.push(name);return new Promise(resolve=>resolves.push(resolve));},
    checkDupsThenRun:(_t,_lat,_lng,_id,fn)=>fn(),
    closeOverlay:name=>closed.push(name), showAppNotification:()=>{},
    renderPins:()=>{},renderSidePanel:()=>{},renderColorFilterUI:()=>{},renderIconFilterUI:()=>{},renderTagFilterUI:()=>{}
  });
  await handlers['edit-save']();
  await handlers['edit-save']();
  handlers['edit-cancel']();
  state.editingPinId='local-pin-b';
  resolves.forEach(resolve=>resolve({ok:true,updatedAt:'audit-time'}));
  await new Promise(setImmediate);
  const result = {requestsWhilePending:calls.length,editingPinAfterOldResponses:state.editingPinId,closeCalls:closed.length};
  assert.equal(result.requestsWhilePending,2);
  assert.equal(result.editingPinAfterOldResponses,null);
  console.log(JSON.stringify({editPendingRace:result},null,2));
})().catch(error=>{console.error(error);process.exitCode=1;});

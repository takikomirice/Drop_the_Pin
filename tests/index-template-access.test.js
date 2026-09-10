const {test}=require('node:test');
const assert=require('node:assert/strict');
const vm=require('node:vm');
const fs=require('node:fs');
function render(token, debug=false) {
  const ctx={};vm.createContext(ctx);
  vm.runInContext(fs.readFileSync('Code.js','utf8'),ctx);
  return vm.runInContext('typeof prepareIndexTemplate_ === "function" ? prepareIndexTemplate_(source, token, debug) : source',Object.assign(ctx,{
    source:'<script>window.__STARTUP_DEBUG__ = false;</script><main>map</main>\n  <? if (editToken) { ?>\n<template data-dtp-audio-editor-boundary="start"></template><div>editor</div><template data-dtp-audio-editor-boundary="end"></template>\n  <? } ?>',token,debug
  }));
}
test('index HTML never emits literal editor directives to an authorized viewer',()=>{
  const html=render('valid-issued-token');assert.ok(html.includes('<div>editor</div>'));assert.ok(!html.includes('<?'));
});
test('startup diagnostics honors the outer URL only with an issued edit token',()=>{
  assert.ok(render('valid',true).includes('window.__STARTUP_DEBUG__ = true;'));
  assert.ok(render('',true).includes('window.__STARTUP_DEBUG__ = false;'));
  assert.ok(render('valid',false).includes('window.__STARTUP_DEBUG__ = false;'));
});
test('index HTML excludes editor code for a viewer without an issued edit token',()=>{
  const html=render('');assert.ok(html.includes('<main>map</main>'));assert.ok(!html.includes('<div>editor</div>'));assert.ok(!html.includes('<?'));
});

const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

test('geocode URL survives the observed GAS line-comment transformation', async () => {
  const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
  const source = html.slice(html.indexOf('    async function geocodeSearch('), html.indexOf('    function setupGeoSearch('));
  // Real HtmlService output truncated the URL after https: in this function.
  const served = source.replace(/\/\/[^\n]*/g, '');
  let requested = '';
  let rendered = false;
  const search = vm.runInNewContext('(' + served + ')', {
    document: {createElement() { return {}; }},
    fetch: async url => { requested = url; return {ok:true, json:async () => []}; },
    renderGeocodeResults() { rendered = true; }
  });
  const panel = {dataset:{}, style:{}, appendChild() {}};
  await search('京都 & 森', () => {}, panel, 'request-1');
  const url = new URL(requested);
  assert.equal(url.origin, 'https://nominatim.openstreetmap.org');
  assert.equal(url.pathname, '/search');
  assert.equal(url.searchParams.get('q'), '京都 & 森');
  assert.equal(url.searchParams.get('accept-language'), 'ja');
  assert.equal(rendered, true);
});

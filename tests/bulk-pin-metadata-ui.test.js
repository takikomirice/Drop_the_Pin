const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.resolve(__dirname, '..', 'shared.html'), 'utf8');

function functionBody(name) {
  const start = indexHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  for (let index = open; index < indexHtml.length; index += 1) {
    if (indexHtml[index] === '{') depth += 1;
    if (indexHtml[index] === '}') depth -= 1;
    if (depth === 0) return indexHtml.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${name}`);
}

test('bulk metadata action is available only for two editable selected pins', () => {
  const body = functionBody('updateBulkBar');
  assert.match(body, /canEdit\(\)/);
  assert.match(body, /count\s*>?=\s*2|count\s*>\s*1/);
  assert.match(indexHtml, /id="bulk-metadata-btn"/);
  assert.equal(sharedHtml.includes('bulkUpdatePinMetadata'), false);
});

test('bulk metadata overlay distinguishes four tag modes and empty replacement', () => {
  for (const mode of ['none', 'add', 'remove', 'replace']) {
    assert.match(indexHtml, new RegExp(`<option value="${mode}"`));
  }
  assert.match(indexHtml, /現在のタグをすべて置き換えます/);
  const payloadBody = functionBody('buildBulkMetadataPayload');
  assert.match(payloadBody, /tagMode/);
  assert.match(payloadBody, /tags/);
  assert.match(payloadBody, /normalizeBulkTagInput/);
});

test('bulk icon choices are generated only from PIN_ICONS and unchanged disables apply', () => {
  const pickerBody = functionBody('renderBulkMetadataIconOptions');
  assert.match(pickerBody, /PIN_ICONS\.forEach/);
  assert.match(pickerBody, /document\.createElement\('option'\)/);
  assert.match(pickerBody, /textContent/);
  const stateBody = functionBody('updateBulkMetadataControls');
  assert.match(stateBody, /tagMode === 'none'/);
  assert.match(stateBody, /icon/);
  assert.match(stateBody, /applyButton\.disabled/);
});

test('bulk metadata save snapshots IDs, includes edit token, and blocks double submission', () => {
  const openBody = functionBody('openBulkMetadataOverlay');
  assert.match(openBody, /canEdit\(\)/);
  assert.match(openBody, /selectedPinIds\.size < 2/);
  assert.match(openBody, /Array\.from\(state\.selectedPinIds\)/);
  const saveBody = functionBody('applyBulkMetadata');
  assert.match(saveBody, /bulkMetadata\.saving/);
  assert.match(saveBody, /withGAS\('bulkUpdatePinMetadata', withEditToken\(payload\)\)/);
  assert.match(saveBody, /bulkMetadata\.saving = true/);
  assert.match(saveBody, /bulkMetadata\.saving = false/);
});

test('bulk metadata applies server updates only after success and leaves failure text-safe', () => {
  const saveBody = functionBody('applyBulkMetadata');
  const successCheck = saveBody.indexOf('if (!result || !result.ok)');
  const merge = saveBody.indexOf('mergeBulkMetadataUpdates(result.updates');
  assert.ok(successCheck !== -1 && merge > successCheck);
  assert.match(saveBody, /renderPins\(\)/);
  assert.match(saveBody, /renderSidePanel\(\)/);
  assert.match(saveBody, /renderIconFilterUI\(\)/);
  assert.match(saveBody, /renderTagFilterUI\(\)/);
  assert.match(saveBody, /errorElement\.textContent/);
  assert.equal(saveBody.slice(saveBody.lastIndexOf('catch')).includes('mergeBulkMetadataUpdates'), false);
});

test('bulk metadata close is disabled while saving and backdrop uses the safe close path', () => {
  const closeBody = functionBody('closeBulkMetadataOverlay');
  assert.match(closeBody, /bulkMetadata\.saving/);
  assert.match(closeBody, /return/);
  const backdropBody = functionBody('closeOverlayFromBackdrop');
  assert.match(backdropBody, /dismissOverlayById\(record\.id\)/);
  const dismissBody = functionBody('dismissOverlayById');
  assert.match(dismissBody, /bulk-metadata-overlay[\s\S]*closeBulkMetadataOverlay/);
});

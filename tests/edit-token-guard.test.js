const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

const guardedServerFunctions = [
  'navigateToFolder',
  'getRootFolderContents',
  'getPinDriveMeta',
  'saveRouteGroup',
  'setRoutePins',
  'putRouteCache',
  'invalidateRouteCacheForPin',
  'invalidateRouteCacheForRoute',
  'deleteRouteGroup',
  'updateRoutesOrder',
  'createShareLink',
  'listShareLinks',
  'setShareLinkEnabled',
  'deleteShareLink',
  'revokeShareLink',
  'saveMapData',
  'duplicatePin',
  'updatePinDetails',
  'movePin',
  'unplacePin',
  'bulkUpdatePinStatus',
  'deletePin',
  'bulkDeletePins',
  'updateAppSettings',
  'updatePin'
];

const unguardedReadOnlyFunctions = [
  'doGet',
  'getMapData',
  'getAppSettings',
  'getRouteGroups',
  'getRouteCache',
  'getSharedViewData',
  'getSharedRoadRouteCache'
];

const clientMutationCalls = [
  'duplicatePin',
  'putRouteCache',
  'setRoutePins',
  'updateRoutesOrder',
  'updateAppSettings',
  'createShareLink',
  'setShareLinkEnabled',
  'deleteShareLink',
  'unplacePin',
  'updatePinDetails',
  'bulkDeletePins',
  'deletePin',
  'movePin',
  'saveRouteGroup',
  'deleteRouteGroup',
  'bulkUpdatePinStatus',
  'getRootFolderContents',
  'navigateToFolder',
  'getPinDriveMeta',
  'saveMapData',
  'listShareLinks'
];

function assertIncludes(source, needle, message) {
  assert.ok(source.includes(needle), message || `Expected source to include ${needle}`);
}

function sourceFunctionBody(source, name) {
  const index = source.indexOf(`function ${name}(`);
  assert.notEqual(index, -1, `Expected function ${name} to exist`);
  const openIndex = source.indexOf('{', index);
  assert.notEqual(openIndex, -1, `Expected function ${name} to have a body`);
  let depth = 0;
  for (let i = openIndex; i < source.length; i += 1) {
    const char = source[i];
    if (char === '{') depth += 1;
    if (char === '}') depth -= 1;
    if (depth === 0) return source.slice(openIndex + 1, i);
  }
  assert.fail(`Could not parse function ${name}`);
}

function withGasCallExpression(source, methodName) {
  const marker = `withGAS('${methodName}'`;
  const start = source.indexOf(marker);
  assert.notEqual(start, -1, `Expected withGAS call for ${methodName}`);
  const end = source.indexOf(');', start);
  assert.notEqual(end, -1, `Expected withGAS call for ${methodName} to terminate`);
  return source.slice(start, end + 2);
}

test('server edit token assertion validates script cache entries', () => {
  assertIncludes(codeJs, 'function getEditTokenCacheKey_(token)');
  const body = sourceFunctionBody(codeJs, 'assertEditToken_');
  assertIncludes(body, 'payload.__editToken');
  assertIncludes(body, 'CacheService.getScriptCache().get(getEditTokenCacheKey_(token))');
  assertIncludes(body, "cached !== '1'");
  assertIncludes(body, '編集権限が確認できません。編集URLを開き直してください。');

  const issueBody = sourceFunctionBody(codeJs, 'issueEditTokenFromRequest_');
  assertIncludes(issueBody, 'CacheService.getScriptCache().put(getEditTokenCacheKey_(editToken)');
  assertIncludes(issueBody, 'EDIT_TOKEN_TTL_SECONDS');
});

test('server mutation and edit-management functions assert edit token near entry', () => {
  guardedServerFunctions.forEach((name) => {
    const body = sourceFunctionBody(codeJs, name);
    const prologue = body.slice(0, 260);
    assertIncludes(prologue, 'assertEditToken_', `${name} should assert edit token near the top`);
  });
});

test('read-only and shared-view server functions remain unguarded', () => {
  unguardedReadOnlyFunctions.forEach((name) => {
    const body = sourceFunctionBody(codeJs, name);
    assert.equal(body.includes('assertEditToken_'), false, `${name} should remain readable without edit token`);
  });
});

test('index mutation calls pass edit token payloads', () => {
  clientMutationCalls.forEach((methodName) => {
    const call = withGasCallExpression(indexHtml, methodName);
    assertIncludes(call, 'withEditToken(', `${methodName} call should include withEditToken`);
  });
});

test('index read-only calls are not forced through edit token', () => {
  ['getMapData', 'getRouteGroups', 'getRouteCache', 'getAppSettings'].forEach((methodName) => {
    const marker = indexHtml.includes(`withGASNoArg('${methodName}'`)
      ? `withGASNoArg('${methodName}'`
      : `withGAS('${methodName}'`;
    const start = indexHtml.indexOf(marker);
    assert.notEqual(start, -1, `Expected read-only call for ${methodName}`);
    const end = indexHtml.indexOf(');', start);
    assert.equal(indexHtml.slice(start, end + 2).includes('withEditToken('), false);
  });
});

test('shared view keeps read-only GAS calls without edit token dependency', () => {
  assertIncludes(sharedHtml, "withGAS('getSharedViewData', state.token)");
  assertIncludes(sharedHtml, "withGAS('getSharedRoadRouteCache', { token: state.token, routeId: routeId })");
  assert.equal(sharedHtml.includes('withEditToken'), false);
  assert.equal(sharedHtml.includes('__EDIT_TOKEN__'), false);
});

test('public client does not expose raw edit keys', () => {
  assert.equal(indexHtml.includes('editKey='), false);
  assert.equal(sharedHtml.includes('editKey='), false);
});

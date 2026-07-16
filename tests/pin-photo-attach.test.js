const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const root = path.resolve(__dirname, '..');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(root, 'shared.html'), 'utf8');

function countId(id) {
  return (indexHtml.match(new RegExp(`id=["']${id}["']`, 'g')) || []).length;
}

test('photo-less editable pin detail alone exposes the existing-pin photo source chooser', () => {
  [
    'pin-detail-photo-add', 'pin-photo-source-overlay', 'pin-photo-source-title',
    'pin-photo-source-target', 'pin-photo-source-local', 'pin-photo-source-drive',
    'pin-photo-source-cancel', 'pin-photo-source-error', 'pin-photo-attach-file-input'
  ].forEach((id) => assert.equal(countId(id), 1, id));
  assert.match(indexHtml, /写真を追加/);
  assert.match(indexHtml, /端末から写真を追加/);
  assert.match(indexHtml, /Driveから写真を追加/);
  assert.match(indexHtml, /!pin\.fileId\s*&&\s*!pin\.imageUrl/);
  assert.match(indexHtml, /canEdit\(\)[\s\S]{0,180}!pin\.fileId\s*&&\s*!pin\.imageUrl/);
  assert.doesNotMatch(
    indexHtml.match(/id="pin-photo-attach-file-input"[^>]*>/)[0],
    /\bmultiple\b/
  );
  assert.equal(sharedHtml.includes('pin-detail-photo-add'), false);
  assert.equal(sharedHtml.includes('pin-photo-source-overlay'), false);
});

test('existing-pin photo state and source flow stay separate and reject every non-single selection', () => {
  assert.match(indexHtml, /pinPhotoAttach:\s*\{/);
  assert.match(indexHtml, /targetPinId:\s*''/);
  assert.match(indexHtml, /expectedUpdatedAt:\s*''/);
  assert.match(indexHtml, /function openPinPhotoSource\(/);
  assert.match(indexHtml, /function startPinPhotoAttachFromFiles\(/);
  assert.match(indexHtml, /fileList\.length\s*!==\s*1/);
  assert.doesNotMatch(indexHtml, /startPinPhotoAttachFromFiles[\s\S]{0,600}\.slice\(0,\s*1\)/);
  assert.match(indexHtml, /getSelectionLimit:\s*function\(\)[\s\S]{0,180}return\s+1/);
  assert.match(indexHtml, /returnToPinPhotoSource/);
});

test('existing-pin attach reuses common preparation preview processor and updates client state in place', () => {
  const workflowSection = indexHtml.slice(
    indexHtml.indexOf('const pinPhotoAttachWorkflow = MultiPhotoImportWorkflow.create'),
    indexHtml.indexOf('const drivePhotoImportController = DrivePhotoImportUI.create')
  );
  assert.match(workflowSection, /builderApi:\s*MultiPhotoImportBuilder/);
  assert.match(workflowSection, /processorApi:\s*ImportPhotoItemProcessor/);
  assert.match(workflowSection, /flowApi:\s*ImportFlowController/);
  assert.match(workflowSection, /onSaved:\s*upsertImportedPin/);
  assert.match(indexHtml, /operationMode:\s*'attach-existing-pin'/);
  assert.match(indexHtml, /readOnlyFields:\s*true/);
  assert.match(indexHtml, /primaryLabel:\s*'このピンに写真を追加'/);
  assert.match(indexHtml, /sourceLabel:[\s\S]{0,180}targetTitle/);
  assert.match(indexHtml, /operation\.sourceKind\s*===\s*'drive'\s*\?\s*'Drive'\s*:\s*'端末'/);
  assert.doesNotMatch(workflowSection, /saveImportPhotoItem\s*\(/);
  assert.doesNotMatch(workflowSection, /location\.reload|location\.replace/);
});

test('attach Preview is read-only for pin metadata but keeps the photo and explicit action', () => {
  assert.match(indexHtml, /previewReadOnlyFields/);
  assert.match(indexHtml, /isEditable\(\)[\s\S]{0,120}!previewReadOnlyFields/);
  assert.match(indexHtml, /import-preview-primary[\s\S]{0,300}primaryLabel/);
  assert.match(indexHtml, /data-import-read-only/);
  assert.match(indexHtml, /このピンに写真を追加/);
});

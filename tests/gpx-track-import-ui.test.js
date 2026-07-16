const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
const sharedHtml = fs.readFileSync(path.join(__dirname, '..', 'shared.html'), 'utf8');
const readme = fs.readFileSync(path.join(__dirname, '..', 'README.md'), 'utf8');

test('Data exposes the edit-only GPX route entry and shared view exposes none', () => {
  for (const id of [
    'gpx-track-import-button', 'gpx-track-file-input',
    'gpx-track-operation-status', 'gpx-track-operation-error'
  ]) assert.match(indexHtml, new RegExp(`id="${id}"`), id);
  assert.match(indexHtml, /id="gpx-track-import-section"[\s\S]*?GPXルート[\s\S]*?id="gpx-track-import-button"[^>]*>[\s\S]*?<span class="action-label">取込<\/span><\/button>/);
  assert.match(indexHtml, /id="gpx-track-file-input"[^>]*type="file"[^>]*accept="\.gpx,application\/gpx\+xml,application\/xml,text\/xml"[^>]*hidden/);
  assert.doesNotMatch(sharedHtml, /gpx-track-|GPXトラックインポート|track-import-preview-overlay/);
});

test('GPX adapter owns validation Core mapping and safe operation targets but never persistence', () => {
  const start = indexHtml.indexOf('const GpxTrackImportAdapter =');
  const end = indexHtml.indexOf('\n    const TrackFileImportUI =', start);
  assert.notEqual(start, -1, 'Expected GpxTrackImportAdapter');
  assert.notEqual(end, -1, 'Expected adapter boundary');
  const source = indexHtml.slice(start, end);
  assert.match(source, /kind:\s*'gpx'/);
  assert.match(source, /label:\s*'GPX'/);
  assert.match(source, /core:\s*GpxTrackInterchangeCore/);
  assert.match(source, /maxFileBytes:\s*5\s*\*\s*1024\s*\*\s*1024/);
  assert.match(source, /operationStatusId:\s*'gpx-track-operation-status'/);
  assert.match(source, /operationErrorId:\s*'gpx-track-operation-error'/);
  assert.match(source, /GPX_FILE_TYPE_INVALID|GPX_FILE_TOO_LARGE|GPX_FILE_READ_FAILED/);
  const safeMessageStart = indexHtml.indexOf('function safeGpxTrackImportMessage(');
  const safeMessageSource = indexHtml.slice(safeMessageStart, start);
  assert.match(safeMessageSource, /GPX_INVALID_XML/);
  assert.match(safeMessageSource, /GPX_NAMESPACE_UNSUPPORTED/);
  assert.match(safeMessageSource, /GPX_TRACK_REQUIRED/);
  assert.match(safeMessageSource, /GPX_POINT_TIME_INVALID/);
  assert.doesNotMatch(source, /saveTrackBundle|callGAS|withEditToken/);
});

test('GPX stats and warnings are rendered as fixed labels and counts only', () => {
  const start = indexHtml.indexOf('const GpxTrackImportAdapter =');
  const end = indexHtml.indexOf('\n    const TrackFileImportUI =', start);
  const source = indexHtml.slice(start, end);
  for (const phrase of [
    'trk要素', 'rte要素', '時刻付きpoint', '標高付きpoint',
    'ウェイポイント', '空の軌跡区間', '時刻のない地点', '標高のない地点'
  ]) assert.match(source, new RegExp(phrase));
  assert.doesNotMatch(source, /innerHTML/);
});

test('README summarizes GPX support, limits, phase history, and manual verification', () => {
  assert.match(readme, /GPXトラック[^\n]*GPX 1\.0 \/ 1\.1[^\n]*trk／rte[^\n]*5MB[^\n]*200 segment[^\n]*20,000 point/);
  assert.match(readme, /写真時刻照合をPhase 8で追加/);
  assert.match(readme, /YAMAP[^\n]*Garmin[^\n]*実GPX/);
  assert.match(readme, /スマートフォン[^\n]*タブレット[^\n]*PC/);
  assert.match(readme, /Apps Script[^\n]*実ブラウザ[^\n]*確認は未実施/);
});

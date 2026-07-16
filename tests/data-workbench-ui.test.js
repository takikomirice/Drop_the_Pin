const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const indexHtml = fs.readFileSync(path.resolve(__dirname, '..', 'index.html'), 'utf8');

function countId(id) {
  const escaped = id.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  return (indexHtml.match(new RegExp(`\\bid=["']${escaped}["']`, 'g')) || []).length;
}

function elementBlock(id, nextId) {
  const start = indexHtml.indexOf(`id="${id}"`);
  assert.notEqual(start, -1, `Expected #${id}`);
  const end = nextId ? indexHtml.indexOf(`id="${nextId}"`, start) : indexHtml.length;
  assert.notEqual(end, -1, `Expected #${nextId} after #${id}`);
  return indexHtml.slice(start, end);
}

function functionBody(name) {
  const start = indexHtml.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `Expected function ${name}`);
  const open = indexHtml.indexOf('{', start);
  let depth = 0;
  let quote = null;
  let escaped = false;
  for (let index = open; index < indexHtml.length; index += 1) {
    const character = indexHtml[index];
    if (quote) {
      if (escaped) escaped = false;
      else if (character === '\\') escaped = true;
      else if (character === quote) quote = null;
      continue;
    }
    if (character === '"' || character === "'" || character === '`') {
      quote = character;
      continue;
    }
    if (character === '{') depth += 1;
    if (character === '}' && --depth === 0) return indexHtml.slice(open + 1, index);
  }
  assert.fail(`Could not parse ${name}`);
}

test('top bar exposes Data independently and keeps the four persistent utilities in Other', () => {
  assert.equal(countId('data-toggle'), 1);
  assert.equal(countId('more-menu-toggle'), 1);
  assert.equal(countId('more-menu'), 1);

  const menu = elementBlock('more-menu', 'mode-badge');
  ['settings-toggle', 'help-open-btn', 'theme-toggle', 'more-root-drive-link']
    .forEach((id) => assert.equal((menu.match(new RegExp(`id="${id}"`, 'g')) || []).length, 1, id));
  assert.match(menu, />設定</);
  assert.match(menu, />使い方</);
  assert.match(menu, />テーマ</);
  assert.match(menu, />保存先Driveを開く</);
});

test('settings and Data are separate workbenches and all legacy operation IDs remain unique', () => {
  const settings = elementBlock('settings-overlay', 'data-overlay');
  const data = elementBlock('data-overlay', 'input-presets-overlay');

  assert.match(settings, /タイトル変更時に写真名も変更/);
  assert.match(settings, /入力プリセット/);
  assert.match(settings, /保存先Driveを開く/);
  assert.doesNotMatch(settings, /csv-interchange-section|geojson-interchange-section|gpx-track-import-section/);

  const operationIds = [
    'csv-import-button', 'csv-export-button',
    'geojson-import-button', 'geojson-export-button',
    'gpx-track-import-button', 'geojson-track-import-button'
  ];
  operationIds.forEach((id) => {
    assert.equal(countId(id), 1, `${id} must exist exactly once`);
    assert.match(data, new RegExp(`id="${id}"`));
  });
  assert.match(data, /ピンデータ/);
  assert.match(data, /ルートデータ/);
  assert.match(data, /GeoJSON Point/);
  assert.doesNotMatch(data, /ルート(?:の)?書出|ルートエクスポート/);
  ['renderCsvInterchangeBusy', 'renderGeoJsonInterchangeBusy', 'renderTrackImportBusy']
    .forEach((name) => assert.match(functionBody(name), /'取込'/, `${name} must restore the compact Data label`));
  assert.doesNotMatch(functionBody('renderTrackImportBusy'), /トラックインポート|トラックを(?:読み込み|保存)/);

  const addMenu = elementBlock('add-menu-overlay', 'upload-overlay');
  assert.doesNotMatch(addMenu, /CSV|GeoJSON|GPX/);
});

test('settings and Data ownership prevents concurrent open and protects busy work', () => {
  const openSettings = functionBody('openSettingsModal');
  const closeSettings = functionBody('closeSettingsModal');
  const openData = functionBody('openDataWorkbench');
  const closeData = functionBody('closeDataWorkbench');

  assert.match(openSettings, /isProductionImportBusy\(\)/);
  assert.match(openSettings, /data-overlay/);
  assert.match(closeSettings, /isProductionImportBusy\(\)|settingsSavePending/);
  assert.match(openData, /isProductionImportBusy\(\)/);
  assert.match(openData, /settings-overlay/);
  assert.match(closeData, /isProductionImportBusy\(\)/);
  assert.match(openData, /openOverlay\('data-overlay'\)/);
  assert.match(indexHtml, /data-overlay[\s\S]*closeDataWorkbench/);
  assert.match(indexHtml, /settings-overlay[\s\S]*closeSettingsModal/);
  assert.match(functionBody('openInputPresetManager'), /settingsSavePending/);
  assert.match(indexHtml, /settingsSavePending[\s\S]*settings-cancel'\)\.disabled = true/);
});

test('preset workbench keeps the 100 item contract, complete operations, and explicit modes', () => {
  const preset = elementBlock('input-presets-overlay', 'location-choice-overlay');
  assert.match(preset, /最大100件/);
  assert.equal(countId('input-presets-count'), 1);
  for (const id of [
    'input-preset-add', 'input-presets-list', 'input-presets-empty',
    'input-preset-save', 'input-preset-cancel', 'input-presets-error'
  ]) assert.match(preset, new RegExp(`id="${id}"`), id);
  for (const field of ['tags', 'status']) {
    assert.match(preset, new RegExp(`id="input-preset-${field}-mode"[\\s\\S]*?value="keep"[\\s\\S]*?value="set"[\\s\\S]*?value="clear"`));
  }
  for (const field of ['color', 'icon']) {
    assert.match(preset, new RegExp(`id="input-preset-${field}-mode"[\\s\\S]*?value="keep"[\\s\\S]*?value="set"`));
    assert.doesNotMatch(preset.match(new RegExp(`<select id="input-preset-${field}-mode"[\\s\\S]*?<\\/select>`))[0], /value="clear"/);
  }
  assert.equal((preset.match(/<option value="set">設定する<\/option>/g) || []).length, 4);
  assert.equal((preset.match(/<option value="clear">空にする<\/option>/g) || []).length, 2);
  assert.doesNotMatch(preset, /指定値|指定色|指定アイコン|>維持<|>削除</);

  const renderer = functionBody('renderInputPresetList');
  ['editInputPreset', 'duplicateInputPreset', 'deleteInputPresetFromUi', 'toggleInputPresetEnabled', 'moveInputPreset']
    .forEach((action) => assert.match(renderer, new RegExp(action), action));
  assert.match(renderer, /input-presets-count/);
  assert.match(renderer, /input-presets-empty/);
  assert.match(functionBody('setInputPresetError'), /style\.display = message \? 'block' : 'none'/);
});

test('GPX and GeoJSON routes share the Pencil 12 route Preview without resume semantics', () => {
  const preview = elementBlock('track-import-preview-overlay', 'dup-warning-overlay');
  assert.match(preview, /ルート取込プレビュー/);
  for (const id of [
    'track-import-preview-name', 'track-import-preview-description', 'track-import-preview-color',
    'track-import-preview-line-style', 'track-import-preview-visible',
    'track-import-preview-source', 'track-import-preview-segments', 'track-import-preview-points',
    'track-import-preview-distance', 'track-import-preview-time', 'track-import-preview-elevation',
    'track-import-preview-warnings', 'track-import-preview-state',
    'track-import-preview-save', 'track-import-preview-retry', 'track-import-preview-discard'
  ]) assert.equal(countId(id), 1, id);
  assert.equal(countId('track-import-preview-line-width'), 0, 'line width stays fixed without an input');
  assert.match(preview, /GPX／GeoJSONルート共通/);
  assert.match(preview, />再試行</);
  assert.doesNotMatch(preview, /失敗分だけ再開|残りを再開/);
  assert.doesNotMatch(preview, />トラック|トラック名|トラックを保存/);
  assert.match(indexHtml, /\.route-preview-grid\s*\{[^}]*flex:\s*1 1 auto;[^}]*overflow:\s*hidden;/);
  assert.match(indexHtml, /\.input-presets-workbench\s*\{[^}]*min-height:\s*0;[^}]*flex:\s*1 1 auto;/);

  assert.match(indexHtml, /trackImportController\s*=\s*TrackFileImportUI\.create/);
  assert.match(indexHtml, /geoJsonTrackImportController\s*=\s*trackImportController\.forKind\('geojson'\)/);
  assert.match(indexHtml, /gpxTrackImportController\s*=\s*trackImportController\.forKind\('gpx'\)/);
});

test('product copy excludes design annotations and deferred route export', () => {
  for (const phrase of ['Decision', 'desktopのみ', '失敗分だけ再開', 'ルート書出は将来']) {
    assert.doesNotMatch(indexHtml, new RegExp(phrase), phrase);
  }
});

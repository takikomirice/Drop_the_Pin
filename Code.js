// ============================================================
//  Drop the Pin! — Google Apps Script バックエンド
// ============================================================

const PinData = (function() {
  const DEFAULT_COLOR = '#e53935';
  const DEFAULT_ICON = 'default';
  const URL_RE = /^https?:\/\/\S+$/i;
  const STATUS_OPTIONS = ['未対応', '対応中', '完了', '保留'];
  const ICON_OPTIONS = ['default', 'photo', 'food', 'hotel', 'nature', 'shop', 'transit', 'warning'];
  const MAX_TAGS = 5;

  function normalizeStatus(value) {
    const s = String(value || '').trim();
    if (s === '') return '';
    if (STATUS_OPTIONS.indexOf(s) === -1) {
      throw new Error('invalid status: ' + s);
    }
    return s;
  }

  function normalizeIcon(value) {
    const icon = String(value || '').trim().toLowerCase();
    if (!icon) return DEFAULT_ICON;
    return ICON_OPTIONS.indexOf(icon) === -1 ? DEFAULT_ICON : icon;
  }

  function normalizeEventAt(value) {
    const raw = String(value || '').trim();
    if (!raw) return '';
    const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (!match) return '';
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = Number(match[4]);
    const minute = Number(match[5]);
    const second = Number(match[6] || '0');
    const date = new Date(year, month - 1, day, hour, minute, second);
    if (
      date.getFullYear() !== year ||
      date.getMonth() !== month - 1 ||
      date.getDate() !== day ||
      date.getHours() !== hour ||
      date.getMinutes() !== minute ||
      date.getSeconds() !== second
    ) {
      return '';
    }
    return match[6] ? raw : raw.slice(0, 16);
  }

  function normalizeTags(values) {
    if (!Array.isArray(values)) return [];
    const seen = {};
    const result = [];
    for (var i = 0; i < values.length; i++) {
      const raw = String(values[i] || '').trim().replace(/^#/, '');
      if (!raw) continue;
      const key = raw.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      result.push(raw);
    }
    if (result.length > MAX_TAGS) {
      throw new Error('tags must be ' + MAX_TAGS + ' or fewer');
    }
    return result;
  }

  function serializeTags(values) {
    return normalizeTags(values).join('|');
  }

  function deserializeTags(value) {
    return String(value || '')
      .split('|')
      .map(function(t) { return t.trim(); })
      .filter(function(t) { return t.length > 0; });
  }

  function normalizeSearchText(value) {
    return String(value || '').toLowerCase().trim();
  }

  function deserializeLinks(value) {
    return String(value || '')
      .split('|')
      .map(function(item) { return item.trim(); })
      .filter(function(item) { return item && URL_RE.test(item); });
  }

  function normalizeLinks(links) {
    if (Array.isArray(links)) {
      return links
        .map(function(item) { return String(item || '').trim(); })
        .filter(function(item) { return item && URL_RE.test(item); });
    }
    return deserializeLinks(links);
  }

  function serializeLinks(links) {
    return normalizeLinks(links).join('|');
  }

  function chooseSpreadsheetId() {
    for (var i = 0; i < arguments.length; i += 1) {
      var value = String(arguments[i] || '').trim();
      if (value) return value;
    }
    return '';
  }

  function buildFileNameForSave(title, originalName, shouldSync) {
    const original = String(originalName || 'image');
    const normalizedTitle = String(title || '').trim();
    if (!shouldSync || !normalizedTitle) return original;
    const extensionMatch = original.match(/(\.[^.]+)$/);
    if (!extensionMatch) return normalizedTitle;
    const titleWithoutExtension = normalizedTitle.replace(/\.(?:jpe?g|png|gif|webp|heic|heif)$/i, '') || 'image';
    return titleWithoutExtension + extensionMatch[1];
  }

  function toNumberOrNull(value) {
    if (value === '' || value == null) return null;
    const num = Number(value);
    return Number.isFinite(num) ? num : null;
  }

  function toBooleanSetting(value) {
    if (value === true) return true;
    return String(value || '').trim().toLowerCase() === 'true';
  }

  function rowToPin(row) {
    return {
      timestamp: row[0] ? (row[0] instanceof Date ? row[0].toISOString() : String(row[0])) : '',
      title: row[1] || '',
      description: row[2] || '',
      lat: toNumberOrNull(row[3]),
      lng: toNumberOrNull(row[4]),
      color: row[5] || DEFAULT_COLOR,
      fileId: row[6] || '',
      imageUrl: row[7] || '',
      id: row[8] || '',
      links: deserializeLinks(row[9] || ''),
      status: String(row[10] || '').trim(),
      tags: deserializeTags(row[11] || ''),
      eventAt: normalizeEventAt(row[12]),
      updatedAt: row[13] ? String(row[13]) : '',
      icon: normalizeIcon(row[14])
    };
  }

  function rowsToPins(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const dataRows = rows[0] && rows[0][8] === 'ID' ? rows.slice(1) : rows;
    return dataRows.filter(function(row) { return row && row[8]; }).map(rowToPin);
  }

  return {
    DEFAULT_COLOR: DEFAULT_COLOR,
    DEFAULT_ICON: DEFAULT_ICON,
    STATUS_OPTIONS: STATUS_OPTIONS,
    ICON_OPTIONS: ICON_OPTIONS.slice(),
    deserializeLinks: deserializeLinks,
    normalizeLinks: normalizeLinks,
    rowToPin: rowToPin,
    rowsToPins: rowsToPins,
    serializeLinks: serializeLinks,
    serializeTags: serializeTags,
    deserializeTags: deserializeTags,
    normalizeTags: normalizeTags,
    normalizeStatus: normalizeStatus,
    normalizeIcon: normalizeIcon,
    normalizeEventAt: normalizeEventAt,
    normalizeSearchText: normalizeSearchText,
    toBooleanSetting: toBooleanSetting,
    chooseSpreadsheetId: chooseSpreadsheetId,
    buildFileNameForSave: buildFileNameForSave
  };
})();

const SHEET_NAME = 'map_info';
const CONFIG_SHEET_NAME = 'config';
const SHARE_LINKS_SHEET_NAME = 'share_links';
const ROUTES_SHEET_NAME = 'routes';
const ROUTE_PINS_SHEET_NAME = 'route_pins';
const ROUTE_CACHE_SHEET_NAME = 'route_cache';
const MAP_INFO_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明',
  '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧',
  '状態', 'タグ', 'イベント時刻', '更新時刻', 'アイコン'
];
const MAP_INFO_COLUMN_WIDTHS = [160, 180, 250, 90, 90, 90, 200, 350, 230, 320, 100, 200, 170, 170, 120];
const MAP_INFO_EVENT_AT_COLUMN = 13;
const MAP_INFO_UPDATED_AT_COLUMN = 14;
const MAP_INFO_ICON_COLUMN = 15;
const MAP_INFO_COLUMN_COUNT = MAP_INFO_HEADERS.length;
const DEFAULT_COLOR = PinData.DEFAULT_COLOR;
const DEFAULT_SHARE_LINK_LABEL = 'Drop the Pin!';
const DEFAULT_ROUTE_COLOR = '#1e88e5';
const MAX_ROUTE_PINS = 100;
const EDIT_KEY_CONFIG_KEY = 'EDIT_KEY';
const WEB_APP_URL_CONFIG_KEY = 'WEB_APP_URL';
const EDIT_TOKEN_TTL_SECONDS = 6 * 60 * 60;
const EDIT_TOKEN_CACHE_PREFIX = 'EDIT_TOKEN_';
const SHARE_LINKS_HEADERS = ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'];
const ROUTES_HEADERS = ['routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'];
const ROUTE_PINS_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];
const SHARED_ROAD_ROUTE_CACHE_PROVIDER = 'osrm';
const SHARE_ROUTE_NONE_SENTINEL = '__share_no_routes__';
const SAFE_COLOR_RE = /^#[0-9a-fA-F]{6}$/;
const ROUTE_LINE_STYLES = { solid: true, dashed: true, dotted: true };
const PIN_TITLE_MAX_LENGTH = 80;
const SPREADSHEET_MUTATION_LOCK_TIMEOUT_MS = 5000;
const SPREADSHEET_MUTATION_BUSY_ERROR = '別の更新処理が実行中です。少し待ってから再試行してください。';
const PIN_DELETE_CONFLICT_ERROR = '削除対象のピンが更新されています。画面を再読み込みしてから再試行してください。';
// RangeList のリクエストサイズを抑え、大量更新時の実行安定性を保つ。
const RANGE_LIST_CHUNK_SIZE = 500;

// ============================================================
//  メニュー / 初期設定
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('設定')
    .addItem('初期設定', 'setupSheet')
    .addItem('WebアプリURLを設定', 'promptWebAppUrl')
    .addItem('編集URLを表示', 'showEditUrlDialog')
    .addItem('編集キーを再生成', 'regenerateEditKeyFromMenu')
    .addToUi();
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  const looksHeader = sheet.getLastRow() > 0 && (
    sheet.getRange('I1').getValue() === 'ID' ||
    sheet.getRange('A1').getValue() === 'タイムスタンプ'
  );
  if (!looksHeader && sheet.getLastRow() > 0) {
    sheet.insertRowBefore(1);
  }

  if (!looksHeader) {
    sheet.getRange(1, 1, 1, MAP_INFO_COLUMN_COUNT).setValues([MAP_INFO_HEADERS]);
  } else {
    const headerValues = sheet.getRange(1, 1, 1, MAP_INFO_COLUMN_COUNT).getValues()[0];
    MAP_INFO_HEADERS.forEach(function(header, index) {
      if (headerValues[index] === '' || headerValues[index] == null) {
        sheet.getRange(1, index + 1).setValue(header);
      }
    });
  }
  sheet.getRange(1, 1, 1, MAP_INFO_COLUMN_COUNT)
    .setBackground('#4caf50')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  MAP_INFO_COLUMN_WIDTHS.forEach((width, index) => {
    sheet.setColumnWidth(index + 1, width);
  });
  sheet.getRange('D:D').setNumberFormat('0.000000');
  sheet.getRange('E:E').setNumberFormat('0.000000');

  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.getRange(1, 1, 1, 3).setValues([['設定項目', '値', '説明']]);
    configSheet.getRange('A1:C1')
      .setBackground('#1565c0')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    configSheet.setFrozenRows(1);
    [200, 350, 300].forEach((width, index) => configSheet.setColumnWidth(index + 1, width));
  }

  ensureConfigEntry_(configSheet, 'IMAGE_DRIVE_URL', '',
    '写真を保存するGoogleドライブフォルダのURL（フォルダを右クリック→共有→リンクをコピー）');
  ensureConfigEntry_(configSheet, 'RENAME_FILE_WITH_TITLE', 'false',
    'true の場合、タイトル編集時に Drive 上の写真名も同じタイトルへ更新');
  ensureEditUrlConfig_(configSheet);
  ensureShareLinksSheet_(ss);
  ensureHeaderSheet_(ss, ROUTES_SHEET_NAME, ROUTES_HEADERS);
  ensureHeaderSheet_(ss, ROUTE_PINS_SHEET_NAME, ROUTE_PINS_HEADERS);
  ensureHeaderSheet_(ss, ROUTE_CACHE_SHEET_NAME, ROUTE_CACHE_HEADERS);

  ui.alert(
    '初期設定完了',
    '"' + SHEET_NAME + '" シート、"' + CONFIG_SHEET_NAME + '" シート、"' + SHARE_LINKS_SHEET_NAME + '" シート、' +
    '"' + ROUTES_SHEET_NAME + '" シート、"' + ROUTE_PINS_SHEET_NAME + '" シート、"' + ROUTE_CACHE_SHEET_NAME + '" シートの準備が整いました。\n\n' +
    '次のステップ:\n' +
    '1. "' + CONFIG_SHEET_NAME + '" シートを開いて IMAGE_DRIVE_URL を設定\n' +
    '2. 必要なら RENAME_FILE_WITH_TITLE を true に変更\n' +
    '3. ウェブアプリとしてデプロイ',
    ui.ButtonSet.OK
  );
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var templateName = params.view === 'shared' ? 'shared' : 'index';
  var template = HtmlService.createTemplateFromFile(templateName);
  template.execUrl = getConfiguredWebAppUrl_();
  template.token = params.token || '';
  template.editToken = templateName === 'index' ? issueEditTokenFromRequest_(params) : '';
  return template.evaluate()
    .setTitle('Drop the Pin!')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
//  設定 / シート読み込み
// ============================================================

function ensureConfigEntry_(sheet, key, value, description) {
  const lastRow = sheet.getLastRow();
  const keys = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  const index = keys.indexOf(key);
  if (index === -1) {
    sheet.appendRow([key, typeof value === 'function' ? value() : value, description]);
    return;
  }

  const row = index + 2;
  if (sheet.getRange(row, 2).getValue() === '') {
    sheet.getRange(row, 2).setValue(typeof value === 'function' ? value() : value);
  }
  if (sheet.getRange(row, 3).getValue() === '') {
    sheet.getRange(row, 3).setValue(description);
  }
}

function generateEditKey_() {
  return 'ed_' + Utilities.getUuid().replace(/[^A-Za-z0-9]/g, '');
}

function generateEditToken_() {
  return 'edt_' + Utilities.getUuid().replace(/[^A-Za-z0-9]/g, '');
}

function getEditTokenCacheKey_(token) {
  return EDIT_TOKEN_CACHE_PREFIX + String(token || '').trim();
}

function assertEditToken_(payload) {
  const token = payload && typeof payload === 'object'
    ? String(payload.__editToken || '')
    : '';
  if (!token) {
    throw new Error('編集権限が確認できません。編集URLを開き直してください。');
  }

  const cached = CacheService.getScriptCache().get(getEditTokenCacheKey_(token));
  if (cached !== '1') {
    throw new Error('編集権限が確認できません。編集URLを開き直してください。');
  }
}

function withSpreadsheetMutationLock_(callback) {
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    acquired = lock.tryLock(SPREADSHEET_MUTATION_LOCK_TIMEOUT_MS);
    if (!acquired) return { ok: false, error: SPREADSHEET_MUTATION_BUSY_ERROR };
    return callback();
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function normalizeWebAppUrl_(url) {
  var normalized = String(url || '').trim();
  if (!normalized) return '';
  normalized = normalized.split('#')[0].split('?')[0].replace(/\/+$/, '');
  return normalized.replace(
    /^https:\/\/script\.google\.com\/a\/([^/]+)\/macros\/s\//,
    'https://script.google.com/a/macros/$1/s/'
  );
}

function ensureEditUrlConfig_(configSheet) {
  const sheet = configSheet || openDataSpreadsheet_().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) throw new Error('config シートが見つかりません');
  ensureConfigEntry_(sheet, EDIT_KEY_CONFIG_KEY, generateEditKey_, '編集URL用の共有キー');
  ensureConfigEntry_(sheet, WEB_APP_URL_CONFIG_KEY, '', 'デプロイ済みWebアプリの /exec URL');
  return sheet;
}

function getConfiguredWebAppUrl_() {
  const config = getAppConfig_();
  return normalizeWebAppUrl_(config[WEB_APP_URL_CONFIG_KEY]) ||
    normalizeWebAppUrl_(ScriptApp.getService().getUrl());
}

function buildEditUrl_() {
  ensureEditUrlConfig_();
  const config = getAppConfig_();
  const editKey = String(config[EDIT_KEY_CONFIG_KEY] || '').trim();
  const webAppUrl = getConfiguredWebAppUrl_();
  if (!editKey) throw new Error('EDIT_KEY が設定されていません');
  if (!webAppUrl) throw new Error('WebアプリURLが取得できません');
  return webAppUrl + '?mode=edit&editKey=' + encodeURIComponent(editKey);
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function promptWebAppUrl() {
  const ui = SpreadsheetApp.getUi();
  ensureEditUrlConfig_();
  const response = ui.prompt(
    'WebアプリURLを設定',
    'デプロイ済みWebアプリの /exec URL を入力してください。空欄にすると ScriptApp.getService().getUrl() を使います。',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const normalizedUrl = normalizeWebAppUrl_(response.getResponseText());
  setConfigValue_(WEB_APP_URL_CONFIG_KEY, normalizedUrl);
  ui.alert('保存しました', normalizedUrl || 'WEB_APP_URL を空欄にしました。', ui.ButtonSet.OK);
}

function showEditUrlDialog() {
  const ui = SpreadsheetApp.getUi();
  const editUrl = buildEditUrl_();
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;line-height:1.6;padding:12px;">' +
      '<p>このURLをClassroomなどで共同編集者に共有してください。</p>' +
      '<p>EDIT_KEYを変更しない限り、同じ編集URLを継続して使えます。</p>' +
      '<p>開きっぱなしで編集できなくなった場合は、このURLを再読み込みしてください。</p>' +
      '<input style="box-sizing:border-box;width:100%;font-size:13px;padding:8px;" readonly value="' + escapeHtml_(editUrl) + '">' +
    '</div>'
  ).setWidth(640).setHeight(260);
  ui.showModalDialog(html, '編集URL');
}

function regenerateEditKeyFromMenu() {
  const ui = SpreadsheetApp.getUi();
  ensureEditUrlConfig_();
  const result = ui.alert(
    '編集キーを再生成',
    '編集キーを再生成すると、これまで共有した編集URLは使えなくなります。続行しますか？',
    ui.ButtonSet.YES_NO
  );
  if (result !== ui.Button.YES) return;
  setConfigValue_(EDIT_KEY_CONFIG_KEY, generateEditKey_());
  ui.alert('編集キーを再生成しました', '新しい編集URLを表示して共有し直してください。', ui.ButtonSet.OK);
}

function issueEditTokenFromRequest_(params) {
  if (String(params && params.mode || '') !== 'edit') return '';
  const config = getAppConfig_();
  const expectedEditKey = String(config[EDIT_KEY_CONFIG_KEY] || '').trim();
  const providedEditKey = String(params && params.editKey || '').trim();
  if (!expectedEditKey || providedEditKey !== expectedEditKey) return '';

  const editToken = generateEditToken_();
  CacheService.getScriptCache().put(getEditTokenCacheKey_(editToken), '1', EDIT_TOKEN_TTL_SECONDS);
  return editToken;
}

function openDataSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getAppConfig_() {
  const sheet = openDataSpreadsheet_().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const config = {};
  data.forEach(function(row) {
    if (row[0]) config[String(row[0])] = String(row[1] || '');
  });
  return config;
}

function setConfigValue_(key, value) {
  const sheet = openDataSpreadsheet_().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) throw new Error('config シートが見つかりません');

  const lastRow = sheet.getLastRow();
  const keys = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat() : [];
  const index = keys.indexOf(key);
  if (index === -1) {
    sheet.appendRow([key, String(value), '']);
    return;
  }
  sheet.getRange(index + 2, 2).setValue(String(value));
}

function extractDriveFolderId_(url) {
  const match = String(url || '').match(/\/folders\/([a-zA-Z0-9_-]{15,})/);
  return match ? match[1] : null;
}

function getDriveFolderUrl_(folderId) {
  return folderId ? 'https://drive.google.com/drive/folders/' + folderId : '';
}

function getRootFolderId_() {
  const config = getAppConfig_();
  return extractDriveFolderId_(config.IMAGE_DRIVE_URL || '');
}

function getRenameFileWithTitle_() {
  const config = getAppConfig_();
  return PinData.toBooleanSetting(config.RENAME_FILE_WITH_TITLE);
}

function openMapInfoSheet_() {
  const sheet = openDataSpreadsheet_().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('map_info シートが見つかりません');
  return sheet;
}

function getRequiredSheet_(sheetName) {
  const sheet = openDataSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(sheetName + ' シートが見つかりません。setupSheet() を実行してください。');
  }
  return sheet;
}

function openShareLinksSheet_() {
  return getRequiredSheet_(SHARE_LINKS_SHEET_NAME);
}

function openRoutesSheet_() {
  return getRequiredSheet_(ROUTES_SHEET_NAME);
}

function openRoutePinsSheet_() {
  return getRequiredSheet_(ROUTE_PINS_SHEET_NAME);
}

function openRouteCacheSheet_() {
  return getRequiredSheet_(ROUTE_CACHE_SHEET_NAME);
}

function readFixedWidthDataRows_(sheet, columnCount) {
  const dataRowCount = Math.max(0, sheet.getLastRow() - 1);
  if (dataRowCount === 0) return [];

  const actualColumnCount = Math.max(columnCount, sheet.getLastColumn());
  return sheet.getRange(2, 1, dataRowCount, actualColumnCount).getValues();
}

function readFormulaPreservingDataRows_(sheet, columnCount) {
  const dataRowCount = Math.max(0, sheet.getLastRow() - 1);
  if (dataRowCount === 0) return { values: [], outputRows: [] };

  const actualColumnCount = Math.max(columnCount, sheet.getLastColumn());
  const range = sheet.getRange(2, 1, dataRowCount, actualColumnCount);
  const values = range.getValues();
  const formulas = range.getFormulas();
  const outputRows = values.map(function(row, rowIndex) {
    return row.map(function(value, columnIndex) {
      return formulas[rowIndex][columnIndex] || value;
    });
  });
  return { values: values, outputRows: outputRows };
}

function rewriteFixedWidthDataRows_(sheet, columnCount, previousRowCount, nextRows) {
  const writeRowCount = Math.max(previousRowCount, nextRows.length);
  if (writeRowCount === 0) return;

  const actualColumnCount = Math.max(columnCount, sheet.getLastColumn());
  const output = nextRows.map(function(row) {
    const copied = new Array(actualColumnCount).fill('');
    for (var i = 0; i < Math.min(row.length, actualColumnCount); i += 1) {
      copied[i] = row[i];
    }
    return copied;
  });
  while (output.length < writeRowCount) {
    output.push(new Array(actualColumnCount).fill(''));
  }

  const requiredRows = writeRowCount + 1;
  if (requiredRows > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  // row 1 はヘッダー。旧件数まで空行を含めた1回の書込みで末尾残留も消す。
  sheet.getRange(2, 1, writeRowCount, actualColumnCount).setValues(output);
}

function filterAndRewriteFixedWidthDataRows_(sheet, columnCount, shouldRemove) {
  const snapshot = readFormulaPreservingDataRows_(sheet, columnCount);
  const keptRows = [];
  const removedRows = [];
  snapshot.values.forEach(function(row, index) {
    if (shouldRemove(row, index)) {
      removedRows.push(row);
    } else {
      keptRows.push(snapshot.outputRows[index]);
    }
  });
  if (removedRows.length > 0) {
    rewriteFixedWidthDataRows_(sheet, columnCount, snapshot.values.length, keptRows);
  }
  return { keptRows: keptRows, removedRows: removedRows };
}

function ensureShareLinksSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHARE_LINKS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHARE_LINKS_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, SHARE_LINKS_HEADERS.length).setValues([SHARE_LINKS_HEADERS]);
  } else {
    const columnCount = Math.max(
      SHARE_LINKS_HEADERS.length,
      typeof sheet.getLastColumn === 'function' ? sheet.getLastColumn() : SHARE_LINKS_HEADERS.length
    );
    const headerValues = sheet.getRange(1, 1, 1, columnCount).getValues()[0];
    SHARE_LINKS_HEADERS.forEach(function(header, index) {
      if (headerValues[index] === '' || headerValues[index] == null) {
        sheet.getRange(1, index + 1).setValue(header);
        headerValues[index] = header;
      }
    });
  }
  sheet.getRange(1, 1, 1, SHARE_LINKS_HEADERS.length)
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  if (typeof sheet.setFrozenRows === 'function') sheet.setFrozenRows(1);
  return sheet;
}

function ensureHeaderSheet_(spreadsheet, sheetName, headers) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#1565c0')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    const headerValues = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    headers.forEach(function(header, index) {
      if (headerValues[index] === '' || headerValues[index] == null) {
        sheet.getRange(1, index + 1).setValue(header);
      }
    });
  }
  return sheet;
}

function normalizeShareLinkLabel_(value) {
  var label = String(value || '').trim();
  return label || DEFAULT_SHARE_LINK_LABEL;
}

function normalizeShareToken_(value) {
  if (value && typeof value === 'object' && value.token != null) {
    value = value.token;
  }
  return String(value || '').trim();
}

function normalizeShareColors_(values) {
  if (!Array.isArray(values)) return [];
  var seen = {};
  var result = [];
  values.forEach(function(value) {
    var color = String(value || '').trim();
    if (!SAFE_COLOR_RE.test(color)) return;
    color = color.toLowerCase();
    var key = color.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    result.push(color);
  });
  return result;
}

function serializeShareColors_(values) {
  return normalizeShareColors_(values).join('|');
}

function deserializeShareColors_(value) {
  return normalizeShareColors_(String(value || '').split('|'));
}

function getExistingRouteIdSetForShare_() {
  try {
    const result = {};
    getRouteGroups().forEach(function(group) {
      const routeId = normalizeRouteId_(group && (group.routeId || group.id));
      if (routeId) result[routeId] = true;
    });
    return result;
  } catch (error) {
    if (typeof Logger !== 'undefined' && Logger.log) {
      Logger.log('share_route_ids_normalize_skipped: ' + (error && error.message ? error.message : error));
    }
    return null;
  }
}

function normalizeShareRouteIdsWithSet_(values, routeIdSet) {
  if (!Array.isArray(values)) return [];
  for (var i = 0; i < values.length; i += 1) {
    if (isShareRouteNoneSentinel_(values[i])) return [SHARE_ROUTE_NONE_SENTINEL];
  }
  var seen = {};
  var result = [];
  values.forEach(function(value) {
    var routeId = normalizeRouteId_(value);
    if (!routeId || seen[routeId]) return;
    if (routeIdSet && !routeIdSet[routeId]) return;
    seen[routeId] = true;
    result.push(routeId);
  });
  return result;
}

function isShareRouteNoneSentinel_(value) {
  return String(value || '').trim() === SHARE_ROUTE_NONE_SENTINEL;
}

function isShareRouteSelectionNone_(routeIds) {
  return Array.isArray(routeIds) && routeIds.length === 1 && isShareRouteNoneSentinel_(routeIds[0]);
}

function normalizeShareRouteIds_(values) {
  return normalizeShareRouteIdsWithSet_(values, getExistingRouteIdSetForShare_());
}

function serializeShareRouteIds_(values) {
  return normalizeShareRouteIds_(values).join('|');
}

function deserializeShareRouteIds_(value) {
  return normalizeShareRouteIdsWithSet_(String(value || '').split('|'), null);
}

function isShareLinkEnabled_(value) {
  if (value === '' || value == null) return true;
  if (value === false) return false;
  return String(value).trim().toLowerCase() === 'true';
}

function shareRowToLink_(row) {
  return {
    createdAt: row[0] ? String(row[0]) : '',
    label: normalizeShareLinkLabel_(row[1]),
    token: row[2] ? String(row[2]) : '',
    tags: PinData.deserializeTags(row[3] || ''),
    tagMode: String(row[4] || 'or') === 'and' ? 'and' : 'or',
    enabled: isShareLinkEnabled_(row[5]),
    revokedAt: row[6] ? String(row[6]) : '',
    colors: deserializeShareColors_(row[7] || ''),
    routeIds: deserializeShareRouteIds_(row[8] || '')
  };
}

function findPinRowIndex_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow === 0) return -1;
  const ids = sheet.getRange(1, 9, lastRow, 1).getValues();
  return ids.findIndex(function(row) {
    return row[0] === id;
  });
}

function currentUpdatedAt_() {
  return new Date().toISOString();
}

// ============================================================
//  フォルダ操作
// ============================================================

function getFolderContents_(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const imageMimeTypes = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];
    const folders = [];
    const subFolders = folder.getFolders();
    while (subFolders.hasNext()) {
      const current = subFolders.next();
      folders.push({
        id: current.getId(),
        name: current.getName(),
        type: 'folder',
        url: getDriveFolderUrl_(current.getId())
      });
    }
    folders.sort(function(a, b) { return a.name.localeCompare(b.name, 'ja'); });

    const images = [];
    const files = folder.getFiles();
    while (files.hasNext()) {
      const currentFile = files.next();
      if (imageMimeTypes.indexOf(currentFile.getMimeType()) === -1) continue;
      images.push({
        id: currentFile.getId(),
        name: currentFile.getName(),
        type: 'image',
        url: currentFile.getUrl()
      });
    }
    images.sort(function(a, b) { return a.name.localeCompare(b.name, 'ja'); });

    return {
      items: folders.concat(images),
      folderId: folderId,
      folderName: folder.getName(),
      folderUrl: getDriveFolderUrl_(folderId)
    };
  } catch (error) {
    return { items: [], error: error.message };
  }
}

function navigateToFolder(payload) {
  assertEditToken_(payload);
  const folderId = payload && typeof payload === 'object' ? payload.folderId : payload;
  return getFolderContents_(folderId);
}

function getRootFolderContents(payload) {
  assertEditToken_(payload);
  const folderId = getRootFolderId_();
  if (!folderId) {
    return { items: [], error: 'IMAGE_DRIVE_URL が config シートに設定されていません' };
  }
  return getFolderContents_(folderId);
}

// ============================================================
//  データ操作
// ============================================================

function getMapData() {
  const sheet = openMapInfoSheet_();
  if (sheet.getLastRow() === 0) return [];

  const pins = PinData.rowsToPins(sheet.getDataRange().getValues());
  return pins.map(function(pin) {
    return enrichPinWithDriveMeta_(pin);
  });
}

function getPinDriveMeta(payload) {
  assertEditToken_(payload);
  const pinId = String(payload && payload.pinId || '').trim();
  if (!pinId) return { ok: false, folderUrl: '', error: 'missing_pin_id' };

  const sheet = openMapInfoSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { ok: false, folderUrl: '', error: 'pin_not_found' };

  const pinDriveRows = sheet.getRange(2, 7, lastRow - 1, 3).getValues();
  const pinDriveRow = pinDriveRows.find(function(row) {
    return row[2] === pinId;
  });
  if (!pinDriveRow) return { ok: false, folderUrl: '', error: 'pin_not_found' };

  const fileId = String(pinDriveRow[0] || '').trim();
  if (!fileId) return { ok: true, folderUrl: '' };

  try {
    return { ok: true, folderUrl: getParentFolderUrlByFileId_(fileId) };
  } catch (error) {
    if (typeof Logger !== 'undefined' && Logger.log) {
      Logger.log('getPinDriveMeta: pinId=' + pinId + ' drive lookup failed: ' +
        (error && error.message ? error.message : error));
    }
    return { ok: false, folderUrl: '', error: 'drive_meta_unavailable' };
  }
}

function normalizeRouteId_(value) {
  return String(value || '').trim();
}

function logRouteNormalize_(reason, routeId) {
  const suffix = routeId ? ' for routeId=' + routeId : '';
  Logger.log('route_normalize: ' + reason + suffix);
}

function normalizeRouteName_(value, routeId) {
  const trimmed = String(value || '').trim();
  if (!trimmed) return { ok: false, error: 'route_name_required' };
  if (trimmed.length > 100) {
    logRouteNormalize_('name truncated to 100 chars', routeId);
    return { ok: true, value: trimmed.slice(0, 100) };
  }
  return { ok: true, value: trimmed };
}

function normalizeRouteColor_(value, routeId) {
  const color = String(value || '').trim();
  if (SAFE_COLOR_RE.test(color)) return color.toLowerCase();
  logRouteNormalize_('color "' + color + '" -> ' + DEFAULT_ROUTE_COLOR, routeId);
  return DEFAULT_ROUTE_COLOR;
}

function normalizeRouteMode_(value, routeId) {
  const mode = String(value || 'straight').trim();
  if (mode === 'straight' || mode === 'road') return mode;
  logRouteNormalize_('routeMode "' + mode + '" -> straight', routeId);
  return 'straight';
}

function normalizeRouteLineStyle_(value, routeId) {
  const style = String(value || '').trim().toLowerCase();
  if (ROUTE_LINE_STYLES[style]) return style;
  if (style) logRouteNormalize_('lineStyle "' + style + '" -> solid', routeId);
  return 'solid';
}

function normalizeRouteClosed_(value) {
  return value === true || String(value || '').trim().toLowerCase() === 'true';
}

function normalizeRouteDisplayEnabled_(value) {
  if (value === '' || value == null) return true;
  if (value === false) return false;
  return String(value).trim().toLowerCase() !== 'false';
}

function normalizeRoutePinId_(value) {
  const id = String(value || '').trim();
  return id || null;
}

function routeRowToGroup_(row, pinIds) {
  const routeId = normalizeRouteId_(row[0]);
  const closed = normalizeRouteClosed_(row[4]);
  const orderIndex = Number(row[9]);
  return {
    id: routeId,
    routeId: routeId,
    name: String(row[1] || ''),
    color: normalizeRouteColor_(row[2], routeId),
    routeMode: normalizeRouteMode_(row[3], routeId),
    closed: closed,
    startPinId: normalizeRoutePinId_(row[5]),
    endPinId: closed ? null : normalizeRoutePinId_(row[6]),
    createdAt: row[7] ? String(row[7]) : '',
    updatedAt: row[8] ? String(row[8]) : '',
    orderIndex: Number.isFinite(orderIndex) ? orderIndex : 0,
    visible: normalizeRouteDisplayEnabled_(row[10]),
    showNumbers: normalizeRouteDisplayEnabled_(row[11]),
    showLine: normalizeRouteDisplayEnabled_(row[12]),
    lineStyle: normalizeRouteLineStyle_(row[13], routeId),
    pinIds: Array.isArray(pinIds) ? pinIds.slice() : []
  };
}

function readRoutePinIdsByRoute_() {
  const sheet = openRoutePinsSheet_();
  if (sheet.getLastRow() < 2) return {};

  const rows = sheet.getDataRange().getValues().slice(1);
  const byRoute = {};
  rows.forEach(function(row, index) {
    const routeId = normalizeRouteId_(row[0]);
    const pinId = normalizeRoutePinId_(row[1]);
    if (!routeId || !pinId) return;
    const pinOrder = Number(row[2]);
    if (!byRoute[routeId]) byRoute[routeId] = [];
    byRoute[routeId].push({
      pinId: pinId,
      pinOrder: Number.isFinite(pinOrder) ? pinOrder : index
    });
  });
  Object.keys(byRoute).forEach(function(routeId) {
    byRoute[routeId].sort(function(a, b) {
      return a.pinOrder - b.pinOrder;
    });
    byRoute[routeId] = byRoute[routeId].map(function(item) {
      return item.pinId;
    });
  });
  return byRoute;
}

function readRouteRows_() {
  const sheet = openRoutesSheet_();
  if (sheet.getLastRow() < 2) return [];

  const pinIdsByRoute = readRoutePinIdsByRoute_();
  return sheet.getDataRange().getValues().slice(1).map(function(row, index) {
    const routeId = normalizeRouteId_(row[0]);
    return {
      rowNumber: index + 2,
      row: row,
      group: routeRowToGroup_(row, pinIdsByRoute[routeId] || [])
    };
  }).filter(function(entry) {
    return entry.group.routeId;
  });
}

function findRouteRow_(routeId) {
  const normalizedRouteId = normalizeRouteId_(routeId);
  const rows = readRouteRows_();
  for (var i = 0; i < rows.length; i += 1) {
    if (rows[i].group.routeId === normalizedRouteId) return rows[i];
  }
  return null;
}

function findRouteSheetRow_(routeId) {
  const normalizedRouteId = normalizeRouteId_(routeId);
  if (!normalizedRouteId) return null;
  const sheet = openRoutesSheet_();
  const rows = readFixedWidthDataRows_(sheet, ROUTES_HEADERS.length);
  for (var i = 0; i < rows.length; i += 1) {
    if (normalizeRouteId_(rows[i][0]) === normalizedRouteId) {
      return { rowNumber: i + 2, row: rows[i] };
    }
  }
  return null;
}

function getRouteGroups() {
  return readRouteRows_().map(function(entry, index) {
    const group = entry.group;
    group.orderIndex = Number.isFinite(Number(entry.row[9])) ? Number(entry.row[9]) : index;
    return group;
  }).sort(function(a, b) {
    return a.orderIndex - b.orderIndex;
  });
}

function routeGroupToRow_(group) {
  return [
    group.routeId,
    group.name,
    group.color,
    group.routeMode,
    group.closed,
    group.startPinId || '',
    group.closed ? '' : (group.endPinId || ''),
    group.createdAt,
    group.updatedAt,
    group.orderIndex,
    group.visible !== false,
    group.showNumbers !== false,
    group.showLine !== false,
    normalizeRouteLineStyle_(group.lineStyle, group.routeId)
  ];
}

function saveRouteGroup(payload) {
  assertEditToken_(payload);
  const data = payload || {};
  const requestedRouteId = normalizeRouteId_(data.routeId || data.id);
  if (data.routeId !== undefined && data.routeId !== null && !requestedRouteId) {
    return { ok: false, error: 'missing_route_id' };
  }

  const routeId = requestedRouteId || Utilities.getUuid();
  const nameResult = normalizeRouteName_(data.name, routeId);
  if (!nameResult.ok) return { ok: false, error: nameResult.error };
  const closed = normalizeRouteClosed_(data.closed);
  const rawStartPinId = normalizeRoutePinId_(data.startPinId);
  const rawEndPinId = normalizeRoutePinId_(data.endPinId);
  if (closed && rawEndPinId) {
    logRouteNormalize_('closed route endPinId cleared', routeId);
  }
  const endPinId = closed ? null : rawEndPinId;
  const orderIndex = Number(data.orderIndex);

  return withSpreadsheetMutationLock_(function() {
    const existing = requestedRouteId ? findRouteRow_(routeId) : null;

    const existingPinIds = existing ? existing.group.pinIds : [];
    const pinIdSet = {};
    existingPinIds.forEach(function(pinId) { pinIdSet[pinId] = true; });
    if (rawStartPinId && !pinIdSet[rawStartPinId]) {
      return { ok: false, error: 'invalid_start_pin', pinId: rawStartPinId };
    }
    if (endPinId && !pinIdSet[endPinId]) {
      return { ok: false, error: 'invalid_end_pin', pinId: endPinId };
    }

    const sheet = openRoutesSheet_();
    const now = currentUpdatedAt_();
    const existingOrderIndex = existing ? Number(existing.row[9]) : NaN;
    const group = {
      routeId: routeId,
      name: nameResult.value,
      color: normalizeRouteColor_(data.color, routeId),
      routeMode: normalizeRouteMode_(data.routeMode, routeId),
      closed: closed,
      startPinId: rawStartPinId,
      endPinId: endPinId,
      createdAt: existing ? String(existing.row[7] || now) : now,
      updatedAt: now,
      orderIndex: Number.isFinite(orderIndex)
        ? orderIndex
        : (Number.isFinite(existingOrderIndex) ? existingOrderIndex : readRouteRows_().length),
      visible: normalizeRouteDisplayEnabled_(data.visible),
      showNumbers: normalizeRouteDisplayEnabled_(data.showNumbers),
      showLine: normalizeRouteDisplayEnabled_(data.showLine),
      lineStyle: normalizeRouteLineStyle_(data.lineStyle, routeId)
    };

    if (existing) {
      sheet.getRange(existing.rowNumber, 1, 1, ROUTES_HEADERS.length).setValues([routeGroupToRow_(group)]);
    } else {
      sheet.appendRow(routeGroupToRow_(group));
    }

    const saved = routeRowToGroup_(routeGroupToRow_(group), existingPinIds);
    return { ok: true, routeGroup: saved };
  });
}

function getExistingPinIdSet_() {
  const sheet = openMapInfoSheet_();
  if (sheet.getLastRow() < 2) return {};

  const rows = sheet.getDataRange().getValues();
  const result = {};
  PinData.rowsToPins(rows).forEach(function(pin) {
    if (pin.id) result[String(pin.id)] = true;
  });
  return result;
}

function validateRoutePinIds_(pinIds) {
  if (!Array.isArray(pinIds)) return { ok: false, error: 'pin_ids_invalid' };
  if (pinIds.length > MAX_ROUTE_PINS) return { ok: false, error: 'too_many_pins' };

  const existingPinIds = getExistingPinIdSet_();
  const seen = {};
  const normalizedPinIds = [];
  for (var i = 0; i < pinIds.length; i += 1) {
    const pinId = normalizeRoutePinId_(pinIds[i]);
    if (!pinId || !existingPinIds[pinId]) {
      return { ok: false, error: 'pin_not_found', pinId: pinId || '' };
    }
    if (seen[pinId]) {
      return { ok: false, error: 'pin_ids_duplicated', pinId: pinId };
    }
    seen[pinId] = true;
    normalizedPinIds.push(pinId);
  }
  return { ok: true, pinIds: normalizedPinIds };
}

function setRoutePins(data) {
  assertEditToken_(data);
  const routeId = normalizeRouteId_(data && data.routeId);
  if (!routeId) return { ok: false, error: 'missing_route_id' };
  return withSpreadsheetMutationLock_(function() {
    if (!findRouteSheetRow_(routeId)) return { ok: false, error: 'route_not_found' };

    const validation = validateRoutePinIds_(data && data.pinIds);
    if (!validation.ok) return validation;

    const sheet = openRoutePinsSheet_();
    const snapshot = readFormulaPreservingDataRows_(sheet, ROUTE_PINS_HEADERS.length);
    const keptRows = snapshot.outputRows.filter(function(_row, index) {
      const row = snapshot.values[index];
      return normalizeRouteId_(row[0]) !== routeId;
    });
    const now = currentUpdatedAt_();
    const replacementRows = validation.pinIds.map(function(pinId, index) {
      return [routeId, pinId, index, now, now];
    });
    rewriteFixedWidthDataRows_(
      sheet,
      ROUTE_PINS_HEADERS.length,
      snapshot.values.length,
      keptRows.concat(replacementRows)
    );
    return { ok: true, routeId: routeId, pinIds: validation.pinIds };
  });
}

function deleteRoutePinsForRoute_(routeId) {
  const normalizedRouteId = normalizeRouteId_(routeId);
  if (!normalizedRouteId) return [];

  const sheet = openRoutePinsSheet_();
  const result = filterAndRewriteFixedWidthDataRows_(
    sheet,
    ROUTE_PINS_HEADERS.length,
    function(row) { return normalizeRouteId_(row[0]) === normalizedRouteId; }
  );
  return result.removedRows.length > 0 ? [normalizedRouteId] : [];
}

function deleteRoutePinsForPinIds_(pinIds) {
  if (!Array.isArray(pinIds) || pinIds.length === 0) return [];

  const pinIdSet = {};
  pinIds.forEach(function(pinId) {
    const normalizedPinId = normalizeRoutePinId_(pinId);
    if (normalizedPinId) pinIdSet[normalizedPinId] = true;
  });
  if (Object.keys(pinIdSet).length === 0) return [];

  const prepared = prepareRoutePinDeletionForPinIdSet_(pinIdSet);
  commitPreparedRoutePinDeletion_(prepared);
  return prepared.affectedRouteIds;
}

function prepareRoutePinDeletionForPinIdSet_(pinIdSet) {
  const affectedRouteIds = {};
  const sheet = openRoutePinsSheet_();
  const snapshot = readFormulaPreservingDataRows_(sheet, ROUTE_PINS_HEADERS.length);
  const keptRows = [];
  let removedCount = 0;
  snapshot.values.forEach(function(row, index) {
    const routeId = normalizeRouteId_(row[0]);
    const pinId = normalizeRoutePinId_(row[1]);
    if (pinId && pinIdSet[pinId]) {
      removedCount += 1;
      if (routeId) affectedRouteIds[routeId] = true;
    } else {
      keptRows.push(snapshot.outputRows[index]);
    }
  });
  return {
    sheet: sheet,
    previousRowCount: snapshot.values.length,
    keptRows: keptRows,
    removedCount: removedCount,
    affectedRouteIds: Object.keys(affectedRouteIds)
  };
}

function commitPreparedRoutePinDeletion_(prepared) {
  if (!prepared || prepared.removedCount === 0) return;
  rewriteFixedWidthDataRows_(
    prepared.sheet,
    ROUTE_PINS_HEADERS.length,
    prepared.previousRowCount,
    prepared.keptRows
  );
}

function deletePinRelationsAndCaches_(pinIds) {
  if (!Array.isArray(pinIds) || pinIds.length === 0) return [];
  const pinIdSet = {};
  pinIds.forEach(function(pinId) {
    const normalizedPinId = normalizeRoutePinId_(pinId);
    if (normalizedPinId) pinIdSet[normalizedPinId] = true;
  });
  if (Object.keys(pinIdSet).length === 0) return [];

  const prepared = prepareRoutePinDeletionForPinIdSet_(pinIdSet);
  // cacheを先に消し、後段失敗時もroute_pinsから再試行対象routeを再構築できるようにする。
  invalidateRouteCacheForRoutes_(prepared.affectedRouteIds);
  commitPreparedRoutePinDeletion_(prepared);
  return prepared.affectedRouteIds;
}

function findRouteIdsByPinIds_(pinIds) {
  if (!Array.isArray(pinIds) || pinIds.length === 0) return [];

  const pinIdSet = {};
  pinIds.forEach(function(pinId) {
    const normalizedPinId = normalizeRoutePinId_(pinId);
    if (normalizedPinId) pinIdSet[normalizedPinId] = true;
  });
  if (Object.keys(pinIdSet).length === 0) return [];

  const sheet = openRoutePinsSheet_();
  if (sheet.getLastRow() < 2) return [];

  const routeIds = {};
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i += 1) {
    const pinId = normalizeRoutePinId_(rows[i][1]);
    const routeId = normalizeRouteId_(rows[i][0]);
    if (pinId && pinIdSet[pinId] && routeId) {
      routeIds[routeId] = true;
    }
  }
  return Object.keys(routeIds);
}

function normalizeRouteCacheKey_(value) {
  return String(value || '').trim();
}

function normalizeRouteCacheProvider_(value) {
  return String(value || '').trim() || 'unknown';
}

function normalizeRouteCacheCoords_(coords) {
  if (!Array.isArray(coords)) return [];
  return coords.map(function(coord) {
    const lat = Array.isArray(coord) ? Number(coord[0]) : Number(coord && coord.lat);
    const lng = Array.isArray(coord) ? Number(coord[1]) : Number(coord && coord.lng);
    return Number.isFinite(lat) && Number.isFinite(lng) ? [lat, lng] : null;
  }).filter(Boolean);
}

function routeCacheRowToEntry_(row) {
  if (!row) return null;
  let coords;
  try {
    coords = JSON.parse(String(row[2] || '[]'));
  } catch (_error) {
    return null;
  }
  const normalizedCoords = normalizeRouteCacheCoords_(coords);
  if (normalizedCoords.length < 2) return null;
  return {
    cacheKey: normalizeRouteCacheKey_(row[0]),
    routeId: normalizeRouteId_(row[1]),
    coords: normalizedCoords,
    provider: normalizeRouteCacheProvider_(row[3]),
    createdAt: row[4] ? String(row[4]) : ''
  };
}

function getRouteCache(data) {
  const cacheKey = normalizeRouteCacheKey_(data && data.cacheKey);
  if (!cacheKey) return { ok: false, error: 'missing_cache_key' };

  const sheet = openRouteCacheSheet_();
  if (sheet.getLastRow() < 2) return { ok: false, miss: true };

  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i += 1) {
    if (normalizeRouteCacheKey_(rows[i][0]) !== cacheKey) continue;
    const entry = routeCacheRowToEntry_(rows[i]);
    if (!entry) return { ok: false, miss: true };
    return {
      ok: true,
      cacheKey: entry.cacheKey,
      routeId: entry.routeId,
      coords: entry.coords,
      provider: entry.provider,
      createdAt: entry.createdAt
    };
  }
  return { ok: false, miss: true };
}

function putRouteCache(data) {
  assertEditToken_(data);
  const cacheKey = normalizeRouteCacheKey_(data && data.cacheKey);
  const routeId = normalizeRouteId_(data && data.routeId);
  const provider = normalizeRouteCacheProvider_(data && data.provider);
  const coords = normalizeRouteCacheCoords_(data && data.coords);
  if (!cacheKey) return { ok: false, error: 'missing_cache_key' };
  if (!routeId) return { ok: false, error: 'missing_route_id' };
  if (coords.length < 2) return { ok: false, error: 'invalid_coords' };

  return withSpreadsheetMutationLock_(function() {
    const sheet = openRouteCacheSheet_();
    const createdAt = currentUpdatedAt_();
    const row = [cacheKey, routeId, JSON.stringify(coords), provider, createdAt, ''];
    const rows = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
    for (var i = 1; i < rows.length; i += 1) {
      if (normalizeRouteCacheKey_(rows[i][0]) === cacheKey) {
        sheet.getRange(i + 1, 1, 1, ROUTE_CACHE_HEADERS.length).setValues([row]);
        return { ok: true, cacheKey: cacheKey, routeId: routeId, provider: provider, createdAt: createdAt };
      }
    }
    sheet.appendRow(row);
    return { ok: true, cacheKey: cacheKey, routeId: routeId, provider: provider, createdAt: createdAt };
  });
}

function deleteRouteCacheRowsForRouteIds_(routeIds) {
  if (!Array.isArray(routeIds) || routeIds.length === 0) return 0;

  const routeIdSet = {};
  routeIds.forEach(function(routeId) {
    const normalizedRouteId = normalizeRouteId_(routeId);
    if (normalizedRouteId) routeIdSet[normalizedRouteId] = true;
  });
  if (Object.keys(routeIdSet).length === 0) return 0;

  const sheet = openRouteCacheSheet_();
  const result = filterAndRewriteFixedWidthDataRows_(
    sheet,
    ROUTE_CACHE_HEADERS.length,
    function(row) { return Boolean(routeIdSet[normalizeRouteId_(row[1])]); }
  );
  return result.removedRows.length;
}

function invalidateRouteCacheForRoutes_(routeIds) {
  return deleteRouteCacheRowsForRouteIds_(routeIds);
}

function invalidateRouteCacheForPins_(pinIds) {
  return invalidateRouteCacheForRoutes_(findRouteIdsByPinIds_(pinIds));
}

function invalidateRouteCacheForPin(data) {
  assertEditToken_(data);
  const pinId = normalizeRoutePinId_(data && data.pinId);
  if (!pinId) return { ok: false, error: 'missing_pin_id' };
  return withSpreadsheetMutationLock_(function() {
    return { ok: true, deleted: invalidateRouteCacheForPins_([pinId]) };
  });
}

function invalidateRouteCacheForRoute(data) {
  assertEditToken_(data);
  const routeId = normalizeRouteId_(data && data.routeId);
  if (!routeId) return { ok: false, error: 'missing_route_id' };
  return withSpreadsheetMutationLock_(function() {
    return { ok: true, deleted: invalidateRouteCacheForRoutes_([routeId]) };
  });
}

function getRouteCacheSheetForRead_() {
  return getRequiredSheet_(ROUTE_CACHE_SHEET_NAME);
}

function parseRouteCacheTimestamp_(value) {
  if (value instanceof Date && Number.isFinite(value.getTime())) return value.getTime();
  const parsed = Date.parse(String(value || ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

function readLatestRouteCacheEntryForRoute_(routeId) {
  const normalizedRouteId = normalizeRouteId_(routeId);
  if (!normalizedRouteId) return null;

  const sheet = getRouteCacheSheetForRead_();
  if (!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getDataRange().getValues();
  let latestRow = null;
  let latestTimestamp = -1;
  let latestIndex = -1;
  for (var i = 1; i < rows.length; i += 1) {
    if (normalizeRouteId_(rows[i][1]) !== normalizedRouteId) continue;
    const timestamp = parseRouteCacheTimestamp_(rows[i][4]);
    if (!latestRow || timestamp > latestTimestamp || (timestamp === latestTimestamp && i > latestIndex)) {
      latestRow = rows[i];
      latestTimestamp = timestamp;
      latestIndex = i;
    }
  }
  return latestRow ? routeCacheRowToEntry_(latestRow) : null;
}

function readLatestRouteCacheEntryByCacheKey_(cacheKey) {
  const normalizedCacheKey = normalizeRouteCacheKey_(cacheKey);
  if (!normalizedCacheKey) return null;

  const sheet = getRouteCacheSheetForRead_();
  if (!sheet || sheet.getLastRow() < 2) return null;

  const rows = sheet.getDataRange().getValues();
  let latestRow = null;
  let latestTimestamp = -1;
  let latestIndex = -1;
  for (var i = 1; i < rows.length; i += 1) {
    if (normalizeRouteCacheKey_(rows[i][0]) !== normalizedCacheKey) continue;
    const timestamp = parseRouteCacheTimestamp_(rows[i][4]);
    if (!latestRow || timestamp > latestTimestamp || (timestamp === latestTimestamp && i > latestIndex)) {
      latestRow = rows[i];
      latestTimestamp = timestamp;
      latestIndex = i;
    }
  }
  return latestRow ? routeCacheRowToEntry_(latestRow) : null;
}

function logSharedRoadRouteCache_(routeId, group, expectedCacheKey, hit, reason) {
  if (typeof Logger === 'undefined' || !Logger.log) return;
  Logger.log('shared_road_route_cache: routeId=' + normalizeRouteId_(routeId)
    + ' routeMode=' + String(group && group.routeMode || '')
    + ' expectedCacheKey=' + normalizeRouteCacheKey_(expectedCacheKey)
    + ' cache ' + (hit ? 'hit' : 'miss')
    + (reason ? ' miss reason=' + reason : ''));
}

function getSharedRoutePinIdsForDisplay_(group) {
  if (!group || !Array.isArray(group.pinIds)) return [];
  const basePinIds = group.pinIds.map(function(pinId) {
    return normalizeRoutePinId_(pinId);
  }).filter(Boolean);

  const seen = {};
  return basePinIds.filter(function(pinId) {
    if (seen[pinId]) return false;
    seen[pinId] = true;
    return true;
  });
}

function roundSharedRouteCacheCoord_(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return '';
  const rounded = Math.round(num * 100000) / 100000;
  return (Object.is(rounded, -0) ? 0 : rounded).toFixed(5);
}

function buildSharedRoadRouteCacheKey_(group, pinById, provider) {
  if (!group || group.routeMode !== 'road') return '';
  const entries = getSharedRoutePinIdsForDisplay_(group).map(function(pinId) {
    const pin = pinById[pinId];
    if (!pin) return null;
    const lat = Number(pin.lat);
    const lng = Number(pin.lng);
    if (!Number.isFinite(lat) || !Number.isFinite(lng)) return null;
    return { pinId: pinId, latLng: [lat, lng] };
  }).filter(Boolean);
  if (entries.length < 2) return '';

  const waypointKey = entries.map(function(entry) {
    return encodeURIComponent(entry.pinId) + ':'
      + roundSharedRouteCacheCoord_(entry.latLng[0]) + ','
      + roundSharedRouteCacheCoord_(entry.latLng[1]);
  }).join('>');
  return [
    normalizeRouteCacheProvider_(provider),
    'road',
    group.closed === true ? 'true' : 'false',
    waypointKey
  ].join('|');
}

function getSharedPinsForShareLink_(shareLink, routeGroups) {
  return getMapPinsForShare_().filter(function(pin) {
    return matchesTagFilter_(pin, shareLink.tags, shareLink.tagMode)
      && matchesColorFilter_(pin, shareLink.colors);
  }).map(function(pin) {
    return toSharedPin_(pin, shareLink.tags);
  });
}

function isShareRouteIdAllowed_(routeId, routeIds) {
  routeId = normalizeRouteId_(routeId);
  if (!routeId) return false;
  if (isShareRouteSelectionNone_(routeIds)) return false;
  if (!routeIds || !routeIds.length) return true;
  return routeIds.indexOf(routeId) !== -1;
}

function filterSharedRouteGroupsForShareLink_(routeGroups, shareLink) {
  const routeIds = shareLink && Array.isArray(shareLink.routeIds) ? shareLink.routeIds : [];
  if (isShareRouteSelectionNone_(routeIds)) return [];
  if (!routeIds.length) return routeGroups;
  return routeGroups.filter(function(group) {
    return isShareRouteIdAllowed_(group && (group.routeId || group.id), routeIds);
  });
}

function buildSharedAllowedPinIdSet_(pins) {
  const allowedPinIdSet = {};
  (Array.isArray(pins) ? pins : []).forEach(function(pin) {
    const pinId = normalizeRoutePinId_(pin && pin.id);
    if (pinId) allowedPinIdSet[pinId] = true;
  });
  return allowedPinIdSet;
}

function indexSharedPinsById_(pins) {
  const pinById = {};
  (Array.isArray(pins) ? pins : []).forEach(function(pin) {
    const pinId = normalizeRoutePinId_(pin && pin.id);
    if (pinId) pinById[pinId] = pin;
  });
  return pinById;
}

function isRouteClosedToAllowedPins_(group, allowedPinIdSet) {
  const pinIds = getSharedRoutePinIdsForDisplay_(group);
  if (pinIds.length < 2) return false;
  for (var i = 0; i < pinIds.length; i += 1) {
    if (!allowedPinIdSet[pinIds[i]]) return false;
  }
  return true;
}

function getSharedRoadRouteCache_(token, routeId) {
  const shareLink = getShareLinkByToken_(token);
  if (!shareLink) return { ok: false };
  if (!shareLink.enabled || shareLink.revokedAt) return { ok: false };

  routeId = normalizeRouteId_(routeId);
  if (!routeId) return { ok: false };
  if (!isShareRouteIdAllowed_(routeId, shareLink.routeIds)) return { ok: false };
  function miss(reason, group, expectedCacheKey) {
    logSharedRoadRouteCache_(routeId, group, expectedCacheKey, false, reason);
    return { ok: false };
  }

  const allRouteGroups = getRouteGroups();
  const sharedPins = getSharedPinsForShareLink_(shareLink, allRouteGroups);
  const allowedPinIdSet = buildSharedAllowedPinIdSet_(sharedPins);
  let rawGroup = null;
  for (var rawIndex = 0; rawIndex < allRouteGroups.length; rawIndex += 1) {
    if (normalizeRouteId_(allRouteGroups[rawIndex].routeId) === routeId) {
      rawGroup = allRouteGroups[rawIndex];
      break;
    }
  }
  if (!rawGroup || !isRouteClosedToAllowedPins_(rawGroup, allowedPinIdSet)) return miss('no_group', rawGroup, '');

  const sharedRouteGroups = getSharedRouteGroups_(sharedPins, allRouteGroups);
  let group = null;
  for (var i = 0; i < sharedRouteGroups.length; i += 1) {
    if (normalizeRouteId_(sharedRouteGroups[i].routeId) === routeId) {
      group = sharedRouteGroups[i];
      break;
    }
  }
  if (!group) return miss('no_group', null, '');
  if (group.routeMode !== 'road') return miss('not_road', group, '');
  if (!isRouteClosedToAllowedPins_(group, allowedPinIdSet)) return miss('no_group', group, '');

  const pinById = indexSharedPinsById_(sharedPins);
  const expectedCacheKey = buildSharedRoadRouteCacheKey_(group, pinById, SHARED_ROAD_ROUTE_CACHE_PROVIDER);
  if (!expectedCacheKey) return miss('no_expected_key', group, expectedCacheKey);

  const entry = readLatestRouteCacheEntryByCacheKey_(expectedCacheKey);
  if (!entry) return miss('no_cache', group, expectedCacheKey);
  if (entry.routeId !== routeId) return miss('route_id_mismatch', group, expectedCacheKey);
  if (!Array.isArray(entry.coords) || entry.coords.length < 2) return miss('invalid_coords', group, expectedCacheKey);

  logSharedRoadRouteCache_(routeId, group, expectedCacheKey, true, '');
  return { ok: true, routeId: routeId, coords: entry.coords };
}

function getSharedRoadRouteCache(data, routeId) {
  const payload = data && typeof data === 'object'
    ? data
    : { token: data, routeId: routeId };
  try {
    return getSharedRoadRouteCache_(payload && payload.token, payload && payload.routeId);
  } catch (error) {
    if (typeof Logger !== 'undefined' && Logger.log) {
      Logger.log('shared_road_route_cache_failed: ' + (error && error.message ? error.message : error));
    }
    return { ok: false };
  }
}

function deleteRouteGroup(data) {
  assertEditToken_(data);
  const routeId = normalizeRouteId_(data && typeof data === 'object' ? (data.routeId || data.id) : data);
  if (!routeId) return { ok: false, error: 'missing_route_id' };

  return withSpreadsheetMutationLock_(function() {
    const sheet = openRoutesSheet_();
    const existing = findRouteSheetRow_(routeId);
    if (!existing) {
      invalidateRouteCacheForRoutes_([routeId]);
      deleteRoutePinsForRoute_(routeId);
      return { ok: false, error: 'route_not_found' };
    }

    sheet.deleteRow(existing.rowNumber);
    invalidateRouteCacheForRoutes_([routeId]);
    deleteRoutePinsForRoute_(routeId);
    return { ok: true };
  });
}

function updateRoutesOrder(data) {
  assertEditToken_(data);
  if (!data || !Array.isArray(data.orderedIds)) {
    return { ok: false, error: 'ordered_ids_required' };
  }

  const seen = {};
  const orderedIds = [];
  for (var i = 0; i < data.orderedIds.length; i += 1) {
    const routeId = normalizeRouteId_(data.orderedIds[i]);
    if (!routeId) return { ok: false, error: 'missing_route_id' };
    if (seen[routeId]) return { ok: false, error: 'duplicate_route_id', routeId: routeId };
    seen[routeId] = true;
    orderedIds.push(routeId);
  }

  return withSpreadsheetMutationLock_(function() {
    const sheet = openRoutesSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      if (orderedIds.length > 0) {
        return { ok: false, error: 'route_not_found', routeId: orderedIds[0] };
      }
      return { ok: true, routeGroups: getRouteGroups() };
    }

    const routeValues = sheet.getRange(2, 1, lastRow - 1, ROUTES_HEADERS.length).getValues();
    const entries = [];
    const byId = {};
    routeValues.forEach(function(row, index) {
      const routeId = normalizeRouteId_(row[0]);
      if (!routeId) return;
      const entry = { routeId: routeId, rowOffset: index };
      entries.push(entry);
      byId[routeId] = entry;
    });
    for (var j = 0; j < orderedIds.length; j += 1) {
      if (!byId[orderedIds[j]]) {
        return { ok: false, error: 'route_not_found', routeId: orderedIds[j] };
      }
    }
    if (entries.length === 0) {
      return { ok: true, routeGroups: getRouteGroups() };
    }

    const orderedEntries = orderedIds.map(function(routeId) {
      return byId[routeId];
    });
    entries.forEach(function(entry) {
      if (!seen[entry.routeId]) orderedEntries.push(entry);
    });

    const orderRange = sheet.getRange(2, 10, lastRow - 1, 1);
    const orderValues = orderRange.getValues();
    const orderFormulas = orderRange.getFormulas();
    const output = orderValues.map(function(row, index) {
      return [orderFormulas[index][0] || row[0]];
    });
    orderedEntries.forEach(function(entry, index) {
      output[entry.rowOffset][0] = index;
    });
    orderRange.setValues(output);

    return { ok: true, routeGroups: getRouteGroups() };
  });
}

function buildSharedViewUrl_(token) {
  return getConfiguredWebAppUrl_() + '?view=shared&token=' + encodeURIComponent(token);
}

function createShareLink(data) {
  assertEditToken_(data);
  const label = normalizeShareLinkLabel_(data && data.label);
  const tags = PinData.normalizeTags(data && data.tags || []);
  const tagMode = String(data && data.tagMode || 'or') === 'and' ? 'and' : 'or';
  const colors = normalizeShareColors_(data && data.colors || []);
  const routeIds = normalizeShareRouteIds_(data && data.routeIds || []);
  const sheet = openShareLinksSheet_();
  const token = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  sheet.appendRow([
    createdAt,
    label,
    token,
    PinData.serializeTags(tags),
    tagMode,
    true,
    '',
    serializeShareColors_(colors),
    serializeShareRouteIds_(routeIds)
  ]);

  return {
    ok: true,
    token: token,
    url: buildSharedViewUrl_(token),
    shareLink: {
      createdAt: createdAt,
      label: label,
      token: token,
      tags: tags,
      tagMode: tagMode,
      colors: colors,
      routeIds: routeIds,
      enabled: true,
      revokedAt: ''
    }
  };
}

function listShareLinks(data) {
  assertEditToken_(data);
  const rows = openShareLinksSheet_().getDataRange().getValues();
  const items = rows.slice(1).map(shareRowToLink_).reverse().map(function(item) {
    item.url = buildSharedViewUrl_(item.token);
    return item;
  });
  return { ok: true, items: items };
}

function getShareLinkByToken_(token) {
  var normalizedToken = normalizeShareToken_(token);
  const rows = openShareLinksSheet_().getDataRange().getValues();
  for (var i = 1; i < rows.length; i += 1) {
    if (String(rows[i][2]) === normalizedToken) {
      return shareRowToLink_(rows[i]);
    }
  }
  return null;
}

function setShareLinkEnabled(data) {
  assertEditToken_(data);
  var normalizedToken = normalizeShareToken_(typeof data === 'object' && data !== null ? data.token : data);
  var enabled = !!(data && typeof data === 'object' ? data.enabled : false);
  const sheet = openShareLinksSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i += 1) {
    if (String(rows[i][2]) !== normalizedToken) continue;
    sheet.getRange(i + 1, 6).setValue(enabled);
    sheet.getRange(i + 1, 7).setValue(enabled ? '' : new Date().toISOString());
    return { ok: true };
  }
  return { ok: false, error: 'token not found' };
}

function deleteShareLink(data) {
  assertEditToken_(data);
  var normalizedToken = normalizeShareToken_(data && typeof data === 'object' ? data.token : data);
  const sheet = openShareLinksSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i += 1) {
    if (String(rows[i][2]) !== normalizedToken) continue;
    sheet.deleteRow(i + 1);
    return { ok: true };
  }
  return { ok: false, error: 'token not found' };
}

function revokeShareLink(data) {
  assertEditToken_(data);
  var token = data && typeof data === 'object' ? data.token : data;
  return setShareLinkEnabled({ token: token, enabled: false, __editToken: data && data.__editToken });
}

function matchesTagFilter_(pin, tags, mode) {
  if (!tags || !tags.length) return true;
  const pinTags = (pin.tags || []).map(function(tag) {
    return String(tag).toLowerCase();
  });
  if (mode === 'and') {
    return tags.every(function(tag) {
      return pinTags.indexOf(String(tag).toLowerCase()) !== -1;
    });
  }
  return tags.some(function(tag) {
    return pinTags.indexOf(String(tag).toLowerCase()) !== -1;
  });
}

function matchesColorFilter_(pin, colors) {
  if (!colors || !colors.length) return true;
  return colors.indexOf(String(pin.color || '').toLowerCase()) !== -1;
}

function getMapPinsForShare_() {
  const sheet = openMapInfoSheet_();
  if (sheet.getLastRow() === 0) return [];
  return PinData.rowsToPins(sheet.getDataRange().getValues());
}

function filterPinTagsForShare_(pin, allowedTags) {
  if (!allowedTags || !allowedTags.length) {
    return (pin.tags || []).slice();
  }
  var allowed = {};
  allowedTags.forEach(function(tag) {
    allowed[String(tag).toLowerCase()] = true;
  });
  return (pin.tags || []).filter(function(tag) {
    return !!allowed[String(tag).toLowerCase()];
  });
}

function toSharedPin_(pin, allowedTags) {
  return {
    id: pin.id,
    title: pin.title || '',
    description: pin.description || '',
    lat: pin.lat,
    lng: pin.lng,
    color: pin.color || '#e53935',
    icon: PinData.normalizeIcon(pin.icon),
    imageUrl: pin.imageUrl || '',
    timestamp: pin.timestamp || '',
    eventAt: PinData.normalizeEventAt(pin.eventAt),
    links: Array.isArray(pin.links) ? pin.links.slice() : [],
    tags: filterPinTagsForShare_(pin, allowedTags)
  };
}

function toSharedRouteGroup_(group, allowedPinIdSet) {
  var routeId = normalizeRouteId_(group && (group.routeId || group.id));
  if (!routeId) return null;

  var filteredPinIds = [];
  var filteredPinIdSet = {};
  (Array.isArray(group.pinIds) ? group.pinIds : []).forEach(function(pinId) {
    var normalizedPinId = normalizeRoutePinId_(pinId);
    if (!normalizedPinId || !allowedPinIdSet[normalizedPinId] || filteredPinIdSet[normalizedPinId]) return;
    filteredPinIdSet[normalizedPinId] = true;
    filteredPinIds.push(normalizedPinId);
  });
  if (filteredPinIds.length === 0) return null;

  var closed = group && group.closed === true;
  var startPinId = normalizeRoutePinId_(group && group.startPinId);
  var endPinId = normalizeRoutePinId_(group && group.endPinId);
  return {
    id: routeId,
    routeId: routeId,
    name: String(group && group.name || ''),
    color: normalizeRouteColor_(group && group.color, routeId),
    visible: group && group.visible === false ? false : true,
    showNumbers: group && group.showNumbers === false ? false : true,
    showLine: group && group.showLine === false ? false : true,
    lineStyle: normalizeRouteLineStyle_(group && group.lineStyle, routeId),
    routeMode: group && group.routeMode === 'road' ? 'road' : 'straight',
    closed: closed,
    startPinId: startPinId && filteredPinIdSet[startPinId] ? startPinId : null,
    endPinId: closed ? null : (endPinId && filteredPinIdSet[endPinId] ? endPinId : null),
    pinIds: filteredPinIds
  };
}

function getSharedRouteGroups_(pins, routeGroups) {
  var allowedPinIdSet = {};
  (Array.isArray(pins) ? pins : []).forEach(function(pin) {
    var pinId = normalizeRoutePinId_(pin && pin.id);
    if (pinId) allowedPinIdSet[pinId] = true;
  });
  if (Object.keys(allowedPinIdSet).length === 0) return [];

  var sourceRouteGroups = Array.isArray(routeGroups) ? routeGroups : getRouteGroups();
  return sourceRouteGroups.map(function(group) {
    return toSharedRouteGroup_(group, allowedPinIdSet);
  }).filter(function(group) {
    return !!group;
  });
}

function getSharedViewData(token) {
  var shareLink = getShareLinkByToken_(token);
  if (!shareLink) return { ok: false, error: 'invalid_share_link' };
  if (!shareLink.enabled) return { ok: false, error: 'revoked_share_link' };

  var allRouteGroups = getRouteGroups();
  var pins = getSharedPinsForShareLink_(shareLink, allRouteGroups);
  var routeGroups = filterSharedRouteGroupsForShareLink_(getSharedRouteGroups_(pins, allRouteGroups), shareLink);
  var noRoutes = isShareRouteSelectionNone_(shareLink.routeIds);

  return {
    ok: true,
    noRoutes: noRoutes,
    shareLink: {
      label: shareLink.label,
      token: shareLink.token,
      tags: shareLink.tags,
      tagMode: shareLink.tagMode,
      colors: shareLink.colors.slice(),
      routeIds: shareLink.routeIds.slice()
    },
    allowedTags: shareLink.tags.slice(),
    allowedColors: shareLink.colors.slice(),
    allowedRouteIds: noRoutes ? [] : shareLink.routeIds.slice(),
    pins: pins,
    routeGroups: routeGroups
  };
}


function saveMapData(data) {
  assertEditToken_(data);
  if (!data || !String(data.title || '').trim()) {
    return { ok: false, error: 'title is required' };
  }

  const title = String(data.title).trim();
  const description = String(data.description || '');
  const color = data.color || DEFAULT_COLOR;
  const icon = PinData.normalizeIcon(data.icon);
  const eventAt = PinData.normalizeEventAt(data.eventAt);
  const links = PinData.normalizeLinks(data.links || data.referenceUrls || []);
  const status = data.status != null ? PinData.normalizeStatus(String(data.status)) : '未対応';
  const tags = PinData.normalizeTags(data.tags || []);
  const sheet = openMapInfoSheet_();
  const id = Utilities.getUuid();
  const now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'
  );

  let fileId = '';
  let imageUrl = '';
  let folderUrl = '';

  if (data.base64) {
    const mimeMatch = String(data.base64).match(/^data:(image\/\w+);base64,/);
    if (!mimeMatch) return { ok: false, error: 'invalid base64 format' };

    const uploadFolderId = data.targetFolderId || getRootFolderId_();
    if (!uploadFolderId) {
      return { ok: false, error: 'フォルダIDが未設定です。config シートの IMAGE_DRIVE_URL を確認してください。' };
    }

    const base64Clean = String(data.base64).replace(/^data:image\/\w+;base64,/, '');
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Clean),
      mimeMatch[1],
      PinData.buildFileNameForSave(title, data.filename, getRenameFileWithTitle_())
    );

    const folder = DriveApp.getFolderById(uploadFolderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileId = file.getId();
    imageUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1920';
    folderUrl = getDriveFolderUrl_(uploadFolderId);
  }

  sheet.appendRow([
    now,
    title,
    description,
    data.lat != null ? Number(data.lat) : '',
    data.lng != null ? Number(data.lng) : '',
    color,
    fileId,
    imageUrl,
    id,
    PinData.serializeLinks(links),
    status,
    PinData.serializeTags(tags),
    eventAt,
    '',
    icon
  ]);

  return {
    ok: true,
    id: id,
    imageUrl: imageUrl,
    fileId: fileId,
    folderUrl: folderUrl,
    links: links
  };
}

function buildDuplicatePinTitle_(title) {
  const suffix = '（コピー）';
  const base = String(title || '').trim() || '無題';
  const maxBaseLength = Math.max(0, PIN_TITLE_MAX_LENGTH - suffix.length);
  return base.slice(0, maxBaseLength) + suffix;
}

function normalizeDuplicateLocation_(mode, sourcePin, data) {
  if (mode === 'unplaced') {
    return { ok: true, lat: null, lng: null };
  }
  if (mode === 'same') {
    return {
      ok: true,
      lat: sourcePin.lat == null || sourcePin.lng == null ? null : sourcePin.lat,
      lng: sourcePin.lat == null || sourcePin.lng == null ? null : sourcePin.lng
    };
  }
  if (mode === 'point') {
    const lat = Number(data && data.lat);
    const lng = Number(data && data.lng);
    if (!Number.isFinite(lat) || !Number.isFinite(lng) || lat < -90 || lat > 90 || lng < -180 || lng > 180) {
      return { ok: false, error: 'invalid_location' };
    }
    return { ok: true, lat: lat, lng: lng };
  }
  return { ok: false, error: 'invalid_location' };
}

function duplicatePin(data) {
  assertEditToken_(data);
  const sourcePinId = String(data && data.sourcePinId || '').trim();
  const mode = String(data && data.mode || 'unplaced').trim();
  const sheet = openMapInfoSheet_();
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex(function(row) {
    return row && String(row[8] || '') === sourcePinId;
  });
  if (rowIndex === -1) return { ok: false, error: 'pin_not_found' };

  const sourcePin = PinData.rowToPin(rows[rowIndex]);
  const location = normalizeDuplicateLocation_(mode, sourcePin, data);
  if (!location.ok) return { ok: false, error: location.error };

  const title = buildDuplicatePinTitle_(sourcePin.title);
  const description = String(sourcePin.description || '');
  const color = SAFE_COLOR_RE.test(String(sourcePin.color || '')) ? sourcePin.color : DEFAULT_COLOR;
  const icon = PinData.normalizeIcon(sourcePin.icon);
  const eventAt = PinData.normalizeEventAt(sourcePin.eventAt);
  const links = PinData.normalizeLinks(sourcePin.links || []);
  const status = sourcePin.status ? PinData.normalizeStatus(String(sourcePin.status)) : '';
  const tags = PinData.normalizeTags(sourcePin.tags || []);
  const id = Utilities.getUuid();
  const now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss'
  );

  const row = [
    now,
    title,
    description,
    location.lat != null ? location.lat : '',
    location.lng != null ? location.lng : '',
    color,
    '',
    '',
    id,
    PinData.serializeLinks(links),
    status,
    PinData.serializeTags(tags),
    eventAt,
    now,
    icon
  ];
  sheet.appendRow(row);

  const pin = PinData.rowToPin(row);
  pin.folderUrl = '';
  return { ok: true, pin: pin };
}

function updatePinDetails(data) {
  assertEditToken_(data);
  if (!data || !data.id) return { ok: false, error: 'missing id' };
  if (!String(data.title || '').trim()) return { ok: false, error: 'title is required' };

  const title = String(data.title).trim();
  let renameFileId = '';
  const result = withSpreadsheetMutationLock_(function() {
    const sheet = openMapInfoSheet_();
    const rowIndex = findPinRowIndex_(sheet, data.id);
    if (rowIndex < 1) return { ok: false, error: 'id not found' };

    const sheetRow = rowIndex + 1;
    const range = sheet.getRange(sheetRow, 1, 1, MAP_INFO_COLUMN_COUNT);
    const row = range.getValues()[0];
    const formulas = range.getFormulas()[0];
    const output = row.map(function(value, index) {
      return formulas[index] || value;
    });
    let links = PinData.normalizeLinks(row[9] || '');

    output[1] = title;
    if (Object.prototype.hasOwnProperty.call(data, 'description')) {
      output[2] = String(data.description || '');
    }
    if (Object.prototype.hasOwnProperty.call(data, 'color')) {
      output[5] = data.color || row[5] || DEFAULT_COLOR;
    }
    if (
      Object.prototype.hasOwnProperty.call(data, 'links') ||
      Object.prototype.hasOwnProperty.call(data, 'referenceUrls')
    ) {
      links = PinData.normalizeLinks(data.links || data.referenceUrls || []);
      output[9] = PinData.serializeLinks(links);
    }
    output[14] = PinData.normalizeIcon(data.icon != null ? data.icon : row[14]);
    if (Object.prototype.hasOwnProperty.call(data, 'eventAt') && data.eventAt != null) {
      output[12] = PinData.normalizeEventAt(data.eventAt);
    }
    if (Object.prototype.hasOwnProperty.call(data, 'status') && data.status != null) {
      output[10] = PinData.normalizeStatus(String(data.status));
    }
    if (Object.prototype.hasOwnProperty.call(data, 'tags') && data.tags != null) {
      output[11] = PinData.serializeTags(data.tags);
    }
    const updatedAt = currentUpdatedAt_();
    output[13] = updatedAt;
    range.setValues([output]);

    if (getRenameFileWithTitle_() && row[6]) {
      renameFileId = String(row[6]);
    }
    return {
      ok: true,
      updatedAt: updatedAt,
      links: links,
      folderUrl: ''
    };
  });

  if (result.ok && renameFileId) {
    renameDriveFileForTitle_(renameFileId, title);
  }
  return result;
}

function movePin(data) {
  assertEditToken_(data);
  if (!data || !data.id) return { ok: false, error: 'missing id' };
  if (data.lat == null || data.lng == null) return { ok: false, error: 'missing lat/lng' };

  return withSpreadsheetMutationLock_(function() {
    const sheet = openMapInfoSheet_();
    const rowIndex = findPinRowIndex_(sheet, data.id);
    if (rowIndex < 1) return { ok: false, error: 'id not found' };

    const sheetRow = rowIndex + 1;
    sheet.getRange(sheetRow, 4, 1, 2).setValues([[Number(data.lat), Number(data.lng)]]);
    sheet.getRange(sheetRow, MAP_INFO_UPDATED_AT_COLUMN).setValue(currentUpdatedAt_());
    invalidateRouteCacheForPins_([data.id]);
    return { ok: true };
  });
}

function unplacePin(data) {
  assertEditToken_(data);
  if (!data || !data.id) return { ok: false, error: 'missing id' };

  return withSpreadsheetMutationLock_(function() {
    const sheet = openMapInfoSheet_();
    const rowIndex = findPinRowIndex_(sheet, data.id);
    if (rowIndex < 1) return { ok: false, error: 'id not found' };

    const sheetRow = rowIndex + 1;
    sheet.getRange(sheetRow, 4, 1, 2).setValues([['', '']]);
    sheet.getRange(sheetRow, MAP_INFO_UPDATED_AT_COLUMN).setValue(currentUpdatedAt_());
    return { ok: true };
  });
}

function bulkUpdatePinStatus(data) {
  assertEditToken_(data);
  if (!data || !Array.isArray(data.ids) || data.ids.length === 0) {
    return { ok: false, error: 'ids must be a non-empty array' };
  }
  let status;
  try {
    status = PinData.normalizeStatus(String(data.status || ''));
  } catch (_e) {
    return { ok: false, error: 'invalid status: ' + data.status };
  }
  if (!status) return { ok: false, error: 'status is required' };
  const targetIds = new Set(data.ids);

  return withSpreadsheetMutationLock_(function() {
    const sheet = openMapInfoSheet_();
    const rows = sheet.getDataRange().getValues();
    const statusRanges = [];
    const updatedAtRanges = [];
    const updatedAt = currentUpdatedAt_();

    for (var rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      if (!targetIds.has(rows[rowIndex][8])) continue;
      const sheetRow = rowIndex + 1;
      statusRanges.push('K' + sheetRow);
      updatedAtRanges.push('N' + sheetRow);
    }
    if (statusRanges.length === 0) return { ok: true, updatedCount: 0 };

    for (var offset = 0; offset < statusRanges.length; offset += RANGE_LIST_CHUNK_SIZE) {
      sheet.getRangeList(statusRanges.slice(offset, offset + RANGE_LIST_CHUNK_SIZE)).setValue(status);
      sheet.getRangeList(updatedAtRanges.slice(offset, offset + RANGE_LIST_CHUNK_SIZE)).setValue(updatedAt);
    }
    return { ok: true, updatedCount: statusRanges.length };
  });
}

function logPinDeleteStage_(pinId, stage) {
  if (typeof Logger === 'undefined' || !Logger.log) return;
  const safePinId = String(pinId || '').replace(/[\r\n]/g, ' ');
  Logger.log('pin_delete: stage=' + String(stage || '') + ' pinId=' + safePinId);
}

function normalizePinDeleteRequestIds_(requestedIds) {
  const seen = {};
  const result = [];
  (Array.isArray(requestedIds) ? requestedIds : []).forEach(function(id) {
    const pinId = String(id == null ? '' : id);
    if (!pinId || seen[pinId]) return;
    seen[pinId] = true;
    result.push(pinId);
  });
  return result;
}

function buildPinDeleteSnapshots_(rows, requestedIds) {
  const targetIds = new Set(normalizePinDeleteRequestIds_(requestedIds));
  const seen = {};
  const snapshots = [];
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const pinId = String(row[8] == null ? '' : row[8]);
    if (!targetIds.has(pinId) || seen[pinId]) continue;
    seen[pinId] = true;
    snapshots.push({
      pinId: pinId,
      fileId: String(row[6] || ''),
      originalRowNumber: rowIndex + 1
    });
  }
  return snapshots.sort(function(a, b) {
    return b.originalRowNumber - a.originalRowNumber;
  });
}

function indexCurrentPinDeleteRows_(rows) {
  const byId = {};
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const pinId = String(row[8] == null ? '' : row[8]);
    if (!pinId || byId[pinId]) continue;
    byId[pinId] = {
      pinId: pinId,
      fileId: String(row[6] || ''),
      rowNumber: rowIndex + 1
    };
  }
  return byId;
}

function groupContiguousPinDeleteRows_(entries) {
  const sorted = entries.slice().sort(function(a, b) {
    return a.rowNumber - b.rowNumber;
  });
  const runs = [];
  sorted.forEach(function(entry) {
    const current = runs[runs.length - 1];
    if (current && entry.rowNumber === current.startRow + current.entries.length) {
      current.entries.push(entry);
      return;
    }
    runs.push({ startRow: entry.rowNumber, entries: [entry] });
  });
  return runs.reverse();
}

function logDriveSuccessesAfterLockFailure_(snapshots) {
  snapshots.forEach(function(snapshot) {
    if (snapshot.fileId) {
      logPinDeleteStage_(snapshot.pinId, 'spreadsheet_lock_failed_after_drive');
    }
  });
}

function deletePin(data) {
  assertEditToken_(data);
  if (!data || !data.id) return { ok: false, error: 'missing id' };

  const sheet = openMapInfoSheet_();
  const rows = sheet.getDataRange().getValues();
  const snapshots = buildPinDeleteSnapshots_(rows, [data.id]);
  if (snapshots.length === 0) {
    const missingPinId = String(data.id);
    return withSpreadsheetMutationLock_(function() {
      const currentRows = sheet.getDataRange().getValues();
      const currentById = indexCurrentPinDeleteRows_(currentRows);
      if (currentById[missingPinId]) {
        logPinDeleteStage_(missingPinId, 'appeared_after_snapshot');
        return { ok: false, error: PIN_DELETE_CONFLICT_ERROR };
      }
      deletePinRelationsAndCaches_([missingPinId]);
      return { ok: false, error: 'id not found' };
    });
  }
  const snapshot = snapshots[0];

  if (snapshot.fileId) {
    try {
      DriveApp.getFileById(snapshot.fileId).setTrashed(true);
    } catch (error) {
      return { ok: false, error: '写真の削除に失敗しました: ' + error.message };
    }
  }

  const result = withSpreadsheetMutationLock_(function() {
    const currentRows = sheet.getDataRange().getValues();
    const currentById = indexCurrentPinDeleteRows_(currentRows);
    const current = currentById[snapshot.pinId];
    if (!current) {
      deletePinRelationsAndCaches_([snapshot.pinId]);
      return { ok: true };
    }
    if (current.fileId !== snapshot.fileId) {
      logPinDeleteStage_(snapshot.pinId, 'file_id_conflict');
      return { ok: false, error: PIN_DELETE_CONFLICT_ERROR };
    }

    sheet.deleteRow(current.rowNumber);
    deletePinRelationsAndCaches_([snapshot.pinId]);
    return { ok: true };
  });
  if (result && result.error === SPREADSHEET_MUTATION_BUSY_ERROR) {
    logDriveSuccessesAfterLockFailure_([snapshot]);
  }
  return result;
}

function bulkDeletePins(data) {
  assertEditToken_(data);
  if (!data || !Array.isArray(data.ids) || data.ids.length === 0) {
    return { ok: false, error: 'ids must be a non-empty array' };
  }

  const sheet = openMapInfoSheet_();
  const rows = sheet.getDataRange().getValues();
  const requestedPinIds = normalizePinDeleteRequestIds_(data.ids);
  const snapshots = buildPinDeleteSnapshots_(rows, data.ids);
  const driveSuccessfulSnapshots = [];
  const failedIdSet = {};

  snapshots.forEach(function(snapshot) {
    try {
      if (snapshot.fileId) {
        DriveApp.getFileById(snapshot.fileId).setTrashed(true);
      }
      driveSuccessfulSnapshots.push(snapshot);
    } catch (_error) {
      failedIdSet[snapshot.pinId] = true;
      logPinDeleteStage_(snapshot.pinId, 'drive_delete_failed');
    }
  });

  const snapshotIdSet = {};
  snapshots.forEach(function(snapshot) { snapshotIdSet[snapshot.pinId] = true; });
  const initiallyMissingPinIds = requestedPinIds.filter(function(pinId) {
    return !snapshotIdSet[pinId];
  });
  if (driveSuccessfulSnapshots.length === 0 && initiallyMissingPinIds.length === 0) {
    return {
      ok: true,
      deletedCount: 0,
      failedIds: snapshots.filter(function(snapshot) {
        return failedIdSet[snapshot.pinId];
      }).map(function(snapshot) { return snapshot.pinId; })
    };
  }

  const result = withSpreadsheetMutationLock_(function() {
    const currentRows = sheet.getDataRange().getValues();
    const currentById = indexCurrentPinDeleteRows_(currentRows);
    const currentEntries = [];
    driveSuccessfulSnapshots.forEach(function(snapshot) {
      const current = currentById[snapshot.pinId];
      if (!current) return;
      if (current.fileId !== snapshot.fileId) {
        failedIdSet[snapshot.pinId] = true;
        logPinDeleteStage_(snapshot.pinId, 'file_id_conflict');
        return;
      }
      currentEntries.push({
        pinId: snapshot.pinId,
        fileId: current.fileId,
        rowNumber: current.rowNumber
      });
    });
    const deletedIds = [];

    groupContiguousPinDeleteRows_(currentEntries).forEach(function(run) {
      try {
        sheet.deleteRows(run.startRow, run.entries.length);
        run.entries.forEach(function(entry) { deletedIds.push(entry.pinId); });
      } catch (_error) {
        run.entries.forEach(function(entry) {
          failedIdSet[entry.pinId] = true;
          logPinDeleteStage_(entry.pinId, 'spreadsheet_delete_failed');
        });
      }
    });

    const cleanupIds = deletedIds.slice();
    requestedPinIds.forEach(function(pinId) {
      if (!currentById[pinId] && cleanupIds.indexOf(pinId) === -1) {
        cleanupIds.push(pinId);
      }
    });
    if (cleanupIds.length > 0) {
      deletePinRelationsAndCaches_(cleanupIds);
    }

    return {
      ok: true,
      deletedCount: deletedIds.length,
      failedIds: snapshots.filter(function(snapshot) {
        return failedIdSet[snapshot.pinId];
      }).map(function(snapshot) { return snapshot.pinId; })
    };
  });
  if (result && result.error === SPREADSHEET_MUTATION_BUSY_ERROR) {
    logDriveSuccessesAfterLockFailure_(driveSuccessfulSnapshots);
  }
  return result;
}

function getAppSettings() {
  const config = getAppConfig_();
  const rootFolderId = extractDriveFolderId_(config.IMAGE_DRIVE_URL || '');
  return {
    ok: true,
    rootFolderId: rootFolderId,
    rootFolderUrl: getDriveFolderUrl_(rootFolderId),
    renameFileWithTitle: PinData.toBooleanSetting(config.RENAME_FILE_WITH_TITLE)
  };
}

function updateAppSettings(data) {
  assertEditToken_(data);
  if (!data) return { ok: false, error: 'missing data' };
  setConfigValue_('RENAME_FILE_WITH_TITLE', data.renameFileWithTitle ? 'true' : 'false');
  return getAppSettings();
}

// 旧フロント互換用
function updatePin(data) {
  assertEditToken_(data);
  const detailResult = updatePinDetails(data);
  if (!detailResult.ok) return detailResult;
  if (data.lat == null || data.lng == null) return detailResult;
  return movePin(data);
}

// ============================================================
//  Drive 補助
// ============================================================

function buildFileNameForSave_(title, originalName, shouldSync) {
  return PinData.buildFileNameForSave(title, originalName, shouldSync);
}

function renameDriveFileForTitle_(fileId, title) {
  const file = DriveApp.getFileById(fileId);
  file.setName(buildFileNameForSave_(title, file.getName(), true));
}

function getParentFolderUrlByFileId_(fileId) {
  const parents = DriveApp.getFileById(fileId).getParents();
  if (!parents.hasNext()) return '';
  return getDriveFolderUrl_(parents.next().getId());
}

function enrichPinWithDriveMeta_(pin) {
  const enriched = {};
  Object.keys(pin).forEach(function(key) {
    enriched[key] = pin[key];
  });
  enriched.folderUrl = '';
  return enriched;
}

// ============================================================
//  テスト補助
// ============================================================

function testSaveMapData() {
  const result = saveMapData({
    base64: 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/wAALCAABAAEBAREA/8QAFAABAAAAAAAAAAAAAAAAAAAACf/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEAAD8AVIP/2Q==',
    filename: 'test.jpg',
    title: 'テスト',
    description: 'テスト投稿',
    lat: 35.6812,
    lng: 139.7671,
    color: '#e53935',
    links: ['https://example.com']
  });
  Logger.log(JSON.stringify(result));
}

function testUpdatePin() {
  const pins = getMapData();
  if (pins.length === 0) {
    Logger.log('no pins');
    return;
  }

  const result = updatePinDetails({
    id: pins[0].id,
    title: '更新テスト',
    description: '説明更新',
    color: '#4caf50',
    links: ['https://example.com/updated']
  });
  Logger.log(JSON.stringify(result));
}

function testRouteCRUD() {
  const pins = getMapData();
  const routeId = 'test-route-' + Utilities.getUuid();
  const createResult = saveRouteGroup({
    routeId: routeId,
    name: '  route CRUD test  ',
    color: 'invalid-color',
    routeMode: 'invalid-mode',
    closed: true,
    endPinId: pins[0] && pins[0].id
  });
  Logger.log('create: ' + JSON.stringify(createResult));

  if (createResult.ok && pins.length > 0) {
    Logger.log('setRoutePins: ' + JSON.stringify(setRoutePins({
      routeId: routeId,
      pinIds: pins.slice(0, Math.min(2, pins.length)).map(function(pin) { return pin.id; })
    })));
  }

  Logger.log('groups: ' + JSON.stringify(getRouteGroups()));
  Logger.log('order: ' + JSON.stringify(updateRoutesOrder({ orderedIds: [routeId] })));
  Logger.log('delete: ' + JSON.stringify(deleteRouteGroup(routeId)));
}

function debugSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty(DATA_SPREADSHEET_ID_KEY);
  Logger.log('ScriptProperties ID: ' + savedId);
  try {
    const ss = resolveDataSpreadsheet_();
    Logger.log('resolved: ' + ss.getName() + ' (' + ss.getId() + ')');
    const sheet = ss.getSheetByName(SHEET_NAME);
    Logger.log('map_info sheet: ' + (sheet ? 'あり' : 'なし'));
    if (sheet) {
      Logger.log('lastRow: ' + sheet.getLastRow());
      if (sheet.getLastRow() > 0) {
        Logger.log('row1: ' + JSON.stringify(sheet.getRange(1, 1, 1, 10).getValues()[0]));
        if (sheet.getLastRow() > 1) {
          Logger.log('row2: ' + JSON.stringify(sheet.getRange(2, 1, 1, 10).getValues()[0]));
        }
      }
    }
  } catch (e) {
    Logger.log('エラー: ' + e.message);
  }
}

if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    PinData: PinData,
    saveMapData: saveMapData,
    updatePinDetails: updatePinDetails,
    movePin: movePin,
    unplacePin: unplacePin,
    bulkUpdatePinStatus: bulkUpdatePinStatus,
    deletePin: deletePin,
    bulkDeletePins: bulkDeletePins,
    createShareLink: createShareLink,
    listShareLinks: listShareLinks,
    revokeShareLink: revokeShareLink,
    setShareLinkEnabled: setShareLinkEnabled,
    deleteShareLink: deleteShareLink,
    getSharedViewData: getSharedViewData,
    getSharedRoadRouteCache: getSharedRoadRouteCache,
    getRouteGroups: getRouteGroups,
    saveRouteGroup: saveRouteGroup,
    deleteRouteGroup: deleteRouteGroup,
    setRoutePins: setRoutePins,
    updateRoutesOrder: updateRoutesOrder,
    getRouteCache: getRouteCache,
    putRouteCache: putRouteCache,
    invalidateRouteCacheForPin: invalidateRouteCacheForPin,
    invalidateRouteCacheForRoute: invalidateRouteCacheForRoute,
    testRouteCRUD: testRouteCRUD,
    setupSheet: setupSheet
  };
}

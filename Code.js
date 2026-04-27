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
    const baseName = shouldSync && String(title || '').trim()
      ? String(title).trim()
      : String(originalName || 'image');
    const extensionMatch = String(originalName || '').match(/(\.[^.]+)$/);
    const extension = extensionMatch ? extensionMatch[1] : '';
    return extension && !/\.[^.]+$/.test(baseName) ? baseName + extension : baseName;
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
const SHARE_LINKS_HEADERS = ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors'];
const ROUTES_HEADERS = ['routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'];
const ROUTE_PINS_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];
const SHARED_ROAD_ROUTE_CACHE_PROVIDER = 'osrm';
const SAFE_COLOR_RE = /^#[0-9a-fA-F]{6}$/;
const ROUTE_LINE_STYLES = { solid: true, dashed: true, dotted: true };

// ============================================================
//  メニュー / 初期設定
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('設定')
    .addItem('初期設定', 'setupSheet')
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
  template.execUrl = ScriptApp.getService().getUrl();
  template.token = params.token || '';
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
    sheet.appendRow([key, value, description]);
    return;
  }

  const row = index + 2;
  if (sheet.getRange(row, 2).getValue() === '') {
    sheet.getRange(row, 2).setValue(value);
  }
  if (sheet.getRange(row, 3).getValue() === '') {
    sheet.getRange(row, 3).setValue(description);
  }
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

function openShareLinksSheet_() {
  return ensureShareLinksSheet_(openDataSpreadsheet_());
}

function openRoutesSheet_() {
  return ensureHeaderSheet_(openDataSpreadsheet_(), ROUTES_SHEET_NAME, ROUTES_HEADERS);
}

function openRoutePinsSheet_() {
  return ensureHeaderSheet_(openDataSpreadsheet_(), ROUTE_PINS_SHEET_NAME, ROUTE_PINS_HEADERS);
}

function openRouteCacheSheet_() {
  return ensureHeaderSheet_(openDataSpreadsheet_(), ROUTE_CACHE_SHEET_NAME, ROUTE_CACHE_HEADERS);
}

function ensureShareLinksSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SHARE_LINKS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHARE_LINKS_SHEET_NAME);
  }
  sheet.getRange(1, 1, 1, SHARE_LINKS_HEADERS.length).setValues([SHARE_LINKS_HEADERS]);
  sheet.getRange('A1:H1')
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
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
    colors: deserializeShareColors_(row[7] || '')
  };
}

function findPinRowIndex_(sheet, id) {
  const rows = sheet.getDataRange().getValues();
  return rows.findIndex(function(row) {
    return row[8] === id;
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

function navigateToFolder(folderId) {
  return getFolderContents_(folderId);
}

function getRootFolderContents() {
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
  const data = payload || {};
  const requestedRouteId = normalizeRouteId_(data.routeId || data.id);
  if (data.routeId !== undefined && data.routeId !== null && !requestedRouteId) {
    return { ok: false, error: 'missing_route_id' };
  }

  const routeId = requestedRouteId || Utilities.getUuid();
  const existing = requestedRouteId ? findRouteRow_(routeId) : null;

  const nameResult = normalizeRouteName_(data.name, routeId);
  if (!nameResult.ok) return { ok: false, error: nameResult.error };

  const closed = normalizeRouteClosed_(data.closed);
  const rawStartPinId = normalizeRoutePinId_(data.startPinId);
  const rawEndPinId = normalizeRoutePinId_(data.endPinId);
  if (closed && rawEndPinId) {
    logRouteNormalize_('closed route endPinId cleared', routeId);
  }
  const endPinId = closed ? null : rawEndPinId;

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
  const orderIndex = Number(data.orderIndex);
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
  const routeId = normalizeRouteId_(data && data.routeId);
  if (!routeId) return { ok: false, error: 'missing_route_id' };
  if (!findRouteRow_(routeId)) return { ok: false, error: 'route_not_found' };

  const validation = validateRoutePinIds_(data && data.pinIds);
  if (!validation.ok) return validation;

  const sheet = openRoutePinsSheet_();
  const rows = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  for (var i = rows.length - 1; i >= 1; i -= 1) {
    if (normalizeRouteId_(rows[i][0]) === routeId) {
      sheet.deleteRow(i + 1);
    }
  }

  const now = currentUpdatedAt_();
  validation.pinIds.forEach(function(pinId, index) {
    sheet.appendRow([routeId, pinId, index, now, now]);
  });
  return { ok: true, routeId: routeId, pinIds: validation.pinIds };
}

function deleteRoutePinsForRoute_(routeId) {
  const normalizedRouteId = normalizeRouteId_(routeId);
  if (!normalizedRouteId) return [];

  const sheet = openRoutePinsSheet_();
  const rows = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  const removedRouteIds = {};
  for (var i = rows.length - 1; i >= 1; i -= 1) {
    if (normalizeRouteId_(rows[i][0]) === normalizedRouteId) {
      removedRouteIds[normalizedRouteId] = true;
      sheet.deleteRow(i + 1);
    }
  }
  return Object.keys(removedRouteIds);
}

function deleteRoutePinsForPinIds_(pinIds) {
  if (!Array.isArray(pinIds) || pinIds.length === 0) return [];

  const pinIdSet = {};
  pinIds.forEach(function(pinId) {
    const normalizedPinId = normalizeRoutePinId_(pinId);
    if (normalizedPinId) pinIdSet[normalizedPinId] = true;
  });
  if (Object.keys(pinIdSet).length === 0) return [];

  const sheet = openRoutePinsSheet_();
  const rows = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  const affectedRouteIds = {};
  for (var i = rows.length - 1; i >= 1; i -= 1) {
    const routeId = normalizeRouteId_(rows[i][0]);
    const pinId = normalizeRoutePinId_(rows[i][1]);
    if (pinId && pinIdSet[pinId]) {
      if (routeId) affectedRouteIds[routeId] = true;
      sheet.deleteRow(i + 1);
    }
  }
  return Object.keys(affectedRouteIds);
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
  const cacheKey = normalizeRouteCacheKey_(data && data.cacheKey);
  const routeId = normalizeRouteId_(data && data.routeId);
  const provider = normalizeRouteCacheProvider_(data && data.provider);
  const coords = normalizeRouteCacheCoords_(data && data.coords);
  if (!cacheKey) return { ok: false, error: 'missing_cache_key' };
  if (!routeId) return { ok: false, error: 'missing_route_id' };
  if (coords.length < 2) return { ok: false, error: 'invalid_coords' };

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
  const rows = sheet.getLastRow() > 0 ? sheet.getDataRange().getValues() : [];
  let deletedCount = 0;
  for (var i = rows.length - 1; i >= 1; i -= 1) {
    if (routeIdSet[normalizeRouteId_(rows[i][1])]) {
      sheet.deleteRow(i + 1);
      deletedCount += 1;
    }
  }
  return deletedCount;
}

function invalidateRouteCacheForRoutes_(routeIds) {
  return deleteRouteCacheRowsForRouteIds_(routeIds);
}

function invalidateRouteCacheForPins_(pinIds) {
  return invalidateRouteCacheForRoutes_(findRouteIdsByPinIds_(pinIds));
}

function invalidateRouteCacheForPin(data) {
  const pinId = normalizeRoutePinId_(data && data.pinId);
  if (!pinId) return { ok: false, error: 'missing_pin_id' };
  return { ok: true, deleted: invalidateRouteCacheForPins_([pinId]) };
}

function invalidateRouteCacheForRoute(data) {
  const routeId = normalizeRouteId_(data && data.routeId);
  if (!routeId) return { ok: false, error: 'missing_route_id' };
  return { ok: true, deleted: invalidateRouteCacheForRoutes_([routeId]) };
}

function getRouteCacheSheetForRead_() {
  return openDataSpreadsheet_().getSheetByName(ROUTE_CACHE_SHEET_NAME);
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
  const basePinIdSet = {};
  basePinIds.forEach(function(pinId) {
    basePinIdSet[pinId] = true;
  });

  const seen = {};
  const pinIds = [];
  function canUsePinId(pinId) {
    return !!pinId && !!basePinIdSet[pinId];
  }
  function addPinId(pinId) {
    if (!canUsePinId(pinId) || seen[pinId]) return;
    seen[pinId] = true;
    pinIds.push(pinId);
  }

  const startPinId = normalizeRoutePinId_(group.startPinId);
  const endPinId = normalizeRoutePinId_(group.endPinId);
  const shouldAppendEndPin = group.closed !== true && canUsePinId(endPinId);
  if (canUsePinId(startPinId)) addPinId(startPinId);
  basePinIds.forEach(function(pinId) {
    if (pinId === startPinId || (shouldAppendEndPin && pinId === endPinId)) return;
    addPinId(pinId);
  });
  if (shouldAppendEndPin) addPinId(endPinId);
  return pinIds;
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

function getSharedPinsForShareLink_(shareLink) {
  return getMapPinsForShare_().filter(function(pin) {
    return matchesTagFilter_(pin, shareLink.tags, shareLink.tagMode)
      && matchesColorFilter_(pin, shareLink.colors);
  }).map(function(pin) {
    return toSharedPin_(pin, shareLink.tags);
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
  function miss(reason, group, expectedCacheKey) {
    logSharedRoadRouteCache_(routeId, group, expectedCacheKey, false, reason);
    return { ok: false };
  }

  const sharedPins = getSharedPinsForShareLink_(shareLink);
  const allowedPinIdSet = buildSharedAllowedPinIdSet_(sharedPins);
  const allRouteGroups = getRouteGroups();
  let rawGroup = null;
  for (var rawIndex = 0; rawIndex < allRouteGroups.length; rawIndex += 1) {
    if (normalizeRouteId_(allRouteGroups[rawIndex].routeId) === routeId) {
      rawGroup = allRouteGroups[rawIndex];
      break;
    }
  }
  if (!rawGroup || !isRouteClosedToAllowedPins_(rawGroup, allowedPinIdSet)) return miss('no_group', rawGroup, '');

  const sharedRouteGroups = getSharedRouteGroups_(sharedPins);
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

function deleteRouteGroup(id) {
  const routeId = normalizeRouteId_(id);
  if (!routeId) return { ok: false, error: 'missing_route_id' };

  const sheet = openRoutesSheet_();
  const existing = findRouteRow_(routeId);
  if (!existing) return { ok: false, error: 'route_not_found' };

  sheet.deleteRow(existing.rowNumber);
  deleteRoutePinsForRoute_(routeId);
  invalidateRouteCacheForRoutes_([routeId]);
  return { ok: true };
}

function updateRoutesOrder(data) {
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

  const sheet = openRoutesSheet_();
  const rows = readRouteRows_();
  const byId = {};
  rows.forEach(function(entry) {
    byId[entry.group.routeId] = entry;
  });
  for (var j = 0; j < orderedIds.length; j += 1) {
    if (!byId[orderedIds[j]]) {
      return { ok: false, error: 'route_not_found', routeId: orderedIds[j] };
    }
  }

  let orderIndex = 0;
  orderedIds.forEach(function(routeId) {
    sheet.getRange(byId[routeId].rowNumber, 10).setValue(orderIndex);
    orderIndex += 1;
  });
  rows.forEach(function(entry) {
    if (seen[entry.group.routeId]) return;
    sheet.getRange(entry.rowNumber, 10).setValue(orderIndex);
    orderIndex += 1;
  });

  return { ok: true, routeGroups: getRouteGroups() };
}

function buildSharedViewUrl_(token) {
  return ScriptApp.getService().getUrl() + '?view=shared&token=' + encodeURIComponent(token);
}

function createShareLink(data) {
  const label = normalizeShareLinkLabel_(data && data.label);
  const tags = PinData.normalizeTags(data && data.tags || []);
  const tagMode = String(data && data.tagMode || 'or') === 'and' ? 'and' : 'or';
  const colors = normalizeShareColors_(data && data.colors || []);
  const sheet = openShareLinksSheet_();
  const token = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  sheet.appendRow([createdAt, label, token, PinData.serializeTags(tags), tagMode, true, '', serializeShareColors_(colors)]);

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
      enabled: true,
      revokedAt: ''
    }
  };
}

function listShareLinks() {
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

function deleteShareLink(token) {
  var normalizedToken = normalizeShareToken_(token);
  const sheet = openShareLinksSheet_();
  const rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i += 1) {
    if (String(rows[i][2]) !== normalizedToken) continue;
    sheet.deleteRow(i + 1);
    return { ok: true };
  }
  return { ok: false, error: 'token not found' };
}

function revokeShareLink(token) {
  return setShareLinkEnabled({ token: token, enabled: false });
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

function getSharedRouteGroups_(pins) {
  var allowedPinIdSet = {};
  (Array.isArray(pins) ? pins : []).forEach(function(pin) {
    var pinId = normalizeRoutePinId_(pin && pin.id);
    if (pinId) allowedPinIdSet[pinId] = true;
  });
  if (Object.keys(allowedPinIdSet).length === 0) return [];

  return getRouteGroups().map(function(group) {
    return toSharedRouteGroup_(group, allowedPinIdSet);
  }).filter(function(group) {
    return !!group;
  });
}

function getSharedViewData(token) {
  var shareLink = getShareLinkByToken_(token);
  if (!shareLink) return { ok: false, error: 'invalid_share_link' };
  if (!shareLink.enabled) return { ok: false, error: 'revoked_share_link' };

  var pins = getMapPinsForShare_().filter(function(pin) {
    return matchesTagFilter_(pin, shareLink.tags, shareLink.tagMode)
      && matchesColorFilter_(pin, shareLink.colors);
  }).map(function(pin) {
    return toSharedPin_(pin, shareLink.tags);
  });

  return {
    ok: true,
    shareLink: {
      label: shareLink.label,
      token: shareLink.token,
      tags: shareLink.tags,
      tagMode: shareLink.tagMode,
      colors: shareLink.colors.slice()
    },
    allowedTags: shareLink.tags.slice(),
    allowedColors: shareLink.colors.slice(),
    pins: pins,
    routeGroups: getSharedRouteGroups_(pins)
  };
}


function saveMapData(data) {
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

function updatePinDetails(data) {
  if (!data || !data.id) return { ok: false, error: 'missing id' };
  if (!String(data.title || '').trim()) return { ok: false, error: 'title is required' };

  const sheet = openMapInfoSheet_();
  const rowIndex = findPinRowIndex_(sheet, data.id);
  if (rowIndex === -1) return { ok: false, error: 'id not found' };

  const sheetRow = rowIndex + 1;
  const row = sheet.getRange(sheetRow, 1, 1, MAP_INFO_COLUMN_COUNT).getValues()[0];
  const title = String(data.title).trim();
  const links = PinData.normalizeLinks(data.links || data.referenceUrls || []);

  sheet.getRange(sheetRow, 2, 1, 2).setValues([[title, String(data.description || '')]]);
  sheet.getRange(sheetRow, 6).setValue(data.color || row[5] || DEFAULT_COLOR);
  sheet.getRange(sheetRow, 10).setValue(PinData.serializeLinks(links));
  sheet.getRange(sheetRow, MAP_INFO_ICON_COLUMN).setValue(PinData.normalizeIcon(data.icon != null ? data.icon : row[14]));
  if (data.eventAt != null) {
    sheet.getRange(sheetRow, MAP_INFO_EVENT_AT_COLUMN).setValue(PinData.normalizeEventAt(data.eventAt));
  }

  if (data.status != null) {
    sheet.getRange(sheetRow, 11).setValue(PinData.normalizeStatus(String(data.status)));
  }
  if (data.tags != null) {
    sheet.getRange(sheetRow, 12).setValue(PinData.serializeTags(data.tags));
  }
  const updatedAt = currentUpdatedAt_();
  sheet.getRange(sheetRow, MAP_INFO_UPDATED_AT_COLUMN).setValue(updatedAt);

  if (getRenameFileWithTitle_() && row[6]) {
    renameDriveFileForTitle_(row[6], title);
  }

  return {
    ok: true,
    updatedAt: updatedAt,
    links: links,
    folderUrl: row[6] ? getParentFolderUrlByFileId_(row[6]) : ''
  };
}

function movePin(data) {
  if (!data || !data.id) return { ok: false, error: 'missing id' };
  if (data.lat == null || data.lng == null) return { ok: false, error: 'missing lat/lng' };

  const sheet = openMapInfoSheet_();
  const rowIndex = findPinRowIndex_(sheet, data.id);
  if (rowIndex === -1) return { ok: false, error: 'id not found' };

  const sheetRow = rowIndex + 1;
  sheet.getRange(sheetRow, 4, 1, 2).setValues([[Number(data.lat), Number(data.lng)]]);
  sheet.getRange(sheetRow, MAP_INFO_UPDATED_AT_COLUMN).setValue(currentUpdatedAt_());
  invalidateRouteCacheForPins_([data.id]);
  return { ok: true };
}

function unplacePin(data) {
  if (!data || !data.id) return { ok: false, error: 'missing id' };

  const sheet = openMapInfoSheet_();
  const rowIndex = findPinRowIndex_(sheet, data.id);
  if (rowIndex === -1) return { ok: false, error: 'id not found' };

  const sheetRow = rowIndex + 1;
  sheet.getRange(sheetRow, 4, 1, 2).setValues([['', '']]);
  sheet.getRange(sheetRow, MAP_INFO_UPDATED_AT_COLUMN).setValue(currentUpdatedAt_());
  return { ok: true };
}

function bulkUpdatePinStatus(data) {
  if (!data || !Array.isArray(data.ids) || data.ids.length === 0) {
    return { ok: false, error: 'ids must be a non-empty array' };
  }
  try {
    PinData.normalizeStatus(String(data.status || ''));
  } catch (_e) {
    return { ok: false, error: 'invalid status: ' + data.status };
  }
  const status = String(data.status).trim();
  if (!status) return { ok: false, error: 'status is required' };

  const sheet = openMapInfoSheet_();
  const rows = sheet.getDataRange().getValues();
  let updatedCount = 0;
  const updatedAt = currentUpdatedAt_();
  data.ids.forEach(function(id) {
    const rowIndex = rows.findIndex(function(row) { return row[8] === id; });
    if (rowIndex === -1) return;
    sheet.getRange(rowIndex + 1, 11).setValue(status);
    sheet.getRange(rowIndex + 1, MAP_INFO_UPDATED_AT_COLUMN).setValue(updatedAt);
    updatedCount += 1;
  });
  return { ok: true, updatedCount: updatedCount };
}

function deletePin(data) {
  if (!data || !data.id) return { ok: false, error: 'missing id' };

  const sheet = openMapInfoSheet_();
  const rowIndex = findPinRowIndex_(sheet, data.id);
  if (rowIndex === -1) return { ok: false, error: 'id not found' };

  const sheetRow = rowIndex + 1;
  const row = sheet.getRange(sheetRow, 1, 1, 10).getValues()[0];
  const fileId = row[6] || '';
  if (fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (error) {
      return { ok: false, error: '写真の削除に失敗しました: ' + error.message };
    }
  }

  sheet.deleteRow(sheetRow);
  const affectedRouteIds = deleteRoutePinsForPinIds_([data.id]);
  invalidateRouteCacheForRoutes_(affectedRouteIds);
  return { ok: true };
}

function bulkDeletePins(data) {
  if (!data || !Array.isArray(data.ids) || data.ids.length === 0) {
    return { ok: false, error: 'ids must be a non-empty array' };
  }

  const sheet = openMapInfoSheet_();
  const rows = sheet.getDataRange().getValues();

  // id → rowIndex のマッピングを作成
  // findPinRowIndex_ は呼び出しごとに getDataRange() を行うため、
  // バッチ処理ではここで一括取得した rows を使って検索する
  const rowIndexMap = {};
  data.ids.forEach(function(id) {
    const rowIndex = rows.findIndex(function(row) { return row[8] === id; });
    if (rowIndex !== -1) {
      rowIndexMap[id] = rowIndex;
    }
  });

  // deleteRow で行番号がずれるため、行番号の大きい順（逆順）に処理する
  const sortedEntries = Object.keys(rowIndexMap)
    .map(function(id) { return [id, rowIndexMap[id]]; })
    .sort(function(a, b) { return b[1] - a[1]; });

  let deletedCount = 0;
  const failedIds = [];

  sortedEntries.forEach(function(entry) {
    const id = entry[0];
    const rowIndex = entry[1];
    const sheetRow = rowIndex + 1;
    try {
      const row = sheet.getRange(sheetRow, 1, 1, 10).getValues()[0];
      const fileId = row[6] || '';
      if (fileId) {
        DriveApp.getFileById(fileId).setTrashed(true);
      }
      sheet.deleteRow(sheetRow);
      deletedCount += 1;
    } catch (error) {
      Logger.log('bulkDeletePins: failed for id=' + id + ' — ' + error.message);
      failedIds.push(id);
    }
  });

  if (deletedCount > 0) {
    const affectedRouteIds = deleteRoutePinsForPinIds_(sortedEntries
      .filter(function(entry) { return failedIds.indexOf(entry[0]) === -1; })
      .map(function(entry) { return entry[0]; }));
    invalidateRouteCacheForRoutes_(affectedRouteIds);
  }

  return { ok: true, deletedCount: deletedCount, failedIds: failedIds };
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
  if (!data) return { ok: false, error: 'missing data' };
  setConfigValue_('RENAME_FILE_WITH_TITLE', data.renameFileWithTitle ? 'true' : 'false');
  return getAppSettings();
}

// 旧フロント互換用
function updatePin(data) {
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
  try {
    const parents = DriveApp.getFileById(fileId).getParents();
    if (!parents.hasNext()) return '';
    return getDriveFolderUrl_(parents.next().getId());
  } catch (_error) {
    return '';
  }
}

function enrichPinWithDriveMeta_(pin) {
  const enriched = {};
  Object.keys(pin).forEach(function(key) {
    enriched[key] = pin[key];
  });
  enriched.folderUrl = pin.fileId ? getParentFolderUrlByFileId_(pin.fileId) : '';
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

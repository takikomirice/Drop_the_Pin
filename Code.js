// ============================================================
//  Drop the Pin! — Google Apps Script バックエンド
// ============================================================

const SPREADSHEET_LITERAL_MARKER_ = '\u200Bdtp-sheet:v1:';
const SPREADSHEET_LITERAL_VALUE_PREFIX_ = SPREADSHEET_LITERAL_MARKER_ + 'v:';
const SPREADSHEET_LITERAL_ESCAPE_PREFIX_ = SPREADSHEET_LITERAL_MARKER_ + 'e:';

function encodeSpreadsheetLiteral_(value) {
  const source = String(value == null ? '' : value);
  if (source.indexOf(SPREADSHEET_LITERAL_MARKER_) === 0) {
    return SPREADSHEET_LITERAL_ESCAPE_PREFIX_ + source;
  }
  return /^[=+\-@\t\r]/.test(source)
    ? SPREADSHEET_LITERAL_VALUE_PREFIX_ + source
    : source;
}

function decodeSpreadsheetLiteral_(value) {
  const source = String(value == null ? '' : value);
  if (source.indexOf(SPREADSHEET_LITERAL_VALUE_PREFIX_) === 0) {
    const decoded = source.slice(SPREADSHEET_LITERAL_VALUE_PREFIX_.length);
    return /^[=+\-@\t\r]/.test(decoded) ? decoded : source;
  }
  if (source.indexOf(SPREADSHEET_LITERAL_ESCAPE_PREFIX_) === 0) {
    const decoded = source.slice(SPREADSHEET_LITERAL_ESCAPE_PREFIX_.length);
    return decoded.indexOf(SPREADSHEET_LITERAL_MARKER_) === 0 ? decoded : source;
  }
  return source;
}

const PinData = (function() {
  const DEFAULT_COLOR = '#e53935';
  const DEFAULT_ICON = 'default';
  const URL_RE = /^https?:\/\/\S+$/i;
  const COLOR_OPTIONS = [
    '#e53935', '#e91e63', '#9c27b0', '#3f51b5', '#2196f3',
    '#00bcd4', '#009688', '#4caf50', '#8bc34a', '#ffeb3b',
    '#ff9800', '#ff5722', '#795548', '#607d8b', '#212121'
  ];
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
    if (year < 1 || month < 1 || month > 12 || hour > 23 || minute > 59 || second > 59) return '';
    const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
    const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (day < 1 || day > days[month - 1]) return '';
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
    const maxLength = 180;
    function safePart(value) {
      return String(value == null ? '' : value)
        .replace(/[\u0000-\u001f\u007f]/g, '')
        .replace(/[<>:"/\\|?*]/g, '_')
        .replace(/[. ]+$/g, '')
        .trim();
    }
    const rawOriginal = String(originalName || 'image');
    const originalExtensionMatch = rawOriginal.match(/(\.[A-Za-z0-9]{1,10})$/);
    const extension = originalExtensionMatch ? originalExtensionMatch[1] : '';
    const safeOriginal = safePart(rawOriginal) || ('image' + extension);
    const normalizedTitle = safePart(
      String(title || '').replace(/\.(?:jpe?g|png|gif|webp|heic|heif)$/i, '')
    );
    let baseName = shouldSync && normalizedTitle
      ? normalizedTitle
      : safePart(extension ? safeOriginal.slice(0, -extension.length) : safeOriginal);
    if (!baseName) baseName = 'image';
    const maxBaseLength = Math.max(1, maxLength - extension.length);
    return baseName.slice(0, maxBaseLength).replace(/[. ]+$/g, '') + extension;
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
    const audioIdColumnIndex = MAP_INFO_HEADERS.indexOf('音声ID');
    return {
      timestamp: row[0] ? (row[0] instanceof Date ? row[0].toISOString() : String(row[0])) : '',
      title: decodeSpreadsheetLiteral_(row[1]),
      description: decodeSpreadsheetLiteral_(row[2]),
      lat: toNumberOrNull(row[3]),
      lng: toNumberOrNull(row[4]),
      color: row[5] || DEFAULT_COLOR,
      fileId: row[6] || '',
      imageUrl: row[7] || '',
      id: row[8] || '',
      links: deserializeLinks(row[9] || ''),
      status: String(row[10] || '').trim(),
      tags: deserializeTags(decodeSpreadsheetLiteral_(row[11])),
      eventAt: normalizeEventAt(row[12]),
      updatedAt: row[13] ? String(row[13]) : '',
      icon: normalizeIcon(row[14]),
      audioId: audioIdColumnIndex === -1 ? '' : String(row[audioIdColumnIndex] || '')
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
    COLOR_OPTIONS: COLOR_OPTIONS.slice(),
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
const TRACKS_SHEET_NAME = 'tracks';
const TRACK_SEGMENTS_SHEET_NAME = 'track_segments';
const INPUT_PRESETS_SHEET_NAME = 'input_presets';
const IMPORT_RECEIPTS_SHEET_NAME = 'import_receipts';
const MAP_INFO_HEADERS = [
  'タイムスタンプ', 'タイトル', '説明',
  '緯度', '経度', 'ピンの色',
  'ファイルID', '画像URL', 'ID', '参考URL一覧',
  '状態', 'タグ', 'イベント時刻', '更新時刻', 'アイコン', '音声ID'
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
const EDIT_URL_CONFIG_KEY = 'EDIT_URL';
const EDIT_TOKEN_TTL_SECONDS = 6 * 60 * 60;
const EDIT_TOKEN_CACHE_PREFIX = 'EDIT_TOKEN_';
const SHARE_LINKS_LEGACY_HEADERS = ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors', 'routeIds'];
const SHARE_ROUTE_TARGETS_HEADER = 'routeTargetsJson';
const SHARE_LINKS_HEADERS = SHARE_LINKS_LEGACY_HEADERS.concat([SHARE_ROUTE_TARGETS_HEADER]);
const SHARED_PAYLOAD_MAX_PINS = 2000;
const SHARED_PAYLOAD_MAX_ROUTES = 100;
const SHARED_PAYLOAD_MAX_COORDINATE_POINTS = 50000;
const SHARED_PAYLOAD_MAX_JSON_BYTES = 5 * 1024 * 1024;
const SHARED_PAYLOAD_TOO_LARGE_CODE = 'SHARED_PAYLOAD_TOO_LARGE';
const SHARED_PAYLOAD_CREATE_ERROR = '共有対象が大きすぎます。ピンまたはルートを減らしてください。';
const SHARE_ROUTE_TARGET_TYPES = {
  'pin-route': true,
  'gpx-route': true,
  'geojson-route': true
};
const ROUTES_HEADERS = ['routeId', 'name', 'color', 'routeMode', 'closed', 'startPinId', 'endPinId', 'createdAt', 'updatedAt', 'orderIndex', 'visible', 'showNumbers', 'showLine', 'lineStyle'];
const ROUTE_PINS_HEADERS = ['routeId', 'pinId', 'pinOrder', 'createdAt', 'updatedAt'];
const ROUTE_CACHE_HEADERS = ['cacheKey', 'routeId', 'coordsJson', 'provider', 'createdAt', 'expiresAt'];
const TRACKS_HEADERS = [
  'trackId', 'name', 'description', 'color', 'sourceType', 'sourceName', 'activeRevision',
  'payloadHash', 'segmentCount', 'pointCount', 'distanceMeters', 'minElevation', 'maxElevation',
  'startTime', 'endTime', 'boundsJson', 'createdAt', 'updatedAt', 'orderIndex', 'visible',
  'lineStyle', 'lineWidth'
];
const TRACK_SEGMENTS_HEADERS = [
  'trackId', 'revisionId', 'segmentIndex', 'chunkIndex', 'pointsJson', 'pointCount',
  'createdAt', 'updatedAt'
];
const INPUT_PRESET_HEADERS = [
  'presetId', 'name', 'enabled', 'orderIndex', 'tagsMode', 'tags', 'colorMode',
  'color', 'iconMode', 'icon', 'statusMode', 'status', 'createdAt', 'updatedAt'
];
const INPUT_PRESET_COLUMN_COUNT = INPUT_PRESET_HEADERS.length;
const IMPORT_RECEIPT_HEADERS = [
  'idempotencyKey', 'jobId', 'itemId', 'payloadHash', 'state', 'leaseOwner',
  'leaseUntil', 'pinId', 'targetFolderId', 'tempFileName', 'fileId', 'imageUrl',
  'folderUrl', 'createdAt', 'updatedAt', 'lastErrorCode', 'sourceDriveFileId',
  'mediaKind', 'operationMode', 'targetPinId', 'cleanupFileId'
];
const LEGACY_IMPORT_RECEIPT_HEADERS = IMPORT_RECEIPT_HEADERS.slice(0, 16);
const PHOTO_IMPORT_RECEIPT_HEADERS = IMPORT_RECEIPT_HEADERS.slice(0, 17);
const IMPORT_RECEIPT_COLUMN_COUNT = IMPORT_RECEIPT_HEADERS.length;
const IMPORT_RECEIPT_STATES = {
  RESERVED: 'reserved',
  FILE_SAVED: 'file_saved',
  LINKED: 'linked',
  COMPLETED: 'completed',
  CLEANUP_PENDING: 'cleanup_pending',
  FAILED: 'failed'
};
const IMPORT_ITEM_LEASE_MS = 7 * 60 * 1000;
const IMPORT_AUDIO_MIN_BYTES = 1024;
const IMPORT_AUDIO_MAX_BYTES = 4 * 1024 * 1024;
const IMPORT_ITEM_ID_MAX_LENGTH = 128;
const IMPORT_IDEMPOTENCY_KEY_MAX_LENGTH = IMPORT_ITEM_ID_MAX_LENGTH * 2 + 1;
const MAX_INPUT_PRESETS = 100;
const MAX_TRACKS = 100;
const MAX_TRACK_SEGMENTS = 200;
const MAX_TRACK_POINTS = 20000;
const TRACK_POINTS_PER_CHUNK = 500;
const TRACK_POINTS_JSON_MAX_LENGTH = 40000;
const TRACK_ID_MAX_LENGTH = 128;
const TRACK_REVISION_ID_MAX_LENGTH = 128;
const TRACK_NAME_MAX_LENGTH = 100;
const TRACK_DESCRIPTION_MAX_LENGTH = 400;
const TRACK_SOURCE_NAME_MAX_LENGTH = 200;
const TRACK_STAGE_PROPERTY_PREFIX_ = 'TRACK_STAGE_V1_';
const TRACK_RETIRED_REVISION_PROPERTY_PREFIX_ = 'TRACK_RETIRED_REVISION_V1_';
const INPUT_PRESET_NAME_MAX_LENGTH = 60;
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
    .addItem('編集URLを更新・開く', 'refreshAndOpenEditUrl')
    .addItem('編集キーを再生成', 'regenerateEditKeyFromMenu')
    .addToUi();
}

function isTrulyBlankSheetCell_(value, formula) {
  return (value === '' || value == null) && !formula;
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const configSheet = prepareConfigSheetFirst_(ss);
  ensureConfigEntry_(configSheet, 'IMAGE_DRIVE_URL', '',
    '写真を保存するGoogleドライブフォルダのURL（フォルダを右クリック→共有→リンクをコピー）');
  ensureConfigEntry_(configSheet, 'RENAME_FILE_WITH_TITLE', 'false',
    'true の場合、タイトル編集時に Drive 上の写真名も同じタイトルへ更新');
  refreshEditUrlConfig_(configSheet, { preserveExistingFormula: true });

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
  if (sheet.getMaxColumns() < MAP_INFO_COLUMN_COUNT) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      MAP_INFO_COLUMN_COUNT - sheet.getMaxColumns()
    );
  }

  if (!looksHeader) {
    sheet.getRange(1, 1, 1, MAP_INFO_COLUMN_COUNT).setValues([MAP_INFO_HEADERS]);
  } else {
    const headerRange = sheet.getRange(1, 1, 1, MAP_INFO_COLUMN_COUNT);
    const headerValues = headerRange.getValues()[0];
    const headerFormulas = headerRange.getFormulas()[0];
    MAP_INFO_HEADERS.forEach(function(header, index) {
      if (isTrulyBlankSheetCell_(headerValues[index], headerFormulas[index])) {
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

  ensureShareLinksSheet_(ss);
  ensureHeaderSheet_(ss, ROUTES_SHEET_NAME, ROUTES_HEADERS);
  ensureHeaderSheet_(ss, ROUTE_PINS_SHEET_NAME, ROUTE_PINS_HEADERS);
  ensureHeaderSheet_(ss, ROUTE_CACHE_SHEET_NAME, ROUTE_CACHE_HEADERS);
  ensureHeaderSheet_(ss, TRACKS_SHEET_NAME, TRACKS_HEADERS);
  ensureHeaderSheet_(ss, TRACK_SEGMENTS_SHEET_NAME, TRACK_SEGMENTS_HEADERS);
  ensureHeaderSheet_(ss, INPUT_PRESETS_SHEET_NAME, INPUT_PRESET_HEADERS);
  ensureImportReceiptsSheet_(ss);

  ui.alert(
    '初期設定完了',
    '"' + SHEET_NAME + '" シート、"' + CONFIG_SHEET_NAME + '" シート、"' + SHARE_LINKS_SHEET_NAME + '" シート、' +
    '"' + ROUTES_SHEET_NAME + '" シート、"' + ROUTE_PINS_SHEET_NAME + '" シート、"' + ROUTE_CACHE_SHEET_NAME + '" シート、' +
    '"' + TRACKS_SHEET_NAME + '" シート、"' + TRACK_SEGMENTS_SHEET_NAME + '" シート、' +
    '"' + INPUT_PRESETS_SHEET_NAME + '" シート、"' + IMPORT_RECEIPTS_SHEET_NAME + '" シートの準備が整いました。\n\n' +
    '次のステップ:\n' +
    '1. "' + CONFIG_SHEET_NAME + '" シートを開いて IMAGE_DRIVE_URL を設定\n' +
    '2. 必要なら RENAME_FILE_WITH_TITLE を true に変更\n' +
    '3. ウェブアプリとしてデプロイ',
    ui.ButtonSet.OK
  );
}

function prepareConfigSheetFirst_(ss) {
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const sheets = typeof ss.getSheets === 'function' ? ss.getSheets() : [];

  if (!configSheet) {
    const initialSheet = sheets.length === 1 ? sheets[0] : null;
    const blankInitialSheet = initialSheet
      && initialSheet.getLastRow() === 0
      && (typeof initialSheet.getLastColumn !== 'function' || initialSheet.getLastColumn() === 0)
      && typeof initialSheet.setName === 'function';
    configSheet = blankInitialSheet
      ? initialSheet.setName(CONFIG_SHEET_NAME)
      : ss.insertSheet(CONFIG_SHEET_NAME);
  }

  if (configSheet.getLastRow() === 0) {
    configSheet.getRange(1, 1, 1, 3).setValues([['設定項目', '値', '説明']]);
    configSheet.getRange('A1:C1')
      .setBackground('#1565c0')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    configSheet.setFrozenRows(1);
    [200, 350, 300].forEach(function(width, index) {
      configSheet.setColumnWidth(index + 1, width);
    });
  }

  moveSheetToFirst_(ss, configSheet);
  return configSheet;
}

function moveSheetToFirst_(ss, sheet) {
  if (typeof ss.getSheets !== 'function'
      || typeof ss.setActiveSheet !== 'function'
      || typeof ss.moveActiveSheet !== 'function') return;
  const sheets = ss.getSheets();
  if (!sheets.length || sheets[0] === sheet) return;
  const previousActive = typeof ss.getActiveSheet === 'function' ? ss.getActiveSheet() : null;
  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(1);
  if (previousActive && previousActive !== sheet) ss.setActiveSheet(previousActive);
}

function doGet(e) {
  var params = (e && e.parameter) || {};
  var templateName = params.view === 'shared' ? 'shared' : 'index';
  var template;
  if (templateName === 'shared') {
    template = HtmlService.createTemplateFromFile('shared');
  } else {
    var rawIndex = HtmlService.createHtmlOutputFromFile('index').getContent();
    var vendorLocation = locateAudioVendorBundleInIndex_(rawIndex);
    var strippedIndex = stripAudioVendorBundleFromIndex_(rawIndex, vendorLocation);
    template = HtmlService.createTemplate(strippedIndex);
  }
  template.execUrl = getConfiguredWebAppUrl_();
  template.token = params.token || '';
  template.editToken = templateName === 'index' ? issueEditTokenFromRequest_(params) : '';
  return template.evaluate()
    .setTitle('Drop the Pin!')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

const AUDIO_VENDOR_VERSION_ = '1.50.8';
const AUDIO_VENDOR_START_SENTINEL_ = 'AUDIO_VENDOR_BUNDLE_START';
const AUDIO_VENDOR_END_SENTINEL_ = 'AUDIO_VENDOR_BUNDLE_END';

function countTextOccurrences_(text, needle) {
  return String(text || '').split(needle).length - 1;
}

function locateAudioVendorBundleInIndex_(rawIndex) {
  const content = String(rawIndex || '');
  if (countTextOccurrences_(content, AUDIO_VENDOR_START_SENTINEL_) !== 1
      || countTextOccurrences_(content, AUDIO_VENDOR_END_SENTINEL_) !== 1) {
    throw new Error('Audio vendor bundle markers are invalid.');
  }

  const startIndex = content.indexOf(AUDIO_VENDOR_START_SENTINEL_);
  const endIndex = content.indexOf(AUDIO_VENDOR_END_SENTINEL_);
  if (startIndex < 0 || endIndex <= startIndex) {
    throw new Error('Audio vendor bundle markers are out of order.');
  }

  if (content.slice(0, startIndex).trim()) {
    throw new Error('Audio vendor bundle is not at the index prefix.');
  }

  const opening = AUDIO_VENDOR_START_SENTINEL_ + '\n<script>\n';
  const closing = '\n</script>\n' + AUDIO_VENDOR_END_SENTINEL_;
  const sourceStartIndex = startIndex + opening.length;
  const sourceEndIndex = endIndex - '\n</script>\n'.length;
  if (content.slice(startIndex, sourceStartIndex) !== opening
      || content.slice(sourceEndIndex, endIndex + AUDIO_VENDOR_END_SENTINEL_.length) !== closing) {
    throw new Error('Audio vendor script wrapper is invalid.');
  }

  const source = content.slice(sourceStartIndex, sourceEndIndex);
  if (!source.trim() || /<\/script/i.test(source)) {
    throw new Error('Audio vendor source is invalid.');
  }
  if (countTextOccurrences_(source, 'globalThis.Mediabunny=') !== 1
      || countTextOccurrences_(source, 'globalThis.MediabunnyMp3Encoder=') !== 1) {
    throw new Error('Audio vendor public API is invalid.');
  }

  const afterEndIndex = endIndex + AUDIO_VENDOR_END_SENTINEL_.length;
  var documentStartIndex;
  if (content.slice(afterEndIndex, afterEndIndex + 2) === '\r\n') {
    documentStartIndex = afterEndIndex + 2;
  } else if (content.charAt(afterEndIndex) === '\n') {
    documentStartIndex = afterEndIndex + 1;
  } else {
    throw new Error('Audio vendor bundle must precede the index document.');
  }
  if (!/^<!DOCTYPE html>/i.test(content.slice(documentStartIndex))) {
    throw new Error('Audio vendor bundle is not followed by the index document.');
  }

  return {
    source: source,
    startIndex: startIndex,
    endIndex: endIndex,
    documentStartIndex: documentStartIndex
  };
}

function stripAudioVendorBundleFromIndex_(rawIndex, vendorLocation) {
  const content = String(rawIndex || '');
  const location = vendorLocation || locateAudioVendorBundleInIndex_(content);
  return content.slice(location.documentStartIndex);
}

function getAudioVendorBundle(payload) {
  assertEditToken_(payload);
  const rawIndex = HtmlService.createHtmlOutputFromFile('index').getContent();
  const vendorLocation = locateAudioVendorBundleInIndex_(rawIndex);
  return {
    version: AUDIO_VENDOR_VERSION_,
    source: vendorLocation.source
  };
}

function getPinAudioData(payload) {
  assertEditToken_(payload);
  let pinId;
  try {
    pinId = normalizeImportIdentifier_(
      payload && payload.pinId,
      'pinId',
      IMPORT_ITEM_ID_MAX_LENGTH
    );
  } catch (_error) {
    throw importItemError_('PIN_AUDIO_NOT_FOUND', 'pin audio is unavailable.', false);
  }
  try {
    return pinAudioDataFromBlob_(readPinAudioBlobByPinId_(pinId));
  } catch (_error) {
    throw importItemError_('PIN_AUDIO_NOT_FOUND', 'pin audio is unavailable.', false);
  }
}

function getPinPhotoData(payload) {
  assertEditToken_(payload);
  let pinId;
  try {
    pinId = normalizeImportIdentifier_(
      payload && payload.pinId,
      'pinId',
      IMPORT_ITEM_ID_MAX_LENGTH
    );
    return pinPhotoDataFromBlob_(readPinPhotoBlobByPinId_(pinId));
  } catch (_error) {
    throw importItemError_('PIN_PHOTO_NOT_FOUND', 'pin photo is unavailable.', false);
  }
}

function pinPhotoDataFromBlob_(blob) {
  const mimeType = blob ? String(blob.getContentType() || '').toLowerCase() : '';
  if (!isPinPhotoReadMimeType_(mimeType)) {
    throw new Error('pin photo is unavailable.');
  }
  const bytes = blob.getBytes();
  const byteLength = bytes && Number.isSafeInteger(bytes.length) ? bytes.length : -1;
  if (byteLength < 1 || byteLength > DRIVE_PHOTO_IMPORT_MAX_FILE_BYTES) {
    throw new Error('pin photo is unavailable.');
  }
  return {
    ok: true,
    mimeType: mimeType,
    byteLength: byteLength,
    base64: Utilities.base64Encode(bytes)
  };
}

function pinAudioDataFromBlob_(blob) {
  if (!blob || String(blob.getContentType()) !== 'audio/mpeg') {
    throw new Error('pin audio is unavailable.');
  }
  const bytes = blob.getBytes();
  const byteLength = bytes && Number.isSafeInteger(bytes.length) ? bytes.length : -1;
  if (byteLength < IMPORT_AUDIO_MIN_BYTES || byteLength > IMPORT_AUDIO_MAX_BYTES) {
    throw new Error('pin audio is unavailable.');
  }
  return {
    ok: true,
    mimeType: 'audio/mpeg',
    byteLength: byteLength,
    base64: Utilities.base64Encode(bytes)
  };
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
  const valueRange = sheet.getRange(row, 2);
  if (valueRange.getValue() === '' && !rangeHasFormula_(valueRange)) {
    valueRange.setValue(typeof value === 'function' ? value() : value);
  }
  const descriptionRange = sheet.getRange(row, 3);
  if (descriptionRange.getValue() === '' && !rangeHasFormula_(descriptionRange)) {
    descriptionRange.setValue(description);
  }
}

function rangeHasFormula_(range) {
  if (!range) return false;
  if (typeof range.getFormula === 'function' && range.getFormula()) return true;
  const value = typeof range.getValue === 'function' ? range.getValue() : '';
  return typeof value === 'string' && value.charAt(0) === '=';
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
  ensureConfigEntry_(sheet, EDIT_URL_CONFIG_KEY, '',
    '編集用WebアプリURL。知っている人は編集できるため、共有範囲に注意してください。');
  return sheet;
}

function getConfiguredWebAppUrl_() {
  const config = getAppConfig_();
  return normalizeWebAppUrl_(config[WEB_APP_URL_CONFIG_KEY]) ||
    normalizeWebAppUrl_(ScriptApp.getService().getUrl());
}

function buildEditUrl_() {
  const result = refreshEditUrlConfig_();
  if (!result.url) throw new Error('WebアプリURLが取得できません');
  return result.url;
}

function refreshEditUrlConfig_(configSheet, options) {
  const sheet = ensureEditUrlConfig_(configSheet);
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const config = {};
  const rowNumbers = {};
  rows.forEach(function(row, index) {
    const key = String(row[0] || '');
    if (!key || rowNumbers[key]) return;
    config[key] = String(row[1] || '');
    rowNumbers[key] = index + 2;
  });

  const editKey = String(config[EDIT_KEY_CONFIG_KEY] || '').trim();
  let webAppUrl = normalizeWebAppUrl_(config[WEB_APP_URL_CONFIG_KEY]);
  if (!webAppUrl) {
    try {
      webAppUrl = normalizeWebAppUrl_(ScriptApp.getService().getUrl());
    } catch (error) {
      webAppUrl = '';
    }
  }

  const editUrl = editKey && webAppUrl
    ? webAppUrl + '?mode=edit&editKey=' + encodeURIComponent(editKey)
    : '';
  const editUrlCell = sheet.getRange(rowNumbers[EDIT_URL_CONFIG_KEY], 2);
  if (options && options.preserveExistingFormula && rangeHasFormula_(editUrlCell)) {
    return { url: String(editUrlCell.getValue() || ''), row: rowNumbers[EDIT_URL_CONFIG_KEY] };
  }
  editUrlCell.setValue(editUrl);
  if (typeof editUrlCell.setShowHyperlink === 'function') {
    editUrlCell.setShowHyperlink(true);
  }
  return { url: editUrl, row: rowNumbers[EDIT_URL_CONFIG_KEY] };
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
  refreshEditUrlConfig_();
  ui.alert('保存しました', normalizedUrl || 'WEB_APP_URL を空欄にしました。', ui.ButtonSet.OK);
}

function refreshAndOpenEditUrl() {
  const ui = SpreadsheetApp.getUi();
  const result = refreshEditUrlConfig_();
  const configSheet = openDataSpreadsheet_().getSheetByName(CONFIG_SHEET_NAME);
  configSheet.activate();
  configSheet.getRange(result.row, 2).activate();
  if (!result.url) {
    ui.alert('編集URLを生成できません', 'WEB_APP_URLを設定するか、Webアプリをデプロイしてください。', ui.ButtonSet.OK);
  }
}

function showEditUrlDialog() {
  refreshAndOpenEditUrl();
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
  refreshEditUrlConfig_();
  ui.alert('編集キーを再生成しました', 'configシートの EDIT_URL を共有し直してください。', ui.ButtonSet.OK);
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
  const match = String(url || '').match(/\/folders\/([A-Za-z0-9_-]{10,200})(?:[/?#]|$)/);
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

function openTracksSheet_() {
  return getRequiredSheet_(TRACKS_SHEET_NAME);
}

function openTrackSegmentsSheet_() {
  return getRequiredSheet_(TRACK_SEGMENTS_SHEET_NAME);
}

function openInputPresetsSheet_() {
  return getRequiredSheet_(INPUT_PRESETS_SHEET_NAME);
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

// ============================================================
//  入力プリセット
// ============================================================

function normalizeInputPresetEnabled_(value) {
  if (value === false) return false;
  const normalized = String(value == null ? '' : value).trim().toLowerCase();
  if (normalized === 'false' || normalized === '0') return false;
  return true;
}

function normalizeInputPresetMode_(value, allowedModes, label) {
  const mode = String(value || '').trim().toLowerCase();
  if (allowedModes.indexOf(mode) === -1) {
    return { ok: false, error: label + 'の動作が不正です。' };
  }
  return { ok: true, value: mode };
}

function normalizeInputPresetForSave_(source) {
  const data = source || {};
  const name = String(data.name || '').trim();
  if (!name) return { ok: false, error: 'プリセット名を入力してください。' };
  if (name.length > INPUT_PRESET_NAME_MAX_LENGTH) {
    return { ok: false, error: 'プリセット名は' + INPUT_PRESET_NAME_MAX_LENGTH + '文字以内にしてください。' };
  }

  const tagsModeResult = normalizeInputPresetMode_(data.tagsMode, ['keep', 'set', 'clear'], 'タグ');
  if (!tagsModeResult.ok) return tagsModeResult;
  const colorModeResult = normalizeInputPresetMode_(data.colorMode, ['keep', 'set'], '色');
  if (!colorModeResult.ok) return colorModeResult;
  const iconModeResult = normalizeInputPresetMode_(data.iconMode, ['keep', 'set'], 'アイコン');
  if (!iconModeResult.ok) return iconModeResult;
  const statusModeResult = normalizeInputPresetMode_(data.statusMode, ['keep', 'set', 'clear'], '状態');
  if (!statusModeResult.ok) return statusModeResult;

  const tagsMode = tagsModeResult.value;
  const colorMode = colorModeResult.value;
  const iconMode = iconModeResult.value;
  const statusMode = statusModeResult.value;
  if (tagsMode === 'keep' && colorMode === 'keep' && iconMode === 'keep' && statusMode === 'keep') {
    return { ok: false, error: '少なくとも1項目の動作を「変更しない」以外にしてください。' };
  }

  let tags = [];
  if (tagsMode === 'set') {
    try {
      tags = PinData.normalizeTags(data.tags);
    } catch (_error) {
      return { ok: false, error: 'タグは5件以内にしてください。' };
    }
  }

  let color = null;
  if (colorMode === 'set') {
    color = String(data.color || '').trim().toLowerCase();
    if (!SAFE_COLOR_RE.test(color)) {
      return { ok: false, error: '色は安全な6桁hex形式で指定してください。' };
    }
  }

  let icon = null;
  if (iconMode === 'set') {
    icon = String(data.icon || '').trim().toLowerCase();
    if (PinData.ICON_OPTIONS.indexOf(icon) === -1) {
      return { ok: false, error: 'アイコンが不正です。' };
    }
  }

  let status = null;
  if (statusMode === 'set') {
    try {
      status = PinData.normalizeStatus(data.status);
    } catch (_error) {
      return { ok: false, error: '状態が不正です。' };
    }
    if (!status) return { ok: false, error: '設定する状態を選択してください。' };
  }

  return {
    ok: true,
    preset: {
      presetId: String(data.presetId || '').trim(),
      name: name,
      enabled: normalizeInputPresetEnabled_(data.enabled),
      orderIndex: Number(data.orderIndex),
      tagsMode: tagsMode,
      tags: tags,
      colorMode: colorMode,
      color: color,
      iconMode: iconMode,
      icon: icon,
      statusMode: statusMode,
      status: status,
      createdAt: data.createdAt ? String(data.createdAt) : '',
      updatedAt: data.updatedAt ? String(data.updatedAt) : ''
    }
  };
}

function inputPresetRowToSource_(row, fallbackOrderIndex) {
  return {
    presetId: row[0],
    name: row[1],
    enabled: row[2],
    orderIndex: Number.isFinite(Number(row[3])) ? Number(row[3]) : fallbackOrderIndex,
    tagsMode: row[4],
    tags: PinData.deserializeTags(row[5] || ''),
    colorMode: row[6],
    color: row[7] || null,
    iconMode: row[8],
    icon: row[9] || null,
    statusMode: row[10],
    status: row[11] || null,
    createdAt: row[12],
    updatedAt: row[13]
  };
}

function inputPresetRowToModel_(row, fallbackOrderIndex) {
  const result = normalizeInputPresetForSave_(inputPresetRowToSource_(row, fallbackOrderIndex));
  if (!result.ok) {
    const presetId = String(row && row[0] || '').trim();
    throw new Error('input_presets シートのプリセット' + (presetId ? '「' + presetId + '」' : '') + 'が不正です: ' + result.error);
  }
  return result.preset;
}

function inputPresetToRow_(preset) {
  return [
    preset.presetId,
    preset.name,
    preset.enabled,
    preset.orderIndex,
    preset.tagsMode,
    preset.tagsMode === 'set' ? PinData.serializeTags(preset.tags) : '',
    preset.colorMode,
    preset.colorMode === 'set' ? preset.color : '',
    preset.iconMode,
    preset.iconMode === 'set' ? preset.icon : '',
    preset.statusMode,
    preset.statusMode === 'set' ? preset.status : '',
    preset.createdAt,
    preset.updatedAt
  ];
}

function readInputPresetRows_(sheet) {
  const rowCount = Math.max(0, sheet.getLastRow() - 1);
  if (rowCount === 0) return [];
  return sheet.getRange(2, 1, rowCount, INPUT_PRESET_COLUMN_COUNT).getValues();
}

function inputPresetEntriesFromRows_(rows) {
  const entries = [];
  const seenIds = {};
  rows.forEach(function(row, index) {
    const presetId = String(row && row[0] || '').trim();
    if (!presetId) return;
    if (seenIds[presetId]) {
      throw new Error('input_presets シートに重複したpresetIdがあります: ' + presetId);
    }
    seenIds[presetId] = true;
    entries.push({ rowIndex: index, rowNumber: index + 2, model: inputPresetRowToModel_(row, index) });
  });
  return entries;
}

function compareInputPresets_(left, right) {
  const orderDifference = Number(left.orderIndex) - Number(right.orderIndex);
  if (orderDifference !== 0) return orderDifference;
  const nameDifference = String(left.name).localeCompare(String(right.name), 'ja');
  if (nameDifference !== 0) return nameDifference;
  return String(left.presetId).localeCompare(String(right.presetId), 'ja');
}

function sortedInputPresetModels_(entries) {
  return entries.map(function(entry) { return entry.model; }).sort(compareInputPresets_);
}

function listInputPresets(data) {
  assertEditToken_(data);
  const sheet = openInputPresetsSheet_();
  const entries = inputPresetEntriesFromRows_(readInputPresetRows_(sheet));
  return { ok: true, presets: sortedInputPresetModels_(entries) };
}

function saveInputPreset(data) {
  assertEditToken_(data);
  const payload = data || {};
  const requestedPresetId = String(payload.presetId || '').trim();

  return withSpreadsheetMutationLock_(function() {
    const sheet = openInputPresetsSheet_();
    const rows = readInputPresetRows_(sheet);
    const entries = inputPresetEntriesFromRows_(rows);
    let existing = null;
    for (var i = 0; i < entries.length; i += 1) {
      if (entries[i].model.presetId === requestedPresetId) {
        existing = entries[i];
        break;
      }
    }
    if (requestedPresetId && !existing) {
      return { ok: false, error: 'プリセットが見つかりません。' };
    }
    if (!existing && entries.length >= MAX_INPUT_PRESETS) {
      return { ok: false, error: '入力プリセットは' + MAX_INPUT_PRESETS + '件まで保存できます。' };
    }

    const source = Object.assign({}, existing ? existing.model : {}, payload);
    const normalized = normalizeInputPresetForSave_(source);
    if (!normalized.ok) return normalized;

    const now = currentUpdatedAt_();
    const preset = normalized.preset;
    preset.presetId = existing ? existing.model.presetId : Utilities.getUuid();
    preset.createdAt = existing ? existing.model.createdAt : now;
    preset.updatedAt = now;
    if (existing) {
      preset.orderIndex = existing.model.orderIndex;
    } else {
      const requestedOrderIndex = Number(payload.orderIndex);
      const tailOrderIndex = entries.reduce(function(maximum, entry) {
        const savedOrderIndex = Number(entry.model.orderIndex);
        return Number.isFinite(savedOrderIndex) ? Math.max(maximum, savedOrderIndex) : maximum;
      }, -1) + 1;
      preset.orderIndex = Number.isInteger(requestedOrderIndex) && requestedOrderIndex >= 0
        ? requestedOrderIndex
        : tailOrderIndex;
    }

    const rowNumber = existing ? existing.rowNumber : rows.length + 2;
    if (rowNumber > sheet.getMaxRows()) {
      sheet.insertRowsAfter(sheet.getMaxRows(), rowNumber - sheet.getMaxRows());
    }
    sheet.getRange(rowNumber, 1, 1, INPUT_PRESET_COLUMN_COUNT).setValues([inputPresetToRow_(preset)]);
    return { ok: true, preset: preset };
  });
}

function deleteInputPreset(data) {
  assertEditToken_(data);
  const presetId = String(data && data.presetId || '').trim();
  if (!presetId) return { ok: false, error: '削除するプリセットを指定してください。' };

  return withSpreadsheetMutationLock_(function() {
    const sheet = openInputPresetsSheet_();
    const snapshot = readFormulaPreservingDataRows_(sheet, INPUT_PRESET_COLUMN_COUNT);
    const entries = inputPresetEntriesFromRows_(snapshot.values);
    let deleted = null;
    entries.forEach(function(entry) {
      entry.outputRow = snapshot.outputRows[entry.rowIndex].slice();
      if (entry.model.presetId === presetId) deleted = entry;
    });
    if (!deleted) return { ok: false, error: 'プリセットが見つかりません。' };

    const keptEntries = entries.filter(function(entry) {
      return entry.model.presetId !== presetId;
    }).sort(function(left, right) {
      return compareInputPresets_(left.model, right.model);
    });
    const outputRows = keptEntries.map(function(entry, index) {
      entry.model.orderIndex = index;
      const fixedWidth = inputPresetToRow_(entry.model);
      fixedWidth.forEach(function(value, columnIndex) {
        entry.outputRow[columnIndex] = value;
      });
      return entry.outputRow;
    });
    rewriteFixedWidthDataRows_(sheet, INPUT_PRESET_COLUMN_COUNT, snapshot.values.length, outputRows);
    return {
      ok: true,
      preset: deleted.model,
      presets: keptEntries.map(function(entry) { return entry.model; })
    };
  });
}

function updateInputPresetOrder(data) {
  assertEditToken_(data);
  return withSpreadsheetMutationLock_(function() {
    const sheet = openInputPresetsSheet_();
    const rows = readInputPresetRows_(sheet);
    const entries = inputPresetEntriesFromRows_(rows);
    const requestedIds = data && data.presetIds;
    if (!Array.isArray(requestedIds)) {
      return { ok: false, error: 'プリセットの並び順が不正です。' };
    }

    const normalizedIds = [];
    const requestedIdSet = {};
    for (var i = 0; i < requestedIds.length; i += 1) {
      const presetId = String(requestedIds[i] || '').trim();
      if (!presetId || requestedIdSet[presetId]) {
        return { ok: false, error: 'プリセットIDは重複なく指定してください。' };
      }
      requestedIdSet[presetId] = true;
      normalizedIds.push(presetId);
    }

    const savedIdSet = {};
    entries.forEach(function(entry) { savedIdSet[entry.model.presetId] = true; });
    if (normalizedIds.length !== entries.length || normalizedIds.some(function(presetId) { return !savedIdSet[presetId]; })) {
      return { ok: false, error: '保存済みプリセットと並べ替え対象が一致しません。' };
    }

    const orderById = {};
    normalizedIds.forEach(function(presetId, index) { orderById[presetId] = index; });
    const orderValues = rows.map(function(row) {
      const presetId = String(row && row[0] || '').trim();
      return [presetId ? orderById[presetId] : row[3]];
    });
    if (orderValues.length > 0) {
      sheet.getRange(2, 4, orderValues.length, 1).setValues(orderValues);
    }
    entries.forEach(function(entry) {
      entry.model.orderIndex = orderById[entry.model.presetId];
    });
    return { ok: true, presets: sortedInputPresetModels_(entries) };
  });
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
      SHARE_LINKS_LEGACY_HEADERS.length,
      typeof sheet.getLastColumn === 'function' ? sheet.getLastColumn() : SHARE_LINKS_LEGACY_HEADERS.length
    );
    const headerValues = sheet.getRange(1, 1, 1, columnCount).getValues()[0];
    SHARE_LINKS_LEGACY_HEADERS.forEach(function(header, index) {
      if (headerValues[index] === '' || headerValues[index] == null) {
        sheet.getRange(1, index + 1).setValue(header);
        headerValues[index] = header;
      }
    });
    if (headerValues.indexOf(SHARE_ROUTE_TARGETS_HEADER) === -1) {
      const targetColumn = columnCount + 1;
      sheet.getRange(1, targetColumn).setValue(SHARE_ROUTE_TARGETS_HEADER);
    }
  }
  const styleColumnCount = Math.max(
    SHARE_LINKS_HEADERS.length,
    typeof sheet.getLastColumn === 'function' ? sheet.getLastColumn() : SHARE_LINKS_HEADERS.length
  );
  sheet.getRange(1, 1, 1, styleColumnCount)
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

function shareLinksHeaderIndexMap_(headerRow) {
  const map = {};
  (Array.isArray(headerRow) ? headerRow : []).forEach(function(header, index) {
    const name = String(header || '').trim();
    if (name && map[name] == null) map[name] = index;
  });
  SHARE_LINKS_LEGACY_HEADERS.forEach(function(header, index) {
    if (map[header] == null) map[header] = index;
  });
  return map;
}

function shareRowValue_(row, headerMap, name) {
  const index = headerMap && headerMap[name];
  return index == null ? '' : row[index];
}

function normalizeShareRouteTargetId_(value) {
  const id = String(value == null ? '' : value).trim();
  if (!id || id.length > TRACK_ID_MAX_LENGTH || /[\u0000-\u001f\u007f]/.test(id)
      || /^[=+\-@]/.test(id)) return '';
  return id;
}

function shareRouteTargetKey_(type, id) {
  return String(type || '') + '\0' + String(id || '');
}

function normalizeShareRouteTargetsWithSet_(values, allowedKeySet) {
  if (!Array.isArray(values)) return [];
  const seen = {};
  const targets = [];
  values.forEach(function(value) {
    const type = String(value && value.type || '').trim();
    const id = normalizeShareRouteTargetId_(value && value.id);
    if (!SHARE_ROUTE_TARGET_TYPES[type] || !id) return;
    const key = shareRouteTargetKey_(type, id);
    if (seen[key] || (allowedKeySet && !allowedKeySet[key])) return;
    seen[key] = true;
    targets.push({ type: type, id: id });
  });
  return targets;
}

function getExistingShareRouteTargetKeySet_() {
  const allowed = {};
  try {
    getRouteGroups().forEach(function(group) {
      const id = normalizeShareRouteTargetId_(group && (group.routeId || group.id));
      if (id) allowed[shareRouteTargetKey_('pin-route', id)] = true;
    });
  } catch (_routeError) {
    // A missing route sheet must not broaden the selection.
  }
  try {
    const sheets = openValidatedTrackSheets_();
    readTrackMetadataEntries_(sheets.tracksSheet).forEach(function(entry) {
      try {
        const metadata = trackMetadataFromRow_(entry.row, entry.rowNumber);
        const type = metadata.sourceType === 'gpx' ? 'gpx-route'
          : (metadata.sourceType === 'geojson' ? 'geojson-route' : '');
        if (type) allowed[shareRouteTargetKey_(type, metadata.trackId)] = true;
      } catch (_metadataError) {
        // Corrupted tracks are never selectable for sharing.
      }
    });
  } catch (_trackError) {
    // Track storage is optional for legacy installations.
  }
  return allowed;
}

function normalizeShareRouteTargets_(values) {
  return normalizeShareRouteTargetsWithSet_(values, getExistingShareRouteTargetKeySet_());
}

function serializeShareRouteTargets_(values) {
  return JSON.stringify({ v: 1, targets: normalizeShareRouteTargets_(values) });
}

function parseShareRouteTargetsJson_(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return { state: 'legacy', targets: [] };
  try {
    const parsed = JSON.parse(raw);
    if (!parsed || parsed.v !== 1 || !Array.isArray(parsed.targets)) {
      return { state: 'invalid', targets: [] };
    }
    const targets = [];
    const seen = {};
    for (var i = 0; i < parsed.targets.length; i += 1) {
      const item = parsed.targets[i];
      const type = String(item && item.type || '').trim();
      const id = normalizeShareRouteTargetId_(item && item.id);
      if (!SHARE_ROUTE_TARGET_TYPES[type] || !id) return { state: 'invalid', targets: [] };
      const key = shareRouteTargetKey_(type, id);
      if (seen[key]) continue;
      seen[key] = true;
      targets.push({ type: type, id: id });
    }
    return { state: 'valid', targets: targets };
  } catch (_error) {
    return { state: 'invalid', targets: [] };
  }
}

function shareRowToLink_(row, headerMap) {
  headerMap = headerMap || shareLinksHeaderIndexMap_(SHARE_LINKS_HEADERS);
  const targetSelection = parseShareRouteTargetsJson_(shareRowValue_(row, headerMap, SHARE_ROUTE_TARGETS_HEADER));
  return {
    createdAt: shareRowValue_(row, headerMap, 'createdAt') ? String(shareRowValue_(row, headerMap, 'createdAt')) : '',
    label: normalizeShareLinkLabel_(shareRowValue_(row, headerMap, 'label')),
    token: shareRowValue_(row, headerMap, 'token') ? String(shareRowValue_(row, headerMap, 'token')) : '',
    tags: PinData.deserializeTags(shareRowValue_(row, headerMap, 'tags') || ''),
    tagMode: String(shareRowValue_(row, headerMap, 'tagMode') || 'or') === 'and' ? 'and' : 'or',
    enabled: isShareLinkEnabled_(shareRowValue_(row, headerMap, 'enabled')),
    revokedAt: shareRowValue_(row, headerMap, 'revokedAt') ? String(shareRowValue_(row, headerMap, 'revokedAt')) : '',
    colors: deserializeShareColors_(shareRowValue_(row, headerMap, 'colors') || ''),
    routeIds: deserializeShareRouteIds_(shareRowValue_(row, headerMap, 'routeIds') || ''),
    routeTargetsState: targetSelection.state,
    routeTargets: targetSelection.targets
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
//  Drive写真取込ソース（Phase 9-1）
// ============================================================

const DRIVE_PHOTO_IMPORT_MAX_ITEMS = 20;
const DRIVE_PHOTO_IMPORT_MAX_FILE_BYTES = 15 * 1024 * 1024;
const DRIVE_PHOTO_IMPORT_MAX_TOTAL_BYTES = 100 * 1024 * 1024;
const DRIVE_PHOTO_IMPORT_MAX_FOLDER_ENTRIES = 500;
const DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_DEPTH = 64;
const DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_NODES = 256;
const DRIVE_PHOTO_IMPORT_ID_PATTERN = /^[A-Za-z0-9_-]{10,200}$/;
const DRIVE_PHOTO_IMPORT_DANGEROUS_IDS = {
  '__proto__': true,
  'constructor': true,
  'prototype': true
};
const DRIVE_PHOTO_IMPORT_SHORTCUT_MIME = 'application/vnd.google-apps.shortcut';
const DRIVE_PHOTO_IMPORT_UNSAFE_DISPLAY_NAME_PATTERN = /[\x00-\x1f\x7f-\x9f\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069\ufeff]/;

const DRIVE_PHOTO_IMPORT_ERROR_MESSAGES = {
  DRIVE_IMPORT_ACCESS_DENIED: 'Drive写真の取込権限を確認してください。',
  DRIVE_IMPORT_ROOT_MISSING: 'Drive写真の取込元フォルダが設定されていません。',
  DRIVE_IMPORT_FOLDER_ID_INVALID: 'Driveフォルダを確認してください。',
  DRIVE_IMPORT_FOLDER_NOT_FOUND: 'Driveフォルダを開けませんでした。',
  DRIVE_IMPORT_FOLDER_READ_FAILED: 'Drive写真の一覧を取得できませんでした。もう一度開いてください。',
  DRIVE_IMPORT_FOLDER_OUTSIDE_ROOT: 'このDriveフォルダは取込元の範囲外です。',
  DRIVE_IMPORT_FOLDER_TOO_LARGE: 'このフォルダの項目数が多すぎます。写真を小さなフォルダへ分けてください。',
  DRIVE_IMPORT_FILE_ID_INVALID: 'Driveの写真を確認してください。',
  DRIVE_IMPORT_FILE_NOT_FOUND: 'Driveの写真を開けませんでした。',
  DRIVE_IMPORT_FILE_OUTSIDE_ROOT: 'このDriveの写真は取込元の範囲外です。',
  DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED: '対応していない写真形式です。',
  DRIVE_IMPORT_FILE_TOO_LARGE: '1枚の写真は15MB以内にしてください。',
  DRIVE_IMPORT_FILE_READ_FAILED: 'Driveの写真を読み込めませんでした。もう一度選択してください。'
};

function drivePhotoImportHasOwn_(value, key) {
  return !!value && Object.prototype.hasOwnProperty.call(value, key);
}

function drivePhotoImportFailure_(code) {
  const safeCode = drivePhotoImportHasOwn_(DRIVE_PHOTO_IMPORT_ERROR_MESSAGES, code)
    ? code : 'DRIVE_IMPORT_FILE_READ_FAILED';
  return {
    ok: false,
    errorCode: safeCode,
    error: DRIVE_PHOTO_IMPORT_ERROR_MESSAGES[safeCode]
  };
}

function drivePhotoImportError_(code) {
  const error = new Error(code);
  error.code = code;
  return error;
}

function drivePhotoImportFailureFromError_(error, fallbackCode) {
  let code = fallbackCode;
  try {
    if (error
        && drivePhotoImportHasOwn_(error, 'code')
        && typeof error.code === 'string'
        && drivePhotoImportHasOwn_(DRIVE_PHOTO_IMPORT_ERROR_MESSAGES, error.code)) {
      code = error.code;
    }
  } catch (_error) {
    code = fallbackCode;
  }
  return drivePhotoImportFailure_(code);
}

function isValidDrivePhotoImportId_(value) {
  if (typeof value !== 'string') return false;
  if (!DRIVE_PHOTO_IMPORT_ID_PATTERN.test(value)) return false;
  return !drivePhotoImportHasOwn_(DRIVE_PHOTO_IMPORT_DANGEROUS_IDS, value);
}

function getDrivePhotoImportRootId_() {
  const rootFolderId = getRootFolderId_();
  if (!isValidDrivePhotoImportId_(rootFolderId)) {
    throw drivePhotoImportError_('DRIVE_IMPORT_ROOT_MISSING');
  }
  return rootFolderId;
}

function drivePhotoImportTrashState_(item) {
  try {
    if (!item || typeof item.isTrashed !== 'function') return 'error';
    return item.isTrashed() ? 'trashed' : 'active';
  } catch (_error) {
    return 'error';
  }
}

function drivePhotoImportItemId_(item) {
  try {
    const id = item && typeof item.getId === 'function' ? item.getId() : '';
    return isValidDrivePhotoImportId_(id) ? id : '';
  } catch (_error) {
    return '';
  }
}

function isDriveItemWithinRoot_(item, rootFolderId, allowRootItem) {
  if (!item || !isValidDrivePhotoImportId_(rootFolderId)
      || drivePhotoImportTrashState_(item) !== 'active') {
    return false;
  }
  const initialId = drivePhotoImportItemId_(item);
  if (!initialId) return false;
  if (allowRootItem && initialId === rootFolderId) return true;

  const queue = [{ item: item, depth: 0 }];
  const visited = Object.create(null);
  let cursor = 0;
  let inspected = 0;
  let reachedRoot = false;
  while (cursor < queue.length) {
    if (inspected >= DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_NODES) return false;
    const entry = queue[cursor++];
    const entryTrashState = drivePhotoImportTrashState_(entry.item);
    if (entryTrashState === 'error') return false;
    if (entryTrashState === 'trashed') continue;
    const currentId = drivePhotoImportItemId_(entry.item);
    if (!currentId || drivePhotoImportHasOwn_(visited, currentId)) continue;
    visited[currentId] = true;
    inspected += 1;
    if (currentId === rootFolderId) {
      reachedRoot = true;
      continue;
    }
    if (entry.depth >= DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_DEPTH) return false;

    let parents;
    try {
      parents = entry.item.getParents();
      while (parents.hasNext()) {
        const parent = parents.next();
        const parentTrashState = drivePhotoImportTrashState_(parent);
        if (parentTrashState === 'error') return false;
        if (parentTrashState === 'trashed') continue;
        const parentId = drivePhotoImportItemId_(parent);
        if (!parentId) return false;
        if (parentId === rootFolderId) {
          reachedRoot = true;
          continue;
        }
        if (!drivePhotoImportHasOwn_(visited, parentId)) {
          queue.push({ item: parent, depth: entry.depth + 1 });
        }
        if (queue.length > DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_NODES) return false;
      }
    } catch (_error) {
      return false;
    }
  }
  return reachedRoot;
}

function isDriveFolderWithinRoot_(folder, rootFolderId) {
  return isDriveItemWithinRoot_(folder, rootFolderId, true);
}

function isDriveFileWithinRoot_(file, rootFolderId) {
  try {
    if (!file || file.getMimeType() === DRIVE_PHOTO_IMPORT_SHORTCUT_MIME) return false;
  } catch (_error) {
    return false;
  }
  return isDriveItemWithinRoot_(file, rootFolderId, false);
}

function drivePhotoImportExtensionOf_(name) {
  const match = String(name || '').trim().match(/\.([^.\s]+)$/);
  return match ? match[1].toLowerCase() : '';
}

function classifyDrivePhotoImportFile_(value) {
  const mimeKinds = {
    'image/jpeg': 'jpeg',
    'image/jpg': 'jpeg',
    'image/png': 'png',
    'image/webp': 'webp',
    'image/heic': 'heic',
    'image/heif': 'heic'
  };
  const extensionKinds = {
    jpg: 'jpeg', jpeg: 'jpeg', png: 'png', webp: 'webp', heic: 'heic', heif: 'heic'
  };
  const deniedExtensions = {
    svg: true, gif: true, bmp: true, tif: true, tiff: true, pdf: true
  };
  const name = value && typeof value.name === 'string' ? value.name : '';
  const type = value && typeof value.type === 'string' ? value.type.trim().toLowerCase() : '';
  const extension = drivePhotoImportExtensionOf_(name);
  const mimeKind = mimeKinds[type] || '';
  const extensionKind = extensionKinds[extension] || '';
  const genericMime = !type || type === 'application/octet-stream' || type === 'binary/octet-stream';
  const explicitlyDenied = !!deniedExtensions[extension]
    || type === 'image/svg+xml'
    || type === 'image/gif'
    || type === 'image/bmp'
    || type === 'image/tiff'
    || type === 'application/pdf'
    || /^video\//.test(type);
  if (explicitlyDenied || (type && !mimeKind && !genericMime)) {
    return { supported: false, kind: '', normalizedMimeType: '', extension: '' };
  }
  if (extension && !extensionKind && !mimeKind) {
    return { supported: false, kind: '', normalizedMimeType: '', extension: '' };
  }
  if (mimeKind && extensionKind && mimeKind !== extensionKind) {
    return { supported: false, kind: '', normalizedMimeType: '', extension: '' };
  }
  const kind = mimeKind || extensionKind;
  if (!kind) return { supported: false, kind: '', normalizedMimeType: '', extension: '' };
  const normalizedMimeType = mimeKind
    ? (type === 'image/jpg' ? 'image/jpeg' : type)
    : (kind === 'jpeg' ? 'image/jpeg' : 'image/' + extension);
  return {
    supported: true,
    kind: kind,
    normalizedMimeType: normalizedMimeType,
    extension: extension === 'jpeg' ? 'jpg' : extension
  };
}

function drivePhotoImportFolderName_(item) {
  try {
    const name = String(item.getName());
    if (!name || name.length > 255 || DRIVE_PHOTO_IMPORT_UNSAFE_DISPLAY_NAME_PATTERN.test(name)) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_NOT_FOUND');
    }
    return name;
  } catch (_error) {
    throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_NOT_FOUND');
  }
}

function isValidDrivePhotoImportFileName_(value) {
  return typeof value === 'string'
    && value.length > 0
    && value.length <= 255
    && !DRIVE_PHOTO_IMPORT_UNSAFE_DISPLAY_NAME_PATTERN.test(value)
    && !/[\\/]/.test(value);
}

function drivePhotoImportSort_(values) {
  values.sort(function(a, b) {
    const left = a.name.toLocaleLowerCase();
    const right = b.name.toLocaleLowerCase();
    if (left < right) return -1;
    if (left > right) return 1;
    return a.id < b.id ? -1 : (a.id > b.id ? 1 : 0);
  });
  return values;
}

function drivePhotoImportModifiedAt_(file) {
  try {
    const value = file.getLastUpdated();
    const date = value instanceof Date ? value : new Date(value);
    return Number.isFinite(date.getTime()) ? date.toISOString() : '';
  } catch (_error) {
    return '';
  }
}

function getDrivePhotoImportParent_(folder, rootFolderId) {
  const folderId = drivePhotoImportItemId_(folder);
  if (!folderId || folderId === rootFolderId) return null;
  try {
    const parents = folder.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      if (!isDriveFolderWithinRoot_(parent, rootFolderId)) continue;
      const parentId = drivePhotoImportItemId_(parent);
      if (!parentId) continue;
      return { id: parentId, name: drivePhotoImportFolderName_(parent) };
    }
  } catch (_error) {
    throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_NOT_FOUND');
  }
  throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_NOT_FOUND');
}

function isDriveOriginalArchiveFolder_(folder, rootFolderId) {
  if (!folder || drivePhotoImportTrashState_(folder) !== 'active'
      || String(folder.getName()) !== 'original') return false;
  const parents = folder.getParents();
  while (parents.hasNext()) {
    if (String(parents.next().getId()) === String(rootFolderId)) return true;
  }
  return false;
}

function resolveDrivePhotoImportFolderId_(payload, rootFolderId) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
    throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_ID_INVALID');
  }
  let inheritedFolderId = false;
  try {
    inheritedFolderId = !drivePhotoImportHasOwn_(payload, 'folderId') && 'folderId' in payload;
  } catch (_error) {
    inheritedFolderId = true;
  }
  if (inheritedFolderId) throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_ID_INVALID');
  if (!drivePhotoImportHasOwn_(payload, 'folderId') || payload.folderId === '') return rootFolderId;
  if (!isValidDrivePhotoImportId_(payload.folderId)) {
    throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_ID_INVALID');
  }
  return payload.folderId;
}

function getDrivePhotoImportAssociations_() {
  const associations = {
    importedSourceIds: new Set(),
    managedFileIds: new Set()
  };
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mapSheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!mapSheet) return associations;

    const mapRows = mapSheet.getDataRange().getValues();
    if (!mapRows.length) return associations;
    const mapHeaders = mapRows[0].map(String);
    const mapPinIdColumn = mapHeaders.indexOf('ID');
    const mapFileIdColumn = mapHeaders.indexOf('ファイルID');
    if (mapPinIdColumn < 0 || mapFileIdColumn < 0) return associations;
    const livePinIds = new Set();
    mapRows.slice(1).forEach(function(row) {
      const pinId = String(row[mapPinIdColumn] || '');
      const managedFileId = String(row[mapFileIdColumn] || '');
      if (pinId) livePinIds.add(pinId);
      if (managedFileId) associations.managedFileIds.add(managedFileId);
    });

    const receiptSheet = spreadsheet.getSheetByName(IMPORT_RECEIPTS_SHEET_NAME);
    if (!receiptSheet) return associations;
    const receiptRows = receiptSheet.getDataRange().getValues();
    if (!receiptRows.length) return associations;
    const receiptHeaders = receiptRows[0].map(String);
    const stateColumn = receiptHeaders.indexOf('state');
    const pinIdColumn = receiptHeaders.indexOf('pinId');
    const sourceIdColumn = receiptHeaders.indexOf('sourceDriveFileId');
    if (stateColumn < 0 || pinIdColumn < 0 || sourceIdColumn < 0) return associations;
    receiptRows.slice(1).forEach(function(row) {
      const sourceId = String(row[sourceIdColumn] || '');
      const pinId = String(row[pinIdColumn] || '');
      if (String(row[stateColumn] || '') === IMPORT_RECEIPT_STATES.COMPLETED
          && isValidDrivePhotoImportId_(sourceId)
          && livePinIds.has(pinId)) {
        associations.importedSourceIds.add(sourceId);
      }
    });
  } catch (_error) {
    throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_READ_FAILED');
  }
  return associations;
}

function legacyListDrivePhotoImportFolder_(payload) {
  try {
    try {
      assertEditToken_(payload);
    } catch (_error) {
      throw drivePhotoImportError_('DRIVE_IMPORT_ACCESS_DENIED');
    }
    const rootFolderId = getDrivePhotoImportRootId_();
    const folderId = resolveDrivePhotoImportFolderId_(payload, rootFolderId);
    let folder;
    try {
      folder = DriveApp.getFolderById(folderId);
    } catch (_error) {
      throw drivePhotoImportError_(
        folderId === rootFolderId ? 'DRIVE_IMPORT_ROOT_MISSING' : 'DRIVE_IMPORT_FOLDER_NOT_FOUND'
      );
    }
    if (!isDriveFolderWithinRoot_(folder, rootFolderId)) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_OUTSIDE_ROOT');
    }
    if (isDriveOriginalArchiveFolder_(folder, rootFolderId)) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_NOT_FOUND');
    }

    const associations = getDrivePhotoImportAssociations_();
    const folders = [];
    const photos = [];
    let ignoredUnsupportedFileCount = 0;
    let scannedEntries = 0;
    const countEntry = function() {
      scannedEntries += 1;
      if (scannedEntries > DRIVE_PHOTO_IMPORT_MAX_FOLDER_ENTRIES) {
        throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_TOO_LARGE');
      }
    };

    const childFolders = folder.getFolders();
    while (childFolders.hasNext()) {
      const child = childFolders.next();
      countEntry();
      const childTrashState = drivePhotoImportTrashState_(child);
      if (childTrashState === 'error') throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_NOT_FOUND');
      if (childTrashState === 'trashed') continue;
      if (folderId === rootFolderId && isDriveOriginalArchiveFolder_(child, rootFolderId)) continue;
      const id = drivePhotoImportItemId_(child);
      if (!id) continue;
      folders.push({ id: id, name: drivePhotoImportFolderName_(child) });
    }

    const childFiles = folder.getFiles();
    while (childFiles.hasNext()) {
      const file = childFiles.next();
      countEntry();
      const fileTrashState = drivePhotoImportTrashState_(file);
      if (fileTrashState === 'error') throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_NOT_FOUND');
      if (fileTrashState === 'trashed') continue;
      let id = '';
      let name = '';
      let type = '';
      let sizeBytes = 0;
      try {
        id = file.getId();
        name = file.getName();
        type = file.getMimeType();
        sizeBytes = Number(file.getSize());
      } catch (_error) {
        ignoredUnsupportedFileCount += 1;
        continue;
      }
      if (associations.managedFileIds.has(String(id))
          && !associations.importedSourceIds.has(String(id))) continue;
      const classification = type === DRIVE_PHOTO_IMPORT_SHORTCUT_MIME
        ? { supported: false }
        : classifyDrivePhotoImportFile_({ name: name, type: type, size: sizeBytes });
      const modifiedAt = drivePhotoImportModifiedAt_(file);
      if (!isValidDrivePhotoImportId_(id)
          || !isValidDrivePhotoImportFileName_(name)
          || !classification.supported
          || !Number.isSafeInteger(sizeBytes)
          || sizeBytes <= 0
          || sizeBytes > DRIVE_PHOTO_IMPORT_MAX_FILE_BYTES
          || !modifiedAt) {
        ignoredUnsupportedFileCount += 1;
        continue;
      }
      photos.push({
        id: id,
        name: String(name),
        mimeType: classification.normalizedMimeType,
        sizeBytes: sizeBytes,
        modifiedAt: modifiedAt,
        kind: classification.kind,
        imported: associations.importedSourceIds.has(id)
      });
    }

    drivePhotoImportSort_(folders);
    drivePhotoImportSort_(photos);
    return {
      ok: true,
      folder: {
        id: folderId,
        name: drivePhotoImportFolderName_(folder),
        isRoot: folderId === rootFolderId
      },
      parent: getDrivePhotoImportParent_(folder, rootFolderId),
      folders: folders,
      photos: photos,
      ignoredUnsupportedFileCount: ignoredUnsupportedFileCount,
      counts: { folders: folders.length, photos: photos.length }
    };
  } catch (error) {
    return drivePhotoImportFailureFromError_(error, 'DRIVE_IMPORT_FOLDER_NOT_FOUND');
  }
}

function legacyReadDrivePhotoImportFile_(payload) {
  try {
    try {
      assertEditToken_(payload);
    } catch (_error) {
      throw drivePhotoImportError_('DRIVE_IMPORT_ACCESS_DENIED');
    }
    const rootFolderId = getDrivePhotoImportRootId_();
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)
        || !drivePhotoImportHasOwn_(payload, 'fileId')
        || !isValidDrivePhotoImportId_(payload.fileId)) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_ID_INVALID');
    }
    let file;
    try {
      file = DriveApp.getFileById(payload.fileId);
    } catch (_error) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_NOT_FOUND');
    }
    const fileTrashState = drivePhotoImportTrashState_(file);
    if (fileTrashState === 'error') {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_READ_FAILED');
    }
    if (fileTrashState === 'trashed') {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_NOT_FOUND');
    }
    if (!isDriveFileWithinRoot_(file, rootFolderId)) {
      let shortcut = false;
      try {
        shortcut = file.getMimeType() === DRIVE_PHOTO_IMPORT_SHORTCUT_MIME;
      } catch (_error) {
        // Fall through to the outside-root failure without exposing Drive details.
      }
      throw drivePhotoImportError_(
        shortcut ? 'DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED' : 'DRIVE_IMPORT_FILE_OUTSIDE_ROOT'
      );
    }

    let id;
    let name;
    let type;
    let modifiedAt;
    try {
      id = file.getId();
      name = file.getName();
      type = file.getMimeType();
      modifiedAt = drivePhotoImportModifiedAt_(file);
    } catch (_error) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_READ_FAILED');
    }
    const classification = classifyDrivePhotoImportFile_({ name: name, type: type });
    if (!classification.supported) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED');
    }
    if (String(id) !== payload.fileId
        || !isValidDrivePhotoImportFileName_(name)
        || !modifiedAt) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_READ_FAILED');
    }

    let bytes;
    try {
      const blob = file.getBlob();
      bytes = blob.getBytes();
    } catch (_error) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_READ_FAILED');
    }
    const sizeBytes = bytes && Number.isSafeInteger(bytes.length) ? bytes.length : -1;
    if (sizeBytes <= 0) {
      bytes = null;
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_READ_FAILED');
    }
    if (sizeBytes > DRIVE_PHOTO_IMPORT_MAX_FILE_BYTES) {
      bytes = null;
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_TOO_LARGE');
    }
    let base64;
    try {
      base64 = Utilities.base64Encode(bytes);
    } catch (_error) {
      bytes = null;
      throw drivePhotoImportError_('DRIVE_IMPORT_FILE_READ_FAILED');
    }
    bytes = null;
    return {
      ok: true,
      file: {
        id: String(id),
        name: String(name),
        mimeType: classification.normalizedMimeType,
        sizeBytes: sizeBytes,
        modifiedAt: modifiedAt,
        kind: classification.kind,
        base64: base64
      }
    };
  } catch (error) {
    return drivePhotoImportFailureFromError_(error, 'DRIVE_IMPORT_FILE_READ_FAILED');
  }
}

function listDrivePhotoImportFolder(payload) {
  try {
    try {
      assertEditToken_(payload);
    } catch (_error) {
      throw drivePhotoImportError_('DRIVE_IMPORT_ACCESS_DENIED');
    }
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_ID_INVALID');
    }
    let inheritedFolderId = false;
    try {
      inheritedFolderId = !drivePhotoImportHasOwn_(payload, 'folderId') && 'folderId' in payload;
    } catch (_error) {
      inheritedFolderId = true;
    }
    if (inheritedFolderId) throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_ID_INVALID');
    if (drivePhotoImportHasOwn_(payload, 'folderId') && payload.folderId !== '') {
      if (!isValidDrivePhotoImportId_(payload.folderId)) {
        throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_ID_INVALID');
      }
      throw drivePhotoImportError_('DRIVE_IMPORT_FOLDER_OUTSIDE_ROOT');
    }
    const inbox = listDriveMediaInbox({
      __editToken: payload.__editToken,
      mediaKind: 'photo'
    });
    if (!inbox || inbox.ok !== true) {
      const mediaCode = inbox && String(inbox.errorCode || '');
      const mappedCode = mediaCode === 'DRIVE_MEDIA_ACCESS_DENIED'
        ? 'DRIVE_IMPORT_ACCESS_DENIED'
        : (mediaCode === 'DRIVE_MEDIA_ROOT_MISSING'
          ? 'DRIVE_IMPORT_ROOT_MISSING'
          : (mediaCode === 'DRIVE_MEDIA_INBOX_TOO_LARGE'
            ? 'DRIVE_IMPORT_FOLDER_TOO_LARGE' : 'DRIVE_IMPORT_FOLDER_READ_FAILED'));
      throw drivePhotoImportError_(mappedCode);
    }
    const photos = Array.isArray(inbox.items) ? inbox.items : [];
    return {
      ok: true,
      folder: { id: '', name: '取込Inbox', isRoot: true },
      parent: null,
      folders: [],
      photos: photos,
      ignoredUnsupportedFileCount: 0,
      counts: { folders: 0, photos: photos.length }
    };
  } catch (error) {
    return drivePhotoImportFailureFromError_(error, 'DRIVE_IMPORT_FOLDER_READ_FAILED');
  }
}

function readDrivePhotoImportFile(payload) {
  try {
    assertEditToken_(payload);
  } catch (_error) {
    return drivePhotoImportFailureFromError_(
      drivePhotoImportError_('DRIVE_IMPORT_ACCESS_DENIED'),
      'DRIVE_IMPORT_FILE_READ_FAILED'
    );
  }
  return readDriveMediaImportFile_(payload, 'photo');
}

// ============================================================
//  データ操作
// ============================================================

function startupTimingNow_() {
  if (typeof performance !== 'undefined' && performance
      && typeof performance.now === 'function') return performance.now();
  return Date.now();
}

function logStartupGasStage_(operation, stage, result, startedAt) {
  const durationMs = Math.max(0, Math.round(startupTimingNow_() - startedAt));
  if (typeof Logger !== 'undefined' && Logger && typeof Logger.log === 'function') {
    Logger.log('[startup] gas=' + operation + ' stage=' + stage
      + ' result=' + result + ' durationMs=' + durationMs);
  }
}

function measureStartupGasStage_(operation, stage, callback) {
  const startedAt = startupTimingNow_();
  try {
    const result = callback();
    logStartupGasStage_(operation, stage, 'success', startedAt);
    return result;
  } catch (caught) {
    logStartupGasStage_(operation, stage, 'failure', startedAt);
    throw caught;
  }
}

function getMapData() {
  const totalStartedAt = startupTimingNow_();
  try {
    const sheetData = measureStartupGasStage_('getMapData', 'sheet-read', function() {
      const sheet = openMapInfoSheet_();
      if (sheet.getLastRow() === 0) return { empty: true, rows: [] };
      return { empty: false, rows: sheet.getDataRange().getValues() };
    });
    const pins = measureStartupGasStage_('getMapData', 'row-convert', function() {
      return sheetData.empty ? [] : PinData.rowsToPins(sheetData.rows);
    });
    const enrichedPins = measureStartupGasStage_('getMapData', 'drive-info', function() {
      return pins.map(function(pin) { return enrichPinWithDriveMeta_(pin); });
    });
    const response = measureStartupGasStage_('getMapData', 'response-build', function() {
      return enrichedPins.map(function(pin) { return toClientPin_(pin); });
    });
    logStartupGasStage_('getMapData', 'total', 'success', totalStartedAt);
    return response;
  } catch (caught) {
    logStartupGasStage_('getMapData', 'total', 'failure', totalStartedAt);
    throw caught;
  }
}

function toClientPin_(pin) {
  const projected = Object.assign({}, pin, { hasAudio: Boolean(pin.audioId) });
  delete projected.audioId;
  return projected;
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
  const startedAt = startupTimingNow_();
  try {
    const response = readRouteRows_().map(function(entry, index) {
      const group = entry.group;
      group.orderIndex = Number.isFinite(Number(entry.row[9])) ? Number(entry.row[9]) : index;
      return group;
    }).sort(function(a, b) {
      return a.orderIndex - b.orderIndex;
    });
    logStartupGasStage_('getRouteGroups', 'total', 'success', startedAt);
    return response;
  } catch (caught) {
    logStartupGasStage_('getRouteGroups', 'total', 'failure', startedAt);
    throw caught;
  }
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

function validateRoutePinIds_(pinIds) {
  if (!Array.isArray(pinIds)) return { ok: false, error: 'pin_ids_invalid' };
  if (pinIds.length > MAX_ROUTE_PINS) return { ok: false, error: 'too_many_pins' };

  const sheet = openMapInfoSheet_();
  const pinsById = Object.create(null);
  if (sheet.getLastRow() >= 2) {
    PinData.rowsToPins(sheet.getDataRange().getValues()).forEach(function(pin) {
      if (pin.id) pinsById[String(pin.id)] = pin;
    });
  }
  const seen = Object.create(null);
  const normalizedPinIds = [];
  for (var i = 0; i < pinIds.length; i += 1) {
    const pinId = normalizeRoutePinId_(pinIds[i]);
    if (!pinId || !Object.prototype.hasOwnProperty.call(pinsById, pinId)) {
      return { ok: false, error: 'pin_not_found', pinId: pinId || '' };
    }
    if (seen[pinId]) {
      return { ok: false, error: 'pin_ids_duplicated', pinId: pinId };
    }
    const pin = pinsById[pinId];
    const lat = pin && pin.lat;
    const lng = pin && pin.lng;
    if (lat == null || lng == null
        || typeof lat !== 'number' || !Number.isFinite(lat) || lat < -90 || lat > 90
        || typeof lng !== 'number' || !Number.isFinite(lng) || lng < -180 || lng > 180) {
      return { ok: false, error: 'pin_unplaced', pinId: pinId };
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

function getSharePinRouteIds_(shareLink) {
  if (!shareLink || shareLink.routeTargetsState === 'invalid') return [SHARE_ROUTE_NONE_SENTINEL];
  if (shareLink.routeTargetsState === 'valid') {
    const ids = shareLink.routeTargets.filter(function(target) {
      return target.type === 'pin-route';
    }).map(function(target) { return target.id; });
    return ids.length ? ids : [SHARE_ROUTE_NONE_SENTINEL];
  }
  return Array.isArray(shareLink.routeIds) ? shareLink.routeIds : [];
}

function isSharePinRouteAllowedForLink_(routeId, shareLink) {
  return isShareRouteIdAllowed_(routeId, getSharePinRouteIds_(shareLink));
}

function filterSharedRouteGroupsForShareLink_(routeGroups, shareLink) {
  const routeIds = getSharePinRouteIds_(shareLink);
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
  if (!isSharePinRouteAllowedForLink_(routeId, shareLink)) return { ok: false };
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

// ============================================================
//  Track storage
// ============================================================

function trackStorageError_(code, message, retryable) {
  const error = new Error(message);
  error.code = code;
  error.retryable = !!retryable;
  error.isTrackStorageError = true;
  return error;
}

function trackStorageFailureFromError_(error) {
  const known = error && error.isTrackStorageError === true;
  const code = known ? String(error.code || '') : 'TRACK_STORAGE_FAILED';
  const messages = {
    INVALID_TRACK_PAYLOAD: 'トラックの入力内容を確認してください。',
    TRACK_REVISION_PAYLOAD_CONFLICT: '同じリビジョンに異なる内容が送信されました。',
    TRACK_LIMIT_EXCEEDED: '保存できるトラック数の上限に達しています。',
    TRACK_STORAGE_BUSY: '別の更新処理が実行中です。少し待ってから再試行してください。',
    TRACK_NOT_FOUND: '対象のトラックが見つかりません。再読み込みしてください。',
    TRACK_SHEETS_MISSING: 'tracks または track_segments シートが見つかりません。setupSheet() を実行してください。',
    TRACK_STORAGE_CORRUPTED: 'トラック保存領域が不整合です。管理者へ連絡してください。',
    TRACK_METADATA_UPDATE_FAILED: 'トラックの表示設定を更新できませんでした。再試行してください。',
    TRACK_METADATA_DELETE_FAILED: 'トラック情報を削除できませんでした。再試行してください。',
    TRACK_SEGMENTS_DELETE_FAILED: 'トラックのルート区間データを削除できませんでした。再試行してください。',
    TRACK_SHARE_REFERENCES_DELETE_FAILED: '共有リンクのトラック参照を整理できませんでした。再読み込みしてください。',
    TRACK_STORAGE_FAILED: 'トラックを保存できませんでした。再試行してください。'
  };
  const retryableCodes = {
    TRACK_STORAGE_BUSY: true,
    TRACK_METADATA_UPDATE_FAILED: true,
    TRACK_METADATA_DELETE_FAILED: true,
    TRACK_SEGMENTS_DELETE_FAILED: true,
    TRACK_SHARE_REFERENCES_DELETE_FAILED: true,
    TRACK_STORAGE_FAILED: true
  };
  const failure = {
    ok: false,
    error: messages[code] || messages.TRACK_STORAGE_FAILED,
    errorCode: code || 'TRACK_STORAGE_FAILED',
    retryable: known && typeof error.retryable === 'boolean'
      ? error.retryable : !!retryableCodes[code]
  };
  if (known && error.serverDeleted === true) failure.serverDeleted = true;
  return failure;
}

function normalizeTrackIdentifier_(value, field, maxLength) {
  if (typeof value !== 'string') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', field + ' is invalid.', false);
  }
  const raw = value;
  const normalized = raw.trim();
  if (!normalized || normalized.length > maxLength
      || /[\u0000-\u001f\u007f]/.test(normalized)
      || /^[=+\-@\t\r]/.test(raw)
      || /^[=+\-@]/.test(normalized)) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', field + ' is invalid.', false);
  }
  return normalized;
}

function normalizeTrackTime_(value) {
  if (value === '') return '';
  if (typeof value !== 'string') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point time is invalid.', false);
  }
  const source = value;
  const match = source.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d{1,9}))?(Z|[+\-]\d{2}:\d{2})$/);
  if (!match) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point time is invalid.', false);
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const hour = Number(match[4]);
  const minute = Number(match[5]);
  const second = Number(match[6]);
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const monthDays = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  const offset = match[8];
  const offsetHour = offset === 'Z' ? 0 : Number(offset.slice(1, 3));
  const offsetMinute = offset === 'Z' ? 0 : Number(offset.slice(4, 6));
  if (year < 1 || month < 1 || month > 12 || day < 1 || day > monthDays[month - 1]
      || hour > 23 || minute > 59 || second > 59
      || offsetHour > 14 || offsetMinute > 59 || (offsetHour === 14 && offsetMinute !== 0)) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point time is invalid.', false);
  }
  const milliseconds = ((match[7] || '') + '000').slice(0, 3);
  const canonical = match[1] + '-' + match[2] + '-' + match[3]
    + 'T' + match[4] + ':' + match[5] + ':' + match[6] + '.' + milliseconds + offset;
  const parsed = new Date(canonical);
  if (!Number.isFinite(parsed.getTime())) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point time is invalid.', false);
  }
  const normalized = parsed.toISOString();
  if (!/^\d{4}-/.test(normalized) || normalized.slice(0, 4) === '0000') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point time is invalid.', false);
  }
  return normalized;
}

function normalizeTrackPoint_(point) {
  if (!point || typeof point !== 'object' || Array.isArray(point)) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point is invalid.', false);
  }
  const lat = point.lat;
  const lng = point.lng;
  if (typeof lat !== 'number' || !Number.isFinite(lat) || lat < -90 || lat > 90
      || typeof lng !== 'number' || !Number.isFinite(lng) || lng < -180 || lng > 180) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point coordinate is invalid.', false);
  }
  let elevation = point.elevation;
  if (elevation !== null && (typeof elevation !== 'number' || !Number.isFinite(elevation))) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point elevation is invalid.', false);
  }
  return {
    lat: lat,
    lng: lng,
    elevation: elevation,
    time: normalizeTrackTime_(point.time)
  };
}

function normalizeTrackSegments_(segments) {
  if (!Array.isArray(segments) || segments.length > MAX_TRACK_SEGMENTS) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track segments are invalid.', false);
  }
  let pointCount = 0;
  return segments.map(function(segment, segmentIndex) {
    if (!segment || typeof segment !== 'object' || Array.isArray(segment)
        || !Array.isArray(segment.points) || segment.points.length === 0) {
      throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track segment is invalid.', false);
    }
    const points = segment.points.map(normalizeTrackPoint_);
    pointCount += points.length;
    if (pointCount > MAX_TRACK_POINTS) {
      throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point limit exceeded.', false);
    }
    return { index: segmentIndex, points: points };
  });
}

function haversineTrackMeters_(a, b) {
  const radius = 6371008.8;
  const toRadians = Math.PI / 180;
  const lat1 = a.lat * toRadians;
  const lat2 = b.lat * toRadians;
  const deltaLat = (b.lat - a.lat) * toRadians;
  const deltaLng = (b.lng - a.lng) * toRadians;
  const sinLat = Math.sin(deltaLat / 2);
  const sinLng = Math.sin(deltaLng / 2);
  const h = sinLat * sinLat + Math.cos(lat1) * Math.cos(lat2) * sinLng * sinLng;
  return radius * 2 * Math.atan2(Math.sqrt(h), Math.sqrt(Math.max(0, 1 - h)));
}

function computeTrackSummary_(segments) {
  let pointCount = 0;
  let distanceMeters = 0;
  let minElevation = null;
  let maxElevation = null;
  let startTime = '';
  let endTime = '';
  let south = null;
  let west = null;
  let north = null;
  let east = null;
  (Array.isArray(segments) ? segments : []).forEach(function(segment) {
    const points = segment && Array.isArray(segment.points) ? segment.points : [];
    points.forEach(function(point, pointIndex) {
      pointCount += 1;
      if (pointIndex > 0) distanceMeters += haversineTrackMeters_(points[pointIndex - 1], point);
      if (point.elevation != null) {
        minElevation = minElevation == null ? point.elevation : Math.min(minElevation, point.elevation);
        maxElevation = maxElevation == null ? point.elevation : Math.max(maxElevation, point.elevation);
      }
      if (point.time) {
        if (!startTime) startTime = point.time;
        endTime = point.time;
      }
      south = south == null ? point.lat : Math.min(south, point.lat);
      north = north == null ? point.lat : Math.max(north, point.lat);
      west = west == null ? point.lng : Math.min(west, point.lng);
      east = east == null ? point.lng : Math.max(east, point.lng);
    });
  });
  return {
    segmentCount: Array.isArray(segments) ? segments.length : 0,
    pointCount: pointCount,
    distanceMeters: distanceMeters,
    minElevation: minElevation,
    maxElevation: maxElevation,
    startTime: startTime,
    endTime: endTime,
    bounds: south == null ? null : { south: south, west: west, north: north, east: east }
  };
}

function isInvalidTrackSourceName_(value) {
  return value.length > TRACK_SOURCE_NAME_MAX_LENGTH
    || /[\u0000-\u001f\u007f]/.test(value)
    || /[\\/]/.test(value)
    || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(value);
}

function normalizeTrackBundle_(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track payload is invalid.', false);
  }
  const trackId = normalizeTrackIdentifier_(data.trackId, 'trackId', TRACK_ID_MAX_LENGTH);
  const revisionId = normalizeTrackIdentifier_(data.revisionId, 'revisionId', TRACK_REVISION_ID_MAX_LENGTH);
  const name = typeof data.name === 'string' ? data.name.trim() : '';
  if (!name || name.length > TRACK_NAME_MAX_LENGTH) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track name is invalid.', false);
  }
  if (data.description != null && typeof data.description !== 'string') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track description is invalid.', false);
  }
  const description = data.description == null ? '' : data.description;
  if (description.length > TRACK_DESCRIPTION_MAX_LENGTH) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track description is invalid.', false);
  }
  if (typeof data.color !== 'string'
      || PinData.COLOR_OPTIONS.indexOf(data.color.trim().toLowerCase()) === -1) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track color is invalid.', false);
  }
  const color = data.color.trim().toLowerCase();
  const sourceType = typeof data.sourceType === 'string' ? data.sourceType.trim().toLowerCase() : '';
  if (['gpx', 'geojson', 'manual'].indexOf(sourceType) === -1) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track sourceType is invalid.', false);
  }
  if (data.sourceName != null && typeof data.sourceName !== 'string') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track sourceName is invalid.', false);
  }
  const sourceName = data.sourceName == null ? '' : data.sourceName.trim();
  if (isInvalidTrackSourceName_(sourceName)) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track sourceName is invalid.', false);
  }
  const visible = data.visible === undefined ? true : data.visible;
  if (typeof visible !== 'boolean') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track visible is invalid.', false);
  }
  const lineStyle = data.lineStyle === undefined ? 'solid'
    : (typeof data.lineStyle === 'string' ? data.lineStyle.trim().toLowerCase() : '');
  if (!ROUTE_LINE_STYLES[lineStyle]) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track lineStyle is invalid.', false);
  }
  let orderIndex = null;
  if (data.orderIndex !== undefined && data.orderIndex !== null && data.orderIndex !== '') {
    if (!Number.isInteger(data.orderIndex) || data.orderIndex < 0) {
      throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track orderIndex is invalid.', false);
    }
    orderIndex = data.orderIndex;
  }
  const segments = normalizeTrackSegments_(data.segments);
  const summary = computeTrackSummary_(segments);
  return {
    id: trackId,
    trackId: trackId,
    revisionId: revisionId,
    name: name,
    description: description,
    color: color,
    sourceType: sourceType,
    sourceName: sourceName,
    segments: segments,
    segmentCount: summary.segmentCount,
    pointCount: summary.pointCount,
    distanceMeters: summary.distanceMeters,
    minElevation: summary.minElevation,
    maxElevation: summary.maxElevation,
    startTime: summary.startTime,
    endTime: summary.endTime,
    bounds: summary.bounds,
    orderIndex: orderIndex,
    visible: visible,
    lineStyle: lineStyle,
    lineWidth: 4
  };
}

function compactTrackPoint_(point) {
  return [point.lat, point.lng, point.elevation == null ? null : point.elevation, point.time || ''];
}

function hashTrackPayload_(normalized) {
  return sha256Hex_(JSON.stringify({
    kind: 'track',
    trackId: normalized.trackId,
    revisionId: normalized.revisionId,
    name: normalized.name,
    description: normalized.description,
    color: normalized.color,
    sourceType: normalized.sourceType,
    sourceName: normalized.sourceName,
    segments: normalized.segments.map(function(segment) {
      return segment.points.map(compactTrackPoint_);
    }),
    visible: normalized.visible,
    lineStyle: normalized.lineStyle,
    lineWidth: normalized.lineWidth
  }));
}

function getTrackStagePropertyKey_(trackId) {
  return TRACK_STAGE_PROPERTY_PREFIX_ + sha256Hex_('track-stage:' + trackId);
}

function getTrackRetiredRevisionPropertyPrefix_(trackId) {
  return TRACK_RETIRED_REVISION_PROPERTY_PREFIX_ + sha256Hex_('track:' + trackId) + '_';
}

function getTrackRetiredRevisionPropertyKey_(trackId, revisionId) {
  return getTrackRetiredRevisionPropertyPrefix_(trackId)
    + sha256Hex_('revision:' + revisionId);
}

function readTrackRetiredRevisionHash_(trackId, revisionId) {
  const value = PropertiesService.getScriptProperties().getProperty(
    getTrackRetiredRevisionPropertyKey_(trackId, revisionId)
  );
  if (!value) return null;
  if (!/^[0-9a-f]{64}$/.test(value)) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'retired track revision is invalid.', false);
  }
  return value;
}

function writeTrackRetiredRevisionHash_(trackId, revisionId, payloadHash) {
  PropertiesService.getScriptProperties().setProperty(
    getTrackRetiredRevisionPropertyKey_(trackId, revisionId),
    payloadHash
  );
}

function clearTrackRetiredRevisionHashes_(trackId) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const prefix = getTrackRetiredRevisionPropertyPrefix_(trackId);
  Object.keys(scriptProperties.getProperties()).forEach(function(key) {
    if (key.indexOf(prefix) === 0) scriptProperties.deleteProperty(key);
  });
}

function readTrackStageJournal_(trackId) {
  const raw = PropertiesService.getScriptProperties().getProperty(getTrackStagePropertyKey_(trackId));
  if (!raw) return null;
  let journal;
  try {
    journal = JSON.parse(raw);
  } catch (_error) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track stage journal is invalid.', false);
  }
  if (!journal || typeof journal !== 'object' || Array.isArray(journal)
      || typeof journal.revisionId !== 'string'
      || typeof journal.payloadHash !== 'string'
      || !/^[0-9a-f]{64}$/.test(journal.payloadHash)) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track stage journal is invalid.', false);
  }
  try {
    journal.revisionId = normalizeTrackIdentifier_(journal.revisionId, 'revisionId', TRACK_REVISION_ID_MAX_LENGTH);
  } catch (_error) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track stage journal is invalid.', false);
  }
  return journal;
}

function writeTrackStageJournal_(trackId, revisionId, payloadHash) {
  PropertiesService.getScriptProperties().setProperty(
    getTrackStagePropertyKey_(trackId),
    JSON.stringify({ revisionId: revisionId, payloadHash: payloadHash })
  );
}

function clearTrackStageJournal_(trackId) {
  PropertiesService.getScriptProperties().deleteProperty(getTrackStagePropertyKey_(trackId));
}

function chunkTrackSegments_(normalized) {
  const chunks = [];
  normalized.segments.forEach(function(segment, segmentIndex) {
    let current = [];
    let chunkIndex = 0;
    function pushCurrent() {
      if (!current.length) return;
      const pointsJson = JSON.stringify(current);
      if (pointsJson.length > TRACK_POINTS_JSON_MAX_LENGTH) {
        throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track chunk is too large.', false);
      }
      chunks.push({
        trackId: normalized.trackId,
        revisionId: normalized.revisionId,
        segmentIndex: segmentIndex,
        chunkIndex: chunkIndex,
        pointsJson: pointsJson,
        pointCount: current.length
      });
      chunkIndex += 1;
      current = [];
    }
    segment.points.forEach(function(point) {
      const compact = compactTrackPoint_(point);
      const singleJson = JSON.stringify([compact]);
      if (singleJson.length > TRACK_POINTS_JSON_MAX_LENGTH) {
        throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track point is too large.', false);
      }
      const candidate = current.concat([compact]);
      if (current.length >= TRACK_POINTS_PER_CHUNK
          || JSON.stringify(candidate).length > TRACK_POINTS_JSON_MAX_LENGTH) {
        pushCurrent();
      }
      current.push(compact);
    });
    pushCurrent();
  });
  return chunks;
}

function validateTrackSheetHeaders_(sheet, headers) {
  if (!sheet || sheet.getLastRow() < 1) {
    throw trackStorageError_('TRACK_SHEETS_MISSING', 'track sheet is missing.', false);
  }
  const values = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (!headers.every(function(header, index) { return values[index] === header; })) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track headers are invalid.', false);
  }
}

function openValidatedTrackSheets_() {
  const spreadsheet = openDataSpreadsheet_();
  const tracksSheet = spreadsheet.getSheetByName(TRACKS_SHEET_NAME);
  const segmentsSheet = spreadsheet.getSheetByName(TRACK_SEGMENTS_SHEET_NAME);
  if (!tracksSheet || !segmentsSheet) {
    throw trackStorageError_('TRACK_SHEETS_MISSING', 'track sheets are missing.', false);
  }
  validateTrackSheetHeaders_(tracksSheet, TRACKS_HEADERS);
  validateTrackSheetHeaders_(segmentsSheet, TRACK_SEGMENTS_HEADERS);
  return { tracksSheet: tracksSheet, segmentsSheet: segmentsSheet };
}

function trackMetadataToRow_(track, payloadHash) {
  return [
    track.trackId,
    encodeSpreadsheetLiteral_(track.name),
    encodeSpreadsheetLiteral_(track.description),
    track.color,
    track.sourceType,
    encodeSpreadsheetLiteral_(track.sourceName),
    track.revisionId,
    payloadHash,
    track.segmentCount,
    track.pointCount,
    track.distanceMeters,
    track.minElevation == null ? '' : track.minElevation,
    track.maxElevation == null ? '' : track.maxElevation,
    track.startTime,
    track.endTime,
    JSON.stringify(track.bounds),
    track.createdAt,
    track.updatedAt,
    track.orderIndex,
    track.visible,
    track.lineStyle,
    track.lineWidth
  ];
}

function safeTrackWarningId_(value) {
  const raw = String(value == null ? '' : value);
  const trimmed = raw.trim();
  if (!trimmed || trimmed.length > TRACK_ID_MAX_LENGTH || /[\u0000-\u001f\u007f]/.test(trimmed)
      || /^[=+\-@]/.test(trimmed)) return '';
  return trimmed;
}

function trackMetadataFromRow_(row, rowNumber) {
  const trackId = normalizeTrackIdentifier_(String(row[0] == null ? '' : row[0]), 'trackId', TRACK_ID_MAX_LENGTH);
  const revisionId = normalizeTrackIdentifier_(String(row[6] == null ? '' : row[6]), 'revisionId', TRACK_REVISION_ID_MAX_LENGTH);
  const name = decodeSpreadsheetLiteral_(row[1]);
  const description = decodeSpreadsheetLiteral_(row[2]);
  const color = String(row[3] || '').toLowerCase();
  const sourceType = String(row[4] || '');
  const sourceName = decodeSpreadsheetLiteral_(row[5]);
  const segmentCount = Number(row[8]);
  const pointCount = Number(row[9]);
  const distanceMeters = Number(row[10]);
  const orderIndex = Number(row[18]);
  const lineWidth = Number(row[21]);
  if (!Number.isInteger(segmentCount) || segmentCount < 0 || segmentCount > MAX_TRACK_SEGMENTS
      || !Number.isInteger(pointCount) || pointCount < 0 || pointCount > MAX_TRACK_POINTS
      || !Number.isFinite(distanceMeters) || distanceMeters < 0
      || !Number.isInteger(orderIndex) || orderIndex < 0
      || !Number.isInteger(lineWidth) || lineWidth < 1 || lineWidth > 10
      || !/^[0-9a-f]{64}$/.test(String(row[7] || ''))
      || !ROUTE_LINE_STYLES[String(row[20] || '')]
      || !name || name !== name.trim() || name.length > TRACK_NAME_MAX_LENGTH
      || description.length > TRACK_DESCRIPTION_MAX_LENGTH
      || sourceName !== sourceName.trim() || isInvalidTrackSourceName_(sourceName)
      || PinData.COLOR_OPTIONS.indexOf(color) === -1
      || ['gpx', 'geojson', 'manual'].indexOf(sourceType) === -1
      || typeof row[19] !== 'boolean') {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track metadata is invalid.', false);
  }
  let bounds = null;
  try {
    bounds = JSON.parse(String(row[15] || 'null'));
  } catch (_error) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track bounds are invalid.', false);
  }
  if (pointCount > 0 && (!bounds || ['south', 'west', 'north', 'east'].some(function(key) {
    return typeof bounds[key] !== 'number' || !Number.isFinite(bounds[key]);
  }))) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track bounds are invalid.', false);
  }
  const minElevation = row[11] === '' || row[11] == null ? null : Number(row[11]);
  const maxElevation = row[12] === '' || row[12] == null ? null : Number(row[12]);
  if ((minElevation != null && !Number.isFinite(minElevation))
      || (maxElevation != null && !Number.isFinite(maxElevation))) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track elevation summary is invalid.', false);
  }
  const lineStyle = String(row[20] || '');
  return {
    rowNumber: rowNumber,
    id: trackId,
    trackId: trackId,
    revisionId: revisionId,
    name: name,
    description: description,
    color: color,
    sourceType: sourceType,
    sourceName: sourceName,
    payloadHash: String(row[7] || ''),
    segmentCount: segmentCount,
    pointCount: pointCount,
    distanceMeters: distanceMeters,
    minElevation: minElevation,
    maxElevation: maxElevation,
    startTime: String(row[13] || ''),
    endTime: String(row[14] || ''),
    bounds: bounds,
    createdAt: String(row[16] || ''),
    updatedAt: String(row[17] || ''),
    orderIndex: orderIndex,
    visible: row[19],
    lineStyle: lineStyle,
    lineWidth: 4
  };
}

function readTrackMetadataEntries_(sheet) {
  const rows = readFixedWidthDataRows_(sheet, TRACKS_HEADERS.length);
  return rows.map(function(row, index) {
    return { row: row, rowNumber: index + 2 };
  }).filter(function(entry) { return String(entry.row[0] || '').trim() !== ''; });
}

function maxTrackOrderIndex_(entries) {
  let max = -1;
  entries.forEach(function(entry) {
    const value = Number(entry.row[18]);
    if (Number.isInteger(value) && value >= 0) max = Math.max(max, value);
  });
  return max;
}

function writeTrackSegmentsRevision_(sheet, normalized, chunks, now) {
  const snapshot = readFormulaPreservingDataRows_(sheet, TRACK_SEGMENTS_HEADERS.length);
  const kept = snapshot.outputRows.filter(function(_outputRow, index) {
    const row = snapshot.values[index];
    return !(String(row[0] || '').trim() === normalized.trackId
      && String(row[1] || '').trim() === normalized.revisionId);
  });
  const replacement = chunks.map(function(chunk) {
    return [
      normalized.trackId, normalized.revisionId, chunk.segmentIndex, chunk.chunkIndex,
      chunk.pointsJson, chunk.pointCount, now, now
    ];
  });
  rewriteFixedWidthDataRows_(sheet, TRACK_SEGMENTS_HEADERS.length, snapshot.values.length, kept.concat(replacement));
}

function upsertTrackMetadata_(sheet, savedTrack, payloadHash) {
  const snapshot = readFormulaPreservingDataRows_(sheet, TRACKS_HEADERS.length);
  const kept = [];
  let existingOutput = null;
  snapshot.values.forEach(function(row, index) {
    if (String(row[0] || '').trim() === savedTrack.trackId) {
      if (!existingOutput) existingOutput = snapshot.outputRows[index];
      return;
    }
    kept.push(snapshot.outputRows[index]);
  });
  let replacement = trackMetadataToRow_(savedTrack, payloadHash);
  if (existingOutput && existingOutput.length > TRACKS_HEADERS.length) {
    replacement = replacement.concat(existingOutput.slice(TRACKS_HEADERS.length));
  }
  rewriteFixedWidthDataRows_(sheet, TRACKS_HEADERS.length, snapshot.values.length, kept.concat([replacement]));
}

function cleanupTrackSegmentRevisions_(sheet, trackId, activeRevision) {
  const result = filterAndRewriteFixedWidthDataRows_(
    sheet,
    TRACK_SEGMENTS_HEADERS.length,
    function(row) {
      return String(row[0] || '').trim() === trackId
        && String(row[1] || '').trim() !== activeRevision;
    }
  );
  return result.removedRows.length;
}

function savedTrackFromNormalized_(normalized, metadata) {
  return {
    id: normalized.trackId,
    trackId: normalized.trackId,
    name: normalized.name,
    description: normalized.description,
    color: normalized.color,
    sourceType: normalized.sourceType,
    sourceName: normalized.sourceName,
    revisionId: normalized.revisionId,
    segments: normalized.segments.map(function(segment) {
      return { index: segment.index, points: segment.points.map(function(point) {
        return { lat: point.lat, lng: point.lng, elevation: point.elevation, time: point.time };
      }) };
    }),
    segmentCount: normalized.segmentCount,
    pointCount: normalized.pointCount,
    distanceMeters: normalized.distanceMeters,
    minElevation: normalized.minElevation,
    maxElevation: normalized.maxElevation,
    startTime: normalized.startTime,
    endTime: normalized.endTime,
    bounds: normalized.bounds,
    createdAt: metadata.createdAt,
    updatedAt: metadata.updatedAt,
    orderIndex: metadata.orderIndex,
    visible: normalized.visible,
    lineStyle: normalized.lineStyle,
    lineWidth: normalized.lineWidth
  };
}

function withTrackStorageLock_(callback) {
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    acquired = lock.tryLock(SPREADSHEET_MUTATION_LOCK_TIMEOUT_MS);
    if (!acquired) {
      throw trackStorageError_('TRACK_STORAGE_BUSY', 'track storage is busy.', true);
    }
    return callback();
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function saveTrackBundle(data) {
  assertEditToken_(data);
  try {
    const normalized = normalizeTrackBundle_(data);
    const payloadHash = hashTrackPayload_(normalized);
    const chunks = chunkTrackSegments_(normalized);
    return withTrackStorageLock_(function() {
      const sheets = openValidatedTrackSheets_();
      const entries = readTrackMetadataEntries_(sheets.tracksSheet);
      const matches = entries.filter(function(entry) {
        return String(entry.row[0] || '').trim() === normalized.trackId;
      });
      if (matches.length > 1) {
        throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'duplicate track metadata.', false);
      }
      const existing = matches.length ? trackMetadataFromRow_(matches[0].row, matches[0].rowNumber) : null;
      if (existing && existing.revisionId === normalized.revisionId) {
        if (existing.payloadHash !== payloadHash) {
          throw trackStorageError_('TRACK_REVISION_PAYLOAD_CONFLICT', 'track revision conflict.', false);
        }
        const removed = cleanupTrackSegmentRevisions_(sheets.segmentsSheet, normalized.trackId, normalized.revisionId);
        if (removed > 0) SpreadsheetApp.flush();
        if (readTrackStageJournal_(normalized.trackId)) clearTrackStageJournal_(normalized.trackId);
        return {
          ok: true,
          deduplicated: true,
          track: savedTrackFromNormalized_(normalized, existing)
        };
      }
      if (!existing) {
        const uniqueIds = Object.create(null);
        entries.forEach(function(entry) {
          const id = safeTrackWarningId_(entry.row[0]);
          if (id) uniqueIds[id] = true;
        });
        if (Object.keys(uniqueIds).length >= MAX_TRACKS) {
          throw trackStorageError_('TRACK_LIMIT_EXCEEDED', 'track limit exceeded.', false);
        }
      }
      const staged = readTrackStageJournal_(normalized.trackId);
      if (staged && staged.revisionId === normalized.revisionId && staged.payloadHash !== payloadHash) {
        throw trackStorageError_('TRACK_REVISION_PAYLOAD_CONFLICT', 'staged track revision conflict.', false);
      }
      if (readTrackRetiredRevisionHash_(normalized.trackId, normalized.revisionId)) {
        throw trackStorageError_('TRACK_REVISION_PAYLOAD_CONFLICT', 'retired track revision cannot be reused.', false);
      }
      if (existing) {
        writeTrackRetiredRevisionHash_(
          normalized.trackId, existing.revisionId, existing.payloadHash
        );
      }
      const now = currentUpdatedAt_();
      const orderIndex = normalized.orderIndex == null
        ? (existing ? existing.orderIndex : maxTrackOrderIndex_(entries) + 1)
        : normalized.orderIndex;
      const metadata = {
        createdAt: existing ? existing.createdAt : now,
        updatedAt: now,
        orderIndex: orderIndex
      };
      const savedTrack = savedTrackFromNormalized_(normalized, metadata);
      writeTrackStageJournal_(normalized.trackId, normalized.revisionId, payloadHash);
      writeTrackSegmentsRevision_(sheets.segmentsSheet, normalized, chunks, now);
      SpreadsheetApp.flush();
      upsertTrackMetadata_(sheets.tracksSheet, savedTrack, payloadHash);
      SpreadsheetApp.flush();
      const removed = cleanupTrackSegmentRevisions_(sheets.segmentsSheet, normalized.trackId, normalized.revisionId);
      if (removed > 0) SpreadsheetApp.flush();
      clearTrackStageJournal_(normalized.trackId);
      return { ok: true, deduplicated: false, track: savedTrack };
    });
  } catch (error) {
    return trackStorageFailureFromError_(error);
  }
}

function compactStoredTrackPoint_(value) {
  if (!Array.isArray(value) || value.length !== 4) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'stored track point is invalid.', false);
  }
  return normalizeTrackPoint_({ lat: value[0], lng: value[1], elevation: value[2], time: value[3] });
}

function reconstructTrackSegments_(metadata, segmentRows) {
  let reconstructedPointCount = 0;
  const relevant = segmentRows.filter(function(row) {
    return String(row[0] || '').trim() === metadata.trackId
      && String(row[1] || '').trim() === metadata.revisionId;
  }).map(function(row) {
    const segmentIndex = Number(row[2]);
    const chunkIndex = Number(row[3]);
    const pointCount = Number(row[5]);
    const pointsJson = String(row[4] || '');
    if (!Number.isInteger(segmentIndex) || segmentIndex < 0 || segmentIndex >= metadata.segmentCount
        || !Number.isInteger(chunkIndex) || chunkIndex < 0
        || !Number.isInteger(pointCount) || pointCount < 1 || pointCount > TRACK_POINTS_PER_CHUNK
        || pointsJson.length > TRACK_POINTS_JSON_MAX_LENGTH) {
      throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'stored track chunk is invalid.', false);
    }
    let compactPoints;
    try {
      compactPoints = JSON.parse(pointsJson);
    } catch (_error) {
      throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'stored track JSON is invalid.', false);
    }
    if (!Array.isArray(compactPoints) || compactPoints.length !== pointCount) {
      throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'stored track count is invalid.', false);
    }
    reconstructedPointCount += pointCount;
    if (reconstructedPointCount > metadata.pointCount || reconstructedPointCount > MAX_TRACK_POINTS) {
      throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'stored track point limit is invalid.', false);
    }
    return {
      segmentIndex: segmentIndex,
      chunkIndex: chunkIndex,
      points: compactPoints.map(compactStoredTrackPoint_)
    };
  }).sort(function(a, b) {
    return a.segmentIndex - b.segmentIndex || a.chunkIndex - b.chunkIndex;
  });
  if (metadata.segmentCount === 0) {
    if (relevant.length !== 0) throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'unexpected track chunks.', false);
    return [];
  }
  const segments = [];
  relevant.forEach(function(chunk) {
    while (segments.length <= chunk.segmentIndex) segments.push(null);
    if (!segments[chunk.segmentIndex]) {
      if (chunk.chunkIndex !== 0) throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track chunk gap.', false);
      segments[chunk.segmentIndex] = { index: chunk.segmentIndex, points: [], nextChunkIndex: 0 };
    }
    const segment = segments[chunk.segmentIndex];
    if (chunk.chunkIndex !== segment.nextChunkIndex) {
      throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track chunk order is invalid.', false);
    }
    segment.points = segment.points.concat(chunk.points);
    segment.nextChunkIndex += 1;
  });
  if (segments.length !== metadata.segmentCount || segments.some(function(segment, index) {
    return !segment || segment.index !== index || segment.points.length === 0;
  })) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track segment count is invalid.', false);
  }
  return segments.map(function(segment) { return { index: segment.index, points: segment.points }; });
}

function trackFromStoredRows_(metadata, segmentRows, startupTimings) {
  const restoreStartedAt = startupTimingNow_();
  let segments;
  try {
    segments = reconstructTrackSegments_(metadata, segmentRows);
  } catch (caught) {
    if (startupTimings) startupTimings.jsonRestoreFailed = true;
    throw caught;
  } finally {
    if (startupTimings) {
      startupTimings.jsonRestoreDurationMs += startupTimingNow_() - restoreStartedAt;
    }
  }
  const summaryStartedAt = startupTimingNow_();
  let summary;
  try {
    summary = computeTrackSummary_(segments);
  } catch (caught) {
    if (startupTimings) startupTimings.summaryCalculateFailed = true;
    throw caught;
  } finally {
    if (startupTimings) {
      startupTimings.summaryCalculateDurationMs += startupTimingNow_() - summaryStartedAt;
    }
  }
  if (summary.segmentCount !== metadata.segmentCount || summary.pointCount !== metadata.pointCount) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track summary count is invalid.', false);
  }
  if (PinData.COLOR_OPTIONS.indexOf(metadata.color) === -1
      || ['gpx', 'geojson', 'manual'].indexOf(metadata.sourceType) === -1) {
    throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track metadata value is invalid.', false);
  }
  return {
    id: metadata.trackId,
    trackId: metadata.trackId,
    name: metadata.name,
    description: metadata.description,
    color: metadata.color,
    sourceType: metadata.sourceType,
    sourceName: metadata.sourceName,
    revisionId: metadata.revisionId,
    segments: segments,
    segmentCount: summary.segmentCount,
    pointCount: summary.pointCount,
    distanceMeters: summary.distanceMeters,
    minElevation: summary.minElevation,
    maxElevation: summary.maxElevation,
    startTime: summary.startTime,
    endTime: summary.endTime,
    bounds: summary.bounds,
    createdAt: metadata.createdAt,
    updatedAt: metadata.updatedAt,
    orderIndex: metadata.orderIndex,
    visible: metadata.visible,
    lineStyle: metadata.lineStyle,
    lineWidth: metadata.lineWidth
  };
}

function getTracks() {
  const totalStartedAt = startupTimingNow_();
  const startupTimings = {
    jsonRestoreDurationMs: 0,
    jsonRestoreFailed: false,
    summaryCalculateDurationMs: 0,
    summaryCalculateFailed: false
  };
  try {
    const sheets = measureStartupGasStage_('getTracks', 'sheets-read', function() {
      return openValidatedTrackSheets_();
    });
    const metadataEntries = measureStartupGasStage_('getTracks', 'tracks-read', function() {
      return readTrackMetadataEntries_(sheets.tracksSheet);
    });
    const segmentRows = measureStartupGasStage_('getTracks', 'track-segments-read', function() {
      return readFixedWidthDataRows_(sheets.segmentsSheet, TRACK_SEGMENTS_HEADERS.length);
    });
    const tracks = [];
    const warnings = [];
    const rawGroups = Object.create(null);
    const rawOrder = [];
    metadataEntries.forEach(function(entry) {
      const warningTrackId = safeTrackWarningId_(entry.row[0]);
      if (!warningTrackId) {
        warnings.push({ code: 'TRACK_METADATA_CORRUPTED', trackId: warningTrackId });
        return;
      }
      if (!rawGroups[warningTrackId]) {
        rawGroups[warningTrackId] = [];
        rawOrder.push(warningTrackId);
      }
      rawGroups[warningTrackId].push(entry);
    });
    rawOrder.forEach(function(trackId) {
      const matches = rawGroups[trackId];
      if (matches.length !== 1) {
        warnings.push({ code: 'TRACK_METADATA_CORRUPTED', trackId: trackId });
        return;
      }
      let metadata;
      try {
        metadata = trackMetadataFromRow_(matches[0].row, matches[0].rowNumber);
      } catch (_error) {
        warnings.push({ code: 'TRACK_METADATA_CORRUPTED', trackId: trackId });
        return;
      }
      try {
        tracks.push(trackFromStoredRows_(metadata, segmentRows, startupTimings));
      } catch (_error) {
        warnings.push({ code: 'TRACK_SEGMENTS_CORRUPTED', trackId: trackId });
      }
    });
    logStartupGasStage_(
      'getTracks', 'json-restore', startupTimings.jsonRestoreFailed ? 'failure' : 'success',
      startupTimingNow_() - startupTimings.jsonRestoreDurationMs
    );
    logStartupGasStage_(
      'getTracks', 'summary-calculate', startupTimings.summaryCalculateFailed ? 'failure' : 'success',
      startupTimingNow_() - startupTimings.summaryCalculateDurationMs
    );
    const response = measureStartupGasStage_('getTracks', 'response-build', function() {
      tracks.sort(function(a, b) {
        const orderDifference = a.orderIndex - b.orderIndex;
        if (orderDifference) return orderDifference;
        return a.trackId < b.trackId ? -1 : (a.trackId > b.trackId ? 1 : 0);
      });
      return { ok: true, tracks: tracks, warnings: warnings };
    });
    logStartupGasStage_('getTracks', 'total', 'success', totalStartedAt);
    return response;
  } catch (error) {
    logStartupGasStage_('getTracks', 'total', 'failure', totalStartedAt);
    return trackStorageFailureFromError_(error);
  }
}

function normalizeTrackDisplaySettings_(data) {
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track settings payload is invalid.', false);
  }
  const trackId = normalizeTrackIdentifier_(data.trackId, 'trackId', TRACK_ID_MAX_LENGTH);
  const name = typeof data.name === 'string' ? data.name.trim() : '';
  if (!name || name.length > TRACK_NAME_MAX_LENGTH) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track name is invalid.', false);
  }
  if (data.description != null && typeof data.description !== 'string') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track description is invalid.', false);
  }
  const description = data.description == null ? '' : data.description;
  if (description.length > TRACK_DESCRIPTION_MAX_LENGTH) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track description is invalid.', false);
  }
  if (typeof data.color !== 'string'
      || PinData.COLOR_OPTIONS.indexOf(data.color.trim().toLowerCase()) === -1) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track color is invalid.', false);
  }
  if (typeof data.visible !== 'boolean') {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track visible is invalid.', false);
  }
  const lineStyle = typeof data.lineStyle === 'string'
    ? data.lineStyle.trim().toLowerCase() : '';
  if (!ROUTE_LINE_STYLES[lineStyle]) {
    throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'track lineStyle is invalid.', false);
  }
  return {
    trackId: trackId,
    name: name,
    description: description,
    color: data.color.trim().toLowerCase(),
    visible: data.visible,
    lineStyle: lineStyle,
    lineWidth: 4
  };
}

function updateTrackDisplaySettings(data) {
  assertEditToken_(data);
  try {
    const settings = normalizeTrackDisplaySettings_(data);
    return withTrackStorageLock_(function() {
      const sheets = openValidatedTrackSheets_();
      const matches = readTrackMetadataEntries_(sheets.tracksSheet).filter(function(entry) {
        return String(entry.row[0] || '').trim() === settings.trackId;
      });
      if (matches.length > 1) {
        throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'duplicate track metadata.', false);
      }
      if (matches.length === 0) {
        throw trackStorageError_('TRACK_NOT_FOUND', 'track was not found.', false);
      }
      const entry = matches[0];
      trackMetadataFromRow_(entry.row, entry.rowNumber);
      const lastColumn = Math.max(TRACKS_HEADERS.length, sheets.tracksSheet.getLastColumn());
      const range = sheets.tracksSheet.getRange(entry.rowNumber, 1, 1, lastColumn);
      const values = range.getValues()[0];
      const formulas = range.getFormulas()[0];
      const output = values.map(function(value, index) {
        return (formulas && formulas[index]) || value;
      });
      const updatedAt = currentUpdatedAt_();
      output[1] = encodeSpreadsheetLiteral_(settings.name);
      output[2] = encodeSpreadsheetLiteral_(settings.description);
      output[3] = settings.color;
      output[17] = updatedAt;
      output[19] = settings.visible;
      output[20] = settings.lineStyle;
      output[21] = 4;
      try {
        range.setValues([output]);
        SpreadsheetApp.flush();
      } catch (_writeError) {
        throw trackStorageError_(
          'TRACK_METADATA_UPDATE_FAILED', 'track metadata update failed.', true
        );
      }
      return {
        ok: true,
        track: {
          trackId: settings.trackId,
          name: settings.name,
          description: settings.description,
          color: settings.color,
          visible: settings.visible,
          lineStyle: settings.lineStyle,
          lineWidth: 4,
          updatedAt: updatedAt
        }
      };
    });
  } catch (error) {
    return trackStorageFailureFromError_(error);
  }
}

function updateTracksOrder(data) {
  assertEditToken_(data);
  try {
    if (!data || !Array.isArray(data.orderedIds)) {
      throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'orderedIds is required.', false);
    }
    const seen = Object.create(null);
    const orderedIds = data.orderedIds.map(function(value) {
      const trackId = normalizeTrackIdentifier_(value, 'trackId', TRACK_ID_MAX_LENGTH);
      if (seen[trackId]) {
        throw trackStorageError_('INVALID_TRACK_PAYLOAD', 'trackId is duplicated.', false);
      }
      seen[trackId] = true;
      return trackId;
    });

    return withTrackStorageLock_(function() {
      const sheets = openValidatedTrackSheets_();
      const metadataEntries = readTrackMetadataEntries_(sheets.tracksSheet);
      const entries = [];
      const byId = Object.create(null);
      metadataEntries.forEach(function(entry) {
        let metadata;
        try {
          metadata = trackMetadataFromRow_(entry.row, entry.rowNumber);
        } catch (_error) {
          throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track metadata is invalid.', false);
        }
        if (byId[metadata.trackId]) {
          throw trackStorageError_('TRACK_STORAGE_CORRUPTED', 'track metadata is duplicated.', false);
        }
        const orderedEntry = {
          trackId: metadata.trackId,
          orderIndex: metadata.orderIndex,
          rowNumber: entry.rowNumber
        };
        entries.push(orderedEntry);
        byId[metadata.trackId] = orderedEntry;
      });
      orderedIds.forEach(function(trackId) {
        if (!byId[trackId]) {
          throw trackStorageError_('TRACK_NOT_FOUND', 'track was not found.', false);
        }
      });
      if (entries.length === 0) return { ok: true, tracks: [] };

      entries.sort(function(left, right) {
        const orderDifference = left.orderIndex - right.orderIndex;
        if (orderDifference) return orderDifference;
        return left.trackId < right.trackId ? -1 : (left.trackId > right.trackId ? 1 : 0);
      });
      const nextEntries = orderedIds.map(function(trackId) { return byId[trackId]; });
      entries.forEach(function(entry) {
        if (!seen[entry.trackId]) nextEntries.push(entry);
      });

      const dataRowCount = sheets.tracksSheet.getLastRow() - 1;
      const orderRange = sheets.tracksSheet.getRange(2, 19, dataRowCount, 1);
      const orderValues = orderRange.getValues();
      const orderFormulas = orderRange.getFormulas();
      const output = orderValues.map(function(row, index) {
        return [(orderFormulas[index] && orderFormulas[index][0]) || row[0]];
      });
      nextEntries.forEach(function(entry, index) {
        output[entry.rowNumber - 2][0] = index;
      });
      orderRange.setValues(output);

      const result = getTracks();
      if (!result || result.ok !== true || !Array.isArray(result.tracks)) {
        throw trackStorageError_('TRACK_STORAGE_FAILED', 'ordered tracks could not be read.', true);
      }
      return { ok: true, tracks: result.tracks };
    });
  } catch (error) {
    return trackStorageFailureFromError_(error);
  }
}

function removeTrackShareRouteTargets_(trackId) {
  const spreadsheet = openDataSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHARE_LINKS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 1) {
    return { removedReferences: 0, updatedRows: 0 };
  }
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return { removedReferences: 0, updatedRows: 0 };
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const targetColumns = [];
  headers.forEach(function(header, index) {
    if (String(header || '').trim() === SHARE_ROUTE_TARGETS_HEADER) targetColumns.push(index + 1);
  });
  if (targetColumns.length === 0) return { removedReferences: 0, updatedRows: 0 };
  if (targetColumns.length !== 1) {
    throw trackStorageError_(
      'TRACK_SHARE_REFERENCES_DELETE_FAILED',
      'duplicate routeTargetsJson headers.',
      true
    );
  }
  const dataRowCount = Math.max(0, sheet.getLastRow() - 1);
  if (dataRowCount === 0) return { removedReferences: 0, updatedRows: 0 };

  const range = sheet.getRange(2, targetColumns[0], dataRowCount, 1);
  const values = range.getValues();
  const formulas = typeof range.getFormulas === 'function' ? range.getFormulas() : [];
  const output = values.map(function(row, index) {
    return [(formulas[index] && formulas[index][0]) || row[0]];
  });
  let removedReferences = 0;
  let updatedRows = 0;
  values.forEach(function(row, index) {
    const raw = String(row[0] == null ? '' : row[0]).trim();
    if (!raw || parseShareRouteTargetsJson_(raw).state !== 'valid') return;
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch (_error) {
      return;
    }
    let rowRemoved = 0;
    const keptTargets = parsed.targets.filter(function(target) {
      const type = String(target && target.type || '').trim();
      const id = normalizeShareRouteTargetId_(target && target.id);
      const matches = id === trackId && (type === 'gpx-route' || type === 'geojson-route');
      if (matches) rowRemoved += 1;
      return !matches;
    });
    if (rowRemoved === 0) return;
    parsed.targets = keptTargets;
    output[index][0] = JSON.stringify(parsed);
    removedReferences += rowRemoved;
    updatedRows += 1;
  });
  if (updatedRows > 0) range.setValues(output);
  return { removedReferences: removedReferences, updatedRows: updatedRows };
}

function deleteTrack(data) {
  assertEditToken_(data);
  try {
    const trackId = normalizeTrackIdentifier_(data && data.trackId, 'trackId', TRACK_ID_MAX_LENGTH);
    return withTrackStorageLock_(function() {
      const sheets = openValidatedTrackSheets_();
      let metadataResult;
      try {
        metadataResult = filterAndRewriteFixedWidthDataRows_(
          sheets.tracksSheet,
          TRACKS_HEADERS.length,
          function(row) { return String(row[0] || '').trim() === trackId; }
        );
        if (metadataResult.removedRows.length > 0) SpreadsheetApp.flush();
      } catch (_metadataError) {
        throw trackStorageError_(
          'TRACK_METADATA_DELETE_FAILED', 'track metadata deletion failed.', true
        );
      }
      let segmentResult;
      try {
        segmentResult = filterAndRewriteFixedWidthDataRows_(
          sheets.segmentsSheet,
          TRACK_SEGMENTS_HEADERS.length,
          function(row) { return String(row[0] || '').trim() === trackId; }
        );
        if (segmentResult.removedRows.length > 0) SpreadsheetApp.flush();
      } catch (_segmentError) {
        throw trackStorageError_(
          'TRACK_SEGMENTS_DELETE_FAILED', 'track segment deletion failed.', true
        );
      }
      if (readTrackStageJournal_(trackId)) clearTrackStageJournal_(trackId);
      clearTrackRetiredRevisionHashes_(trackId);
      let shareCleanup;
      try {
        shareCleanup = removeTrackShareRouteTargets_(trackId);
        if (shareCleanup.updatedRows > 0) SpreadsheetApp.flush();
      } catch (_shareError) {
        const error = trackStorageError_(
          'TRACK_SHARE_REFERENCES_DELETE_FAILED', 'track share cleanup failed.', true
        );
        error.serverDeleted = true;
        throw error;
      }
      return {
        ok: true,
        deleted: metadataResult.removedRows.length > 0 || segmentResult.removedRows.length > 0,
        removedShareReferences: shareCleanup.removedReferences
      };
    });
  } catch (error) {
    return trackStorageFailureFromError_(error);
  }
}

function buildSharedViewUrl_(token) {
  return getConfiguredWebAppUrl_() + '?view=shared&token=' + encodeURIComponent(token);
}

function utf8JsonByteLength_(value) {
  const text = String(value == null ? '' : value);
  let byteLength = 0;
  for (var i = 0; i < text.length; i += 1) {
    const code = text.charCodeAt(i);
    if (code <= 0x7f) {
      byteLength += 1;
    } else if (code <= 0x7ff) {
      byteLength += 2;
    } else if (code >= 0xd800 && code <= 0xdbff
        && i + 1 < text.length
        && text.charCodeAt(i + 1) >= 0xdc00 && text.charCodeAt(i + 1) <= 0xdfff) {
      byteLength += 4;
      i += 1;
    } else {
      byteLength += 3;
    }
  }
  return byteLength;
}

function sharedPayloadBudgetFailure_() {
  return { ok: false, errorCode: SHARED_PAYLOAD_TOO_LARGE_CODE };
}

function validateSharedPayloadBudget_(dto, metrics) {
  const counts = metrics || {};
  const pinCount = Number(counts.pinCount);
  const routeCount = Number(counts.routeCount);
  const coordinateCount = Number(counts.coordinateCount);
  if (!Number.isSafeInteger(pinCount) || pinCount < 0 || pinCount > SHARED_PAYLOAD_MAX_PINS
      || !Number.isSafeInteger(routeCount) || routeCount < 0 || routeCount > SHARED_PAYLOAD_MAX_ROUTES
      || !Number.isSafeInteger(coordinateCount) || coordinateCount < 0
      || coordinateCount > SHARED_PAYLOAD_MAX_COORDINATE_POINTS) {
    return sharedPayloadBudgetFailure_();
  }
  if (dto == null) return { ok: true, jsonBytes: null };

  let json;
  try {
    json = JSON.stringify(dto);
  } catch (_error) {
    return sharedPayloadBudgetFailure_();
  }
  if (typeof json !== 'string') return sharedPayloadBudgetFailure_();
  const jsonBytes = utf8JsonByteLength_(json);
  if (jsonBytes > SHARED_PAYLOAD_MAX_JSON_BYTES) return sharedPayloadBudgetFailure_();
  return { ok: true, jsonBytes: jsonBytes };
}

function sharedPayloadTooLargeReadResult_() {
  return {
    ok: false,
    error: 'shared_payload_too_large',
    errorCode: SHARED_PAYLOAD_TOO_LARGE_CODE
  };
}

function createShareLink(data) {
  assertEditToken_(data);
  const label = normalizeShareLinkLabel_(data && data.label);
  const tags = PinData.normalizeTags(data && data.tags || []);
  const tagMode = String(data && data.tagMode || 'or') === 'and' ? 'and' : 'or';
  const colors = normalizeShareColors_(data && data.colors || []);
  let requestedTargets;
  if (data && Array.isArray(data.routeTargets)) {
    requestedTargets = data.routeTargets;
  } else {
    const legacyRouteIds = normalizeShareRouteIds_(data && data.routeIds || []);
    if (isShareRouteSelectionNone_(legacyRouteIds)) {
      requestedTargets = [];
    } else if (legacyRouteIds.length) {
      requestedTargets = legacyRouteIds.map(function(routeId) { return { type: 'pin-route', id: routeId }; });
    } else {
      requestedTargets = getRouteGroups().map(function(group) {
        return { type: 'pin-route', id: group && (group.routeId || group.id) };
      });
    }
  }
  const routeTargets = normalizeShareRouteTargets_(requestedTargets);
  const selectedPinRouteIds = routeTargets.filter(function(target) {
    return target.type === 'pin-route';
  }).map(function(target) { return target.id; });
  const routeIds = selectedPinRouteIds.length ? selectedPinRouteIds : [SHARE_ROUTE_NONE_SENTINEL];
  const token = Utilities.getUuid();
  const createdAt = new Date().toISOString();
  const shareLink = {
    createdAt: createdAt,
    label: label,
    token: token,
    tags: tags,
    tagMode: tagMode,
    colors: colors,
    routeIds: routeIds,
    routeTargets: routeTargets,
    routeTargetsState: 'valid',
    enabled: true,
    revokedAt: ''
  };
  const sharedViewResult = buildSharedViewDataForLink_(shareLink);
  if (!sharedViewResult.ok && sharedViewResult.errorCode === SHARED_PAYLOAD_TOO_LARGE_CODE) {
    return {
      ok: false,
      errorCode: SHARED_PAYLOAD_TOO_LARGE_CODE,
      error: SHARED_PAYLOAD_CREATE_ERROR
    };
  }

  const spreadsheet = openDataSpreadsheet_();
  const sheet = ensureShareLinksSheet_(spreadsheet);
  const currentRows = sheet.getDataRange().getValues();
  const headerRow = currentRows.length ? currentRows[0] : SHARE_LINKS_HEADERS;
  const headerMap = shareLinksHeaderIndexMap_(headerRow);
  const row = new Array(Math.max(headerRow.length, SHARE_LINKS_HEADERS.length)).fill('');
  row[headerMap.createdAt] = createdAt;
  row[headerMap.label] = label;
  row[headerMap.token] = token;
  row[headerMap.tags] = PinData.serializeTags(tags);
  row[headerMap.tagMode] = tagMode;
  row[headerMap.enabled] = true;
  row[headerMap.revokedAt] = '';
  row[headerMap.colors] = serializeShareColors_(colors);
  row[headerMap.routeIds] = routeIds.join('|');
  row[headerMap[SHARE_ROUTE_TARGETS_HEADER]] = JSON.stringify({ v: 1, targets: routeTargets });
  sheet.appendRow(row);

  return {
    ok: true,
    token: token,
    url: buildSharedViewUrl_(token),
    shareLink: shareLink
  };
}

function listShareLinks(data) {
  assertEditToken_(data);
  const rows = openShareLinksSheet_().getDataRange().getValues();
  const headerMap = shareLinksHeaderIndexMap_(rows[0] || SHARE_LINKS_HEADERS);
  const items = rows.slice(1).map(function(row) { return shareRowToLink_(row, headerMap); }).reverse().map(function(item) {
    item.url = buildSharedViewUrl_(item.token);
    return item;
  });
  return { ok: true, items: items };
}

function getShareLinkByToken_(token) {
  var normalizedToken = normalizeShareToken_(token);
  const rows = openShareLinksSheet_().getDataRange().getValues();
  const headerMap = shareLinksHeaderIndexMap_(rows[0] || SHARE_LINKS_HEADERS);
  for (var i = 1; i < rows.length; i += 1) {
    if (String(shareRowValue_(rows[i], headerMap, 'token')) === normalizedToken) {
      return shareRowToLink_(rows[i], headerMap);
    }
  }
  return null;
}

function resolveShareLinkMutationTarget_(rows, requiredHeaders, normalizedToken) {
  const headerRow = Array.isArray(rows) && Array.isArray(rows[0]) ? rows[0] : [];
  const headerIndexes = {};
  const headerCounts = {};
  headerRow.forEach(function(header, index) {
    const name = String(header || '').trim();
    if (!name) return;
    headerCounts[name] = (headerCounts[name] || 0) + 1;
    if (headerIndexes[name] == null) headerIndexes[name] = index;
  });
  for (var i = 0; i < requiredHeaders.length; i += 1) {
    const requiredHeader = requiredHeaders[i];
    if (headerCounts[requiredHeader] !== 1) {
      return { ok: false, error: 'share_links header is invalid' };
    }
  }

  const matchingRowNumbers = [];
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    if (normalizeShareToken_(rows[rowIndex][headerIndexes.token]) === normalizedToken) {
      matchingRowNumbers.push(rowIndex + 1);
    }
  }
  if (matchingRowNumbers.length === 0) return { ok: false, error: 'token not found' };
  if (matchingRowNumbers.length > 1) return { ok: false, error: 'duplicate token' };
  return { ok: true, rowNumber: matchingRowNumbers[0], headerIndexes: headerIndexes };
}

function setShareLinkEnabled(data) {
  assertEditToken_(data);
  var normalizedToken = normalizeShareToken_(typeof data === 'object' && data !== null ? data.token : data);
  var enabled = !!(data && typeof data === 'object' ? data.enabled : false);
  if (!normalizedToken) return { ok: false, error: 'token not found' };
  return withSpreadsheetMutationLock_(function() {
    const sheet = openShareLinksSheet_();
    const rows = sheet.getDataRange().getValues();
    const target = resolveShareLinkMutationTarget_(rows, ['token', 'enabled', 'revokedAt'], normalizedToken);
    if (!target.ok) return target;
    sheet.getRange(target.rowNumber, target.headerIndexes.enabled + 1).setValue(enabled);
    sheet.getRange(target.rowNumber, target.headerIndexes.revokedAt + 1).setValue(enabled ? '' : new Date().toISOString());
    return { ok: true };
  });
}

function deleteShareLink(data) {
  assertEditToken_(data);
  var normalizedToken = normalizeShareToken_(data && typeof data === 'object' ? data.token : data);
  if (!normalizedToken) return { ok: false, error: 'token not found' };
  return withSpreadsheetMutationLock_(function() {
    const sheet = openShareLinksSheet_();
    const rows = sheet.getDataRange().getValues();
    const target = resolveShareLinkMutationTarget_(rows, ['token'], normalizedToken);
    if (!target.ok) return target;
    sheet.deleteRow(target.rowNumber);
    return { ok: true };
  });
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
  const clientPin = toClientPin_(pin);
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
    tags: filterPinTagsForShare_(pin, allowedTags),
    hasAudio: clientPin.hasAudio
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

function toSharedPinRoute_(group) {
  return {
    type: 'pin-route',
    id: group.routeId || group.id,
    routeId: group.routeId || group.id,
    name: group.name,
    color: group.color,
    visible: group.visible,
    showNumbers: group.showNumbers,
    showLine: group.showLine,
    lineStyle: group.lineStyle,
    routeMode: group.routeMode,
    closed: group.closed,
    startPinId: group.startPinId,
    endPinId: group.endPinId,
    pinIds: Array.isArray(group.pinIds) ? group.pinIds.slice() : []
  };
}

function toSharedTrackRoute_(track, type) {
  return {
    type: type,
    id: track.trackId || track.id,
    name: track.name,
    description: track.description,
    color: track.color,
    visible: track.visible !== false,
    lineStyle: track.lineStyle,
    lineWidth: 4,
    distanceMeters: track.distanceMeters,
    bounds: track.bounds,
    segments: (Array.isArray(track.segments) ? track.segments : []).map(function(segment) {
      return (Array.isArray(segment.points) ? segment.points : []).map(function(point) {
        return { lat: Number(point.lat), lng: Number(point.lng) };
      });
    })
  };
}

function prepareSharedTrackRoutes_(targets) {
  const selectedKeySet = {};
  (Array.isArray(targets) ? targets : []).forEach(function(target) {
    if (target.type === 'gpx-route' || target.type === 'geojson-route') {
      selectedKeySet[shareRouteTargetKey_(target.type, target.id)] = true;
    }
  });
  if (Object.keys(selectedKeySet).length === 0) return { entries: [], segmentsSheet: null };

  try {
    const sheets = openValidatedTrackSheets_();
    const metadataEntries = readTrackMetadataEntries_(sheets.tracksSheet);
    const groupedEntries = Object.create(null);
    metadataEntries.forEach(function(entry) {
      const id = safeTrackWarningId_(entry.row[0]);
      if (!id) return;
      if (!groupedEntries[id]) groupedEntries[id] = [];
      groupedEntries[id].push(entry);
    });
    const orderedMetadata = [];
    Object.keys(groupedEntries).forEach(function(id) {
      const entries = groupedEntries[id];
      if (entries.length !== 1) return;
      try {
        const metadata = trackMetadataFromRow_(entries[0].row, entries[0].rowNumber);
        const type = metadata.sourceType === 'gpx' ? 'gpx-route'
          : (metadata.sourceType === 'geojson' ? 'geojson-route' : '');
        const key = shareRouteTargetKey_(type, metadata.trackId);
        if (type) orderedMetadata.push({ key: key, metadata: metadata, type: type });
      } catch (_metadataError) {
        // Invalid storage never becomes public data.
      }
    });
    orderedMetadata.sort(function(left, right) {
      const orderDifference = left.metadata.orderIndex - right.metadata.orderIndex;
      if (orderDifference) return orderDifference;
      return left.metadata.trackId < right.metadata.trackId
        ? -1 : (left.metadata.trackId > right.metadata.trackId ? 1 : 0);
    });
    const selectedMetadata = orderedMetadata.filter(function(entry) {
      return selectedKeySet[entry.key] === true;
    });
    return { entries: selectedMetadata, segmentsSheet: sheets.segmentsSheet };
  } catch (_trackError) {
    // Existing installations without valid track sheets remain pin-route compatible.
    return { entries: [], segmentsSheet: null };
  }
}

function sharedTrackRevisionKey_(trackId, revisionId) {
  return String(trackId || '') + '\0' + String(revisionId || '');
}

function prepareSharedTrackSegmentIndex_(sheet, entries) {
  const groups = Object.create(null);
  const dataRowCount = Math.max(0, sheet.getLastRow() - 1);
  if (dataRowCount === 0 || !entries.length) return groups;

  const selectedKeys = Object.create(null);
  entries.forEach(function(entry) {
    const metadata = entry.metadata;
    selectedKeys[sharedTrackRevisionKey_(metadata.trackId, metadata.revisionId)] = true;
  });
  const identityRows = sheet.getRange(2, 1, dataRowCount, 4).getValues();
  const pointCountRows = sheet.getRange(2, 6, dataRowCount, 1).getValues();
  identityRows.forEach(function(identity, index) {
    const key = sharedTrackRevisionKey_(String(identity[0] || '').trim(), String(identity[1] || '').trim());
    if (!selectedKeys[key]) return;
    if (!groups[key]) groups[key] = [];
    groups[key].push({
      rowNumber: index + 2,
      identity: identity.slice(),
      pointCount: pointCountRows[index][0]
    });
  });
  return groups;
}

function readSharedTrackSegmentRows_(sheet, indexedRows) {
  const rows = [];
  const source = Array.isArray(indexedRows) ? indexedRows : [];
  for (var index = 0; index < source.length;) {
    const runStart = index;
    while (index + 1 < source.length
        && source[index + 1].rowNumber === source[index].rowNumber + 1) {
      index += 1;
    }
    const runLength = index - runStart + 1;
    const jsonValues = sheet.getRange(source[runStart].rowNumber, 5, runLength, 1).getValues();
    for (var offset = 0; offset < runLength; offset += 1) {
      const indexed = source[runStart + offset];
      rows.push(indexed.identity.concat([jsonValues[offset][0], indexed.pointCount]));
    }
    index += 1;
  }
  return rows;
}

function getSharedTrackRouteMap_(prepared, metrics) {
  const result = Object.create(null);
  const entries = prepared && Array.isArray(prepared.entries) ? prepared.entries : [];
  if (entries.length === 0 || !prepared.segmentsSheet) {
    return { ok: true, routesByKey: result, coordinateCount: metrics.coordinateCount };
  }

  let segmentIndex;
  try {
    segmentIndex = prepareSharedTrackSegmentIndex_(prepared.segmentsSheet, entries);
  } catch (_trackError) {
    return { ok: true, routesByKey: result, coordinateCount: metrics.coordinateCount };
  }
  let coordinateCount = metrics.coordinateCount;
  for (var i = 0; i < entries.length; i += 1) {
    const entry = entries[i];
    const nextCoordinateCount = coordinateCount + entry.metadata.pointCount;
    const nextBudget = validateSharedPayloadBudget_(null, {
      pinCount: metrics.pinCount,
      routeCount: metrics.routeCount,
      coordinateCount: nextCoordinateCount
    });
    if (!nextBudget.ok) return { ok: false };
    try {
      const revisionKey = sharedTrackRevisionKey_(entry.metadata.trackId, entry.metadata.revisionId);
      const segmentRows = readSharedTrackSegmentRows_(prepared.segmentsSheet, segmentIndex[revisionKey]);
      const track = trackFromStoredRows_(entry.metadata, segmentRows);
      result[entry.key] = toSharedTrackRoute_(track, entry.type);
      coordinateCount = nextCoordinateCount;
    } catch (_segmentError) {
      // An incomplete selected track is omitted instead of leaking partial storage.
    }
  }
  return { ok: true, routesByKey: result, coordinateCount: coordinateCount };
}

function getSharedPinRouteMap_(routeGroups) {
  const pinRouteMap = Object.create(null);
  (Array.isArray(routeGroups) ? routeGroups : []).forEach(function(group) {
    const route = toSharedPinRoute_(group);
    pinRouteMap[shareRouteTargetKey_('pin-route', route.id)] = route;
  });
  return pinRouteMap;
}

function buildSharedTypedRoutes_(shareLink, pinRouteMap, trackRouteMap) {
  if (!shareLink || shareLink.routeTargetsState === 'invalid') return [];
  if (shareLink.routeTargetsState !== 'valid') {
    return Object.keys(pinRouteMap).map(function(key) { return pinRouteMap[key]; });
  }
  const selectedKeySet = Object.create(null);
  shareLink.routeTargets.forEach(function(target) {
    selectedKeySet[shareRouteTargetKey_(target.type, target.id)] = true;
  });
  const pinRoutes = Object.keys(pinRouteMap).filter(function(key) {
    return selectedKeySet[key] === true;
  }).map(function(key) { return pinRouteMap[key]; });
  const trackRoutes = Object.keys(trackRouteMap).filter(function(key) {
    return selectedKeySet[key] === true;
  }).map(function(key) { return trackRouteMap[key]; });
  return pinRoutes.concat(trackRoutes);
}

function sharedPinHasCoordinates_(pin) {
  return !!pin && pin.lat != null && pin.lng != null
    && Number.isFinite(Number(pin.lat)) && Number.isFinite(Number(pin.lng));
}

function buildSharedViewDataForLink_(shareLink) {
  var allRouteGroups = getRouteGroups();
  var pins = getSharedPinsForShareLink_(shareLink, allRouteGroups);
  var coordinateCount = pins.reduce(function(count, pin) {
    return count + (sharedPinHasCoordinates_(pin) ? 1 : 0);
  }, 0);
  var metrics = { pinCount: pins.length, routeCount: 0, coordinateCount: coordinateCount };
  if (!validateSharedPayloadBudget_(null, metrics).ok) return sharedPayloadTooLargeReadResult_();

  var routeGroups = filterSharedRouteGroupsForShareLink_(getSharedRouteGroups_(pins, allRouteGroups), shareLink);
  var pinRouteMap = getSharedPinRouteMap_(routeGroups);
  var preparedTracks = shareLink && shareLink.routeTargetsState === 'valid'
    ? prepareSharedTrackRoutes_(shareLink.routeTargets)
    : { entries: [], segmentsSheet: null };
  metrics.routeCount = Object.keys(pinRouteMap).length + preparedTracks.entries.length;
  if (!validateSharedPayloadBudget_(null, metrics).ok) return sharedPayloadTooLargeReadResult_();

  var trackResult = getSharedTrackRouteMap_(preparedTracks, metrics);
  if (!trackResult.ok) return sharedPayloadTooLargeReadResult_();
  metrics.coordinateCount = trackResult.coordinateCount;
  var routes = buildSharedTypedRoutes_(shareLink, pinRouteMap, trackResult.routesByKey);
  metrics.routeCount = routes.length;

  var noRoutes = shareLink.routeTargetsState === 'invalid'
    || (shareLink.routeTargetsState === 'valid' && shareLink.routeTargets.length === 0)
    || (shareLink.routeTargetsState === 'legacy' && isShareRouteSelectionNone_(shareLink.routeIds));
  var allowedPinRouteIds = getSharePinRouteIds_(shareLink);
  var dto = {
    ok: true,
    noRoutes: noRoutes,
    shareLink: {
      label: shareLink.label,
      token: shareLink.token,
      tags: shareLink.tags,
      tagMode: shareLink.tagMode,
      colors: shareLink.colors.slice(),
      routeIds: shareLink.routeIds.slice(),
      routeTargets: shareLink.routeTargetsState === 'valid' ? shareLink.routeTargets.slice() : null
    },
    allowedTags: shareLink.tags.slice(),
    allowedColors: shareLink.colors.slice(),
    allowedRouteIds: noRoutes || isShareRouteSelectionNone_(allowedPinRouteIds) ? [] : allowedPinRouteIds.slice(),
    pins: pins,
    routeGroups: routeGroups,
    routes: routes
  };
  if (!validateSharedPayloadBudget_(dto, metrics).ok) return sharedPayloadTooLargeReadResult_();
  return dto;
}

function resolveSharedProjection_(shareToken) {
  var shareLink = getShareLinkByToken_(shareToken);
  if (!shareLink) {
    return { ok: false, result: { ok: false, error: 'invalid_share_link' } };
  }
  if (!shareLink.enabled || shareLink.revokedAt) {
    return { ok: false, result: { ok: false, error: 'revoked_share_link' } };
  }
  var data = buildSharedViewDataForLink_(shareLink);
  if (!data || data.ok !== true) {
    return { ok: false, result: data || { ok: false, error: 'invalid_share_link' } };
  }
  var pinIds = new Set();
  data.pins.forEach(function(pin) {
    var pinId = String(pin && pin.id || '').trim();
    if (pinId) pinIds.add(pinId);
  });
  return { ok: true, data: data, pinIds: pinIds };
}

function createSafeSharedAudioError_() {
  return importItemError_(
    'SHARED_PIN_AUDIO_UNAVAILABLE',
    'shared pin audio is unavailable.',
    false
  );
}

function getSharedPinAudioData(payload) {
  try {
    if (!payload || typeof payload !== 'object') throw new Error('invalid payload');
    var projection = resolveSharedProjection_(payload.shareToken);
    if (!projection.ok) throw new Error('invalid projection');
    var pinId = normalizeImportIdentifier_(
      payload.pinId,
      'pinId',
      IMPORT_ITEM_ID_MAX_LENGTH
    );
    if (!projection.pinIds.has(pinId)) throw new Error('pin is outside projection');
    var sharedPin = projection.data.pins.find(function(pin) {
      return String(pin && pin.id || '') === pinId;
    });
    if (!sharedPin || sharedPin.hasAudio !== true) throw new Error('pin audio is unavailable');
    return pinAudioDataFromBlob_(readPinAudioBlobByPinId_(pinId));
  } catch (_error) {
    throw createSafeSharedAudioError_();
  }
}

function getSharedViewData(token) {
  var projection = resolveSharedProjection_(token);
  return projection.ok ? projection.data : projection.result;
}

function ensureImportReceiptsSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(IMPORT_RECEIPTS_SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(IMPORT_RECEIPTS_SHEET_NAME);
  const hasRows = sheet.getLastRow() > 0;
  const hasHeader = hasRows
    && String(sheet.getRange(1, 1).getValue() || '') === IMPORT_RECEIPT_HEADERS[0];
  if (hasRows && !hasHeader) sheet.insertRowBefore(1);
  if (sheet.getMaxColumns() < IMPORT_RECEIPT_COLUMN_COUNT) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      IMPORT_RECEIPT_COLUMN_COUNT - sheet.getMaxColumns()
    );
  }
  if (!hasHeader) {
    sheet.getRange(1, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT).setValues([IMPORT_RECEIPT_HEADERS]);
  } else {
    const headerRange = sheet.getRange(1, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT);
    const headerValues = headerRange.getValues()[0];
    const headerFormulas = headerRange.getFormulas()[0];
    let matchingPrefixLength = 0;
    while (matchingPrefixLength < IMPORT_RECEIPT_COLUMN_COUNT
        && headerValues[matchingPrefixLength] === IMPORT_RECEIPT_HEADERS[matchingPrefixLength]) {
      matchingPrefixLength += 1;
    }
    const missingSuffix = headerValues.slice(matchingPrefixLength).every(function(value, index) {
      return isTrulyBlankSheetCell_(value, headerFormulas[matchingPrefixLength + index]);
    });
    if (matchingPrefixLength < IMPORT_RECEIPT_COLUMN_COUNT && missingSuffix) {
      sheet.getRange(
        1,
        matchingPrefixLength + 1,
        1,
        IMPORT_RECEIPT_COLUMN_COUNT - matchingPrefixLength
      ).setValues([IMPORT_RECEIPT_HEADERS.slice(matchingPrefixLength)]);
    } else {
      IMPORT_RECEIPT_HEADERS.forEach(function(header, index) {
        if (headerValues[index] !== header && !headerFormulas[index]) {
          sheet.getRange(1, index + 1).setValue(header);
        }
      });
    }
  }
  sheet.getRange(1, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT)
    .setBackground('#1565c0')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function importItemError_(code, message, retryable) {
  const error = new Error(message);
  error.code = code;
  error.retryable = !!retryable;
  error.isImportItemError = true;
  return error;
}

function isImportItemError_(error) {
  return !!(error && error.isImportItemError === true && error.code);
}

function importItemFailure_(code, message, retryable) {
  return { ok: false, error: message, errorCode: code, retryable: !!retryable };
}

function importItemFailureFromError_(error) {
  const code = isImportItemError_(error) ? String(error.code) : 'IMPORT_ITEM_SAVE_FAILED';
  const messages = {
    INVALID_IMPORT_PAYLOAD: 'インポート項目の入力内容を確認してください。',
    INVALID_AUDIO_PAYLOAD: '音声の入力内容を確認してください。',
    INVALID_IDEMPOTENCY_KEY: 'インポート項目の識別情報が一致しません。',
    IDEMPOTENCY_PAYLOAD_CONFLICT: '同じ項目に異なる内容が送信されました。',
    IMPORT_RECEIPT_SHEET_MISSING: 'import_receipts シートが見つかりません。setupSheet() を実行してください。',
    IMPORT_ITEM_IN_PROGRESS: '同じ項目を保存中です。少し待ってから再試行してください。',
    IMPORT_ITEM_LEASE_LOST: '保存処理の所有権が変わりました。再試行してください。',
    IMPORT_RECEIPT_CORRUPTED: 'インポートの保存記録が不整合です。管理者へ連絡してください。',
    IMPORT_DRIVE_FILE_FAILED: '写真ファイルを保存できませんでした。再試行してください。',
    IMPORT_DRIVE_SOURCE_INVALID: '選択したDrive写真を確認できませんでした。もう一度選択してください。',
    IMPORT_DRIVE_SOURCE_DELETE_FAILED: '以前のDrive取込の保存状態を確認できませんでした。再読み込みしてください。',
    DRIVE_SOURCE_NOT_EDITABLE: '選択したDrive写真を表示用ファイルとして利用できません。写真を選び直してください。',
    DRIVE_SOURCE_CHECK_FAILED: '選択したDrive写真を確認できませんでした。再試行してください。',
    DRIVE_LINK_SHARING_DENIED: '組織のGoogle Drive共有ポリシーにより写真を公開できません。公開可能な保存先を設定するか、Google Workspace管理者へリンク共有設定を確認してください。',
    DRIVE_LINK_SHARING_FAILED: '管理用写真のリンク共有を確認できませんでした。再試行してください。',
    DRIVE_MANAGED_COPY_CREATE_FAILED: '管理用の写真コピーを作成できませんでした。保存先Driveの作成権限を確認してください。',
    DRIVE_MANAGED_COPY_FINALIZE_FAILED: '管理用の写真コピーを確定できませんでした。再試行してください。',
    DRIVE_MEDIA_STRUCTURE_AMBIGUOUS: 'Driveメディアの保存先に同名フォルダが複数あります。整理してから再試行してください。',
    DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT: 'Driveメディアの保存先に同名の項目があります。整理してから再試行してください。',
    DRIVE_MEDIA_STRUCTURE_PARENT_INVALID: 'Driveメディアの保存先階層を確認してください。',
    DRIVE_MEDIA_STRUCTURE_FAILED: 'Driveメディアの保存先を準備できませんでした。Driveの作成権限を確認してください。',
    DRIVE_SOURCE_ALREADY_LINKED: 'このDrive写真は既に別のピンへ紐づいています。',
    DRIVE_ORIGINAL_FOLDER_CREATE_FAILED: 'originalフォルダを作成できませんでした。Driveの作成権限を確認してください。',
    DRIVE_ORIGINAL_FOLDER_AMBIGUOUS: 'originalフォルダが複数あります。1つに整理してから再試行してください。',
    DRIVE_SOURCE_MOVE_FAILED: 'ピンは登録しましたが、元写真をoriginal/photosへ移動できませんでした。Driveの移動権限を確認して再試行してください。',
    DRIVE_SOURCE_MOVE_VERIFY_FAILED: 'ピンは登録しましたが、元写真のoriginal/photosへの移動結果を確認できませんでした。Driveを確認して再試行してください。',
    IMPORT_MAP_ROW_FAILED_AFTER_SOURCE_MOVE: '元写真の整理は完了しましたが、ピン情報を保存できませんでした。再試行してください。',
    PIN_PHOTO_ATTACH_TARGET_NOT_FOUND: '写真を追加するピンが見つかりません。画面を再読み込みしてください。',
    PIN_PHOTO_ATTACH_ALREADY_HAS_PHOTO: 'このピンには既に写真が登録されています。',
    PIN_PHOTO_ATTACH_CONFLICT: 'ピンが別の操作で更新されました。画面を再読み込みしてから再試行してください。',
    PIN_PHOTO_ATTACH_FILE_CREATE_FAILED: '表示用の写真を作成できませんでした。Driveの保存権限を確認してください。',
    PIN_PHOTO_ATTACH_SOURCE_MOVE_FAILED: '写真は追加しましたが、元写真をoriginal/photosへ移動できませんでした。再試行してください。',
    PIN_PHOTO_ATTACH_MAP_UPDATE_FAILED: '写真ファイルは準備できましたが、ピンへ設定できませんでした。再試行してください。',
    PIN_AUDIO_TARGET_NOT_FOUND: '音声を設定するピンが見つかりません。画面を再読み込みしてください。',
    PIN_AUDIO_ALREADY_ATTACHED: 'このピンには既に音声が登録されています。',
    PIN_AUDIO_MISSING: 'このピンには差し替える音声がありません。',
    PIN_AUDIO_CONFLICT: 'ピンが別の操作で更新されました。画面を再読み込みしてから再試行してください。',
    IMPORT_AUDIO_SOURCE_INVALID: '選択したDrive音声を確認できませんでした。もう一度選択してください。',
    IMPORT_AUDIO_SOURCE_CHECK_FAILED: '選択したDrive音声を確認できませんでした。再試行してください。',
    IMPORT_AUDIO_SOURCE_ARCHIVE_FAILED: '音声は登録しましたが、元ファイルをoriginal/audioへ移動できませんでした。再試行してください。',
    IMPORT_AUDIO_FILE_INVALID: '管理用音声ファイルを確認できませんでした。管理者へ連絡してください。',
    IMPORT_AUDIO_FILE_SAVE_FAILED: '管理用音声ファイルを保存できませんでした。再試行してください。',
    IMPORT_AUDIO_MAP_UPDATE_FAILED: '音声ファイルは準備できましたが、ピンへ設定できませんでした。再試行してください。',
    IMPORT_AUDIO_CLEANUP_FAILED: '音声は登録しましたが、以前の音声ファイルを整理できませんでした。再試行してください。',
    PIN_AUDIO_NOT_FOUND: '音声を確認できませんでした。',
    IMPORT_MAP_ROW_FAILED: 'ピン情報を保存できませんでした。再試行してください。',
    IMPORT_ITEM_SAVE_FAILED: 'インポート項目を保存できませんでした。再試行してください。'
  };
  const retryableCodes = {
    IMPORT_ITEM_IN_PROGRESS: true,
    IMPORT_ITEM_LEASE_LOST: true,
    IMPORT_DRIVE_FILE_FAILED: true,
    IMPORT_DRIVE_SOURCE_DELETE_FAILED: true,
    DRIVE_SOURCE_CHECK_FAILED: true,
    DRIVE_LINK_SHARING_FAILED: true,
    DRIVE_MANAGED_COPY_CREATE_FAILED: true,
    DRIVE_MANAGED_COPY_FINALIZE_FAILED: true,
    DRIVE_MEDIA_STRUCTURE_PARENT_INVALID: true,
    DRIVE_MEDIA_STRUCTURE_FAILED: true,
    DRIVE_ORIGINAL_FOLDER_CREATE_FAILED: true,
    DRIVE_SOURCE_MOVE_FAILED: true,
    DRIVE_SOURCE_MOVE_VERIFY_FAILED: true,
    IMPORT_MAP_ROW_FAILED_AFTER_SOURCE_MOVE: true,
    PIN_PHOTO_ATTACH_FILE_CREATE_FAILED: true,
    PIN_PHOTO_ATTACH_SOURCE_MOVE_FAILED: true,
    PIN_PHOTO_ATTACH_MAP_UPDATE_FAILED: true,
    IMPORT_AUDIO_SOURCE_CHECK_FAILED: true,
    IMPORT_AUDIO_SOURCE_ARCHIVE_FAILED: true,
    IMPORT_AUDIO_FILE_SAVE_FAILED: true,
    IMPORT_AUDIO_MAP_UPDATE_FAILED: true,
    IMPORT_AUDIO_CLEANUP_FAILED: true,
    IMPORT_MAP_ROW_FAILED: true,
    IMPORT_ITEM_SAVE_FAILED: true
  };
  return importItemFailure_(
    code,
    messages[code] || messages.IMPORT_ITEM_SAVE_FAILED,
    isImportItemError_(error) && typeof error.retryable === 'boolean'
      ? error.retryable : !!retryableCodes[code]
  );
}

function importProviderErrorField_(error, field) {
  try {
    const value = error && error[field];
    return value == null ? '' : String(value);
  } catch (_error) {
    return '';
  }
}

function isDriveProviderPermissionDenied_(error) {
  const detail = (
    importProviderErrorField_(error, 'name') + ' '
    + importProviderErrorField_(error, 'message')
  ).toLowerCase();
  return /access denied|permission denied|insufficient permissions?|do not have permission|not authorized|not permitted|forbidden|polic(?:y|ies)|disabled by (?:your )?administrator|shar(?:e|ing)[^\n]{0,40}(?:disabled|prohibited|not allowed)|organization[^\n]{0,40}(?:restrict|does not allow)|権限|ポリシー|管理者[^\n]{0,20}(?:無効|禁止|制限)|共有[^\n]{0,30}(?:禁止|制限|許可され)|アクセス[^\s。]*(?:拒否|許可され)|許可されていません/i.test(detail);
}

function sanitizeProviderMessage_(value) {
  try {
    const message = String(value == null ? '' : value).toLowerCase();
    if (!message) return '';
    if (/access denied|permission denied|insufficient permissions?|do not have permission|not authorized|not permitted|forbidden|polic(?:y|ies)|disabled by (?:your )?administrator|shar(?:e|ing)[^\n]{0,40}(?:disabled|prohibited|not allowed)|organization[^\n]{0,40}(?:restrict|does not allow)|権限|ポリシー|管理者[^\n]{0,20}(?:無効|禁止|制限)|共有[^\n]{0,30}(?:禁止|制限|許可され)|許可されていません/i.test(message)) {
      return 'permission_or_policy_denied';
    }
    if (/quota|rate limit|too many requests|resource exhausted|容量|上限/i.test(message)) {
      return 'quota_or_rate_limited';
    }
    if (/timeout|timed out|temporar|service unavailable|internal error|再試行/i.test(message)) {
      return 'transient_provider_failure';
    }
    if (/not found|missing|unavailable|見つかりません|利用できません/i.test(message)) {
      return 'resource_unavailable';
    }
    return 'provider_error';
  } catch (_error) {
    return '';
  }
}

function sanitizeProviderErrorName_(value) {
  const name = importProviderErrorField_({ value: value }, 'value');
  if (!name) return '';
  return /^(?:Error|Exception|TypeError|RangeError|ServiceException)$/.test(name)
    ? name : 'ProviderError';
}

function logDrivePhotoImportFailure_(stage, error) {
  try {
    console.error(JSON.stringify({
      operation: 'drive-photo-import',
      stage: String(stage || 'unknown'),
      errorName: sanitizeProviderErrorName_(importProviderErrorField_(error, 'name')),
      errorMessage: sanitizeProviderMessage_(importProviderErrorField_(error, 'message'))
    }));
  } catch (_error) {
    // Diagnostics must never change the import result.
  }
}

function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(value),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(byte) {
    return ((Number(byte) + 256) % 256).toString(16).padStart(2, '0');
  }).join('');
}

function normalizeImportIdentifier_(value, field, maxLength) {
  const normalized = String(value == null ? '' : value).trim();
  if (!normalized || normalized.length > maxLength || /[\u0000-\u001f\u007f]/.test(normalized)
      || /^[=+\-@]/.test(normalized)) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', field + ' is invalid.', false);
  }
  return normalized;
}

function isValidImportHttpUrl_(value) {
  const source = String(value || '');
  if (!source || /[\u0000-\u001f\u007f-\u009f]/.test(source)
      || !/^https?:\/\/\S+$/i.test(source)) return false;
  if (typeof URL === 'function') {
    try {
      const parsed = new URL(source);
      return (parsed.protocol === 'http:' || parsed.protocol === 'https:') && !!parsed.hostname;
    } catch (_error) {
      return false;
    }
  }
  const match = source.match(/^https?:\/\/([^\s\/?#]+)(?:[\/?#][^\s]*)?$/i);
  if (!match) return false;
  let authority = String(match[1] || '');
  const atIndex = authority.lastIndexOf('@');
  if (atIndex !== -1) {
    if (atIndex === authority.length - 1) return false;
    authority = authority.slice(atIndex + 1);
  }
  let hostname = authority;
  let port = '';
  let bracketedIpv6 = false;
  if (authority.charAt(0) === '[') {
    bracketedIpv6 = true;
    const closing = authority.indexOf(']');
    if (closing <= 1) return false;
    hostname = authority.slice(1, closing);
    const suffix = authority.slice(closing + 1);
    if (suffix) {
      if (suffix.charAt(0) !== ':') return false;
      port = suffix.slice(1);
    }
    let ipv6ForValidation = hostname;
    const embeddedIpv4 = hostname.match(/^(.*:)(\d{1,3}(?:\.\d{1,3}){3})$/);
    if (embeddedIpv4) {
      const ipv4Parts = embeddedIpv4[2].split('.');
      if (ipv4Parts.some(function(part) {
        return !/^(?:0|[1-9]\d{0,2})$/.test(part) || Number(part) > 255;
      })) return false;
      ipv6ForValidation = embeddedIpv4[1] + '0:0';
    }
    if (!ipv6ForValidation.includes(':') || !/^[0-9a-f:]+$/i.test(ipv6ForValidation)) return false;
    const halves = ipv6ForValidation.split('::');
    if (halves.length > 2) return false;
    const groups = halves.reduce(function(result, half) {
      return result.concat(half ? half.split(':') : []);
    }, []);
    if (groups.some(function(group) { return !/^[0-9a-f]{1,4}$/i.test(group); })) return false;
    if (halves.length === 1 ? groups.length !== 8 : groups.length >= 8) return false;
  } else {
    const colon = authority.lastIndexOf(':');
    if (colon !== -1) {
      if (authority.indexOf(':') !== colon) return false;
      hostname = authority.slice(0, colon);
      port = authority.slice(colon + 1);
    }
  }
  if (!hostname) return false;
  if (port && (!/^\d+$/.test(port) || Number(port) > 65535)) return false;
  if (authority.endsWith(':') && !port) return false;
  if (/%(?![0-9a-f]{2})/i.test(hostname)) return false;
  let decodedHostname = hostname;
  if (!bracketedIpv6) {
    try {
      decodedHostname = decodeURIComponent(hostname);
    } catch (_error) {
      return false;
    }
    if (/[\u0000-\u0020\u007f-\u009f<>\\^|]/.test(decodedHostname)
        || /[#\/:?@\[\]]/.test(decodedHostname)) return false;
  }
  if (/^[0-9.]+$/.test(decodedHostname)) {
    const parts = decodedHostname.split('.');
    if (parts.length !== 4 || parts.some(function(part) {
      return !/^\d{1,3}$/.test(part) || Number(part) > 255;
    })) return false;
  }
  if (!bracketedIpv6) {
    const labels = decodedHostname.replace(/\.$/, '').split('.');
    if (labels.some(function(label) { return !label || /[\[\]@:\s]/.test(label); })) return false;
  }
  return true;
}

function normalizeNewPinPayload_(data, options) {
  const source = data || {};
  const config = options || {};
  const title = String(source.title == null ? '' : source.title).trim();
  if (!title || (config.strict && title.length > PIN_TITLE_MAX_LENGTH)) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'title is invalid.', false);
  }
  const description = String(source.description == null ? '' : source.description);
  if (config.strict && description.length > 400) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'description is invalid.', false);
  }
  const color = config.strict ? source.color : (source.color || DEFAULT_COLOR);
  if (config.strict && (typeof color !== 'string'
      || !SAFE_COLOR_RE.test(color)
      || PinData.COLOR_OPTIONS.indexOf(color.toLowerCase()) === -1)) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'color is invalid.', false);
  }
  const rawIcon = String(source.icon == null ? '' : source.icon).trim().toLowerCase();
  if (config.strict && PinData.ICON_OPTIONS.indexOf(rawIcon) === -1) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'icon is invalid.', false);
  }
  const normalizedColor = config.strict && config.normalizeColor === true
    ? color.toLowerCase() : color;
  const icon = config.strict ? rawIcon : PinData.normalizeIcon(source.icon);
  const hasStatus = Object.prototype.hasOwnProperty.call(source, 'status');
  let status;
  try {
    status = hasStatus
      ? PinData.normalizeStatus(String(source.status == null ? '' : source.status))
      : (config.defaultMissingStatus == null ? '' : String(config.defaultMissingStatus));
  } catch (error) {
    if (!config.strict) throw error;
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'status is invalid.', false);
  }
  const latMissing = source.lat == null;
  const lngMissing = source.lng == null;
  let lat = latMissing ? null : Number(source.lat);
  let lng = lngMissing ? null : Number(source.lng);
  if (config.strict) {
    if (latMissing !== lngMissing
        || (!latMissing && (typeof source.lat !== 'number' || typeof source.lng !== 'number'
          || !Number.isFinite(lat) || !Number.isFinite(lng)
          || lat < -90 || lat > 90 || lng < -180 || lng > 180))) {
      throw importItemError_('INVALID_IMPORT_PAYLOAD', 'coordinates are invalid.', false);
    }
  }
  let tags;
  try {
    if (config.strict && !Array.isArray(source.tags)) {
      throw new Error('tags must be an array');
    }
    tags = PinData.normalizeTags(source.tags || []);
  } catch (error) {
    if (!config.strict) throw error;
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'tags are invalid.', false);
  }
  let links;
  if (config.strict) {
    if (!Array.isArray(source.links)) {
      throw importItemError_('INVALID_IMPORT_PAYLOAD', 'links are invalid.', false);
    }
    const seenLinks = {};
    links = [];
    source.links.forEach(function(value) {
      if (typeof value !== 'string') {
        throw importItemError_('INVALID_IMPORT_PAYLOAD', 'links are invalid.', false);
      }
      const link = value.trim();
      if (!link) return;
      if (!isValidImportHttpUrl_(link)) {
        throw importItemError_('INVALID_IMPORT_PAYLOAD', 'links are invalid.', false);
      }
      if (!seenLinks[link]) {
        seenLinks[link] = true;
        links.push(link);
      }
    });
  } else {
    links = PinData.normalizeLinks(source.links || source.referenceUrls || []);
  }
  const rawEventAt = String(source.eventAt == null ? '' : source.eventAt).trim();
  const eventAt = PinData.normalizeEventAt(rawEventAt);
  if (config.strict && rawEventAt && !eventAt) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'eventAt is invalid.', false);
  }
  return {
    title: title,
    description: description,
    lat: lat,
    lng: lng,
    color: normalizedColor,
    icon: icon,
    status: status,
    tags: tags,
    links: links,
    eventAt: eventAt
  };
}

function buildMapInfoRow_(normalized, resultMeta) {
  const meta = resultMeta || {};
  return [
    meta.timestamp || '',
    normalized.title,
    normalized.description,
    normalized.lat == null ? '' : normalized.lat,
    normalized.lng == null ? '' : normalized.lng,
    normalized.color,
    meta.fileId || '',
    meta.imageUrl || '',
    meta.pinId || '',
    PinData.serializeLinks(normalized.links),
    normalized.status,
    PinData.serializeTags(normalized.tags),
    normalized.eventAt,
    meta.updatedAt || '',
    normalized.icon,
    ''
  ];
}

function spreadsheetLiteral_(value) {
  return encodeSpreadsheetLiteral_(value);
}

function appendMapInfoRow_(sheet, row) {
  const storageRow = row.slice();
  storageRow[1] = spreadsheetLiteral_(storageRow[1]);
  storageRow[2] = spreadsheetLiteral_(storageRow[2]);
  storageRow[11] = spreadsheetLiteral_(storageRow[11]);
  sheet.appendRow(storageRow);
}

function mapInfoRowToPinResult_(row, folderUrl) {
  const pin = PinData.rowToPin(row || []);
  pin.folderUrl = String(folderUrl || '');
  return toClientPin_(pin);
}

function normalizeImportPhotoPayload_(data) {
  const source = data || {};
  const operationMode = source.operationMode == null || source.operationMode === ''
    ? 'create-pin' : String(source.operationMode);
  if (operationMode !== 'create-pin' && operationMode !== 'attach-existing-pin') {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'operationMode is invalid.', false);
  }
  const jobId = normalizeImportIdentifier_(source.jobId, 'jobId', IMPORT_ITEM_ID_MAX_LENGTH);
  const itemId = normalizeImportIdentifier_(source.itemId, 'itemId', IMPORT_ITEM_ID_MAX_LENGTH);
  const idempotencyKey = normalizeImportIdentifier_(
    source.idempotencyKey,
    'idempotencyKey',
    IMPORT_IDEMPOTENCY_KEY_MAX_LENGTH
  );
  if (idempotencyKey !== jobId + ':' + itemId) {
    throw importItemError_('INVALID_IDEMPOTENCY_KEY', 'idempotency key mismatch.', false);
  }
  const base64 = String(source.base64 || '');
  if (!/^data:image\/jpeg;base64,[A-Za-z0-9+/]+={0,2}$/.test(base64)) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'base64 must be JPEG.', false);
  }
  let jpegBytes;
  try {
    jpegBytes = Utilities.base64Decode(base64.replace(/^data:image\/jpeg;base64,/, ''));
  } catch (_error) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'base64 must be JPEG.', false);
  }
  const signature = (jpegBytes || []).slice(0, 3).map(function(byte) {
    return (Number(byte) + 256) % 256;
  });
  if (signature.length !== 3 || signature[0] !== 0xff || signature[1] !== 0xd8 || signature[2] !== 0xff) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'base64 must be JPEG.', false);
  }
  const filename = String(source.filename == null ? '' : source.filename).trim();
  if (!filename || filename.length > 255 || /[\u0000-\u001f\u007f]/.test(filename)) {
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'filename is invalid.', false);
  }
  let pin;
  let targetPinId = '';
  let expectedUpdatedAt = '';
  if (operationMode === 'attach-existing-pin') {
    targetPinId = normalizeImportIdentifier_(
      source.targetPinId,
      'targetPinId',
      IMPORT_ITEM_ID_MAX_LENGTH
    );
    if (typeof source.expectedUpdatedAt !== 'string'
        || source.expectedUpdatedAt.length > 64
        || /[\u0000-\u001f\u007f]/.test(source.expectedUpdatedAt)) {
      throw importItemError_('INVALID_IMPORT_PAYLOAD', 'expectedUpdatedAt is invalid.', false);
    }
    expectedUpdatedAt = source.expectedUpdatedAt;
    pin = {
      title: '', description: '', eventAt: '', lat: null, lng: null,
      color: DEFAULT_COLOR, icon: 'default', links: [], status: '', tags: []
    };
  } else {
    pin = normalizeNewPinPayload_(source, { strict: true, defaultMissingStatus: '' });
  }
  let sourceDriveFileId = '';
  let hasSourceDriveFileId = false;
  try {
    hasSourceDriveFileId = Object.prototype.hasOwnProperty.call(source, 'sourceDriveFileId');
    if (!hasSourceDriveFileId && 'sourceDriveFileId' in Object(source)) {
      throw importItemError_('INVALID_IMPORT_PAYLOAD', 'sourceDriveFileId is invalid.', false);
    }
  } catch (error) {
    if (isImportItemError_(error)) throw error;
    throw importItemError_('INVALID_IMPORT_PAYLOAD', 'sourceDriveFileId is invalid.', false);
  }
  if (hasSourceDriveFileId) {
    if (typeof source.sourceDriveFileId !== 'string') {
      throw importItemError_('INVALID_IMPORT_PAYLOAD', 'sourceDriveFileId is invalid.', false);
    }
    sourceDriveFileId = source.sourceDriveFileId.trim();
    if (sourceDriveFileId && !isValidDrivePhotoImportId_(sourceDriveFileId)) {
      throw importItemError_('INVALID_IMPORT_PAYLOAD', 'sourceDriveFileId is invalid.', false);
    }
  }
  return Object.assign({}, pin, {
    jobId: jobId,
    itemId: itemId,
    idempotencyKey: idempotencyKey,
    idempotencyKeyHash: sha256Hex_(idempotencyKey),
    base64: base64,
    jpegBytes: jpegBytes,
    filename: filename,
    targetFolderId: String(source.targetFolderId || '').trim(),
    sourceDriveFileId: sourceDriveFileId,
    operationMode: operationMode,
    targetPinId: targetPinId,
    expectedUpdatedAt: expectedUpdatedAt
  });
}

function hashImportPayload_(normalized) {
  if (normalized.operationMode === 'attach-existing-pin') {
    const attachPayload = {
      operationMode: normalized.operationMode,
      jobId: normalized.jobId,
      itemId: normalized.itemId,
      targetPinId: normalized.targetPinId,
      expectedUpdatedAt: normalized.expectedUpdatedAt,
      base64: normalized.base64,
      filename: normalized.filename,
      targetFolderId: normalized.targetFolderId
    };
    if (normalized.sourceDriveFileId) {
      attachPayload.sourceDriveFileId = normalized.sourceDriveFileId;
    }
    return sha256Hex_(JSON.stringify(attachPayload));
  }
  const hashPayload = {
    jobId: normalized.jobId,
    itemId: normalized.itemId,
    base64: normalized.base64,
    filename: normalized.filename,
    title: normalized.title,
    description: normalized.description,
    eventAt: normalized.eventAt,
    lat: normalized.lat,
    lng: normalized.lng,
    color: normalized.color,
    icon: normalized.icon,
    links: normalized.links.slice(),
    targetFolderId: normalized.targetFolderId,
    status: normalized.status,
    tags: normalized.tags.slice()
  };
  if (normalized.sourceDriveFileId) {
    hashPayload.sourceDriveFileId = normalized.sourceDriveFileId;
  }
  return sha256Hex_(JSON.stringify(hashPayload));
}

function normalizeImportPinPayload_(data) {
  const source = data || {};
  const jobId = normalizeImportIdentifier_(source.jobId, 'jobId', IMPORT_ITEM_ID_MAX_LENGTH);
  const itemId = normalizeImportIdentifier_(source.itemId, 'itemId', IMPORT_ITEM_ID_MAX_LENGTH);
  const idempotencyKey = normalizeImportIdentifier_(
    source.idempotencyKey,
    'idempotencyKey',
    IMPORT_IDEMPOTENCY_KEY_MAX_LENGTH
  );
  if (idempotencyKey !== jobId + ':' + itemId) {
    throw importItemError_('INVALID_IDEMPOTENCY_KEY', 'idempotency key mismatch.', false);
  }
  const pin = normalizeNewPinPayload_(source, {
    strict: true,
    normalizeColor: true,
    defaultMissingStatus: ''
  });
  return Object.assign({}, pin, {
    jobId: jobId,
    itemId: itemId,
    idempotencyKey: idempotencyKey,
    idempotencyKeyHash: sha256Hex_(idempotencyKey)
  });
}

function hashImportPinPayload_(normalized) {
  return sha256Hex_(JSON.stringify({
    kind: 'pin',
    jobId: normalized.jobId,
    itemId: normalized.itemId,
    title: normalized.title,
    description: normalized.description,
    eventAt: normalized.eventAt,
    lat: normalized.lat,
    lng: normalized.lng,
    color: normalized.color,
    icon: normalized.icon,
    links: normalized.links.slice(),
    status: normalized.status,
    tags: normalized.tags.slice()
  }));
}

function getImportNow_() {
  return new Date();
}

function formatMapTimestamp_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
}

function openImportReceiptsSheet_() {
  try {
    const sheet = getRequiredSheet_(IMPORT_RECEIPTS_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT).getValues()[0];
    const valid = IMPORT_RECEIPT_HEADERS.every(function(header, index) {
      return headers[index] === header;
    });
    if (!valid) {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt headers are invalid.', false);
    }
    return sheet;
  } catch (_error) {
    if (_error && _error.code === 'IMPORT_RECEIPT_CORRUPTED') throw _error;
    throw importItemError_(
      'IMPORT_RECEIPT_SHEET_MISSING',
      'import_receipts sheet is missing.',
      false
    );
  }
}

function inspectImportReceiptSchemaForSave_(sheet) {
  if (!sheet) return { state: 'missing', sheet: null };
  const maxColumns = sheet.getMaxColumns();
  const lastColumn = sheet.getLastColumn();
  const readColumnCount = Math.max(1, Math.min(maxColumns, Math.max(lastColumn, IMPORT_RECEIPT_COLUMN_COUNT)));
  const headerRange = sheet.getRange(1, 1, 1, readColumnCount);
  const headers = headerRange.getValues()[0].map(function(value) {
    return String(value == null ? '' : value);
  });
  const headerFormulas = headerRange.getFormulas()[0];
  while (headers.length < IMPORT_RECEIPT_COLUMN_COUNT) headers.push('');
  while (headerFormulas.length < IMPORT_RECEIPT_COLUMN_COUNT) headerFormulas.push('');
  if (lastColumn > IMPORT_RECEIPT_COLUMN_COUNT) {
    return { state: 'corrupted', sheet: sheet };
  }
  const current = IMPORT_RECEIPT_HEADERS.every(function(header, index) {
    return headers[index] === header;
  });
  if (current) return { state: 'current', sheet: sheet, headerCount: IMPORT_RECEIPT_COLUMN_COUNT };
  const legacyHeaderCounts = [PHOTO_IMPORT_RECEIPT_HEADERS.length, LEGACY_IMPORT_RECEIPT_HEADERS.length];
  for (var legacyIndex = 0; legacyIndex < legacyHeaderCounts.length; legacyIndex += 1) {
    const headerCount = legacyHeaderCounts[legacyIndex];
    const expectedHeaders = IMPORT_RECEIPT_HEADERS.slice(0, headerCount);
    const prefixMatches = expectedHeaders.every(function(header, index) {
      return headers[index] === header;
    });
    const appendedHeadersEmpty = headers.slice(headerCount, IMPORT_RECEIPT_COLUMN_COUNT).every(function(header, index) {
      return isTrulyBlankSheetCell_(header, headerFormulas[headerCount + index]);
    });
    if (prefixMatches && appendedHeadersEmpty) {
      return { state: 'legacy', sheet: sheet, headerCount: headerCount };
    }
  }
  return { state: 'corrupted', sheet: sheet };
}

function findImportReceiptForSavePreflight_(sheet, keyHash, columnCount) {
  if (!sheet || sheet.getLastRow() < 2) return null;
  const keys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const matchingRows = [];
  keys.forEach(function(row, index) {
    if (String(row[0] || '') === String(keyHash || '')) matchingRows.push(index + 2);
  });
  if (matchingRows.length > 1) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'duplicate receipt keys.', false);
  }
  if (!matchingRows.length) return null;
  const rowNumber = matchingRows[0];
  const row = sheet.getRange(rowNumber, 1, 1, columnCount).getValues()[0];
  return importReceiptFromRow_(row, rowNumber);
}

function inspectExistingImportReceiptForSave_(keyHash) {
  const spreadsheet = openDataSpreadsheet_();
  const inspection = inspectImportReceiptSchemaForSave_(
    spreadsheet.getSheetByName(IMPORT_RECEIPTS_SHEET_NAME)
  );
  if (inspection.state === 'corrupted') {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt headers are invalid.', false);
  }
  if (inspection.state === 'missing') return null;
  return findImportReceiptForSavePreflight_(
    inspection.sheet,
    keyHash,
    Number(inspection.headerCount || IMPORT_RECEIPT_COLUMN_COUNT)
  );
}

function ensureImportReceiptSchemaForSave_() {
  const spreadsheet = openDataSpreadsheet_();
  let inspection = inspectImportReceiptSchemaForSave_(
    spreadsheet.getSheetByName(IMPORT_RECEIPTS_SHEET_NAME)
  );
  if (inspection.state === 'current') return inspection.sheet;
  if (inspection.state === 'corrupted') {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt headers are invalid.', false);
  }
  return withImportReceiptLock_(function() {
    inspection = inspectImportReceiptSchemaForSave_(
      spreadsheet.getSheetByName(IMPORT_RECEIPTS_SHEET_NAME)
    );
    if (inspection.state === 'current') return inspection.sheet;
    if (inspection.state === 'corrupted') {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt headers are invalid.', false);
    }

    let sheet = inspection.sheet;
    if (!sheet) {
      sheet = spreadsheet.insertSheet(IMPORT_RECEIPTS_SHEET_NAME);
      if (sheet.getMaxColumns() < IMPORT_RECEIPT_COLUMN_COUNT) {
        sheet.insertColumnsAfter(
          sheet.getMaxColumns(),
          IMPORT_RECEIPT_COLUMN_COUNT - sheet.getMaxColumns()
        );
      }
      sheet.getRange(1, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT).setValues([IMPORT_RECEIPT_HEADERS]);
      sheet.getRange(1, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT)
        .setBackground('#1565c0')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
      return sheet;
    }

    const legacyHeaderCount = inspection.headerCount;
    const appendedColumnCount = IMPORT_RECEIPT_COLUMN_COUNT - legacyHeaderCount;
    if (sheet.getMaxColumns() >= IMPORT_RECEIPT_COLUMN_COUNT && sheet.getLastRow() > 1) {
      const appendedValues = sheet.getRange(
        2,
        legacyHeaderCount + 1,
        sheet.getLastRow() - 1,
        appendedColumnCount
      ).getValues();
      if (appendedValues.some(function(row) {
        return row.some(function(value) { return String(value || '') !== ''; });
      })) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'legacy receipt data has an unexpected column.', false);
      }
    }
    if (sheet.getMaxColumns() < IMPORT_RECEIPT_COLUMN_COUNT) {
      sheet.insertColumnsAfter(
        sheet.getMaxColumns(),
        IMPORT_RECEIPT_COLUMN_COUNT - sheet.getMaxColumns()
      );
    }
    sheet.getRange(1, legacyHeaderCount + 1, 1, appendedColumnCount).setValues([
      IMPORT_RECEIPT_HEADERS.slice(legacyHeaderCount)
    ]);
    sheet.getRange(1, legacyHeaderCount + 1, 1, appendedColumnCount)
      .setBackground('#1565c0')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    return sheet;
  });
}

function importReceiptFromRow_(row, rowNumber) {
  const source = row || [];
  const receipt = { rowNumber: rowNumber };
  IMPORT_RECEIPT_HEADERS.forEach(function(header, index) { receipt[header] = source[index] || ''; });
  if (!receipt.mediaKind) receipt.mediaKind = 'photo';
  return receipt;
}

function importReceiptToRow_(receipt) {
  return IMPORT_RECEIPT_HEADERS.map(function(header) {
    return receipt[header] == null ? '' : receipt[header];
  });
}

function findImportReceipt_(sheet, keyHash) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const matchingRows = [];
  keys.forEach(function(row, index) {
    if (String(row[0] || '') === keyHash) matchingRows.push(index + 2);
  });
  if (matchingRows.length > 1) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'duplicate receipts.', false);
  }
  if (matchingRows.length === 0) return null;
  const rowNumber = matchingRows[0];
  const row = sheet.getRange(rowNumber, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT).getValues()[0];
  return importReceiptFromRow_(row, rowNumber);
}

function writeImportReceipt_(sheet, receipt) {
  const row = importReceiptToRow_(receipt);
  if (receipt.rowNumber) {
    sheet.getRange(receipt.rowNumber, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT).setValues([row]);
  } else {
    sheet.appendRow(row);
    receipt.rowNumber = sheet.getLastRow();
  }
  return receipt;
}

function findMapInfoRowByPinId_(sheet, pinId) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const ids = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
  let foundRow = 0;
  ids.forEach(function(row, index) {
    if (String(row[0] || '') === pinId) {
      if (foundRow) throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'duplicate map pin ids.', false);
      foundRow = index + 2;
    }
  });
  return foundRow
    ? { rowNumber: foundRow, row: sheet.getRange(foundRow, 1, 1, MAP_INFO_COLUMN_COUNT).getValues()[0] }
    : null;
}

function importLeaseIsActive_(receipt, now) {
  if (!receipt.leaseOwner || !receipt.leaseUntil) return false;
  const until = new Date(receipt.leaseUntil);
  return Number.isFinite(until.getTime()) && until.getTime() > now.getTime();
}

function withImportReceiptLock_(callback) {
  const lock = LockService.getScriptLock();
  let acquired = false;
  try {
    acquired = lock.tryLock(SPREADSHEET_MUTATION_LOCK_TIMEOUT_MS);
    if (!acquired) {
      throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'receipt lock is busy.', true);
    }
    let result;
    let callbackError = null;
    try {
      result = callback();
    } catch (error) {
      callbackError = error;
    }
    try {
      SpreadsheetApp.flush();
    } catch (flushError) {
      if (!callbackError) callbackError = flushError;
    }
    if (callbackError) throw callbackError;
    return result;
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function importTempFileName_(pinId) {
  return '__drop_pin_import_' + String(pinId) + '.jpg';
}

function importPhotoAttachTempFileName_(idempotencyKeyHash) {
  const digest = String(idempotencyKeyHash || '').toLowerCase();
  if (!/^[0-9a-f]{64}$/.test(digest)) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'photo attach key is invalid.', false);
  }
  return '__drop_pin_attach_' + digest.slice(0, 32) + '.jpg';
}

function isImportPhotoAttachTempFileName_(value) {
  return /^__drop_pin_attach_[0-9a-f]{32}\.jpg$/.test(String(value || ''));
}

function assertImportReceiptStorage_(receipt, expected) {
  if (String(receipt.targetFolderId) !== expected.targetFolderId
      || String(receipt.tempFileName) !== expected.tempFileName
      || String(receipt.sourceDriveFileId || '') !== String(expected.sourceDriveFileId || '')) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt storage metadata changed.', false);
  }
}

function assertImportPinReceiptStorage_(receipt) {
  const storageColumns = [
    'targetFolderId', 'tempFileName', 'fileId', 'imageUrl', 'folderUrl', 'sourceDriveFileId'
  ];
  if (storageColumns.some(function(column) { return String(receipt[column] || '') !== ''; })) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'photo-less receipt contains storage metadata.', false);
  }
}

function assertImportPinMapRow_(mapResult) {
  if (mapResult && (String(mapResult.row[6] || '') || String(mapResult.row[7] || ''))) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'photo-less map row contains storage metadata.', false);
  }
}

function assertImportReceiptOwner_(receipt, expected) {
  if (!receipt || !receipt.pinId) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt pin is missing.', false);
  }
  if (String(receipt.payloadHash) !== expected.payloadHash) {
    throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'payload conflict.', false);
  }
  if (String(receipt.pinId) !== expected.pinId) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt pin changed.', false);
  }
  assertImportReceiptStorage_(receipt, expected);
  if (String(receipt.leaseOwner) !== expected.leaseOwner) {
    throw importItemError_('IMPORT_ITEM_LEASE_LOST', 'lease owner changed.', true);
  }
}

function assertDriveSourceAvailableForReceipt_(receiptSheet, mapSheet, normalized, now) {
  const sourceDriveFileId = String(normalized.sourceDriveFileId || '');
  if (!sourceDriveFileId) return null;
  const receiptLastRow = receiptSheet.getLastRow();
  const receiptRows = receiptLastRow < 2 ? [] : receiptSheet
    .getRange(2, 1, receiptLastRow - 1, IMPORT_RECEIPT_COLUMN_COUNT)
    .getValues();
  const mapLastRow = mapSheet.getLastRow();
  const livePinIds = new Set();
  if (mapLastRow >= 2) {
    mapSheet.getRange(2, 7, mapLastRow - 1, 3).getValues().forEach(function(row) {
      const pinId = String(row[2] || '');
      if (pinId) livePinIds.add(pinId);
    });
  }
  let recoverableReceipt = null;
  receiptRows.forEach(function(row, index) {
    const receipt = importReceiptFromRow_(row, index + 2);
    if (String(receipt.idempotencyKey || '') === normalized.idempotencyKeyHash
        || String(receipt.sourceDriveFileId || '') !== sourceDriveFileId) return;
    const state = String(receipt.state || '');
    if (state === IMPORT_RECEIPT_STATES.COMPLETED) {
      if (livePinIds.has(String(receipt.pinId || ''))) {
        throw importItemError_(
          'DRIVE_SOURCE_ALREADY_LINKED',
          'Drive source is already linked.',
          false
        );
      }
      return;
    }
    const activeLease = importLeaseIsActive_(receipt, now);
    if (activeLease) {
      throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'Drive source is in progress.', true);
    }
    const hasDurableFile = !!String(receipt.fileId || '');
    const recoverableState = hasDurableFile
      ? (state === IMPORT_RECEIPT_STATES.FAILED
        || state === IMPORT_RECEIPT_STATES.FILE_SAVED)
      : (state === IMPORT_RECEIPT_STATES.FAILED
        || state === IMPORT_RECEIPT_STATES.RESERVED);
    if (recoverableState) {
      if (recoverableReceipt) {
        throw importItemError_(
          'IMPORT_RECEIPT_CORRUPTED',
          'multiple stale Drive source owners exist.',
          false
        );
      }
      recoverableReceipt = receipt;
      return;
    }
    if (hasDurableFile || state === IMPORT_RECEIPT_STATES.FILE_SAVED) {
      throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'Drive source ownership is unresolved.', true);
    }
  });
  return recoverableReceipt;
}

function claimImportReceipt_(receiptSheet, mapSheet, normalized, payloadHash, leaseOwner) {
  return withImportReceiptLock_(function() {
    const now = getImportNow_();
    let receipt = findImportReceipt_(receiptSheet, normalized.idempotencyKeyHash);
    let recoverableReceipt = assertDriveSourceAvailableForReceipt_(
      receiptSheet,
      mapSheet,
      normalized,
      now
    );
    if (receipt && recoverableReceipt) {
      if (String(receipt.jobId) !== normalized.jobId
          || String(receipt.itemId) !== normalized.itemId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt ids mismatch.', false);
      }
      if (String(receipt.payloadHash) !== payloadHash) {
        throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'payload conflict.', false);
      }
      if (!receipt.pinId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt pin is missing.', false);
      }
      assertImportReceiptStorage_(receipt, {
        targetFolderId: normalized.targetFolderId,
        tempFileName: importTempFileName_(receipt.pinId),
        sourceDriveFileId: normalized.sourceDriveFileId
      });
      const currentMap = findMapInfoRowByPinId_(mapSheet, String(receipt.pinId));
      if (!currentMap) {
        if (importLeaseIsActive_(receipt, now)) {
          throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'item is in progress.', true);
        }
        const currentState = String(receipt.state || '');
        const recoverableHasFile = !!String(recoverableReceipt.fileId || '');
        if (!recoverableHasFile && String(receipt.fileId || '')) {
          recoverableReceipt = null;
        } else if (!recoverableHasFile) {
          throw importItemError_(
            'IMPORT_RECEIPT_CORRUPTED',
            'multiple unjournaled Drive files may exist.',
            false
          );
        } else if (String(receipt.fileId || '')
            || (currentState !== IMPORT_RECEIPT_STATES.RESERVED
              && currentState !== IMPORT_RECEIPT_STATES.FAILED)) {
          throw importItemError_(
            'IMPORT_RECEIPT_CORRUPTED',
            'multiple Drive source owners cannot be reconciled.',
            false
          );
        } else {
          receiptSheet.getRange(
            receipt.rowNumber,
            1,
            1,
            IMPORT_RECEIPT_COLUMN_COUNT
          ).setValues([new Array(IMPORT_RECEIPT_COLUMN_COUNT).fill('')]);
          receipt = null;
        }
      }
    }
    if (!receipt && recoverableReceipt) {
      receipt = recoverableReceipt;
      const sourceDriveFileId = String(normalized.sourceDriveFileId || '');
      const recoveredFileId = String(receipt.fileId || '');
      if (!receipt.pinId
          || String(receipt.tempFileName || '') !== importTempFileName_(String(receipt.pinId))
          || String(receipt.sourceDriveFileId || '') !== sourceDriveFileId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'stale Drive receipt is invalid.', false);
      }
      const recoveredTargetFolderId = String(receipt.targetFolderId || '');
      if (recoveredFileId === sourceDriveFileId) {
        receipt.fileId = '';
        receipt.imageUrl = '';
        receipt.folderUrl = '';
        receipt.targetFolderId = normalized.targetFolderId;
      } else if (recoveredTargetFolderId) {
        normalized.targetFolderId = recoveredTargetFolderId;
      } else {
        receipt.targetFolderId = normalized.targetFolderId;
      }
      payloadHash = hashImportPayload_(normalized);
      receipt.idempotencyKey = normalized.idempotencyKeyHash;
      receipt.jobId = normalized.jobId;
      receipt.itemId = normalized.itemId;
      receipt.payloadHash = payloadHash;
      receipt.sourceDriveFileId = sourceDriveFileId;
    }
    if (receipt) {
      if (String(receipt.jobId) !== normalized.jobId || String(receipt.itemId) !== normalized.itemId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt ids mismatch.', false);
      }
      if (String(receipt.payloadHash) !== payloadHash) {
        throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'payload conflict.', false);
      }
      if (!receipt.pinId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt pin is missing.', false);
      }
      assertImportReceiptStorage_(receipt, {
        targetFolderId: normalized.targetFolderId,
        tempFileName: importTempFileName_(receipt.pinId),
        sourceDriveFileId: normalized.sourceDriveFileId
      });
      const existingMap = findMapInfoRowByPinId_(mapSheet, String(receipt.pinId));
      if (existingMap) {
        const mapFileId = existingMap.row[6] || receipt.fileId || '';
        const mapImageUrl = existingMap.row[7] || receipt.imageUrl || '';
        const needsReceiptRepair = receipt.state !== IMPORT_RECEIPT_STATES.COMPLETED
          || String(receipt.leaseOwner || '') !== ''
          || String(receipt.leaseUntil || '') !== ''
          || String(receipt.lastErrorCode || '') !== ''
          || String(receipt.fileId || '') !== String(mapFileId)
          || String(receipt.imageUrl || '') !== String(mapImageUrl);
        if (needsReceiptRepair) {
          receipt.state = IMPORT_RECEIPT_STATES.COMPLETED;
          receipt.leaseOwner = '';
          receipt.leaseUntil = '';
          receipt.fileId = mapFileId;
          receipt.imageUrl = mapImageUrl;
          receipt.updatedAt = now.toISOString();
          receipt.lastErrorCode = '';
          writeImportReceipt_(receiptSheet, receipt);
        }
        return {
          completed: true,
          receipt: receipt,
          payloadHash: payloadHash,
          pin: mapInfoRowToPinResult_(existingMap.row, receipt.folderUrl)
        };
      }
      if (importLeaseIsActive_(receipt, now) && String(receipt.leaseOwner) !== leaseOwner) {
        throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'item is in progress.', true);
      }
    } else {
      const pinId = Utilities.getUuid();
      receipt = {
        rowNumber: 0,
        idempotencyKey: normalized.idempotencyKeyHash,
        jobId: normalized.jobId,
        itemId: normalized.itemId,
        payloadHash: payloadHash,
        state: IMPORT_RECEIPT_STATES.RESERVED,
        leaseOwner: '',
        leaseUntil: '',
        pinId: pinId,
        targetFolderId: normalized.targetFolderId,
        tempFileName: importTempFileName_(pinId),
        fileId: '', imageUrl: '', folderUrl: '',
        createdAt: now.toISOString(), updatedAt: now.toISOString(), lastErrorCode: '',
        sourceDriveFileId: normalized.sourceDriveFileId,
        mediaKind: 'photo', operationMode: 'create-pin', targetPinId: '', cleanupFileId: ''
      };
    }
    receipt.state = receipt.fileId ? IMPORT_RECEIPT_STATES.FILE_SAVED : IMPORT_RECEIPT_STATES.RESERVED;
    receipt.leaseOwner = leaseOwner;
    receipt.leaseUntil = new Date(now.getTime() + IMPORT_ITEM_LEASE_MS).toISOString();
    receipt.updatedAt = now.toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return { completed: false, receipt: receipt, payloadHash: payloadHash };
  });
}

function photoAttachTarget_(mapSheet, normalized, receipt) {
  const target = findMapInfoRowByPinId_(mapSheet, String(normalized.targetPinId || ''));
  if (!target) {
    throw importItemError_(
      'PIN_PHOTO_ATTACH_TARGET_NOT_FOUND',
      'photo attach target is missing.',
      false
    );
  }
  const fileId = String(target.row[6] || '');
  const imageUrl = String(target.row[7] || '');
  const receiptMatchesPhoto = !!receipt
    && !!String(receipt.fileId || '')
    && fileId === String(receipt.fileId || '')
    && imageUrl === String(receipt.imageUrl || '');
  if (fileId || imageUrl) {
    if (receiptMatchesPhoto) {
      return { target: target, attachedByReceipt: true };
    }
    throw importItemError_(
      'PIN_PHOTO_ATTACH_ALREADY_HAS_PHOTO',
      'photo attach target already has a photo.',
      false
    );
  }
  if (String(target.row[13] || '') !== String(normalized.expectedUpdatedAt || '')) {
    throw importItemError_(
      'PIN_PHOTO_ATTACH_CONFLICT',
      'photo attach target was updated.',
      false
    );
  }
  return { target: target, attachedByReceipt: false };
}

function claimImportPhotoAttachReceipt_(receiptSheet, mapSheet, normalized, payloadHash, leaseOwner) {
  return withImportReceiptLock_(function() {
    const now = getImportNow_();
    let receipt = findImportReceipt_(receiptSheet, normalized.idempotencyKeyHash);
    const recoverableReceipt = assertDriveSourceAvailableForReceipt_(
      receiptSheet,
      mapSheet,
      normalized,
      now
    );
    if (!receipt && recoverableReceipt) {
      if (String(recoverableReceipt.pinId || '') !== normalized.targetPinId
          || !isImportPhotoAttachTempFileName_(recoverableReceipt.tempFileName)) {
        throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'Drive source belongs to another item.', true);
      }
      receipt = recoverableReceipt;
      receipt.idempotencyKey = normalized.idempotencyKeyHash;
      receipt.jobId = normalized.jobId;
      receipt.itemId = normalized.itemId;
      receipt.payloadHash = payloadHash;
      receipt.targetFolderId = normalized.targetFolderId;
      receipt.sourceDriveFileId = normalized.sourceDriveFileId;
    } else if (receipt && recoverableReceipt
        && recoverableReceipt.rowNumber !== receipt.rowNumber) {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'multiple photo attach receipts exist.', false);
    }
    if (receipt) {
      if (String(receipt.jobId) !== normalized.jobId || String(receipt.itemId) !== normalized.itemId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt ids mismatch.', false);
      }
      if (String(receipt.payloadHash) !== payloadHash) {
        throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'payload conflict.', false);
      }
      if (String(receipt.pinId || '') !== normalized.targetPinId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'photo attach target changed.', false);
      }
      assertImportReceiptStorage_(receipt, {
        targetFolderId: normalized.targetFolderId,
        tempFileName: String(receipt.tempFileName || ''),
        sourceDriveFileId: normalized.sourceDriveFileId
      });
      if (!isImportPhotoAttachTempFileName_(receipt.tempFileName)) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'photo attach temp file is invalid.', false);
      }
    }
    const targetState = photoAttachTarget_(mapSheet, normalized, receipt);
    normalized.title = decodeSpreadsheetLiteral_(targetState.target.row[1]);
    if (targetState.attachedByReceipt) {
      let updatedAt = String(targetState.target.row[13] || '');
      if (updatedAt === normalized.expectedUpdatedAt) {
        updatedAt = currentUpdatedAt_();
        try {
          mapSheet.getRange(targetState.target.rowNumber, 14, 1, 1).setValue(updatedAt);
          targetState.target.row[13] = updatedAt;
        } catch (error) {
          logDrivePhotoImportFailure_('pin-photo-attach-map', error);
          throw importItemError_(
            'PIN_PHOTO_ATTACH_MAP_UPDATE_FAILED',
            'photo attach timestamp repair failed.',
            true
          );
        }
      }
      receipt.state = IMPORT_RECEIPT_STATES.COMPLETED;
      receipt.leaseOwner = '';
      receipt.leaseUntil = '';
      receipt.updatedAt = now.toISOString();
      receipt.lastErrorCode = '';
      writeImportReceipt_(receiptSheet, receipt);
      return {
        completed: true,
        receipt: receipt,
        payloadHash: payloadHash,
        pin: mapInfoRowToPinResult_(targetState.target.row, receipt.folderUrl)
      };
    }
    if (receipt && receipt.state === IMPORT_RECEIPT_STATES.COMPLETED) {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'completed attachment is missing.', false);
    }
    if (receipt && importLeaseIsActive_(receipt, now)
        && String(receipt.leaseOwner || '') !== leaseOwner) {
      throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'item is in progress.', true);
    }
    if (!receipt) {
      receipt = {
        rowNumber: 0,
        idempotencyKey: normalized.idempotencyKeyHash,
        jobId: normalized.jobId,
        itemId: normalized.itemId,
        payloadHash: payloadHash,
        state: IMPORT_RECEIPT_STATES.RESERVED,
        leaseOwner: '', leaseUntil: '',
        pinId: normalized.targetPinId,
        targetFolderId: normalized.targetFolderId,
        tempFileName: importPhotoAttachTempFileName_(normalized.idempotencyKeyHash),
        fileId: '', imageUrl: '', folderUrl: '',
        createdAt: now.toISOString(), updatedAt: now.toISOString(), lastErrorCode: '',
        sourceDriveFileId: normalized.sourceDriveFileId,
        mediaKind: 'photo', operationMode: normalized.operationMode,
        targetPinId: normalized.targetPinId, cleanupFileId: ''
      };
    }
    receipt.state = receipt.fileId ? IMPORT_RECEIPT_STATES.FILE_SAVED : IMPORT_RECEIPT_STATES.RESERVED;
    receipt.leaseOwner = leaseOwner;
    receipt.leaseUntil = new Date(now.getTime() + IMPORT_ITEM_LEASE_MS).toISOString();
    receipt.updatedAt = now.toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return {
      completed: false,
      receipt: receipt,
      payloadHash: payloadHash,
      target: targetState.target
    };
  });
}

function assertPhotoAttachTargetReady_(receiptSheet, mapSheet, expected, normalized) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    assertImportReceiptOwner_(receipt, expected);
    const targetState = photoAttachTarget_(mapSheet, normalized, null);
    normalized.title = decodeSpreadsheetLiteral_(targetState.target.row[1]);
    return targetState.target;
  });
}

function finalizeImportPhotoAttach_(receiptSheet, mapSheet, expected, normalized) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    assertImportReceiptOwner_(receipt, expected);
    const targetState = photoAttachTarget_(mapSheet, normalized, receipt);
    const target = targetState.target;
    let updatedAt = String(target.row[13] || '');
    if (!targetState.attachedByReceipt) {
      updatedAt = currentUpdatedAt_();
      try {
        mapSheet.getRange(target.rowNumber, 7, 1, 2).setValues([[
          String(receipt.fileId || ''),
          String(receipt.imageUrl || '')
        ]]);
        target.row[6] = String(receipt.fileId || '');
        target.row[7] = String(receipt.imageUrl || '');
        mapSheet.getRange(target.rowNumber, 14, 1, 1).setValue(updatedAt);
        target.row[13] = updatedAt;
      } catch (error) {
        logDrivePhotoImportFailure_('pin-photo-attach-map', error);
        throw importItemError_(
          'PIN_PHOTO_ATTACH_MAP_UPDATE_FAILED',
          'photo attach map update failed.',
          true
        );
      }
    } else if (updatedAt === normalized.expectedUpdatedAt) {
      updatedAt = currentUpdatedAt_();
      try {
        mapSheet.getRange(target.rowNumber, 14, 1, 1).setValue(updatedAt);
        target.row[13] = updatedAt;
      } catch (error) {
        logDrivePhotoImportFailure_('pin-photo-attach-map', error);
        throw importItemError_(
          'PIN_PHOTO_ATTACH_MAP_UPDATE_FAILED',
          'photo attach timestamp update failed.',
          true
        );
      }
    }
    receipt.state = IMPORT_RECEIPT_STATES.COMPLETED;
    receipt.leaseOwner = '';
    receipt.leaseUntil = '';
    receipt.updatedAt = getImportNow_().toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return mapInfoRowToPinResult_(target.row, receipt.folderUrl);
  });
}

function resolveImportDriveFile_(receipt, normalized) {
  let file = null;
  let folder = null;
  let created = false;
  if (receipt.fileId) {
    try {
      file = DriveApp.getFileById(String(receipt.fileId));
      const parents = file.getParents();
      let belongsToTarget = false;
      while (parents.hasNext()) {
        if (String(parents.next().getId()) === String(receipt.targetFolderId)) {
          belongsToTarget = true;
          break;
        }
      }
      if (!belongsToTarget) {
        throw importItemError_(
          'DRIVE_MANAGED_COPY_FINALIZE_FAILED',
          'saved managed file parent changed.',
          true
        );
      }
    } catch (error) {
      if (isImportItemError_(error)) throw error;
      logDrivePhotoImportFailure_('managed-copy-finalize', error);
      throw importItemError_(
        'DRIVE_MANAGED_COPY_FINALIZE_FAILED',
        'saved managed file is unavailable.',
        !isDriveProviderPermissionDenied_(error)
      );
    }
  } else {
    try {
      folder = DriveApp.getFolderById(String(receipt.targetFolderId));
      const iterator = folder.getFilesByName(String(receipt.tempFileName));
      const matches = [];
      while (iterator.hasNext() && matches.length < 2) matches.push(iterator.next());
      if (matches.length > 1) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'multiple temporary files.', false);
      }
      if (matches.length === 1) {
        file = matches[0];
      } else {
        const blob = Utilities.newBlob(
          normalized.jpegBytes,
          'image/jpeg',
          String(receipt.tempFileName)
        );
        file = folder.createFile(blob);
        created = true;
      }
    } catch (error) {
      if (isImportItemError_(error)) throw error;
      logDrivePhotoImportFailure_('managed-copy-create', error);
      throw importItemError_(
        'DRIVE_MANAGED_COPY_CREATE_FAILED',
        'managed file operation failed.',
        !isDriveProviderPermissionDenied_(error)
      );
    }
  }
  const fileId = String(file.getId());
  if (normalized.sourceDriveFileId && fileId === String(normalized.sourceDriveFileId)) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'managed file matches Drive source.', false);
  }
  return {
    file: file,
    created: created,
    managed: true,
    fileId: fileId,
    imageUrl: 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1920',
    folderUrl: getDriveFolderUrl_(receipt.targetFolderId),
    parentFolderId: String(receipt.targetFolderId)
  };
}

function driveSourceProviderError_(stage, error) {
  logDrivePhotoImportFailure_(stage, error);
  const permissionDenied = isDriveProviderPermissionDenied_(error);
  return importItemError_(
    permissionDenied ? 'DRIVE_SOURCE_NOT_EDITABLE' : 'DRIVE_SOURCE_CHECK_FAILED',
    permissionDenied ? 'Drive source access was denied.' : 'Drive source check failed.',
    !permissionDenied
  );
}

function isDriveItemWithinRootForImport_(item, rootFolderId, allowRootItem) {
  if (!item || !isValidDrivePhotoImportId_(rootFolderId) || item.isTrashed()) return false;
  const initialId = String(item.getId());
  if (!isValidDrivePhotoImportId_(initialId)) return false;
  if (allowRootItem && initialId === rootFolderId) return true;

  const queue = [{ item: item, depth: 0 }];
  const visited = Object.create(null);
  let cursor = 0;
  let inspected = 0;
  let reachedRoot = false;
  while (cursor < queue.length) {
    if (inspected >= DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_NODES) return false;
    const entry = queue[cursor++];
    if (entry.item.isTrashed()) continue;
    const currentId = String(entry.item.getId());
    if (!isValidDrivePhotoImportId_(currentId)
        || drivePhotoImportHasOwn_(visited, currentId)) continue;
    visited[currentId] = true;
    inspected += 1;
    if (currentId === rootFolderId) {
      reachedRoot = true;
      continue;
    }
    if (entry.depth >= DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_DEPTH) return false;

    const parents = entry.item.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      if (parent.isTrashed()) continue;
      const parentId = String(parent.getId());
      if (!isValidDrivePhotoImportId_(parentId)) return false;
      if (parentId === rootFolderId) {
        reachedRoot = true;
        continue;
      }
      if (!drivePhotoImportHasOwn_(visited, parentId)) {
        queue.push({ item: parent, depth: entry.depth + 1 });
      }
      if (queue.length > DRIVE_PHOTO_IMPORT_MAX_ANCESTRY_NODES) return false;
    }
  }
  return reachedRoot;
}

function getDriveSourceParentForImport_(file, rootFolderId) {
  try {
    const parents = file.getParents();
    const directParents = [];
    while (parents.hasNext()) {
      const parent = parents.next();
      const parentId = String(parent.getId());
      if (!isValidDrivePhotoImportId_(parentId) || parent.isTrashed()) {
        throw importItemError_('IMPORT_DRIVE_SOURCE_INVALID', 'Drive source parent is invalid.', false);
      }
      directParents.push({ folder: parent, folderId: parentId });
    }
    if (directParents.length === 1 && directParents[0].folderId === String(rootFolderId)) {
      return directParents[0];
    }
  } catch (error) {
    if (isImportItemError_(error)) throw error;
    throw driveSourceProviderError_('source-validation', error);
  }
  throw importItemError_('IMPORT_DRIVE_SOURCE_INVALID', 'Drive source parent is invalid.', false);
}

function validateDriveSourceForImport_(sourceDriveFileId) {
  let rootFolderId;
  let file;
  try {
    rootFolderId = getDrivePhotoImportRootId_();
    file = DriveApp.getFileById(String(sourceDriveFileId));
  } catch (error) {
    if (error && error.code === 'DRIVE_IMPORT_ROOT_MISSING') {
      throw importItemError_('IMPORT_DRIVE_SOURCE_INVALID', 'Drive import root is unavailable.', false);
    }
    throw driveSourceProviderError_('source-validation', error);
  }
  let id;
  let name;
  let mimeType;
  let sizeBytes;
  let parent;
  let trashed;
  try {
    id = String(file.getId());
    name = String(file.getName());
    mimeType = String(file.getMimeType());
    sizeBytes = Number(file.getSize());
    trashed = file.isTrashed();
    parent = getDriveSourceParentForImport_(file, rootFolderId);
  } catch (error) {
    if (isImportItemError_(error)) throw error;
    throw driveSourceProviderError_('source-validation', error);
  }
  const classification = classifyDrivePhotoImportFile_({
    name: name,
    type: mimeType,
    size: sizeBytes
  });
  if (trashed
      || mimeType === DRIVE_PHOTO_IMPORT_SHORTCUT_MIME
      || id !== String(sourceDriveFileId)
      || !isValidDrivePhotoImportFileName_(name)
      || !classification.supported
      || !Number.isSafeInteger(sizeBytes)
      || sizeBytes <= 0
      || sizeBytes > DRIVE_PHOTO_IMPORT_MAX_FILE_BYTES) {
    throw importItemError_('IMPORT_DRIVE_SOURCE_INVALID', 'Drive source metadata is invalid.', false);
  }
  return {
    file: file,
    fileId: id,
    name: name,
    kind: classification.kind,
    parentFolder: parent.folder,
    parentFolderId: parent.folderId,
    rootFolderId: rootFolderId
  };
}

function validateUnownedDriveSourceForImport_(sourceDriveFileId) {
  try {
    const associations = getMediaDriveAssociations_();
    const sourceId = String(sourceDriveFileId || '');
    if (associations.managedPhotoIds.has(sourceId)
        || associations.completedSourceIds.has(sourceId)) {
      throw importItemError_(
        'DRIVE_SOURCE_ALREADY_LINKED',
        'Drive source is already managed or completed.',
        false
      );
    }
  } catch (error) {
    if (isImportItemError_(error)) throw error;
    throw driveSourceProviderError_('source-association', error);
  }
  return validateDriveSourceForImport_(sourceDriveFileId);
}

function driveItemDirectParentIds_(item) {
  const ids = [];
  const seen = Object.create(null);
  const parents = item.getParents();
  while (parents.hasNext()) {
    const id = String(parents.next().getId());
    if (!isValidDrivePhotoImportId_(id) || drivePhotoImportHasOwn_(seen, id)) continue;
    seen[id] = true;
    ids.push(id);
  }
  return ids;
}

function resolveDriveOriginalFolder_(rootFolderId) {
  try {
    const rootFolder = DriveApp.getFolderById(String(rootFolderId));
    const folders = rootFolder.getFolders();
    const matches = [];
    while (folders.hasNext()) {
      const folder = folders.next();
      if (folder.isTrashed() || String(folder.getName()) !== 'original') continue;
      const folderId = String(folder.getId());
      if (!isValidDrivePhotoImportId_(folderId)
          || driveItemDirectParentIds_(folder).indexOf(String(rootFolderId)) === -1) continue;
      matches.push(folder);
      if (matches.length > 1) {
        throw importItemError_(
          'DRIVE_ORIGINAL_FOLDER_AMBIGUOUS',
          'multiple original folders exist.',
          false
        );
      }
    }
    if (matches.length === 1) return matches[0];

    const created = rootFolder.createFolder('original');
    if (!created || created.isTrashed() || String(created.getName()) !== 'original'
        || driveItemDirectParentIds_(created).indexOf(String(rootFolderId)) === -1) {
      throw new Error('created original folder could not be verified');
    }
    return created;
  } catch (error) {
    if (isImportItemError_(error)) throw error;
    logDrivePhotoImportFailure_('original-folder', error);
    throw importItemError_(
      'DRIVE_ORIGINAL_FOLDER_CREATE_FAILED',
      'original folder resolution failed.',
      true
    );
  }
}

function moveDriveSourceToOriginal_(source, originalFolder) {
  if (!source || !source.file || !originalFolder) {
    throw importItemError_('DRIVE_SOURCE_MOVE_FAILED', 'Drive source move input is invalid.', true);
  }
  const sourceFile = source.file;
  const sourceFileId = String(source.fileId || '');
  const originalFolderId = String(originalFolder.getId());
  let parentsBefore;
  try {
    parentsBefore = driveItemDirectParentIds_(sourceFile);
  } catch (error) {
    logDrivePhotoImportFailure_('source-move-check', error);
    throw importItemError_(
      'DRIVE_SOURCE_MOVE_VERIFY_FAILED',
      'Drive source parents could not be checked.',
      true
    );
  }
  if (parentsBefore.indexOf(originalFolderId) === -1) {
    try {
      sourceFile.moveTo(originalFolder);
    } catch (error) {
      logDrivePhotoImportFailure_('source-move', error);
      throw importItemError_(
        'DRIVE_SOURCE_MOVE_FAILED',
        'Drive source move failed.',
        true
      );
    }
  }
  try {
    if (sourceFile.isTrashed() || String(sourceFile.getId()) !== sourceFileId
        || driveItemDirectParentIds_(sourceFile).indexOf(originalFolderId) === -1) {
      throw new Error('Drive source parent did not change');
    }
  } catch (error) {
    logDrivePhotoImportFailure_('source-move-verify', error);
    throw importItemError_(
      'DRIVE_SOURCE_MOVE_VERIFY_FAILED',
      'Drive source move verification failed.',
      true
    );
  }
  return { moved: parentsBefore.indexOf(originalFolderId) === -1 };
}

function archiveCompletedDriveSourceIfPending_(sourceDriveFileId, mediaStructure) {
  const sourceId = String(sourceDriveFileId || '');
  const rootFolderId = String(mediaStructure && mediaStructure.root || '');
  const originalPhotosId = String(mediaStructure && mediaStructure.originalPhotos || '');
  if (!isValidDrivePhotoImportId_(sourceId)
      || !isValidDrivePhotoImportId_(rootFolderId)
      || !isValidDrivePhotoImportId_(originalPhotosId)) {
    throw importItemError_('IMPORT_DRIVE_SOURCE_INVALID', 'completed Drive source archive input is invalid.', false);
  }
  let file;
  let parentIds;
  try {
    file = DriveApp.getFileById(sourceId);
    if (!file || file.isTrashed() || String(file.getId()) !== sourceId) return { moved: false };
    parentIds = driveItemDirectParentIds_(file);
  } catch (error) {
    if (sanitizeProviderMessage_(importProviderErrorField_(error, 'message')) === 'resource_unavailable') {
      return { moved: false };
    }
    throw driveSourceProviderError_('source-archive-replay', error);
  }
  if (parentIds.length === 1 && parentIds[0] === originalPhotosId) return { moved: false };
  if (parentIds.length !== 1 || parentIds[0] !== rootFolderId) return { moved: false };
  const source = validateDriveSourceForImport_(sourceId);
  return moveDriveSourceToOriginal_(
    source,
    DriveApp.getFolderById(originalPhotosId)
  );
}

function finalizeImportDisplayFile_(driveResult, normalized) {
  if (!driveResult || !driveResult.managed) return driveResult;
  let sharingWasConfirmedDenied = false;
  try {
    let access = driveResult.file.getSharingAccess();
    let hasAnonymousViewAccess = (
      access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK
    ) && driveResult.file.getSharingPermission() === DriveApp.Permission.VIEW;
    if (!hasAnonymousViewAccess) {
      driveResult.file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      access = driveResult.file.getSharingAccess();
      hasAnonymousViewAccess = (
        access === DriveApp.Access.ANYONE || access === DriveApp.Access.ANYONE_WITH_LINK
      ) && driveResult.file.getSharingPermission() === DriveApp.Permission.VIEW;
    }
    if (!hasAnonymousViewAccess) {
      sharingWasConfirmedDenied = true;
      throw new Error('anonymous view sharing was not applied');
    }
  } catch (error) {
    logDrivePhotoImportFailure_('sharing', error);
    const sharingWasDenied = sharingWasConfirmedDenied || isDriveProviderPermissionDenied_(error);
    throw importItemError_(
      sharingWasDenied ? 'DRIVE_LINK_SHARING_DENIED' : 'DRIVE_LINK_SHARING_FAILED',
      sharingWasDenied
        ? 'managed JPEG sharing denied.' : 'managed JPEG sharing verification failed.',
      !sharingWasDenied
    );
  }

  return driveResult;
}

function renameManagedImportDisplayFile_(driveResult, normalized) {
  if (!driveResult || !driveResult.managed) return driveResult;
  try {
    const finalName = PinData.buildFileNameForSave(
      normalized.title,
      normalized.filename,
      getRenameFileWithTitle_()
    );
    driveResult.file.setName(finalName);
  } catch (error) {
    logDrivePhotoImportFailure_('managed-copy-finalize', error);
    throw importItemError_(
      'DRIVE_MANAGED_COPY_FINALIZE_FAILED',
      'managed JPEG finalize failed.',
      !isDriveProviderPermissionDenied_(error)
    );
  }
  return driveResult;
}

function compensateUnpersistedImportDriveFile_(receiptSheet, expected, driveResult) {
  if (!driveResult || !driveResult.created || !driveResult.file) return false;
  try {
    return withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash) !== expected.payloadHash
          || String(receipt.pinId) !== expected.pinId
          || String(receipt.leaseOwner) !== expected.leaseOwner
          || String(receipt.state) !== IMPORT_RECEIPT_STATES.RESERVED
          || String(receipt.fileId || '')) return false;
      if (String(receipt.targetFolderId) !== expected.targetFolderId
          || String(receipt.tempFileName) !== expected.tempFileName) return false;

      const file = driveResult.file;
      if (String(file.getId()) !== String(driveResult.fileId)
          || String(file.getName()) !== expected.tempFileName
          || !driveResult.managed) return false;
      const parents = file.getParents();
      let belongsToTarget = false;
      while (parents.hasNext()) {
        if (String(parents.next().getId()) === String(driveResult.parentFolderId || expected.targetFolderId)) {
          belongsToTarget = true;
          break;
        }
      }
      if (!belongsToTarget) return false;
      file.setTrashed(true);
      return true;
    });
  } catch (error) {
    logDrivePhotoImportFailure_('managed-copy-compensation', error);
    return false;
  }
}

function compensatePersistedImportDriveFile_(receiptSheet, expected, driveResult, errorCode) {
  if (!driveResult || !driveResult.file || !driveResult.managed) return false;
  let owned = false;
  try {
    owned = withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash) !== expected.payloadHash
          || String(receipt.pinId) !== expected.pinId
          || String(receipt.leaseOwner) !== expected.leaseOwner
          || String(receipt.state) !== IMPORT_RECEIPT_STATES.FILE_SAVED
          || String(receipt.fileId || '') !== String(driveResult.fileId)) return false;
      return true;
    });
  } catch (_error) {
    return false;
  }
  if (!owned) return false;
  try {
    driveResult.file.setTrashed(true);
  } catch (error) {
    logDrivePhotoImportFailure_('managed-copy-compensation', error);
    return false;
  }
  try {
    return withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash) !== expected.payloadHash
          || String(receipt.pinId) !== expected.pinId
          || String(receipt.leaseOwner) !== expected.leaseOwner
          || String(receipt.state) !== IMPORT_RECEIPT_STATES.FILE_SAVED
          || String(receipt.fileId || '') !== String(driveResult.fileId)) return false;
      receipt.state = IMPORT_RECEIPT_STATES.FAILED;
      receipt.leaseOwner = '';
      receipt.leaseUntil = '';
      receipt.fileId = '';
      receipt.imageUrl = '';
      receipt.folderUrl = '';
      receipt.updatedAt = getImportNow_().toISOString();
      receipt.lastErrorCode = String(errorCode || 'IMPORT_ITEM_SAVE_FAILED');
      writeImportReceipt_(receiptSheet, receipt);
      return true;
    });
  } catch (error) {
    logDrivePhotoImportFailure_('managed-copy-compensation-journal', error);
    return false;
  }
}

function persistImportDriveResult_(receiptSheet, expected, driveResult) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    assertImportReceiptOwner_(receipt, expected);
    const now = getImportNow_();
    receipt.state = IMPORT_RECEIPT_STATES.FILE_SAVED;
    receipt.fileId = driveResult.fileId;
    receipt.imageUrl = driveResult.imageUrl;
    receipt.folderUrl = driveResult.folderUrl;
    receipt.updatedAt = now.toISOString();
    receipt.leaseUntil = new Date(now.getTime() + IMPORT_ITEM_LEASE_MS).toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return receipt;
  });
}

function finalizeImportItem_(receiptSheet, mapSheet, expected, normalized) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    assertImportReceiptOwner_(receipt, expected);
    let existingMap = findMapInfoRowByPinId_(mapSheet, expected.pinId);
    if (!existingMap) {
      const row = buildMapInfoRow_(normalized, {
        timestamp: formatMapTimestamp_(getImportNow_()),
        pinId: expected.pinId,
        fileId: receipt.fileId,
        imageUrl: receipt.imageUrl,
        updatedAt: ''
      });
      try {
        appendMapInfoRow_(mapSheet, row);
      } catch (error) {
        logDrivePhotoImportFailure_('map-row', error);
        throw importItemError_('IMPORT_MAP_ROW_FAILED', 'map row append failed.', true);
      }
      existingMap = { rowNumber: mapSheet.getLastRow(), row: row };
    }
    receipt.state = IMPORT_RECEIPT_STATES.COMPLETED;
    receipt.leaseOwner = '';
    receipt.leaseUntil = '';
    receipt.updatedAt = getImportNow_().toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return mapInfoRowToPinResult_(existingMap.row, receipt.folderUrl);
  });
}

function markImportReceiptFailed_(receiptSheet, expected, errorCode) {
  try {
    withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash) !== expected.payloadHash
          || String(receipt.pinId) !== expected.pinId
          || String(receipt.leaseOwner) !== expected.leaseOwner) return;
      receipt.state = IMPORT_RECEIPT_STATES.FAILED;
      receipt.leaseOwner = '';
      receipt.leaseUntil = '';
      receipt.updatedAt = getImportNow_().toISOString();
      receipt.lastErrorCode = String(errorCode || 'IMPORT_ITEM_SAVE_FAILED');
      writeImportReceipt_(receiptSheet, receipt);
    });
  } catch (_error) {
    // The original safe error remains the response when failure journaling is unavailable.
  }
}

function isImportSourceMoveFailureCode_(errorCode) {
  return errorCode === 'DRIVE_SOURCE_MOVE_FAILED'
    || errorCode === 'DRIVE_SOURCE_MOVE_VERIFY_FAILED'
    || errorCode === 'PIN_PHOTO_ATTACH_SOURCE_MOVE_FAILED';
}

function markCompletedImportSourceMoveFailed_(receiptSheet, expected, errorCode) {
  if (!isImportSourceMoveFailureCode_(errorCode)) return;
  try {
    withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash) !== expected.payloadHash
          || String(receipt.pinId) !== expected.pinId
          || String(receipt.state) !== IMPORT_RECEIPT_STATES.COMPLETED
          || String(receipt.sourceDriveFileId || '') !== expected.sourceDriveFileId) return;
      receipt.updatedAt = getImportNow_().toISOString();
      receipt.lastErrorCode = String(errorCode);
      writeImportReceipt_(receiptSheet, receipt);
    });
  } catch (_error) {
    // The completed pin remains authoritative if cleanup journaling is unavailable.
  }
}

function claimImportPinReceipt_(receiptSheet, mapSheet, normalized, payloadHash, leaseOwner) {
  return withImportReceiptLock_(function() {
    const now = getImportNow_();
    let receipt = findImportReceipt_(receiptSheet, normalized.idempotencyKeyHash);
    const wasCompleted = !!receipt && receipt.state === IMPORT_RECEIPT_STATES.COMPLETED;
    if (receipt) {
      if (String(receipt.jobId) !== normalized.jobId || String(receipt.itemId) !== normalized.itemId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt ids mismatch.', false);
      }
      if (String(receipt.payloadHash) !== payloadHash) {
        throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'payload conflict.', false);
      }
      if (!receipt.pinId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt pin is missing.', false);
      }
      const validStates = [
        IMPORT_RECEIPT_STATES.RESERVED,
        IMPORT_RECEIPT_STATES.FAILED,
        IMPORT_RECEIPT_STATES.COMPLETED
      ];
      if (validStates.indexOf(String(receipt.state)) === -1) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'photo-less receipt state is invalid.', false);
      }
      assertImportPinReceiptStorage_(receipt);
      const existingMap = findMapInfoRowByPinId_(mapSheet, String(receipt.pinId));
      if (existingMap) {
        assertImportPinMapRow_(existingMap);
        const needsReceiptRepair = !wasCompleted
          || String(receipt.leaseOwner || '') !== ''
          || String(receipt.leaseUntil || '') !== ''
          || String(receipt.lastErrorCode || '') !== '';
        if (needsReceiptRepair) {
          receipt.state = IMPORT_RECEIPT_STATES.COMPLETED;
          receipt.leaseOwner = '';
          receipt.leaseUntil = '';
          receipt.updatedAt = now.toISOString();
          receipt.lastErrorCode = '';
          writeImportReceipt_(receiptSheet, receipt);
        }
        return {
          completed: true,
          deduplicated: true,
          receipt: receipt,
          pin: mapInfoRowToPinResult_(existingMap.row, '')
        };
      }
      if (importLeaseIsActive_(receipt, now) && String(receipt.leaseOwner) !== leaseOwner) {
        throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'item is in progress.', true);
      }
    } else {
      receipt = {
        rowNumber: 0,
        idempotencyKey: normalized.idempotencyKeyHash,
        jobId: normalized.jobId,
        itemId: normalized.itemId,
        payloadHash: payloadHash,
        state: IMPORT_RECEIPT_STATES.RESERVED,
        leaseOwner: '',
        leaseUntil: '',
        pinId: Utilities.getUuid(),
        targetFolderId: '', tempFileName: '', fileId: '', imageUrl: '', folderUrl: '',
        createdAt: now.toISOString(), updatedAt: now.toISOString(), lastErrorCode: '',
        sourceDriveFileId: '',
        mediaKind: 'photo', operationMode: 'create-pin', targetPinId: '', cleanupFileId: ''
      };
    }
    receipt.state = IMPORT_RECEIPT_STATES.RESERVED;
    receipt.leaseOwner = leaseOwner;
    receipt.leaseUntil = new Date(now.getTime() + IMPORT_ITEM_LEASE_MS).toISOString();
    receipt.updatedAt = now.toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return { completed: false, deduplicated: wasCompleted, receipt: receipt };
  });
}

function assertImportPinReceiptOwner_(receipt, expected) {
  if (!receipt || !receipt.pinId) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt pin is missing.', false);
  }
  if (String(receipt.payloadHash) !== expected.payloadHash) {
    throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'payload conflict.', false);
  }
  if (String(receipt.pinId) !== expected.pinId) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt pin changed.', false);
  }
  assertImportPinReceiptStorage_(receipt);
  if (String(receipt.leaseOwner) !== expected.leaseOwner) {
    throw importItemError_('IMPORT_ITEM_LEASE_LOST', 'lease owner changed.', true);
  }
}

function finalizeImportPinItem_(receiptSheet, mapSheet, expected, normalized) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    assertImportPinReceiptOwner_(receipt, expected);
    let existingMap = findMapInfoRowByPinId_(mapSheet, expected.pinId);
    assertImportPinMapRow_(existingMap);
    if (!existingMap) {
      const row = buildMapInfoRow_(normalized, {
        timestamp: formatMapTimestamp_(getImportNow_()),
        pinId: expected.pinId,
        fileId: '', imageUrl: '', updatedAt: ''
      });
      try {
        appendMapInfoRow_(mapSheet, row);
      } catch (_error) {
        throw importItemError_('IMPORT_MAP_ROW_FAILED', 'map row append failed.', true);
      }
      existingMap = { rowNumber: mapSheet.getLastRow(), row: row };
    }
    receipt.state = IMPORT_RECEIPT_STATES.COMPLETED;
    receipt.leaseOwner = '';
    receipt.leaseUntil = '';
    receipt.updatedAt = getImportNow_().toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return mapInfoRowToPinResult_(existingMap.row, '');
  });
}

function saveImportPinItem(data) {
  assertEditToken_(data);
  let normalized;
  let receiptSheet;
  let expected = null;
  try {
    normalized = normalizeImportPinPayload_(data);
    const payloadHash = hashImportPinPayload_(normalized);
    receiptSheet = openImportReceiptsSheet_();
    const mapSheet = openMapInfoSheet_();
    const leaseOwner = Utilities.getUuid();
    const claim = claimImportPinReceipt_(receiptSheet, mapSheet, normalized, payloadHash, leaseOwner);
    if (claim.completed) {
      return { ok: true, deduplicated: true, pin: claim.pin };
    }
    expected = {
      idempotencyKeyHash: normalized.idempotencyKeyHash,
      payloadHash: payloadHash,
      pinId: String(claim.receipt.pinId),
      leaseOwner: leaseOwner
    };
    const pin = finalizeImportPinItem_(receiptSheet, mapSheet, expected, normalized);
    return { ok: true, deduplicated: claim.deduplicated === true, pin: pin };
  } catch (error) {
    const code = isImportItemError_(error) ? String(error.code) : 'IMPORT_ITEM_SAVE_FAILED';
    if (receiptSheet && expected && code !== 'IMPORT_ITEM_LEASE_LOST') {
      markImportReceiptFailed_(receiptSheet, expected, code);
    }
    return importItemFailureFromError_(error);
  }
}

function saveImportPhotoItem(data) {
  assertEditToken_(data);
  let normalized;
  let receiptSheet;
  let expected = null;
  let unpersistedDriveResult = null;
  let persistedDriveResult = null;
  let sourceMoveCompleted = false;
  let pinLinked = false;
  let finalized = false;
  try {
    normalized = normalizeImportPhotoPayload_(data);
    const preflightReceipt = inspectExistingImportReceiptForSave_(
      normalized.idempotencyKeyHash
    );
    let preflightDriveSource = null;
    if (normalized.sourceDriveFileId
        && (!preflightReceipt || !String(preflightReceipt.fileId || ''))) {
      preflightDriveSource = validateUnownedDriveSourceForImport_(
        normalized.sourceDriveFileId
      );
    }
    receiptSheet = ensureImportReceiptSchemaForSave_();
    const existingReceipt = findImportReceipt_(receiptSheet, normalized.idempotencyKeyHash);
    let mediaStructure;
    try {
      mediaStructure = ensureMediaDriveStructure_();
    } catch (error) {
      const structureCode = error && typeof error.code === 'string'
        ? String(error.code) : 'DRIVE_MEDIA_STRUCTURE_FAILED';
      const knownStructureCodes = {
        DRIVE_MEDIA_STRUCTURE_AMBIGUOUS: true,
        DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT: true,
        DRIVE_MEDIA_STRUCTURE_PARENT_INVALID: true,
        DRIVE_MEDIA_STRUCTURE_FAILED: true
      };
      throw importItemError_(
        knownStructureCodes[structureCode] ? structureCode : 'DRIVE_MEDIA_STRUCTURE_FAILED',
        'media Drive structure is unavailable.',
        structureCode !== 'DRIVE_MEDIA_STRUCTURE_AMBIGUOUS'
          && structureCode !== 'DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT'
      );
    }
    const managedPhotosFolderId = String(mediaStructure.photos || '');
    normalized.targetFolderId = managedPhotosFolderId;
    if (existingReceipt && String(existingReceipt.targetFolderId || '')) {
      const receiptTargetFolderId = String(existingReceipt.targetFolderId);
      if (receiptTargetFolderId !== managedPhotosFolderId) {
        const managedTargetPayloadHash = hashImportPayload_(normalized);
        normalized.targetFolderId = receiptTargetFolderId;
        if (String(existingReceipt.payloadHash || '') === managedTargetPayloadHash) {
          normalized.targetFolderId = managedPhotosFolderId;
        }
      }
    }
    if (!normalized.targetFolderId) {
      throw importItemError_('INVALID_IMPORT_PAYLOAD', 'target folder is missing.', false);
    }
    const mapSheet = openMapInfoSheet_();
    let driveSource = preflightDriveSource;
    if (normalized.sourceDriveFileId && !driveSource
        && (!existingReceipt || !String(existingReceipt.fileId || ''))) {
      driveSource = validateUnownedDriveSourceForImport_(normalized.sourceDriveFileId);
    }
    let payloadHash = hashImportPayload_(normalized);
    const leaseOwner = Utilities.getUuid();
    const claim = normalized.operationMode === 'attach-existing-pin'
      ? claimImportPhotoAttachReceipt_(
        receiptSheet, mapSheet, normalized, payloadHash, leaseOwner
      )
      : claimImportReceipt_(receiptSheet, mapSheet, normalized, payloadHash, leaseOwner);
    payloadHash = String(claim.payloadHash || payloadHash);
    if (claim.completed) {
      if (normalized.sourceDriveFileId) {
        const archiveResult = archiveCompletedDriveSourceIfPending_(
          normalized.sourceDriveFileId,
          mediaStructure
        );
        sourceMoveCompleted = archiveResult.moved === true;
      }
      return { ok: true, deduplicated: true, pin: claim.pin };
    }
    if (normalized.sourceDriveFileId && !driveSource
        && !String(claim.receipt.fileId || '')) {
      driveSource = validateUnownedDriveSourceForImport_(normalized.sourceDriveFileId);
    }
    expected = {
      idempotencyKeyHash: normalized.idempotencyKeyHash,
      payloadHash: payloadHash,
      pinId: String(claim.receipt.pinId),
      leaseOwner: leaseOwner,
      targetFolderId: normalized.targetFolderId,
      tempFileName: String(claim.receipt.tempFileName || ''),
      sourceDriveFileId: normalized.sourceDriveFileId
    };
    const driveResult = resolveImportDriveFile_(claim.receipt, normalized);
    unpersistedDriveResult = driveResult;
    finalizeImportDisplayFile_(driveResult, normalized);
    const receipt = persistImportDriveResult_(receiptSheet, expected, driveResult);
    unpersistedDriveResult = null;
    persistedDriveResult = driveResult;
    renameManagedImportDisplayFile_(driveResult, normalized);
    expected.pinId = String(receipt.pinId);
    if (normalized.operationMode === 'attach-existing-pin') {
      assertPhotoAttachTargetReady_(receiptSheet, mapSheet, expected, normalized);
    }
    let pin;
    pin = normalized.operationMode === 'attach-existing-pin'
      ? finalizeImportPhotoAttach_(receiptSheet, mapSheet, expected, normalized)
      : finalizeImportItem_(receiptSheet, mapSheet, expected, normalized);
    pinLinked = true;
    if (normalized.sourceDriveFileId) {
      if (driveSource) {
        moveDriveSourceToOriginal_(
          driveSource,
          DriveApp.getFolderById(String(mediaStructure.originalPhotos))
        );
        sourceMoveCompleted = true;
      } else {
        const archiveResult = archiveCompletedDriveSourceIfPending_(
          normalized.sourceDriveFileId,
          mediaStructure
        );
        sourceMoveCompleted = archiveResult.moved === true;
      }
    }
    finalized = true;
    return { ok: true, deduplicated: false, pin: pin };
  } catch (error) {
    let mappedError = error;
    let code = isImportItemError_(mappedError)
      ? String(mappedError.code) : 'IMPORT_ITEM_SAVE_FAILED';
    if (normalized && normalized.operationMode === 'attach-existing-pin') {
      if (code === 'DRIVE_MANAGED_COPY_CREATE_FAILED') {
        mappedError = importItemError_(
          'PIN_PHOTO_ATTACH_FILE_CREATE_FAILED',
          'photo attach managed file creation failed.',
          !!(error && error.retryable)
        );
      } else if (code === 'DRIVE_SOURCE_MOVE_FAILED'
          || code === 'DRIVE_SOURCE_MOVE_VERIFY_FAILED') {
        mappedError = importItemError_(
          'PIN_PHOTO_ATTACH_SOURCE_MOVE_FAILED',
          'photo attach source move failed.',
          true
        );
      }
      code = isImportItemError_(mappedError)
        ? String(mappedError.code) : 'IMPORT_ITEM_SAVE_FAILED';
    }
    if (receiptSheet && expected && pinLinked && isImportSourceMoveFailureCode_(code)) {
      markCompletedImportSourceMoveFailed_(receiptSheet, expected, code);
    }
    if (receiptSheet && expected && unpersistedDriveResult) {
      compensateUnpersistedImportDriveFile_(receiptSheet, expected, unpersistedDriveResult);
    }
    const preMoveCompensationCodes = {
      DRIVE_ORIGINAL_FOLDER_CREATE_FAILED: true,
      DRIVE_ORIGINAL_FOLDER_AMBIGUOUS: true,
      DRIVE_SOURCE_MOVE_FAILED: true,
      DRIVE_SOURCE_MOVE_VERIFY_FAILED: true,
      PIN_PHOTO_ATTACH_TARGET_NOT_FOUND: true,
      PIN_PHOTO_ATTACH_ALREADY_HAS_PHOTO: true,
      PIN_PHOTO_ATTACH_CONFLICT: true,
      PIN_PHOTO_ATTACH_SOURCE_MOVE_FAILED: true
    };
    if (receiptSheet && expected && persistedDriveResult && !sourceMoveCompleted && !pinLinked
        && preMoveCompensationCodes[code]) {
      compensatePersistedImportDriveFile_(
        receiptSheet,
        expected,
        persistedDriveResult,
        code
      );
    }
    if (receiptSheet && expected && !finalized && code !== 'IMPORT_ITEM_LEASE_LOST') {
      markImportReceiptFailed_(receiptSheet, expected, code);
    }
    return importItemFailureFromError_(mappedError);
  }
}

function normalizeImportAudioPayload_(data) {
  const source = data || {};
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'audio payload is invalid.', false);
  }
  const jobId = normalizeImportIdentifier_(source.jobId, 'jobId', IMPORT_ITEM_ID_MAX_LENGTH);
  const itemId = normalizeImportIdentifier_(source.itemId, 'itemId', IMPORT_ITEM_ID_MAX_LENGTH);
  const idempotencyKey = normalizeImportIdentifier_(
    source.idempotencyKey,
    'idempotencyKey',
    IMPORT_IDEMPOTENCY_KEY_MAX_LENGTH
  );
  if (idempotencyKey !== jobId + ':' + itemId) {
    throw importItemError_('INVALID_IDEMPOTENCY_KEY', 'idempotency key mismatch.', false);
  }
  const operationMode = String(source.operationMode || '');
  if (['create-pin', 'attach-existing-pin', 'replace-existing-audio'].indexOf(operationMode) === -1) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'operationMode is invalid.', false);
  }
  const sourceKind = String(source.sourceKind || '');
  if (sourceKind !== 'local' && sourceKind !== 'drive') {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'sourceKind is invalid.', false);
  }
  if (source.audioMimeType !== 'audio/mpeg') {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'audioMimeType is invalid.', false);
  }
  const audioBase64 = String(source.audioBase64 || '');
  if (!audioBase64 || audioBase64.length % 4 !== 0
      || !/^[A-Za-z0-9+/]*={0,2}$/.test(audioBase64)) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'audioBase64 is invalid.', false);
  }
  let audioBytes;
  try {
    audioBytes = Utilities.base64Decode(audioBase64);
    if (Utilities.base64Encode(audioBytes) !== audioBase64) throw new Error('non-canonical base64');
  } catch (_error) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'audioBase64 is invalid.', false);
  }
  const byteLength = audioBytes && Number.isSafeInteger(audioBytes.length) ? audioBytes.length : -1;
  const first = byteLength > 0 ? ((Number(audioBytes[0]) + 256) % 256) : -1;
  const second = byteLength > 1 ? ((Number(audioBytes[1]) + 256) % 256) : -1;
  const third = byteLength > 2 ? ((Number(audioBytes[2]) + 256) % 256) : -1;
  const hasId3 = first === 0x49 && second === 0x44 && third === 0x33;
  const hasMpegFrameSync = first === 0xff && (second & 0xe0) === 0xe0;
  if (byteLength < IMPORT_AUDIO_MIN_BYTES || byteLength > IMPORT_AUDIO_MAX_BYTES
      || (!hasId3 && !hasMpegFrameSync)) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'MP3 bytes are invalid.', false);
  }
  const sourceFileName = String(source.sourceFileName == null ? '' : source.sourceFileName).trim();
  if (!sourceFileName || sourceFileName.length > 255
      || /[\u0000-\u001f\u007f]/.test(sourceFileName)) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'sourceFileName is invalid.', false);
  }
  const sourceDriveFileId = String(source.sourceDriveFileId || '').trim();
  if ((sourceKind === 'local' && sourceDriveFileId)
      || (sourceKind === 'drive' && !isValidDrivePhotoImportId_(sourceDriveFileId))) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'sourceDriveFileId is invalid.', false);
  }

  let targetPinId = '';
  let expectedUpdatedAt = '';
  let pin;
  if (operationMode === 'create-pin') {
    if (!source.pin || typeof source.pin !== 'object' || Array.isArray(source.pin)) {
      throw importItemError_('INVALID_AUDIO_PAYLOAD', 'pin is invalid.', false);
    }
    const pinSource = Object.assign({}, source.pin);
    if (pinSource.lat === '') pinSource.lat = null;
    if (pinSource.lng === '') pinSource.lng = null;
    pinSource.eventAt = pinSource.eventTime == null ? '' : pinSource.eventTime;
    try {
      pin = normalizeNewPinPayload_(pinSource, {
        strict: true,
        normalizeColor: true,
        defaultMissingStatus: ''
      });
    } catch (_error) {
      throw importItemError_('INVALID_AUDIO_PAYLOAD', 'pin is invalid.', false);
    }
  } else {
    targetPinId = normalizeImportIdentifier_(
      source.targetPinId,
      'targetPinId',
      IMPORT_ITEM_ID_MAX_LENGTH
    );
    if (typeof source.expectedUpdatedAt !== 'string'
        || source.expectedUpdatedAt.length > 64
        || /[\u0000-\u001f\u007f]/.test(source.expectedUpdatedAt)) {
      throw importItemError_('INVALID_AUDIO_PAYLOAD', 'expectedUpdatedAt is invalid.', false);
    }
    expectedUpdatedAt = source.expectedUpdatedAt;
    pin = {
      title: '', description: '', lat: null, lng: null, color: DEFAULT_COLOR,
      icon: 'default', status: '', tags: [], links: [], eventAt: ''
    };
  }
  return Object.assign({}, pin, {
    jobId: jobId,
    itemId: itemId,
    idempotencyKey: idempotencyKey,
    idempotencyKeyHash: sha256Hex_(idempotencyKey),
    operationMode: operationMode,
    targetPinId: targetPinId,
    expectedUpdatedAt: expectedUpdatedAt,
    sourceKind: sourceKind,
    sourceDriveFileId: sourceDriveFileId,
    sourceFileName: sourceFileName,
    audioMimeType: 'audio/mpeg',
    audioBase64: audioBase64,
    audioBytes: audioBytes
  });
}

function hashImportAudioPayload_(normalized) {
  const hashPayload = {
    mediaKind: 'audio',
    jobId: normalized.jobId,
    itemId: normalized.itemId,
    operationMode: normalized.operationMode,
    targetPinId: normalized.targetPinId,
    expectedUpdatedAt: normalized.expectedUpdatedAt,
    sourceKind: normalized.sourceKind,
    sourceDriveFileId: normalized.sourceDriveFileId,
    sourceFileName: normalized.sourceFileName,
    audioMimeType: normalized.audioMimeType,
    audioBase64: normalized.audioBase64
  };
  if (normalized.operationMode === 'create-pin') {
    hashPayload.pin = {
      title: normalized.title,
      description: normalized.description,
      lat: normalized.lat,
      lng: normalized.lng,
      color: normalized.color,
      icon: normalized.icon,
      status: normalized.status,
      tags: normalized.tags.slice(),
      links: normalized.links.slice(),
      eventAt: normalized.eventAt
    };
  }
  return sha256Hex_(JSON.stringify(hashPayload));
}

function preflightImportAudioTarget_(mapSheet, normalized, preflightReceipt, payloadHash) {
  if (preflightReceipt) {
    if (String(preflightReceipt.payloadHash || '') !== payloadHash) {
      throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'audio payload conflict.', false);
    }
    if (String(preflightReceipt.mediaKind || '') !== 'audio') {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt media kind changed.', false);
    }
  }
  if (normalized.operationMode === 'create-pin') return;
  audioReceiptMapTarget_(mapSheet, normalized, preflightReceipt || null);
}

function importAudioTempFileName_(idempotencyKeyHash) {
  const digest = String(idempotencyKeyHash || '').toLowerCase();
  if (!/^[0-9a-f]{64}$/.test(digest)) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio receipt key is invalid.', false);
  }
  return '__drop_pin_audio_' + digest.slice(0, 32) + '.mp3';
}

function audioReceiptMapTarget_(mapSheet, normalized, receipt) {
  const pinId = normalized.operationMode === 'create-pin'
    ? String(receipt && receipt.pinId || '') : normalized.targetPinId;
  const target = pinId ? findMapInfoRowByPinId_(mapSheet, pinId) : null;
  const receiptFileId = String(receipt && receipt.fileId || '');
  if (target && receiptFileId && String(target.row[15] || '') === receiptFileId) {
    return { target: target, linked: true, cleanupFileId: String(receipt.cleanupFileId || '') };
  }
  if (normalized.operationMode === 'create-pin') {
    if (target) {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio create pin already exists.', false);
    }
    return { target: null, linked: false, cleanupFileId: '' };
  }
  if (!target) {
    throw importItemError_('PIN_AUDIO_TARGET_NOT_FOUND', 'audio target is missing.', false);
  }
  const currentAudioId = String(target.row[15] || '');
  if (String(target.row[13] || '') !== normalized.expectedUpdatedAt) {
    throw importItemError_('PIN_AUDIO_CONFLICT', 'audio target was updated.', false);
  }
  if (normalized.operationMode === 'attach-existing-pin' && currentAudioId) {
    throw importItemError_('PIN_AUDIO_ALREADY_ATTACHED', 'audio target already has audio.', false);
  }
  if (normalized.operationMode === 'replace-existing-audio' && !currentAudioId) {
    throw importItemError_('PIN_AUDIO_MISSING', 'audio target has no audio.', false);
  }
  return {
    target: target,
    linked: false,
    cleanupFileId: normalized.operationMode === 'replace-existing-audio' ? currentAudioId : ''
  };
}

function assertAudioReceiptMetadata_(receipt, normalized, targetFolderId) {
  if (!receipt
      || String(receipt.mediaKind || '') !== 'audio'
      || String(receipt.operationMode || '') !== normalized.operationMode
      || String(receipt.targetPinId || '') !== normalized.targetPinId
      || String(receipt.targetFolderId || '') !== String(targetFolderId || '')
      || String(receipt.tempFileName || '') !== importAudioTempFileName_(normalized.idempotencyKeyHash)
      || String(receipt.sourceDriveFileId || '') !== normalized.sourceDriveFileId
      || String(receipt.imageUrl || '') !== '') {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio receipt metadata changed.', false);
  }
}

function assertAudioSourceReceiptAvailable_(receiptSheet, normalized) {
  if (!normalized.sourceDriveFileId || receiptSheet.getLastRow() < 2) return;
  const rows = receiptSheet.getRange(
    2, 1, receiptSheet.getLastRow() - 1, IMPORT_RECEIPT_COLUMN_COUNT
  ).getValues();
  rows.forEach(function(row, index) {
    const receipt = importReceiptFromRow_(row, index + 2);
    if (String(receipt.idempotencyKey || '') === normalized.idempotencyKeyHash
        || String(receipt.sourceDriveFileId || '') !== normalized.sourceDriveFileId) return;
    if (String(receipt.state || '') !== IMPORT_RECEIPT_STATES.FAILED) {
      throw importItemError_('DRIVE_SOURCE_ALREADY_LINKED', 'audio source is already owned.', false);
    }
  });
}

function claimImportAudioReceipt_(receiptSheet, mapSheet, normalized, payloadHash, leaseOwner, targetFolderId) {
  return withImportReceiptLock_(function() {
    const now = getImportNow_();
    assertAudioSourceReceiptAvailable_(receiptSheet, normalized);
    let receipt = findImportReceipt_(receiptSheet, normalized.idempotencyKeyHash);
    if (receipt) {
      if (String(receipt.jobId || '') !== normalized.jobId
          || String(receipt.itemId || '') !== normalized.itemId) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio receipt ids changed.', false);
      }
      if (String(receipt.payloadHash || '') !== payloadHash) {
        throw importItemError_('IDEMPOTENCY_PAYLOAD_CONFLICT', 'audio payload conflict.', false);
      }
      if (!String(receipt.pinId || '')) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio receipt pin is missing.', false);
      }
      assertAudioReceiptMetadata_(receipt, normalized, targetFolderId);
      if (importLeaseIsActive_(receipt, now) && String(receipt.leaseOwner || '') !== leaseOwner) {
        throw importItemError_('IMPORT_ITEM_IN_PROGRESS', 'audio item is in progress.', true);
      }
    } else {
      const targetState = audioReceiptMapTarget_(mapSheet, normalized, null);
      receipt = {
        rowNumber: 0,
        idempotencyKey: normalized.idempotencyKeyHash,
        jobId: normalized.jobId,
        itemId: normalized.itemId,
        payloadHash: payloadHash,
        state: IMPORT_RECEIPT_STATES.RESERVED,
        leaseOwner: '', leaseUntil: '',
        pinId: normalized.operationMode === 'create-pin' ? Utilities.getUuid() : normalized.targetPinId,
        targetFolderId: String(targetFolderId || ''),
        tempFileName: importAudioTempFileName_(normalized.idempotencyKeyHash),
        fileId: '', imageUrl: '', folderUrl: '',
        createdAt: now.toISOString(), updatedAt: now.toISOString(), lastErrorCode: '',
        sourceDriveFileId: normalized.sourceDriveFileId,
        mediaKind: 'audio', operationMode: normalized.operationMode,
        targetPinId: normalized.targetPinId,
        cleanupFileId: targetState.cleanupFileId
      };
    }
    const targetState = audioReceiptMapTarget_(mapSheet, normalized, receipt);
    const currentState = String(receipt.state || '');
    if (targetState.linked) {
      if (normalized.operationMode !== 'create-pin'
          && String(targetState.target.row[13] || '') === normalized.expectedUpdatedAt) {
        const repairedUpdatedAt = currentUpdatedAt_();
        try {
          mapSheet.getRange(targetState.target.rowNumber, 14, 1, 1).setValue(repairedUpdatedAt);
          targetState.target.row[13] = repairedUpdatedAt;
        } catch (_error) {
          throw importItemError_(
            'IMPORT_AUDIO_MAP_UPDATE_FAILED',
            'audio timestamp repair failed.',
            true
          );
        }
      }
      if (currentState === IMPORT_RECEIPT_STATES.COMPLETED) {
        return {
          completed: true,
          cleanup: false,
          receipt: receipt,
          pin: mapInfoRowToPinResult_(targetState.target.row, ''),
          deduplicated: true
        };
      }
      if ([
        IMPORT_RECEIPT_STATES.FILE_SAVED,
        IMPORT_RECEIPT_STATES.LINKED,
        IMPORT_RECEIPT_STATES.CLEANUP_PENDING,
        IMPORT_RECEIPT_STATES.FAILED
      ].indexOf(currentState) === -1) {
        throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio linked receipt state is invalid.', false);
      }
      receipt.state = IMPORT_RECEIPT_STATES.LINKED;
      receipt.leaseOwner = '';
      receipt.leaseUntil = '';
      receipt.updatedAt = now.toISOString();
      receipt.lastErrorCode = '';
      writeImportReceipt_(receiptSheet, receipt);
      return {
        completed: false,
        cleanup: true,
        receipt: receipt,
        pin: mapInfoRowToPinResult_(targetState.target.row, ''),
        deduplicated: true
      };
    }
    if (currentState === IMPORT_RECEIPT_STATES.COMPLETED
        || currentState === IMPORT_RECEIPT_STATES.LINKED
        || currentState === IMPORT_RECEIPT_STATES.CLEANUP_PENDING) {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio receipt linkage is missing.', false);
    }
    if (String(receipt.cleanupFileId || '') !== targetState.cleanupFileId) {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio cleanup owner changed.', false);
    }
    receipt.state = receipt.fileId ? IMPORT_RECEIPT_STATES.FILE_SAVED : IMPORT_RECEIPT_STATES.RESERVED;
    receipt.leaseOwner = leaseOwner;
    receipt.leaseUntil = new Date(now.getTime() + IMPORT_ITEM_LEASE_MS).toISOString();
    receipt.updatedAt = now.toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return {
      completed: false,
      cleanup: false,
      receipt: receipt,
      target: targetState.target,
      deduplicated: false
    };
  });
}

function assertAudioReceiptOwner_(receipt, expected) {
  if (!receipt
      || String(receipt.payloadHash || '') !== expected.payloadHash
      || String(receipt.pinId || '') !== expected.pinId
      || String(receipt.leaseOwner || '') !== expected.leaseOwner
      || String(receipt.mediaKind || '') !== 'audio') {
    throw importItemError_('IMPORT_ITEM_LEASE_LOST', 'audio receipt ownership changed.', true);
  }
}

function persistImportAudioFile_(receiptSheet, expected, driveResult) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    assertAudioReceiptOwner_(receipt, expected);
    receipt.state = IMPORT_RECEIPT_STATES.FILE_SAVED;
    receipt.fileId = String(driveResult.fileId || '');
    receipt.imageUrl = '';
    receipt.folderUrl = '';
    receipt.updatedAt = getImportNow_().toISOString();
    receipt.leaseUntil = new Date(getImportNow_().getTime() + IMPORT_ITEM_LEASE_MS).toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return receipt;
  });
}

function finalizeImportAudioLink_(receiptSheet, mapSheet, expected, normalized) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    assertAudioReceiptOwner_(receipt, expected);
    const targetState = audioReceiptMapTarget_(mapSheet, normalized, receipt);
    let mapResult = targetState.target;
    if (!targetState.linked) {
      if (normalized.operationMode === 'create-pin') {
        const row = buildMapInfoRow_(normalized, {
          timestamp: formatMapTimestamp_(getImportNow_()),
          pinId: expected.pinId,
          fileId: '', imageUrl: '', updatedAt: ''
        });
        row[15] = String(receipt.fileId || '');
        try {
          appendMapInfoRow_(mapSheet, row);
        } catch (_error) {
          throw importItemError_('IMPORT_AUDIO_MAP_UPDATE_FAILED', 'audio pin append failed.', true);
        }
        mapResult = { rowNumber: mapSheet.getLastRow(), row: row };
      } else {
        const target = targetState.target;
        const updatedAt = currentUpdatedAt_();
        try {
          mapSheet.getRange(target.rowNumber, 16, 1, 1).setValue(String(receipt.fileId || ''));
          target.row[15] = String(receipt.fileId || '');
          mapSheet.getRange(target.rowNumber, 14, 1, 1).setValue(updatedAt);
          target.row[13] = updatedAt;
        } catch (_error) {
          throw importItemError_('IMPORT_AUDIO_MAP_UPDATE_FAILED', 'audio map update failed.', true);
        }
        mapResult = target;
      }
    }
    receipt.state = IMPORT_RECEIPT_STATES.LINKED;
    receipt.leaseOwner = '';
    receipt.leaseUntil = '';
    receipt.updatedAt = getImportNow_().toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return { pin: mapInfoRowToPinResult_(mapResult.row, ''), receipt: receipt };
  });
}

function completeImportAudioReceipt_(receiptSheet, expected) {
  return withImportReceiptLock_(function() {
    const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
    if (!receipt
        || String(receipt.payloadHash || '') !== expected.payloadHash
        || String(receipt.pinId || '') !== expected.pinId
        || String(receipt.mediaKind || '') !== 'audio') {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio completion owner changed.', false);
    }
    receipt.state = IMPORT_RECEIPT_STATES.COMPLETED;
    receipt.leaseOwner = '';
    receipt.leaseUntil = '';
    receipt.updatedAt = getImportNow_().toISOString();
    receipt.lastErrorCode = '';
    writeImportReceipt_(receiptSheet, receipt);
    return receipt;
  });
}

function markImportAudioCleanupPending_(receiptSheet, expected, errorCode) {
  try {
    withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash || '') !== expected.payloadHash
          || String(receipt.pinId || '') !== expected.pinId
          || String(receipt.mediaKind || '') !== 'audio') return;
      receipt.state = IMPORT_RECEIPT_STATES.CLEANUP_PENDING;
      receipt.leaseOwner = '';
      receipt.leaseUntil = '';
      receipt.updatedAt = getImportNow_().toISOString();
      receipt.lastErrorCode = String(errorCode || 'IMPORT_AUDIO_CLEANUP_FAILED');
      writeImportReceipt_(receiptSheet, receipt);
    });
  } catch (_error) {
    // The linked map row remains authoritative if cleanup journaling is unavailable.
  }
}

function runImportAudioCleanup_(receipt, structure) {
  if (String(receipt.sourceDriveFileId || '')) {
    archiveAudioDriveSourceIfPending_(String(receipt.sourceDriveFileId), structure);
  }
  if (String(receipt.cleanupFileId || '')) {
    trashManagedAudioFileIfOwned_(String(receipt.cleanupFileId), structure);
  }
}

function audioReceiptFileIsReferenced_(mapSheet, fileId) {
  const targetId = String(fileId || '');
  if (!targetId || mapSheet.getLastRow() < 2) return false;
  return mapSheet.getRange(2, 16, mapSheet.getLastRow() - 1, 1).getValues().some(function(row) {
    return String(row[0] || '') === targetId;
  });
}

function compensateUnlinkedImportAudioFile_(receiptSheet, mapSheet, expected, driveResult, structure, errorCode) {
  if (!driveResult || !driveResult.created || !driveResult.fileId) return false;
  let owned = false;
  try {
    owned = withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash || '') !== expected.payloadHash
          || String(receipt.pinId || '') !== expected.pinId
          || String(receipt.fileId || '') !== String(driveResult.fileId)
          || String(receipt.mediaKind || '') !== 'audio'
          || audioReceiptFileIsReferenced_(mapSheet, driveResult.fileId)) return false;
      return true;
    });
  } catch (_error) {
    return false;
  }
  if (!owned) return false;
  try {
    trashManagedAudioFileIfOwned_(String(driveResult.fileId), structure);
  } catch (_error) {
    return false;
  }
  try {
    withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt || String(receipt.fileId || '') !== String(driveResult.fileId)
          || audioReceiptFileIsReferenced_(mapSheet, driveResult.fileId)) return;
      receipt.state = IMPORT_RECEIPT_STATES.FAILED;
      receipt.fileId = '';
      receipt.imageUrl = '';
      receipt.folderUrl = '';
      receipt.leaseOwner = '';
      receipt.leaseUntil = '';
      receipt.updatedAt = getImportNow_().toISOString();
      receipt.lastErrorCode = String(errorCode || 'IMPORT_AUDIO_MAP_UPDATE_FAILED');
      writeImportReceipt_(receiptSheet, receipt);
    });
  } catch (_error) {
    return false;
  }
  return true;
}

function markImportAudioFailedIfUnlinked_(receiptSheet, mapSheet, expected, errorCode) {
  try {
    withImportReceiptLock_(function() {
      const receipt = findImportReceipt_(receiptSheet, expected.idempotencyKeyHash);
      if (!receipt
          || String(receipt.payloadHash || '') !== expected.payloadHash
          || String(receipt.pinId || '') !== expected.pinId
          || String(receipt.mediaKind || '') !== 'audio') return;
      if (receipt.fileId && audioReceiptFileIsReferenced_(mapSheet, receipt.fileId)) {
        receipt.state = IMPORT_RECEIPT_STATES.CLEANUP_PENDING;
      } else {
        receipt.state = IMPORT_RECEIPT_STATES.FAILED;
      }
      receipt.leaseOwner = '';
      receipt.leaseUntil = '';
      receipt.updatedAt = getImportNow_().toISOString();
      receipt.lastErrorCode = String(errorCode || 'IMPORT_ITEM_SAVE_FAILED');
      writeImportReceipt_(receiptSheet, receipt);
    });
  } catch (_error) {
    // Preserve the original safe failure if receipt journaling is unavailable.
  }
}

function saveImportAudioItem(data) {
  assertEditToken_(data);
  let normalized;
  let receiptSheet;
  let mapSheet;
  let expected = null;
  let driveResult = null;
  let structure = null;
  let currentPin = null;
  try {
    normalized = normalizeImportAudioPayload_(data);
    const payloadHash = hashImportAudioPayload_(normalized);
    const preflightReceipt = inspectExistingImportReceiptForSave_(normalized.idempotencyKeyHash);
    mapSheet = openMapInfoSheet_();
    preflightImportAudioTarget_(mapSheet, normalized, preflightReceipt, payloadHash);
    structure = audioStorageStructureForImport_();
    if (normalized.sourceKind === 'drive'
        && (!preflightReceipt || !String(preflightReceipt.fileId || ''))) {
      validateAudioDriveSourceForImport_(normalized.sourceDriveFileId, structure);
    }
    receiptSheet = ensureImportReceiptSchemaForSave_();
    const leaseOwner = Utilities.getUuid();
    const claim = claimImportAudioReceipt_(
      receiptSheet,
      mapSheet,
      normalized,
      payloadHash,
      leaseOwner,
      String(structure.audio || '')
    );
    currentPin = claim.target ? mapInfoRowToPinResult_(claim.target.row, '') : claim.pin;
    expected = {
      idempotencyKeyHash: normalized.idempotencyKeyHash,
      payloadHash: payloadHash,
      pinId: String(claim.receipt.pinId || ''),
      leaseOwner: leaseOwner
    };
    if (claim.completed) {
      return {
        ok: true,
        deduplicated: true,
        cleanupRequired: false,
        pin: claim.pin
      };
    }
    let linkedPin = claim.pin || null;
    let cleanupReceipt = claim.receipt;
    if (!claim.cleanup) {
      driveResult = resolveManagedAudioFile_(claim.receipt, normalized, structure);
      persistImportAudioFile_(receiptSheet, expected, driveResult);
      const linked = finalizeImportAudioLink_(receiptSheet, mapSheet, expected, normalized);
      linkedPin = linked.pin;
      cleanupReceipt = linked.receipt;
    }
    try {
      runImportAudioCleanup_(cleanupReceipt, structure);
      completeImportAudioReceipt_(receiptSheet, expected);
      return {
        ok: true,
        deduplicated: claim.deduplicated === true,
        cleanupRequired: false,
        pin: linkedPin
      };
    } catch (cleanupError) {
      const cleanupCode = isImportItemError_(cleanupError)
        ? String(cleanupError.code) : 'IMPORT_AUDIO_CLEANUP_FAILED';
      markImportAudioCleanupPending_(receiptSheet, expected, cleanupCode);
      return {
        ok: true,
        deduplicated: claim.deduplicated === true,
        cleanupRequired: true,
        pin: linkedPin
      };
    }
  } catch (error) {
    const code = isImportItemError_(error) ? String(error.code) : 'IMPORT_ITEM_SAVE_FAILED';
    if (receiptSheet && mapSheet && expected && driveResult && structure) {
      compensateUnlinkedImportAudioFile_(
        receiptSheet, mapSheet, expected, driveResult, structure, code
      );
    }
    if (receiptSheet && mapSheet && expected && code !== 'IMPORT_ITEM_LEASE_LOST') {
      markImportAudioFailedIfUnlinked_(receiptSheet, mapSheet, expected, code);
    }
    const failure = importItemFailureFromError_(error);
    if (normalized && normalized.operationMode === 'replace-existing-audio' && currentPin) {
      failure.pin = currentPin;
    }
    return failure;
  }
}

function normalizePinAudioMutationPayload_(data) {
  const source = data || {};
  if (!source || typeof source !== 'object' || Array.isArray(source)) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'audio mutation payload is invalid.', false);
  }
  const pinId = normalizeImportIdentifier_(source.pinId, 'pinId', IMPORT_ITEM_ID_MAX_LENGTH);
  if (typeof source.expectedUpdatedAt !== 'string'
      || source.expectedUpdatedAt.length > 64
      || /[\u0000-\u001f\u007f]/.test(source.expectedUpdatedAt)) {
    throw importItemError_('INVALID_AUDIO_PAYLOAD', 'expectedUpdatedAt is invalid.', false);
  }
  return { pinId: pinId, expectedUpdatedAt: source.expectedUpdatedAt };
}

function audioCleanupOwnerDigest_(domain, value) {
  return sha256Hex_(
    '\u0000audio-cleanup-owner:' + String(domain || '') + ':' + String(value || '')
  );
}

function audioCleanupFileDigest_(fileId) {
  return audioCleanupOwnerDigest_('file', String(fileId || ''));
}

function audioCleanupJournalIdentity_(operationMode, pinId, ownerValue) {
  const payload = {
    kind: 'audio-cleanup',
    operationMode: String(operationMode || ''),
    pinId: String(pinId || ''),
    ownerValue: String(ownerValue || '')
  };
  const payloadHash = sha256Hex_(JSON.stringify(payload));
  return {
    operationMode: payload.operationMode,
    pinId: payload.pinId,
    ownerValue: payload.ownerValue,
    receiptJobId: 'audio-cleanup',
    receiptItemId: payload.ownerValue,
    payloadHash: payloadHash,
    // External import identifiers reject NUL, so a client cannot generate this hash preimage.
    idempotencyKeyHash: sha256Hex_('\u0000audio-cleanup-key:' + JSON.stringify(payload))
  };
}

function newPinDeleteAudioCleanupIdentity_(pinId, cleanupFileId) {
  const identity = audioCleanupJournalIdentity_(
    'delete-pin-audio',
    pinId,
    audioCleanupOwnerDigest_('delete-attempt', Utilities.getUuid())
  );
  identity.receiptJobId = 'audio-cleanup-delete:' + audioCleanupFileDigest_(cleanupFileId);
  return identity;
}

function pinDeleteAudioCleanupIdentityFromReceipt_(receipt) {
  const pinId = String(receipt && receipt.targetPinId || '');
  const cleanupFileId = String(receipt && receipt.cleanupFileId || '');
  const ownerValue = String(receipt && receipt.itemId || '');
  if (!pinId || !cleanupFileId || !/^[0-9a-f]{64}$/.test(ownerValue)) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio cleanup generation is invalid.', false);
  }
  const identity = audioCleanupJournalIdentity_('delete-pin-audio', pinId, ownerValue);
  identity.receiptJobId = 'audio-cleanup-delete:' + audioCleanupFileDigest_(cleanupFileId);
  return identity;
}

function assertAudioCleanupJournal_(receipt, identity) {
  if (!receipt
      || String(receipt.payloadHash || '') !== identity.payloadHash
      || String(receipt.mediaKind || '') !== 'audio'
      || String(receipt.operationMode || '') !== identity.operationMode
      || String(receipt.pinId || '') !== identity.pinId
      || String(receipt.targetPinId || '') !== identity.pinId
      || String(receipt.jobId || '') !== identity.receiptJobId
      || String(receipt.itemId || '') !== identity.receiptItemId
      || String(receipt.fileId || '') !== ''
      || String(receipt.imageUrl || '') !== ''
      || String(receipt.sourceDriveFileId || '') !== '') {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio cleanup journal changed.', false);
  }
  return receipt;
}

function findAudioCleanupJournal_(receiptSheet, identity) {
  const receipt = findImportReceipt_(receiptSheet, identity.idempotencyKeyHash);
  return receipt ? assertAudioCleanupJournal_(receipt, identity) : null;
}

function claimAudioCleanupJournal_(receiptSheet, identity, cleanupFileId) {
  let receipt = findAudioCleanupJournal_(receiptSheet, identity);
  const normalizedCleanupFileId = String(cleanupFileId || '');
  if (receipt) {
    if (normalizedCleanupFileId
        && String(receipt.cleanupFileId || '') !== normalizedCleanupFileId) {
      throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio cleanup target changed.', false);
    }
    return receipt;
  }
  if (!normalizedCleanupFileId) {
    throw importItemError_('PIN_AUDIO_MISSING', 'audio cleanup target is missing.', false);
  }
  const now = getImportNow_().toISOString();
  receipt = {
    rowNumber: 0,
    idempotencyKey: identity.idempotencyKeyHash,
    jobId: identity.receiptJobId,
    itemId: identity.receiptItemId,
    payloadHash: identity.payloadHash,
    state: IMPORT_RECEIPT_STATES.RESERVED,
    leaseOwner: '', leaseUntil: '',
    pinId: identity.pinId,
    targetFolderId: '', tempFileName: '', fileId: '', imageUrl: '', folderUrl: '',
    createdAt: now, updatedAt: now, lastErrorCode: '', sourceDriveFileId: '',
    mediaKind: 'audio', operationMode: identity.operationMode,
    targetPinId: identity.pinId, cleanupFileId: normalizedCleanupFileId
  };
  writeImportReceipt_(receiptSheet, receipt);
  return receipt;
}

function markAudioCleanupJournalState_(receiptSheet, identity, state, errorCode) {
  const receipt = findAudioCleanupJournal_(receiptSheet, identity);
  if (!receipt) {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'audio cleanup journal is missing.', false);
  }
  receipt.state = state;
  receipt.leaseOwner = '';
  receipt.leaseUntil = '';
  receipt.updatedAt = getImportNow_().toISOString();
  receipt.lastErrorCode = String(errorCode || '');
  writeImportReceipt_(receiptSheet, receipt);
  return receipt;
}

function commitRemovePinAudio_(receiptSheet, mapSheet, normalized, identity) {
  return withImportReceiptLock_(function() {
    const target = findMapInfoRowByPinId_(mapSheet, normalized.pinId);
    if (!target) {
      throw importItemError_('PIN_AUDIO_TARGET_NOT_FOUND', 'audio target is missing.', false);
    }
    const currentAudioId = String(target.row[15] || '');
    let journal = findAudioCleanupJournal_(receiptSheet, identity);
    if (journal && currentAudioId
        && currentAudioId !== String(journal.cleanupFileId || '')) {
      if (String(journal.state || '') !== IMPORT_RECEIPT_STATES.COMPLETED) {
        journal = markAudioCleanupJournalState_(
          receiptSheet,
          identity,
          IMPORT_RECEIPT_STATES.CLEANUP_PENDING,
          ''
        );
      }
      return { journal: journal, pin: mapInfoRowToPinResult_(target.row, '') };
    }
    if (currentAudioId) {
      if (String(target.row[13] || '') !== normalized.expectedUpdatedAt) {
        throw importItemError_('PIN_AUDIO_CONFLICT', 'audio target was updated.', false);
      }
      journal = claimAudioCleanupJournal_(receiptSheet, identity, currentAudioId);
      mapSheet.getRange(target.rowNumber, 16, 1, 1).setValue('');
      target.row[15] = '';
    } else if (!journal) {
      throw importItemError_('PIN_AUDIO_MISSING', 'audio target has no audio.', false);
    }
    if (String(target.row[13] || '') === normalized.expectedUpdatedAt) {
      const updatedAt = currentUpdatedAt_();
      mapSheet.getRange(target.rowNumber, 14, 1, 1).setValue(updatedAt);
      target.row[13] = updatedAt;
    }
    if (String(journal.state || '') !== IMPORT_RECEIPT_STATES.COMPLETED) {
      journal = markAudioCleanupJournalState_(
        receiptSheet,
        identity,
        IMPORT_RECEIPT_STATES.CLEANUP_PENDING,
        ''
      );
    }
    return { journal: journal, pin: mapInfoRowToPinResult_(target.row, '') };
  });
}

function drainAudioCleanupJournal_(receiptSheet, identity, structure) {
  let receipt = findAudioCleanupJournal_(receiptSheet, identity);
  if (!receipt || String(receipt.state || '') === IMPORT_RECEIPT_STATES.COMPLETED) {
    return { cleanupRequired: false, receipt: receipt };
  }
  try {
    trashManagedAudioFileIfOwned_(String(receipt.cleanupFileId || ''), structure);
  } catch (_error) {
    try {
      withImportReceiptLock_(function() {
        markAudioCleanupJournalState_(
          receiptSheet,
          identity,
          IMPORT_RECEIPT_STATES.CLEANUP_PENDING,
          'IMPORT_AUDIO_CLEANUP_FAILED'
        );
      });
    } catch (_journalError) {
      // The existing cleanup journal remains authoritative for replay.
    }
    return { cleanupRequired: true, receipt: receipt };
  }
  receipt = withImportReceiptLock_(function() {
    return markAudioCleanupJournalState_(
      receiptSheet,
      identity,
      IMPORT_RECEIPT_STATES.COMPLETED,
      ''
    );
  });
  return { cleanupRequired: false, receipt: receipt };
}

function preparePinDeleteAudioCleanup_(snapshots, requestedPinIds) {
  if (typeof audioStorageStructureForImport_ !== 'function'
      || typeof trashManagedAudioFileIfOwned_ !== 'function') return null;
  const audioByPinId = Object.create(null);
  (snapshots || []).forEach(function(snapshot) {
    if (snapshot && snapshot.pinId) {
      audioByPinId[String(snapshot.pinId)] = String(snapshot.audioId || '');
    }
  });
  const requestedIds = normalizePinDeleteRequestIds_(requestedPinIds);
  const requestedIdSet = Object.create(null);
  requestedIds.forEach(function(pinId) { requestedIdSet[pinId] = true; });
  const entries = requestedIds.map(function(pinId) {
    return {
      pinId: pinId,
      pendingJournals: [],
      completedJournals: [],
      hadJournal: false
    };
  });
  const byPinId = Object.create(null);
  entries.forEach(function(entry) { byPinId[entry.pinId] = entry; });
  let hasWork = requestedIds.some(function(pinId) { return !!audioByPinId[pinId]; });
  const spreadsheet = openDataSpreadsheet_();
  const inspection = inspectImportReceiptSchemaForSave_(
    spreadsheet.getSheetByName(IMPORT_RECEIPTS_SHEET_NAME)
  );
  if (inspection.state === 'corrupted') {
    throw importItemError_('IMPORT_RECEIPT_CORRUPTED', 'receipt headers are invalid.', false);
  }
  if (inspection.state === 'current' && inspection.sheet.getLastRow() >= 2) {
    inspection.sheet.getRange(
      2, 1, inspection.sheet.getLastRow() - 1, IMPORT_RECEIPT_COLUMN_COUNT
    ).getValues().forEach(function(row, index) {
      const receipt = importReceiptFromRow_(row, index + 2);
      if (String(receipt.mediaKind || '') !== 'audio'
          || String(receipt.operationMode || '') !== 'delete-pin-audio'
          || !requestedIdSet[String(receipt.targetPinId || '')]) return;
      const identity = pinDeleteAudioCleanupIdentityFromReceipt_(receipt);
      assertAudioCleanupJournal_(receipt, identity);
      const entry = byPinId[identity.pinId];
      const journal = { identity: identity, receipt: receipt };
      if (String(receipt.state || '') === IMPORT_RECEIPT_STATES.COMPLETED) {
        entry.completedJournals.push(journal);
        if (!Object.prototype.hasOwnProperty.call(audioByPinId, identity.pinId)) {
          entry.hadJournal = true;
        }
      } else {
        entry.pendingJournals.push(journal);
        entry.hadJournal = true;
        hasWork = true;
      }
    });
  }
  const hasReplayEvidence = entries.some(function(entry) {
    return entry.completedJournals.length > 0
      && !Object.prototype.hasOwnProperty.call(audioByPinId, entry.pinId);
  });
  const canDiscoverCurrentAudio = (snapshots || []).length > 0;
  if (!hasWork && !hasReplayEvidence && !canDiscoverCurrentAudio) return null;
  const structure = hasWork ? audioStorageStructureForImport_() : null;
  const receiptSheet = (hasWork || canDiscoverCurrentAudio)
    ? ensureImportReceiptSchemaForSave_()
    : inspection.sheet;
  return {
    structure: structure,
    receiptSheet: receiptSheet,
    entries: entries,
    byPinId: byPinId
  };
}

function claimPinDeleteAudioCleanup_(preparation, pinId, cleanupFileId) {
  if (!preparation || !preparation.receiptSheet) return null;
  const entry = preparation.byPinId[String(pinId || '')];
  if (!entry) return null;
  const cleanupDigest = audioCleanupFileDigest_(cleanupFileId);
  let journal = entry.pendingJournals.find(function(candidate) {
    return candidate.identity.receiptJobId === 'audio-cleanup-delete:' + cleanupDigest;
  }) || null;
  const identity = journal
    ? journal.identity
    : newPinDeleteAudioCleanupIdentity_(String(pinId || ''), cleanupFileId);
  const receipt = claimAudioCleanupJournal_(
    preparation.receiptSheet,
    identity,
    cleanupFileId
  );
  journal = { identity: identity, receipt: receipt };
  if (!entry.pendingJournals.some(function(candidate) {
    return candidate.identity.idempotencyKeyHash === identity.idempotencyKeyHash;
  })) {
    entry.pendingJournals.push(journal);
  }
  entry.hadJournal = true;
  return journal;
}

function stagePinDeleteAudioCleanup_(preparation, journal) {
  if (!preparation || !journal) return null;
  journal.receipt = markAudioCleanupJournalState_(
    preparation.receiptSheet,
    journal.identity,
    IMPORT_RECEIPT_STATES.CLEANUP_PENDING,
    ''
  );
  return journal;
}

function reconcilePinDeleteAudioCleanup_(preparation, identity) {
  if (!preparation || !identity) return false;
  return withImportReceiptLock_(function() {
    const receipt = findAudioCleanupJournal_(preparation.receiptSheet, identity);
    if (!receipt || String(receipt.state || '') === IMPORT_RECEIPT_STATES.COMPLETED) return !!receipt;
    const target = findMapInfoRowByPinId_(openMapInfoSheet_(), identity.pinId);
    if (target && String(target.row[15] || '') === String(receipt.cleanupFileId || '')) {
      return false;
    }
    markAudioCleanupJournalState_(
      preparation.receiptSheet,
      identity,
      IMPORT_RECEIPT_STATES.CLEANUP_PENDING,
      ''
    );
    return true;
  });
}

function drainPreparedPinDeleteAudioCleanup_(preparation, identities) {
  if (!preparation) return { cleanupRequired: false, attempted: false };
  const unique = Object.create(null);
  const queue = [];
  (identities || []).concat([].concat.apply([], preparation.entries.map(function(entry) {
    return entry.pendingJournals.map(function(journal) { return journal.identity; });
  }))).forEach(function(identity) {
    if (!identity || unique[identity.idempotencyKeyHash]) return;
    unique[identity.idempotencyKeyHash] = true;
    queue.push(identity);
  });
  let cleanupRequired = false;
  let attempted = false;
  queue.forEach(function(identity) {
    let ready = false;
    try {
      ready = reconcilePinDeleteAudioCleanup_(preparation, identity);
    } catch (_error) {
      cleanupRequired = true;
      return;
    }
    if (!ready) return;
    attempted = true;
    if (!preparation.structure) {
      try {
        preparation.structure = audioStorageStructureForImport_();
      } catch (_error) {
        cleanupRequired = true;
        return;
      }
    }
    const result = drainAudioCleanupJournal_(
      preparation.receiptSheet,
      identity,
      preparation.structure
    );
    if (result.cleanupRequired) cleanupRequired = true;
  });
  return { cleanupRequired: cleanupRequired, attempted: attempted };
}

function preflightPinAudioMutation_(mapSheet, normalized) {
  const target = findMapInfoRowByPinId_(mapSheet, normalized.pinId);
  if (!target) {
    throw importItemError_('PIN_AUDIO_TARGET_NOT_FOUND', 'audio target is missing.', false);
  }
  if (String(target.row[13] || '') !== normalized.expectedUpdatedAt) {
    throw importItemError_('PIN_AUDIO_CONFLICT', 'audio target was updated.', false);
  }
  if (!String(target.row[15] || '')) {
    throw importItemError_('PIN_AUDIO_MISSING', 'audio target has no audio.', false);
  }
  return target;
}

function removePinAudio(data) {
  assertEditToken_(data);
  try {
    const normalized = normalizePinAudioMutationPayload_(data);
    const mapSheet = openMapInfoSheet_();
    const identity = audioCleanupJournalIdentity_(
      'remove-pin-audio',
      normalized.pinId,
      audioCleanupOwnerDigest_('remove-request', normalized.expectedUpdatedAt)
    );
    const preflightReceipt = inspectExistingImportReceiptForSave_(identity.idempotencyKeyHash);
    if (preflightReceipt) {
      assertAudioCleanupJournal_(preflightReceipt, identity);
    } else {
      preflightPinAudioMutation_(mapSheet, normalized);
    }
    const structure = audioStorageStructureForImport_();
    const receiptSheet = ensureImportReceiptSchemaForSave_();
    let mutation;
    try {
      mutation = commitRemovePinAudio_(receiptSheet, mapSheet, normalized, identity);
    } catch (error) {
      if (!findAudioCleanupJournal_(receiptSheet, identity)) throw error;
      mutation = commitRemovePinAudio_(receiptSheet, mapSheet, normalized, identity);
    }
    const cleanup = drainAudioCleanupJournal_(receiptSheet, identity, structure);
    return { ok: true, cleanupRequired: cleanup.cleanupRequired, pin: mutation.pin };
  } catch (error) {
    return importItemFailureFromError_(error);
  }
}


function saveMapData(data) {
  assertEditToken_(data);
  if (!data || !String(data.title || '').trim()) {
    return { ok: false, error: 'title is required' };
  }

  const normalized = normalizeNewPinPayload_(data, {
    strict: false,
    defaultMissingStatus: '未対応'
  });
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
      PinData.buildFileNameForSave(normalized.title, data.filename, getRenameFileWithTitle_())
    );

    const folder = DriveApp.getFolderById(uploadFolderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    fileId = file.getId();
    imageUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w1920';
    folderUrl = getDriveFolderUrl_(uploadFolderId);
  }

  appendMapInfoRow_(sheet, buildMapInfoRow_(normalized, {
    timestamp: now,
    pinId: id,
    fileId: fileId,
    imageUrl: imageUrl,
    updatedAt: ''
  }));

  return {
    ok: true,
    id: id,
    imageUrl: imageUrl,
    fileId: fileId,
    folderUrl: folderUrl,
    links: normalized.links
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
    icon,
    ''
  ];
  sheet.appendRow(row);

  const pin = toClientPin_(PinData.rowToPin(row));
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
    const sourceProtection = inspectDriveImportSourceProtection_([{
      pinId: String(data.id),
      fileId: renameFileId
    }]);
    if (!pinDriveFileIsProtectedSource_(
      { pinId: String(data.id), fileId: renameFileId },
      sourceProtection
    )) {
      renameDriveFileForTitle_(renameFileId, title);
    }
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
    invalidateRouteCacheForPins_([data.id]);
    return { ok: true };
  });
}

function normalizeBulkPinMetadataIds_(ids) {
  const seen = {};
  const result = [];
  (Array.isArray(ids) ? ids : []).forEach(function(id) {
    const pinId = String(id == null ? '' : id);
    if (!pinId || seen[pinId]) return;
    seen[pinId] = true;
    result.push(pinId);
  });
  return result;
}

function writeBulkPinMetadataColumnRuns_(sheet, rows, rowIndexes, columnIndex) {
  if (!rowIndexes.length) return;
  const sortedRowIndexes = rowIndexes.slice().sort(function(a, b) { return a - b; });
  let runStart = sortedRowIndexes[0];
  let previous = runStart;

  function writeRun(start, end) {
    const values = [];
    for (var rowIndex = start; rowIndex <= end; rowIndex += 1) {
      values.push([rows[rowIndex][columnIndex]]);
    }
    sheet.getRange(start + 1, columnIndex + 1, values.length, 1).setValues(values);
  }

  for (var index = 1; index < sortedRowIndexes.length; index += 1) {
    const current = sortedRowIndexes[index];
    if (current === previous + 1) {
      previous = current;
      continue;
    }
    writeRun(runStart, previous);
    runStart = current;
    previous = current;
  }
  writeRun(runStart, previous);
}

function bulkUpdatePinMetadata(data) {
  assertEditToken_(data);
  if (!data || !Array.isArray(data.ids) || data.ids.length === 0) {
    return { ok: false, error: 'ids must be a non-empty array' };
  }

  const tagMode = String(data && data.tagMode || 'none').trim().toLowerCase();
  if (['none', 'add', 'remove', 'replace'].indexOf(tagMode) === -1) {
    return { ok: false, error: 'invalid tagMode: ' + tagMode };
  }

  let requestedTags = [];
  if (tagMode !== 'none') {
    try {
      requestedTags = PinData.normalizeTags(data && data.tags);
    } catch (error) {
      return { ok: false, error: error && error.message ? error.message : String(error) };
    }
    if ((tagMode === 'add' || tagMode === 'remove') && requestedTags.length === 0) {
      return { ok: false, error: tagMode + ' requires at least one tag' };
    }
  }

  const hasIcon = !!(data && Object.prototype.hasOwnProperty.call(data, 'icon') && data.icon != null);
  let requestedIcon = null;
  if (hasIcon) {
    const rawIcon = String(data.icon || '').trim().toLowerCase();
    if (PinData.ICON_OPTIONS.indexOf(rawIcon) === -1) {
      return { ok: false, error: 'invalid icon: ' + String(data.icon) };
    }
    requestedIcon = PinData.normalizeIcon(rawIcon);
  }

  const hasStatus = !!(data && Object.prototype.hasOwnProperty.call(data, 'status') && data.status != null);
  let requestedStatus = null;
  if (hasStatus) {
    try {
      requestedStatus = PinData.normalizeStatus(String(data.status));
    } catch (_error) {
      return { ok: false, error: 'invalid status: ' + data.status };
    }
    if (!requestedStatus) return { ok: false, error: 'status is required' };
  }
  if (tagMode === 'none' && !hasIcon && !hasStatus) {
    return { ok: false, error: 'no metadata changes requested' };
  }

  return withSpreadsheetMutationLock_(function() {
    const targetIds = normalizeBulkPinMetadataIds_(data.ids);
    if (targetIds.length === 0) {
      return { ok: false, error: 'ids must be a non-empty array' };
    }
    const sheet = openMapInfoSheet_();
    const lastRow = sheet.getLastRow();
    const rows = lastRow > 0
      ? sheet.getRange(1, 1, lastRow, MAP_INFO_COLUMN_COUNT).getValues()
      : [];
    const requestedSet = new Set(targetIds);
    const rowIndexById = {};

    for (var rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
      const pinId = String(rows[rowIndex][8] == null ? '' : rows[rowIndex][8]);
      if (!requestedSet.has(pinId) || Object.prototype.hasOwnProperty.call(rowIndexById, pinId)) continue;
      rowIndexById[pinId] = rowIndex;
    }

    const missingIds = targetIds.filter(function(pinId) {
      return !Object.prototype.hasOwnProperty.call(rowIndexById, pinId);
    });
    if (missingIds.length > 0) {
      return { ok: false, error: 'pin ids not found', missingIds: missingIds };
    }

    const changedTagRows = [];
    const changedIconRows = [];
    const changedStatusRows = [];
    const changedRowIndexes = [];
    const changedRowSet = {};
    const normalizedTagsById = {};
    let validationError = null;

    targetIds.forEach(function(pinId) {
      if (validationError) return;
      const currentRowIndex = rowIndexById[pinId];
      const row = rows[currentRowIndex];
      let currentTags;
      try {
        currentTags = PinData.normalizeTags(PinData.deserializeTags(row[11] || ''));
      } catch (error) {
        validationError = error && error.message ? error.message : String(error);
        return;
      }
      normalizedTagsById[pinId] = currentTags;

      if (tagMode !== 'none') {
        let nextTags;
        try {
          if (tagMode === 'replace') {
            nextTags = requestedTags.slice();
          } else if (tagMode === 'add') {
            nextTags = PinData.normalizeTags(currentTags.concat(requestedTags));
          } else {
            const removeSet = {};
            requestedTags.forEach(function(tag) { removeSet[tag.toLowerCase()] = true; });
            nextTags = PinData.normalizeTags(currentTags.filter(function(tag) {
              return !removeSet[tag.toLowerCase()];
            }));
          }
        } catch (error) {
          validationError = error && error.message ? error.message : String(error);
          return;
        }
        normalizedTagsById[pinId] = nextTags;
        const serializedTags = PinData.serializeTags(nextTags);
        if (String(row[11] || '') !== serializedTags) {
          row[11] = serializedTags;
          changedTagRows.push(currentRowIndex);
          changedRowSet[currentRowIndex] = true;
        }
      }

      if (hasIcon && String(row[14] || '').trim().toLowerCase() !== requestedIcon) {
        row[14] = requestedIcon;
        changedIconRows.push(currentRowIndex);
        changedRowSet[currentRowIndex] = true;
      }

      if (hasStatus && String(row[10] || '').trim() !== requestedStatus) {
        row[10] = requestedStatus;
        changedStatusRows.push(currentRowIndex);
        changedRowSet[currentRowIndex] = true;
      }
    });

    if (validationError) return { ok: false, error: validationError };

    const updatedAt = currentUpdatedAt_();
    Object.keys(changedRowSet).map(Number).sort(function(a, b) { return a - b; }).forEach(function(rowIndex) {
      rows[rowIndex][13] = updatedAt;
      changedRowIndexes.push(rowIndex);
    });

    writeBulkPinMetadataColumnRuns_(sheet, rows, changedStatusRows, 10);
    writeBulkPinMetadataColumnRuns_(sheet, rows, changedTagRows, 11);
    writeBulkPinMetadataColumnRuns_(sheet, rows, changedRowIndexes, 13);
    writeBulkPinMetadataColumnRuns_(sheet, rows, changedIconRows, 14);

    const updates = targetIds.map(function(pinId) {
      const row = rows[rowIndexById[pinId]];
      return {
        id: pinId,
        tags: normalizedTagsById[pinId].slice(),
        icon: PinData.normalizeIcon(row[14]),
        status: String(row[10] || '').trim(),
        updatedAt: row[13] ? String(row[13]) : ''
      };
    });
    return {
      ok: true,
      updatedCount: changedRowIndexes.length,
      unchangedCount: targetIds.length - changedRowIndexes.length,
      updates: updates
    };
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
      audioId: String(row[15] || ''),
      originalRowNumber: rowIndex + 1
    });
  }
  return snapshots.sort(function(a, b) {
    return b.originalRowNumber - a.originalRowNumber;
  });
}

function driveFileHasSurvivingPinReference_(rows, fileId, deletingPinIds) {
  const targetFileId = String(fileId || '');
  if (!targetFileId) return false;
  const deleting = deletingPinIds instanceof Set ? deletingPinIds : new Set(deletingPinIds || []);
  for (var rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];
    const pinId = String(row[8] || '');
    if (pinId && !deleting.has(pinId) && String(row[6] || '') === targetFileId) return true;
  }
  return false;
}

function inspectDriveImportSourceProtection_(pinFiles) {
  const protection = {
    protectAll: false,
    fileIds: Object.create(null),
    managedFileIds: Object.create(null)
  };
  const requestedFileIds = Object.create(null);
  (pinFiles || []).forEach(function(pinFile) {
    if (!pinFile || !pinFile.fileId) return;
    requestedFileIds[String(pinFile.fileId)] = true;
  });
  if (Object.keys(requestedFileIds).length === 0) return protection;

  const sheet = openDataSpreadsheet_().getSheetByName(IMPORT_RECEIPTS_SHEET_NAME);
  if (!sheet) {
    logDrivePhotoImportFailure_(
      'source-ownership-check',
      new Error('import receipt ownership sheet is unavailable')
    );
    protection.protectAll = true;
    return protection;
  }
  try {
    const headers = sheet.getRange(1, 1, 1, IMPORT_RECEIPT_COLUMN_COUNT).getValues()[0];
    const validHeaders = IMPORT_RECEIPT_HEADERS.every(function(header, index) {
      return headers[index] === header;
    });
    if (!validHeaders) throw new Error('import receipt ownership schema is unavailable');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return protection;

    sheet.getRange(2, 8, lastRow - 1, 10).getValues().forEach(function(row) {
      const targetFolderId = String(row[1] || '');
      const fileId = String(row[3] || '');
      const sourceDriveFileId = String(row[9] || '');
      const fileIdValid = isValidDrivePhotoImportId_(fileId);
      const sourceDriveFileIdValid = isValidDrivePhotoImportId_(sourceDriveFileId);
      const targetFolderIdValid = isValidDrivePhotoImportId_(targetFolderId);
      if (sourceDriveFileIdValid && requestedFileIds[sourceDriveFileId]) {
        protection.fileIds[sourceDriveFileId] = true;
      }
      const managedFileConfirmed = fileIdValid && (sourceDriveFileId
        ? (sourceDriveFileIdValid
          && fileId !== sourceDriveFileId
          && (!targetFolderId || targetFolderIdValid))
        : targetFolderIdValid);
      if (fileId && requestedFileIds[fileId] && managedFileConfirmed) {
        protection.managedFileIds[fileId] = true;
      }
    });
  } catch (error) {
    logDrivePhotoImportFailure_('source-ownership-check', error);
    protection.protectAll = true;
  }
  return protection;
}

function pinDriveFileIsProtectedSource_(pinFile, protection) {
  if (!pinFile || !pinFile.fileId || !protection) return false;
  const fileId = String(pinFile.fileId);
  return protection.protectAll === true
    || protection.fileIds[fileId] === true
    || protection.managedFileIds[fileId] !== true;
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
      audioId: String(row[15] || ''),
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
  let audioCleanupPreparation = null;
  try {
    audioCleanupPreparation = preparePinDeleteAudioCleanup_(snapshots, [data.id]);
  } catch (error) {
    return importItemFailureFromError_(error);
  }
  if (snapshots.length === 0) {
    const missingPinId = String(data.id);
    let missingResult;
    let missingError = null;
    try {
      missingResult = withSpreadsheetMutationLock_(function() {
        const currentRows = sheet.getDataRange().getValues();
        const currentById = indexCurrentPinDeleteRows_(currentRows);
        if (currentById[missingPinId]) {
          logPinDeleteStage_(missingPinId, 'appeared_after_snapshot');
          return { ok: false, error: PIN_DELETE_CONFLICT_ERROR };
        }
        deletePinRelationsAndCaches_([missingPinId]);
        return { ok: false, error: 'id not found' };
      });
    } catch (error) {
      missingError = error;
    }
    const missingDrain = drainPreparedPinDeleteAudioCleanup_(audioCleanupPreparation, []);
    const hadJournal = !!(audioCleanupPreparation
      && audioCleanupPreparation.entries.some(function(entry) { return entry.hadJournal; }));
    if (missingError) throw missingError;
    if (hadJournal && missingResult && missingResult.error !== SPREADSHEET_MUTATION_BUSY_ERROR) {
      missingResult = { ok: true, cleanupRequired: missingDrain.cleanupRequired };
    }
    return missingResult;
  }
  const snapshot = snapshots[0];
  let deletedAudioId = '';
  const audioCleanupIdentities = [];
  const sourceProtection = inspectDriveImportSourceProtection_(snapshots);

  if (snapshot.fileId
      && !pinDriveFileIsProtectedSource_(snapshot, sourceProtection)
      && !driveFileHasSurvivingPinReference_(rows, snapshot.fileId, new Set([snapshot.pinId]))) {
    try {
      DriveApp.getFileById(snapshot.fileId).setTrashed(true);
    } catch (_error) {
      return { ok: false, error: '写真の削除に失敗しました。もう一度お試しください。' };
    }
  }

  let result;
  let mutationError = null;
  try {
    result = withSpreadsheetMutationLock_(function() {
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

      let audioJournal = null;
      if (current.audioId && audioCleanupPreparation) {
        audioJournal = claimPinDeleteAudioCleanup_(
          audioCleanupPreparation,
          current.pinId,
          current.audioId
        );
        audioCleanupIdentities.push(audioJournal.identity);
      } else {
        deletedAudioId = current.audioId;
      }
      sheet.deleteRow(current.rowNumber);
      if (audioJournal) stagePinDeleteAudioCleanup_(audioCleanupPreparation, audioJournal);
      deletePinRelationsAndCaches_([snapshot.pinId]);
      return { ok: true };
    });
  } catch (error) {
    mutationError = error;
  }
  if (result && result.error === SPREADSHEET_MUTATION_BUSY_ERROR) {
    logDriveSuccessesAfterLockFailure_([snapshot]);
  }
  const journalDrain = drainPreparedPinDeleteAudioCleanup_(
    audioCleanupPreparation,
    audioCleanupIdentities
  );
  if (result && result.ok && deletedAudioId
      && typeof audioStorageStructureForImport_ === 'function'
      && typeof trashManagedAudioFileIfOwned_ === 'function') {
    try {
      const structure = audioStorageStructureForImport_();
      trashManagedAudioFileIfOwned_(deletedAudioId, structure);
      result.cleanupRequired = false;
    } catch (_error) {
      result.cleanupRequired = true;
    }
  }
  const hadAudioCleanupJournal = !!(audioCleanupPreparation
    && audioCleanupPreparation.entries.some(function(entry) { return entry.hadJournal; }));
  if (result && result.ok && hadAudioCleanupJournal) {
    result.cleanupRequired = journalDrain.cleanupRequired;
  }
  if (mutationError) throw mutationError;
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
  const requestedPinIdSet = new Set(requestedPinIds);
  const snapshots = buildPinDeleteSnapshots_(rows, data.ids);
  let audioCleanupPreparation = null;
  try {
    audioCleanupPreparation = preparePinDeleteAudioCleanup_(snapshots, requestedPinIds);
  } catch (error) {
    return importItemFailureFromError_(error);
  }
  const sourceProtection = inspectDriveImportSourceProtection_(snapshots);
  const driveSuccessfulSnapshots = [];
  const deletedAudioSnapshots = [];
  const audioCleanupIdentities = [];
  const failedIdSet = {};
  const driveFileResults = Object.create(null);

  snapshots.forEach(function(snapshot) {
    try {
      if (snapshot.fileId && !pinDriveFileIsProtectedSource_(snapshot, sourceProtection)) {
        if (driveFileHasSurvivingPinReference_(rows, snapshot.fileId, requestedPinIdSet)) {
          driveSuccessfulSnapshots.push(snapshot);
          return;
        }
        if (!Object.prototype.hasOwnProperty.call(driveFileResults, snapshot.fileId)) {
          DriveApp.getFileById(snapshot.fileId).setTrashed(true);
          driveFileResults[snapshot.fileId] = true;
        } else if (driveFileResults[snapshot.fileId] !== true) {
          throw new Error('Drive delete failed.');
        }
      }
      driveSuccessfulSnapshots.push(snapshot);
    } catch (_error) {
      if (snapshot.fileId) driveFileResults[snapshot.fileId] = false;
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

  let result;
  let mutationError = null;
  try {
    result = withSpreadsheetMutationLock_(function() {
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
        let audioJournal = null;
        if (current.audioId && audioCleanupPreparation) {
          audioJournal = claimPinDeleteAudioCleanup_(
            audioCleanupPreparation,
            current.pinId,
            current.audioId
          );
          audioCleanupIdentities.push(audioJournal.identity);
        }
        currentEntries.push({
          pinId: snapshot.pinId,
          fileId: current.fileId,
          audioId: current.audioId,
          audioJournal: audioJournal,
          rowNumber: current.rowNumber
        });
      });
      const deletedIds = [];

      groupContiguousPinDeleteRows_(currentEntries).forEach(function(run) {
        try {
          sheet.deleteRows(run.startRow, run.entries.length);
        } catch (_error) {
          run.entries.forEach(function(entry) {
            failedIdSet[entry.pinId] = true;
            logPinDeleteStage_(entry.pinId, 'spreadsheet_delete_failed');
          });
          return;
        }
        run.entries.forEach(function(entry) {
          deletedIds.push(entry.pinId);
          if (entry.audioJournal) {
            stagePinDeleteAudioCleanup_(audioCleanupPreparation, entry.audioJournal);
          } else if (entry.audioId) {
            deletedAudioSnapshots.push(entry);
          }
        });
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
  } catch (error) {
    mutationError = error;
  }
  if (result && result.error === SPREADSHEET_MUTATION_BUSY_ERROR) {
    logDriveSuccessesAfterLockFailure_(driveSuccessfulSnapshots);
  }
  const journalDrain = drainPreparedPinDeleteAudioCleanup_(
    audioCleanupPreparation,
    audioCleanupIdentities
  );
  if (result && result.ok && deletedAudioSnapshots.length
      && typeof audioStorageStructureForImport_ === 'function'
      && typeof trashManagedAudioFileIfOwned_ === 'function') {
    let cleanupRequired = false;
    try {
      const structure = audioStorageStructureForImport_();
      deletedAudioSnapshots.forEach(function(snapshot) {
        try {
          trashManagedAudioFileIfOwned_(snapshot.audioId, structure);
        } catch (_error) {
          cleanupRequired = true;
        }
      });
    } catch (_error) {
      cleanupRequired = true;
    }
    result.cleanupRequired = cleanupRequired;
  }
  const hadAudioCleanupJournal = !!(audioCleanupPreparation
    && audioCleanupPreparation.entries.some(function(entry) { return entry.hadJournal; }));
  if (result && result.ok && hadAudioCleanupJournal) {
    result.cleanupRequired = journalDrain.cleanupRequired;
  }
  if (mutationError) throw mutationError;
  return result;
}

function getAppSettings() {
  const startedAt = startupTimingNow_();
  try {
    const config = getAppConfig_();
    const rootFolderId = extractDriveFolderId_(config.IMAGE_DRIVE_URL || '');
    const response = {
      ok: true,
      rootFolderId: rootFolderId,
      rootFolderUrl: getDriveFolderUrl_(rootFolderId),
      renameFileWithTitle: PinData.toBooleanSetting(config.RENAME_FILE_WITH_TITLE)
    };
    logStartupGasStage_('getAppSettings', 'total', 'success', startedAt);
    return response;
  } catch (caught) {
    logStartupGasStage_('getAppSettings', 'total', 'failure', startedAt);
    throw caught;
  }
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
    listDrivePhotoImportFolder: listDrivePhotoImportFolder,
    readDrivePhotoImportFile: readDrivePhotoImportFile,
    isDriveFolderWithinRoot_: isDriveFolderWithinRoot_,
    isDriveFileWithinRoot_: isDriveFileWithinRoot_,
    saveMapData: saveMapData,
    saveImportPhotoItem: saveImportPhotoItem,
    updatePinDetails: updatePinDetails,
    movePin: movePin,
    unplacePin: unplacePin,
    bulkUpdatePinMetadata: bulkUpdatePinMetadata,
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

var MEDIA_DRIVE_NAMES = Object.freeze({
  photos: 'photos',
  audio: 'audio',
  original: 'original',
  guide: 'ここに直接ファイルを入れてください.txt'
});
var DRIVE_AUDIO_MAX_BYTES = 15 * 1024 * 1024;
var MANAGED_AUDIO_MIN_BYTES = 1024;
var MANAGED_AUDIO_MAX_BYTES = 4 * 1024 * 1024;

var DRIVE_MEDIA_ERROR_MESSAGES = Object.freeze({
  DRIVE_MEDIA_ACCESS_DENIED: 'Driveメディアの取込権限を確認してください。',
  DRIVE_MEDIA_ROOT_MISSING: 'Driveメディアの取込元フォルダが設定されていません。',
  DRIVE_MEDIA_KIND_INVALID: 'Driveメディアの種類を確認してください。',
  DRIVE_MEDIA_STRUCTURE_AMBIGUOUS: 'Driveメディアの保存先に同名フォルダまたは案内ファイルが複数あります。',
  DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT: 'Driveメディアの保存先に同名の項目があります。',
  DRIVE_MEDIA_STRUCTURE_PARENT_INVALID: 'Driveメディアの保存先階層を確認してください。',
  DRIVE_MEDIA_STRUCTURE_FAILED: 'Driveメディアの保存先を準備できませんでした。',
  DRIVE_MEDIA_INBOX_READ_FAILED: 'Driveメディアの一覧を取得できませんでした。',
  DRIVE_MEDIA_INBOX_TOO_LARGE: '取込Inboxの項目数が多すぎます。',
  DRIVE_AUDIO_FILE_ID_INVALID: 'Driveの音声を確認してください。',
  DRIVE_AUDIO_FILE_NOT_FOUND: 'Driveの音声を開けませんでした。',
  DRIVE_AUDIO_FILE_OUTSIDE_INBOX: 'このDrive音声は取込Inboxの直下にありません。',
  DRIVE_AUDIO_FILE_TYPE_UNSUPPORTED: '対応していない音声形式です。',
  DRIVE_AUDIO_FILE_EMPTY: '空の音声ファイルは取り込めません。',
  DRIVE_AUDIO_FILE_TOO_LARGE: '音声ファイルは15MB以内にしてください。',
  DRIVE_AUDIO_FILE_READ_FAILED: 'Driveの音声を読み込めませんでした。'
});

function mediaDriveHasOwn_(value, key) {
  return !!value && Object.prototype.hasOwnProperty.call(value, key);
}

function mediaDriveError_(code) {
  var error = new Error(code);
  error.code = code;
  return error;
}

function mediaDriveFailure_(code) {
  var safeCode = mediaDriveHasOwn_(DRIVE_MEDIA_ERROR_MESSAGES, code)
    ? code : 'DRIVE_MEDIA_STRUCTURE_FAILED';
  return {
    ok: false,
    errorCode: safeCode,
    error: DRIVE_MEDIA_ERROR_MESSAGES[safeCode]
  };
}

function mediaDriveFailureFromError_(error, fallbackCode) {
  var code = fallbackCode;
  try {
    if (error && mediaDriveHasOwn_(error, 'code')
        && mediaDriveHasOwn_(DRIVE_MEDIA_ERROR_MESSAGES, error.code)) {
      code = error.code;
    }
  } catch (_error) {
    code = fallbackCode;
  }
  return mediaDriveFailure_(code);
}

function assertMediaDriveEditToken_(payload) {
  try {
    assertEditToken_(payload);
  } catch (_error) {
    throw mediaDriveError_('DRIVE_MEDIA_ACCESS_DENIED');
  }
}

function getMediaDriveRootId_() {
  var rootFolderId = getRootFolderId_();
  if (!isValidDrivePhotoImportId_(rootFolderId)) {
    throw mediaDriveError_('DRIVE_MEDIA_ROOT_MISSING');
  }
  return String(rootFolderId);
}

function mediaDriveDirectParentIds_(item) {
  var ids = [];
  var parents = item.getParents();
  while (parents.hasNext()) ids.push(String(parents.next().getId()));
  return ids;
}

function mediaDriveIsExactDirectChild_(item, parentId) {
  var ids = mediaDriveDirectParentIds_(item);
  return ids.length === 1 && ids[0] === String(parentId);
}

function mediaDriveIsActive_(item) {
  return !!item && typeof item.isTrashed === 'function' && !item.isTrashed();
}

function mediaDriveInspectFolder_(parent, name) {
  var parentId = String(parent.getId());
  var fileMatches = [];
  var files = parent.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (!mediaDriveIsActive_(file) || String(file.getName()) !== name) continue;
    if (!mediaDriveIsExactDirectChild_(file, parentId)) {
      throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
    }
    fileMatches.push(file);
  }
  if (fileMatches.length) throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT');

  var matches = [];
  var folders = parent.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    if (!mediaDriveIsActive_(folder) || String(folder.getName()) !== name) continue;
    if (!mediaDriveIsExactDirectChild_(folder, parentId)) {
      throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
    }
    matches.push(folder);
    if (matches.length > 1) throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_AMBIGUOUS');
  }
  return matches.length === 1 ? matches[0] : null;
}

function mediaDriveCreateFolder_(parent, name) {
  var parentId = String(parent.getId());
  var created = parent.createFolder(name);
  if (!mediaDriveIsActive_(created)
      || String(created.getName()) !== name
      || !isValidDrivePhotoImportId_(String(created.getId()))
      || !mediaDriveIsExactDirectChild_(created, parentId)) {
    throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
  }
  return created;
}

function mediaDriveInspectGuide_(rootFolder) {
  var rootId = String(rootFolder.getId());
  var folderMatches = [];
  var folders = rootFolder.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    if (mediaDriveIsActive_(folder) && String(folder.getName()) === MEDIA_DRIVE_NAMES.guide) {
      folderMatches.push(folder);
    }
  }
  if (folderMatches.length) throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_NAME_CONFLICT');

  var matches = [];
  var files = rootFolder.getFilesByName(MEDIA_DRIVE_NAMES.guide);
  while (files.hasNext()) {
    var file = files.next();
    if (!mediaDriveIsActive_(file)) continue;
    if (!mediaDriveIsExactDirectChild_(file, rootId)) {
      throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
    }
    matches.push(file);
    if (matches.length > 1) throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_AMBIGUOUS');
  }
  return matches.length === 1 ? matches[0] : null;
}

function mediaDriveCreateGuide_(rootFolder) {
  var rootId = String(rootFolder.getId());
  var blob = Utilities.newBlob([], 'text/plain', MEDIA_DRIVE_NAMES.guide);
  var created = rootFolder.createFile(blob);
  if (!mediaDriveIsActive_(created)
      || String(created.getName()) !== MEDIA_DRIVE_NAMES.guide
      || !mediaDriveIsExactDirectChild_(created, rootId)) {
    throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
  }
  return created;
}

function mediaDriveInspectStructure_(rootFolder) {
  var photosFolder = mediaDriveInspectFolder_(rootFolder, MEDIA_DRIVE_NAMES.photos);
  var audioFolder = mediaDriveInspectFolder_(rootFolder, MEDIA_DRIVE_NAMES.audio);
  var originalFolder = mediaDriveInspectFolder_(rootFolder, MEDIA_DRIVE_NAMES.original);
  var guideFile = mediaDriveInspectGuide_(rootFolder);
  var originalPhotosFolder = null;
  var originalAudioFolder = null;
  if (originalFolder) {
    originalPhotosFolder = mediaDriveInspectFolder_(originalFolder, MEDIA_DRIVE_NAMES.photos);
    originalAudioFolder = mediaDriveInspectFolder_(originalFolder, MEDIA_DRIVE_NAMES.audio);
  }
  return {
    photosFolder: photosFolder,
    audioFolder: audioFolder,
    originalFolder: originalFolder,
    originalPhotosFolder: originalPhotosFolder,
    originalAudioFolder: originalAudioFolder,
    guideFile: guideFile
  };
}

function mediaDriveCreateMissingStructure_(rootFolder, inspected) {
  var photosFolder = inspected.photosFolder
    || mediaDriveCreateFolder_(rootFolder, MEDIA_DRIVE_NAMES.photos);
  var audioFolder = inspected.audioFolder
    || mediaDriveCreateFolder_(rootFolder, MEDIA_DRIVE_NAMES.audio);
  var originalFolder = inspected.originalFolder
    || mediaDriveCreateFolder_(rootFolder, MEDIA_DRIVE_NAMES.original);
  var originalPhotosFolder = inspected.originalPhotosFolder
    || mediaDriveCreateFolder_(originalFolder, MEDIA_DRIVE_NAMES.photos);
  var originalAudioFolder = inspected.originalAudioFolder
    || mediaDriveCreateFolder_(originalFolder, MEDIA_DRIVE_NAMES.audio);
  var guideFile = inspected.guideFile || mediaDriveCreateGuide_(rootFolder);
  return {
    photosFolder: photosFolder,
    audioFolder: audioFolder,
    originalFolder: originalFolder,
    originalPhotosFolder: originalPhotosFolder,
    originalAudioFolder: originalAudioFolder,
    guideFile: guideFile
  };
}

function mediaDriveStructureIds_(rootId, inspected) {
  if (!inspected.photosFolder || !inspected.audioFolder || !inspected.originalFolder
      || !inspected.originalPhotosFolder || !inspected.originalAudioFolder
      || !inspected.guideFile) {
    throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
  }
  return {
    root: rootId,
    photos: String(inspected.photosFolder.getId()),
    audio: String(inspected.audioFolder.getId()),
    original: String(inspected.originalFolder.getId()),
    originalPhotos: String(inspected.originalPhotosFolder.getId()),
    originalAudio: String(inspected.originalAudioFolder.getId())
  };
}

function withMediaDriveStructureLock_(callback) {
  var lock = LockService.getScriptLock();
  var acquired = false;
  try {
    acquired = lock.tryLock(SPREADSHEET_MUTATION_LOCK_TIMEOUT_MS);
    if (!acquired) throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_FAILED');
    return callback();
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function mediaDriveAddOwner_(owners, fileId, pinId) {
  var normalizedFileId = String(fileId || '');
  var normalizedPinId = String(pinId || '');
  if (!isValidDrivePhotoImportId_(normalizedFileId) || !normalizedPinId) return;
  if (!mediaDriveHasOwn_(owners, normalizedFileId)) owners[normalizedFileId] = Object.create(null);
  owners[normalizedFileId][normalizedPinId] = true;
}

function getMediaDriveAssociations_() {
  var result = {
    managedPhotoIds: new Set(),
    managedAudioIds: new Set(),
    completedSourceIds: new Set(),
    photoOwners: Object.create(null)
  };
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var mapSheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (mapSheet) {
      var mapRows = mapSheet.getDataRange().getValues();
      if (mapRows.length) {
        var mapHeaders = mapRows[0].map(String);
        var mapPinIdColumn = mapHeaders.indexOf('ID');
        var mapPhotoIdColumn = mapHeaders.indexOf('ファイルID');
        var mapAudioIdColumn = mapHeaders.indexOf('音声ID');
        mapRows.slice(1).forEach(function(row) {
          var pinId = mapPinIdColumn >= 0 ? String(row[mapPinIdColumn] || '') : '';
          var photoId = mapPhotoIdColumn >= 0 ? String(row[mapPhotoIdColumn] || '') : '';
          var audioId = mapAudioIdColumn >= 0 ? String(row[mapAudioIdColumn] || '') : '';
          if (isValidDrivePhotoImportId_(photoId)) {
            result.managedPhotoIds.add(photoId);
            mediaDriveAddOwner_(result.photoOwners, photoId, pinId);
          }
          if (isValidDrivePhotoImportId_(audioId)) result.managedAudioIds.add(audioId);
        });
      }
    }

    var receiptSheet = spreadsheet.getSheetByName(IMPORT_RECEIPTS_SHEET_NAME);
    if (receiptSheet) {
      var receiptRows = receiptSheet.getDataRange().getValues();
      if (receiptRows.length) {
        var receiptHeaders = receiptRows[0].map(String);
        var stateColumn = receiptHeaders.indexOf('state');
        var pinIdColumn = receiptHeaders.indexOf('pinId');
        var fileIdColumn = receiptHeaders.indexOf('fileId');
        var sourceIdColumn = receiptHeaders.indexOf('sourceDriveFileId');
        var kindColumn = receiptHeaders.indexOf('mediaKind');
        receiptRows.slice(1).forEach(function(row) {
          if (stateColumn < 0 || String(row[stateColumn] || '') !== 'completed') return;
          var kind = kindColumn >= 0 ? String(row[kindColumn] || '') : '';
          var fileId = fileIdColumn >= 0 ? String(row[fileIdColumn] || '') : '';
          var sourceId = sourceIdColumn >= 0 ? String(row[sourceIdColumn] || '') : '';
          var pinId = pinIdColumn >= 0 ? String(row[pinIdColumn] || '') : '';
          if (isValidDrivePhotoImportId_(sourceId)) result.completedSourceIds.add(sourceId);
          if (kind === 'audio') {
            if (isValidDrivePhotoImportId_(fileId)) result.managedAudioIds.add(fileId);
          } else if (isValidDrivePhotoImportId_(fileId)) {
            result.managedPhotoIds.add(fileId);
            mediaDriveAddOwner_(result.photoOwners, fileId, pinId);
          }
        });
      }
    }
    return result;
  } catch (_error) {
    throw mediaDriveError_('DRIVE_MEDIA_INBOX_READ_FAILED');
  }
}

function migrateLegacyManagedPhotos_(rootFolder, photosFolder) {
  var associations = getMediaDriveAssociations_();
  var rootId = String(rootFolder.getId());
  var files = rootFolder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (!mediaDriveIsActive_(file)) continue;
    var fileId = String(file.getId());
    var owners = associations.photoOwners[fileId];
    if (!owners || Object.keys(owners).length !== 1) continue;
    if (!mediaDriveIsExactDirectChild_(file, rootId)) continue;
    var classification = classifyDrivePhotoImportFile_({
      name: String(file.getName()),
      type: String(file.getMimeType())
    });
    if (!classification.supported || classification.kind !== 'jpeg') continue;
    file.moveTo(photosFolder);
    if (!mediaDriveIsExactDirectChild_(file, String(photosFolder.getId()))) {
      throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_PARENT_INVALID');
    }
  }
}

function ensureMediaDriveStructure_() {
  var rootId = getMediaDriveRootId_();
  try {
    return withMediaDriveStructureLock_(function() {
      var rootFolder = DriveApp.getFolderById(rootId);
      if (!mediaDriveIsActive_(rootFolder) || String(rootFolder.getId()) !== rootId) {
        throw mediaDriveError_('DRIVE_MEDIA_ROOT_MISSING');
      }
      var inspected = mediaDriveInspectStructure_(rootFolder);
      mediaDriveCreateMissingStructure_(rootFolder, inspected);
      var verified = mediaDriveInspectStructure_(rootFolder);
      var structure = mediaDriveStructureIds_(rootId, verified);
      migrateLegacyManagedPhotos_(rootFolder, verified.photosFolder);
      return structure;
    });
  } catch (error) {
    if (error && mediaDriveHasOwn_(DRIVE_MEDIA_ERROR_MESSAGES, error.code)) throw error;
    throw mediaDriveError_('DRIVE_MEDIA_STRUCTURE_FAILED');
  }
}

function ensureMediaDriveStructure(payload) {
  try {
    assertMediaDriveEditToken_(payload);
    ensureMediaDriveStructure_();
    return { ok: true };
  } catch (error) {
    return mediaDriveFailureFromError_(error, 'DRIVE_MEDIA_STRUCTURE_FAILED');
  }
}

function classifyDriveAudioImportFile_(value) {
  var name = value && typeof value.name === 'string' ? value.name : '';
  var type = value && typeof value.type === 'string' ? value.type.trim().toLowerCase() : '';
  var extension = drivePhotoImportExtensionOf_(name);
  var expected = {
    m4a: { kind: 'm4a', mimeType: 'audio/mp4' },
    mp3: { kind: 'mp3', mimeType: 'audio/mpeg' },
    wav: { kind: 'wav', mimeType: 'audio/wav' }
  }[extension];
  if (!expected || type !== expected.mimeType) {
    return { supported: false, kind: '', normalizedMimeType: '' };
  }
  return { supported: true, kind: expected.kind, normalizedMimeType: expected.mimeType };
}

function classifyDriveMediaImportFile_(mediaKind, value) {
  return mediaKind === 'audio'
    ? classifyDriveAudioImportFile_(value)
    : classifyDrivePhotoImportFile_(value);
}

function resolveDriveMediaKind_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)
      || !mediaDriveHasOwn_(payload, 'mediaKind')
      || (payload.mediaKind !== 'photo' && payload.mediaKind !== 'audio')) {
    throw mediaDriveError_('DRIVE_MEDIA_KIND_INVALID');
  }
  return payload.mediaKind;
}

function listDriveMediaInbox(payload) {
  try {
    assertMediaDriveEditToken_(payload);
    var mediaKind = resolveDriveMediaKind_(payload);
    var structure = ensureMediaDriveStructure_();
    var rootFolder = DriveApp.getFolderById(structure.root);
    var associations = getMediaDriveAssociations_();
    var excludedManagedIds = mediaKind === 'audio'
      ? associations.managedAudioIds : associations.managedPhotoIds;
    var items = [];
    var scannedEntries = 0;
    var files = rootFolder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      if (String(file.getName()) === MEDIA_DRIVE_NAMES.guide) continue;
      scannedEntries += 1;
      if (scannedEntries > DRIVE_PHOTO_IMPORT_MAX_FOLDER_ENTRIES) {
        throw mediaDriveError_('DRIVE_MEDIA_INBOX_TOO_LARGE');
      }
      if (!mediaDriveIsActive_(file) || !mediaDriveIsExactDirectChild_(file, structure.root)) continue;
      var id = String(file.getId());
      var name = String(file.getName());
      var mimeType = String(file.getMimeType());
      var sizeBytes = Number(file.getSize());
      if (excludedManagedIds.has(id) || associations.completedSourceIds.has(id)
          || mimeType === DRIVE_PHOTO_IMPORT_SHORTCUT_MIME) continue;
      var classification = classifyDriveMediaImportFile_(mediaKind, { name: name, type: mimeType });
      var modifiedAt = drivePhotoImportModifiedAt_(file);
      if (!isValidDrivePhotoImportId_(id)
          || !isValidDrivePhotoImportFileName_(name)
          || !classification.supported
          || !Number.isSafeInteger(sizeBytes)
          || sizeBytes <= 0
          || sizeBytes > DRIVE_AUDIO_MAX_BYTES
          || !modifiedAt) continue;
      items.push({
        id: id,
        name: name,
        mimeType: classification.normalizedMimeType,
        sizeBytes: sizeBytes,
        modifiedAt: modifiedAt,
        kind: classification.kind
      });
    }
    drivePhotoImportSort_(items);
    return { ok: true, items: items };
  } catch (error) {
    return mediaDriveFailureFromError_(error, 'DRIVE_MEDIA_INBOX_READ_FAILED');
  }
}

function driveMediaReadCodes_(mediaKind) {
  return mediaKind === 'audio' ? {
    invalid: 'DRIVE_AUDIO_FILE_ID_INVALID',
    missing: 'DRIVE_AUDIO_FILE_NOT_FOUND',
    outside: 'DRIVE_AUDIO_FILE_OUTSIDE_INBOX',
    unsupported: 'DRIVE_AUDIO_FILE_TYPE_UNSUPPORTED',
    empty: 'DRIVE_AUDIO_FILE_EMPTY',
    large: 'DRIVE_AUDIO_FILE_TOO_LARGE',
    failed: 'DRIVE_AUDIO_FILE_READ_FAILED'
  } : {
    invalid: 'DRIVE_IMPORT_FILE_ID_INVALID',
    missing: 'DRIVE_IMPORT_FILE_NOT_FOUND',
    outside: 'DRIVE_IMPORT_FILE_OUTSIDE_ROOT',
    unsupported: 'DRIVE_IMPORT_FILE_TYPE_UNSUPPORTED',
    empty: 'DRIVE_IMPORT_FILE_READ_FAILED',
    large: 'DRIVE_IMPORT_FILE_TOO_LARGE',
    failed: 'DRIVE_IMPORT_FILE_READ_FAILED'
  };
}

function readDriveMediaImportFile_(payload, mediaKind) {
  var codes = driveMediaReadCodes_(mediaKind);
  try {
    try {
      assertEditToken_(payload);
    } catch (_error) {
      if (mediaKind === 'audio') throw mediaDriveError_('DRIVE_MEDIA_ACCESS_DENIED');
      throw drivePhotoImportError_('DRIVE_IMPORT_ACCESS_DENIED');
    }
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)
        || !mediaDriveHasOwn_(payload, 'fileId')
        || !isValidDrivePhotoImportId_(payload.fileId)) {
      throw mediaKind === 'audio' ? mediaDriveError_(codes.invalid) : drivePhotoImportError_(codes.invalid);
    }
    var structure = ensureMediaDriveStructure_();
    var associations = getMediaDriveAssociations_();
    var managedIds = mediaKind === 'audio'
      ? associations.managedAudioIds : associations.managedPhotoIds;
    if (managedIds.has(String(payload.fileId))
        || associations.completedSourceIds.has(String(payload.fileId))) {
      throw mediaKind === 'audio' ? mediaDriveError_(codes.outside) : drivePhotoImportError_(codes.outside);
    }
    var file;
    try {
      file = DriveApp.getFileById(payload.fileId);
    } catch (_error) {
      throw mediaKind === 'audio' ? mediaDriveError_(codes.missing) : drivePhotoImportError_(codes.missing);
    }
    if (!mediaDriveIsActive_(file)) {
      throw mediaKind === 'audio' ? mediaDriveError_(codes.missing) : drivePhotoImportError_(codes.missing);
    }
    var id = String(file.getId());
    var name = String(file.getName());
    var mimeType = String(file.getMimeType());
    var classification = classifyDriveMediaImportFile_(mediaKind, { name: name, type: mimeType });
    if (!classification.supported || mimeType === DRIVE_PHOTO_IMPORT_SHORTCUT_MIME) {
      throw mediaKind === 'audio'
        ? mediaDriveError_(codes.unsupported) : drivePhotoImportError_(codes.unsupported);
    }
    if (!mediaDriveIsExactDirectChild_(file, structure.root)) {
      throw mediaKind === 'audio' ? mediaDriveError_(codes.outside) : drivePhotoImportError_(codes.outside);
    }
    var modifiedAt = drivePhotoImportModifiedAt_(file);
    if (id !== payload.fileId || !isValidDrivePhotoImportFileName_(name) || !modifiedAt) {
      throw mediaKind === 'audio' ? mediaDriveError_(codes.failed) : drivePhotoImportError_(codes.failed);
    }
    var bytes;
    try {
      bytes = file.getBlob().getBytes();
    } catch (_error) {
      throw mediaKind === 'audio' ? mediaDriveError_(codes.failed) : drivePhotoImportError_(codes.failed);
    }
    var sizeBytes = bytes && Number.isSafeInteger(bytes.length) ? bytes.length : -1;
    if (sizeBytes <= 0) {
      bytes = null;
      throw mediaKind === 'audio' ? mediaDriveError_(codes.empty) : drivePhotoImportError_(codes.empty);
    }
    if (sizeBytes > DRIVE_AUDIO_MAX_BYTES) {
      bytes = null;
      throw mediaKind === 'audio' ? mediaDriveError_(codes.large) : drivePhotoImportError_(codes.large);
    }
    var base64;
    try {
      base64 = Utilities.base64Encode(bytes);
    } catch (_error) {
      bytes = null;
      throw mediaKind === 'audio' ? mediaDriveError_(codes.failed) : drivePhotoImportError_(codes.failed);
    }
    bytes = null;
    return {
      ok: true,
      file: {
        id: id,
        name: name,
        mimeType: classification.normalizedMimeType,
        sizeBytes: sizeBytes,
        modifiedAt: modifiedAt,
        kind: classification.kind,
        base64: base64
      }
    };
  } catch (error) {
    return mediaKind === 'audio'
      ? mediaDriveFailureFromError_(error, codes.failed)
      : drivePhotoImportFailureFromError_(error, codes.failed);
  }
}

function readDriveAudioImportFile(payload) {
  return readDriveMediaImportFile_(payload, 'audio');
}

function audioStorageImportError_(code, message, retryable) {
  if (typeof importItemError_ === 'function') {
    return importItemError_(code, message, retryable === true);
  }
  var error = new Error(String(message || code || 'Audio storage failed.'));
  error.code = String(code || 'AUDIO_STORAGE_FAILED');
  error.retryable = retryable === true;
  return error;
}

function audioStorageStructureForRead_() {
  var rootId = getMediaDriveRootId_();
  try {
    var rootFolder = DriveApp.getFolderById(rootId);
    if (!mediaDriveIsActive_(rootFolder) || String(rootFolder.getId()) !== rootId) {
      throw new Error('media root is unavailable');
    }
    var audioFolder = mediaDriveInspectFolder_(rootFolder, MEDIA_DRIVE_NAMES.audio);
    if (!audioFolder) throw new Error('managed audio folder is unavailable');
    return { root: rootId, audio: String(audioFolder.getId()) };
  } catch (_error) {
    throw audioStorageImportError_(
      'PIN_AUDIO_NOT_FOUND',
      'pin audio is unavailable.',
      false
    );
  }
}

function audioStorageStructureForImport_() {
  try {
    return ensureMediaDriveStructure_();
  } catch (error) {
    var code = error && typeof error.code === 'string'
      ? String(error.code) : 'DRIVE_MEDIA_STRUCTURE_FAILED';
    throw audioStorageImportError_(code, 'media Drive structure is unavailable.', true);
  }
}

function audioStorageFileIsDirectManaged_(file, structure) {
  return !!file
    && mediaDriveIsActive_(file)
    && String(file.getMimeType()) === 'audio/mpeg'
    && mediaDriveIsExactDirectChild_(file, String(structure && structure.audio || ''));
}

function validateAudioDriveSourceForImport_(sourceDriveFileId, structure) {
  var sourceId = String(sourceDriveFileId || '');
  if (!isValidDrivePhotoImportId_(sourceId)) {
    throw audioStorageImportError_('IMPORT_AUDIO_SOURCE_INVALID', 'audio source is invalid.', false);
  }
  try {
    var associations = getMediaDriveAssociations_();
    if (associations.managedAudioIds.has(sourceId)
        || associations.completedSourceIds.has(sourceId)) {
      throw audioStorageImportError_(
        'DRIVE_SOURCE_ALREADY_LINKED',
        'audio source is already linked.',
        false
      );
    }
    var file = DriveApp.getFileById(sourceId);
    var name = String(file.getName());
    var mimeType = String(file.getMimeType());
    var sizeBytes = Number(file.getSize());
    var classification = classifyDriveAudioImportFile_({ name: name, type: mimeType });
    if (!mediaDriveIsActive_(file)
        || String(file.getId()) !== sourceId
        || !mediaDriveIsExactDirectChild_(file, String(structure && structure.root || ''))
        || !isValidDrivePhotoImportFileName_(name)
        || !classification.supported
        || !Number.isSafeInteger(sizeBytes)
        || sizeBytes <= 0
        || sizeBytes > DRIVE_AUDIO_MAX_BYTES) {
      throw audioStorageImportError_('IMPORT_AUDIO_SOURCE_INVALID', 'audio source is invalid.', false);
    }
    return { file: file, fileId: sourceId };
  } catch (error) {
    if (error && typeof error.code === 'string') throw error;
    throw audioStorageImportError_(
      'IMPORT_AUDIO_SOURCE_CHECK_FAILED',
      'audio source could not be verified.',
      true
    );
  }
}

function resolveManagedAudioFile_(receipt, normalized, structure) {
  var audioFolderId = String(structure && structure.audio || '');
  var audioFolder;
  try {
    audioFolder = DriveApp.getFolderById(audioFolderId);
    var file = null;
    var created = false;
    if (String(receipt.fileId || '')) {
      file = DriveApp.getFileById(String(receipt.fileId));
      if (!audioStorageFileIsDirectManaged_(file, structure)) {
        throw audioStorageImportError_(
          'IMPORT_AUDIO_FILE_INVALID',
          'managed audio file is invalid.',
          false
        );
      }
    } else {
      var matches = [];
      var iterator = audioFolder.getFilesByName(String(receipt.tempFileName || ''));
      while (iterator.hasNext() && matches.length < 2) matches.push(iterator.next());
      if (matches.length > 1) {
        throw audioStorageImportError_(
          'IMPORT_RECEIPT_CORRUPTED',
          'multiple managed audio files exist.',
          false
        );
      }
      if (matches.length === 1) {
        file = matches[0];
      } else {
        file = audioFolder.createFile(Utilities.newBlob(
          normalized.audioBytes,
          'audio/mpeg',
          String(receipt.tempFileName || '')
        ));
        created = true;
      }
    }
    if (!audioStorageFileIsDirectManaged_(file, structure)
        || String(file.getId()) === String(normalized.sourceDriveFileId || '')) {
      throw audioStorageImportError_(
        'IMPORT_AUDIO_FILE_INVALID',
        'managed audio file could not be verified.',
        false
      );
    }
    return {
      file: file,
      fileId: String(file.getId()),
      created: created,
      parentFolderId: audioFolderId
    };
  } catch (error) {
    if (error && typeof error.code === 'string') throw error;
    throw audioStorageImportError_(
      'IMPORT_AUDIO_FILE_SAVE_FAILED',
      'managed audio file could not be saved.',
      true
    );
  }
}

function archiveAudioDriveSourceIfPending_(sourceDriveFileId, structure) {
  var sourceId = String(sourceDriveFileId || '');
  if (!sourceId) return { moved: false };
  try {
    var file = DriveApp.getFileById(sourceId);
    if (!file || file.isTrashed() || String(file.getId()) !== sourceId) {
      return { moved: false };
    }
    if (mediaDriveIsExactDirectChild_(file, String(structure.originalAudio || ''))) {
      return { moved: false };
    }
    if (!mediaDriveIsExactDirectChild_(file, String(structure.root || ''))) {
      throw audioStorageImportError_(
        'IMPORT_AUDIO_SOURCE_ARCHIVE_FAILED',
        'audio source archive location is invalid.',
        true
      );
    }
    validateAudioDriveSourceForImport_(sourceId, structure);
    file.moveTo(DriveApp.getFolderById(String(structure.originalAudio || '')));
    if (!mediaDriveIsExactDirectChild_(file, String(structure.originalAudio || ''))) {
      throw new Error('audio source archive could not be verified');
    }
    return { moved: true };
  } catch (error) {
    if (error && error.code === 'IMPORT_AUDIO_SOURCE_ARCHIVE_FAILED') throw error;
    throw audioStorageImportError_(
      'IMPORT_AUDIO_SOURCE_ARCHIVE_FAILED',
      'audio source could not be archived.',
      true
    );
  }
}

function managedAudioHasLiveReference_(fileId) {
  var targetId = String(fileId || '');
  var sheet = openMapInfoSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;
  return sheet.getRange(2, 16, lastRow - 1, 1).getValues().some(function(row) {
    return String(row[0] || '') === targetId;
  });
}

function trashManagedAudioFileIfOwned_(fileId, structure) {
  var targetId = String(fileId || '');
  if (!targetId || managedAudioHasLiveReference_(targetId)) return false;
  try {
    var file = DriveApp.getFileById(targetId);
    if (file.isTrashed()) return true;
    if (!audioStorageFileIsDirectManaged_(file, structure)) return false;
    file.setTrashed(true);
    return true;
  } catch (error) {
    throw audioStorageImportError_(
      'IMPORT_AUDIO_CLEANUP_FAILED',
      'managed audio cleanup failed.',
      true
    );
  }
}

function isPinPhotoReadMimeType_(mimeType) {
  return mimeType === 'image/jpeg'
    || mimeType === 'image/png'
    || mimeType === 'image/gif'
    || mimeType === 'image/webp';
}

function readPinPhotoBlobByPinId_(pinId) {
  var normalizedPinId = String(pinId || '');
  if (!normalizedPinId) {
    throw importItemError_('PIN_PHOTO_NOT_FOUND', 'pin photo is unavailable.', false);
  }
  var target = findMapInfoRowByPinId_(openMapInfoSheet_(), normalizedPinId);
  var fileId = target ? String(target.row[6] || '') : '';
  if (!fileId) {
    throw importItemError_('PIN_PHOTO_NOT_FOUND', 'pin photo is unavailable.', false);
  }
  try {
    var file = DriveApp.getFileById(fileId);
    var mimeType = String(file.getMimeType() || '').toLowerCase();
    var sizeBytes = Number(file.getSize());
    if (file.isTrashed()
        || !isPinPhotoReadMimeType_(mimeType)
        || !Number.isSafeInteger(sizeBytes)
        || sizeBytes < 1
        || sizeBytes > DRIVE_PHOTO_IMPORT_MAX_FILE_BYTES) {
      throw new Error('pin photo validation failed');
    }
    return file.getBlob();
  } catch (_error) {
    throw importItemError_('PIN_PHOTO_NOT_FOUND', 'pin photo is unavailable.', false);
  }
}

function readPinAudioBlobByPinId_(pinId) {
  var normalizedPinId = String(pinId || '');
  if (!normalizedPinId) {
    throw audioStorageImportError_('PIN_AUDIO_NOT_FOUND', 'pin audio is unavailable.', false);
  }
  var structure = audioStorageStructureForRead_();
  var target = findMapInfoRowByPinId_(openMapInfoSheet_(), normalizedPinId);
  var audioId = target ? String(target.row[15] || '') : '';
  if (!audioId) {
    throw audioStorageImportError_('PIN_AUDIO_NOT_FOUND', 'pin audio is unavailable.', false);
  }
  try {
    var file = DriveApp.getFileById(audioId);
    if (!audioStorageFileIsDirectManaged_(file, structure)) {
      throw new Error('managed audio ownership check failed');
    }
    var sizeBytes = Number(file.getSize());
    if (!Number.isSafeInteger(sizeBytes)
        || sizeBytes < MANAGED_AUDIO_MIN_BYTES
        || sizeBytes > MANAGED_AUDIO_MAX_BYTES) {
      throw new Error('managed audio size check failed');
    }
    return file.getBlob();
  } catch (_error) {
    throw audioStorageImportError_('PIN_AUDIO_NOT_FOUND', 'pin audio is unavailable.', false);
  }
}

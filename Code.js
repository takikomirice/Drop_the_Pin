// ============================================================
//  Drop the Pin! — Google Apps Script バックエンド
// ============================================================

const PinData = (function() {
  const DEFAULT_COLOR = '#e53935';
  const URL_RE = /^https?:\/\/\S+$/i;
  const STATUS_OPTIONS = ['未対応', '対応中', '完了', '保留'];
  const MAX_TAGS = 5;

  function normalizeStatus(value) {
    const s = String(value || '').trim();
    if (s === '') return '';
    if (STATUS_OPTIONS.indexOf(s) === -1) {
      throw new Error('invalid status: ' + s);
    }
    return s;
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
      tags: deserializeTags(row[11] || '')
    };
  }

  function rowsToPins(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const dataRows = rows[0] && rows[0][8] === 'ID' ? rows.slice(1) : rows;
    return dataRows.filter(function(row) { return row && row[8]; }).map(rowToPin);
  }

  return {
    DEFAULT_COLOR: DEFAULT_COLOR,
    STATUS_OPTIONS: STATUS_OPTIONS,
    deserializeLinks: deserializeLinks,
    normalizeLinks: normalizeLinks,
    rowToPin: rowToPin,
    rowsToPins: rowsToPins,
    serializeLinks: serializeLinks,
    serializeTags: serializeTags,
    deserializeTags: deserializeTags,
    normalizeTags: normalizeTags,
    normalizeStatus: normalizeStatus,
    normalizeSearchText: normalizeSearchText,
    toBooleanSetting: toBooleanSetting,
    chooseSpreadsheetId: chooseSpreadsheetId,
    buildFileNameForSave: buildFileNameForSave
  };
})();

const SHEET_NAME = 'map_info';
const CONFIG_SHEET_NAME = 'config';
const SHARE_LINKS_SHEET_NAME = 'share_links';
const DEFAULT_COLOR = PinData.DEFAULT_COLOR;
const DEFAULT_SHARE_LINK_LABEL = 'Drop the Pin!';
const SHARE_LINKS_HEADERS = ['createdAt', 'label', 'token', 'tags', 'tagMode', 'enabled', 'revokedAt', 'colors'];
const SAFE_COLOR_RE = /^#[0-9a-fA-F]{6}$/;

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

  const headers = [
    'タイムスタンプ', 'タイトル', '説明',
    '緯度', '経度', 'ピンの色',
    'ファイルID', '画像URL', 'ID', '参考URL一覧',
    '状態', 'タグ'
  ];
  const looksHeader = sheet.getLastRow() > 0 && (
    sheet.getRange('I1').getValue() === 'ID' ||
    sheet.getRange('A1').getValue() === 'タイムスタンプ'
  );
  if (!looksHeader && sheet.getLastRow() > 0) {
    sheet.insertRowBefore(1);
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  // K/L 列が不足している既存シートを補完する
  if (sheet.getLastColumn() < 12) {
    sheet.getRange(1, 11).setValue('状態');
    sheet.getRange(1, 12).setValue('タグ');
  }
  sheet.getRange('A1:L1')
    .setBackground('#4caf50')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  [160, 180, 250, 90, 90, 90, 200, 350, 230, 320, 100, 200].forEach((width, index) => {
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

  ui.alert(
    '初期設定完了',
    '"' + SHEET_NAME + '" シート、"' + CONFIG_SHEET_NAME + '" シート、"' + SHARE_LINKS_SHEET_NAME + '" シートの準備が整いました。\n\n' +
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
    imageUrl: pin.imageUrl || '',
    timestamp: pin.timestamp || '',
    links: Array.isArray(pin.links) ? pin.links.slice() : [],
    tags: filterPinTagsForShare_(pin, allowedTags)
  };
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
    pins: pins
  };
}


function saveMapData(data) {
  if (!data || !String(data.title || '').trim()) {
    return { ok: false, error: 'title is required' };
  }

  const title = String(data.title).trim();
  const description = String(data.description || '');
  const color = data.color || DEFAULT_COLOR;
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
    PinData.serializeTags(tags)
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
  const row = sheet.getRange(sheetRow, 1, 1, 12).getValues()[0];
  const title = String(data.title).trim();
  const links = PinData.normalizeLinks(data.links || data.referenceUrls || []);

  sheet.getRange(sheetRow, 2, 1, 2).setValues([[title, String(data.description || '')]]);
  sheet.getRange(sheetRow, 6).setValue(data.color || row[5] || DEFAULT_COLOR);
  sheet.getRange(sheetRow, 10).setValue(PinData.serializeLinks(links));

  if (data.status != null) {
    sheet.getRange(sheetRow, 11).setValue(PinData.normalizeStatus(String(data.status)));
  }
  if (data.tags != null) {
    sheet.getRange(sheetRow, 12).setValue(PinData.serializeTags(data.tags));
  }

  if (getRenameFileWithTitle_() && row[6]) {
    renameDriveFileForTitle_(row[6], title);
  }

  return {
    ok: true,
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
  return { ok: true };
}

function unplacePin(data) {
  if (!data || !data.id) return { ok: false, error: 'missing id' };

  const sheet = openMapInfoSheet_();
  const rowIndex = findPinRowIndex_(sheet, data.id);
  if (rowIndex === -1) return { ok: false, error: 'id not found' };

  const sheetRow = rowIndex + 1;
  sheet.getRange(sheetRow, 4, 1, 2).setValues([['', '']]);
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
  data.ids.forEach(function(id) {
    const rowIndex = rows.findIndex(function(row) { return row[8] === id; });
    if (rowIndex === -1) return;
    sheet.getRange(rowIndex + 1, 11).setValue(status);
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
  return { ok: true };
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
    unplacePin: unplacePin,
    bulkUpdatePinStatus: bulkUpdatePinStatus,
    createShareLink: createShareLink,
    listShareLinks: listShareLinks,
    revokeShareLink: revokeShareLink,
    setShareLinkEnabled: setShareLinkEnabled,
    deleteShareLink: deleteShareLink,
    getSharedViewData: getSharedViewData,
    setupSheet: setupSheet
  };
}

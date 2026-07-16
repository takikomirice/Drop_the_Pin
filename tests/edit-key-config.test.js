const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function createSheet(rows = [], name = '') {
  return {
    sheetName: name,
    rows: rows.map((row) => row.slice()),
    activated: false,
    activatedRanges: [],
    hyperlinkRows: [],
    getLastRow() {
      return this.rows.length;
    },
    getLastColumn() {
      return this.rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
    },
    getName() {
      return this.sheetName;
    },
    setName(nextName) {
      this.sheetName = nextName;
      return this;
    },
    getRange(row, column, numRows = 1, numColumns = 1) {
      if (typeof row === 'string') {
        const match = row.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
        if (!match) {
          return {
            setNumberFormat() { return this; },
            setBackground() { return this; },
            setFontColor() { return this; },
            setFontWeight() { return this; }
          };
        }
        const columnNumber = (letters) => letters.split('').reduce(
          (total, letter) => total * 26 + letter.charCodeAt(0) - 64,
          0
        );
        row = Number(match[2]);
        column = columnNumber(match[1]);
        numRows = match[4] ? Number(match[4]) - row + 1 : 1;
        numColumns = match[3] ? columnNumber(match[3]) - column + 1 : 1;
      }
      const range = {
        getValues: () => {
          const values = [];
          for (let r = 0; r < numRows; r += 1) {
            const source = this.rows[row - 1 + r] || [];
            values.push(source.slice(column - 1, column - 1 + numColumns));
          }
          return values;
        },
        setValues: (values) => {
          for (let r = 0; r < numRows; r += 1) {
            const targetRow = row - 1 + r;
            if (!this.rows[targetRow]) this.rows[targetRow] = [];
            for (let c = 0; c < numColumns; c += 1) {
              this.rows[targetRow][column - 1 + c] = values[r][c];
            }
          }
          return range;
        },
        getValue: () => (this.rows[row - 1] || [])[column - 1],
        setValue: (value) => {
          if (!this.rows[row - 1]) this.rows[row - 1] = [];
          this.rows[row - 1][column - 1] = value;
          return range;
        },
        setBackground: () => range,
        setFontColor: () => range,
        setFontWeight: () => range,
        setNumberFormat: () => range,
        setShowHyperlink: () => {
          this.hyperlinkRows.push([row, column]);
          return range;
        },
        activate: () => {
          this.activatedRanges.push([row, column]);
          return range;
        }
      };
      return range;
    },
    appendRow(row) {
      this.rows.push(row.slice());
    },
    setFrozenRows() {},
    setColumnWidth() {},
    insertRowBefore(row) {
      this.rows.splice(row - 1, 0, []);
    },
    activate() {
      this.activated = true;
      return this;
    }
  };
}

function loadApi(configRows = [['設定項目', '値', '説明']], options = {}) {
  const sheets = {};
  const sheetOrder = [];
  const initialSheets = options.initialSheets || [{ name: 'config', rows: configRows }];
  initialSheets.forEach((definition) => {
    const sheet = createSheet(definition.rows || [], definition.name);
    sheets[definition.name] = sheet;
    sheetOrder.push(sheet);
  });
  const menuItems = [];
  const cache = {};
  const templates = [];
  const ui = {
    Button: { OK: 'OK', YES: 'YES', NO: 'NO' },
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' },
    createMenu: (name) => ({
      name,
      addItem(label, functionName) {
        menuItems.push({ label, functionName });
        return this;
      },
      addToUi() {
        return this;
      }
    }),
    alerts: [],
    alert(...args) {
      this.alerts.push(args);
      return options.alertResult || 'OK';
    },
    prompt: () => ({
      getSelectedButton: () => options.promptButton || 'CANCEL',
      getResponseText: () => options.promptText || ''
    })
  };
  let activeSheet = sheetOrder[0] || null;
  const spreadsheet = {
    getSheetByName: (name) => sheetOrder.find((sheet) => sheet.getName() === name) || null,
    getSheets: () => sheetOrder.slice(),
    insertSheet: (name) => {
      const sheet = createSheet([], name);
      sheets[name] = sheet;
      sheetOrder.push(sheet);
      return sheet;
    },
    getActiveSheet: () => activeSheet,
    setActiveSheet: (sheet) => {
      activeSheet = sheet;
      return sheet;
    },
    moveActiveSheet: (position) => {
      const currentIndex = sheetOrder.indexOf(activeSheet);
      if (currentIndex === -1) return;
      sheetOrder.splice(currentIndex, 1);
      sheetOrder.splice(Math.max(0, Math.min(position - 1, sheetOrder.length)), 0, activeSheet);
    }
  };
  const context = {
    Logger: { log() {} },
    Utilities: { getUuid: () => '11111111-2222-4333-8444-555555555555' },
    ScriptApp: { getService: () => ({
      getUrl: () => {
        if (options.serviceUrlError) throw new Error('not deployed');
        return options.serviceUrl === undefined
          ? 'https://script.google.com/macros/s/deploy/exec'
          : options.serviceUrl;
      }
    }) },
    SpreadsheetApp: {
      getUi: () => ui,
      getActiveSpreadsheet: () => spreadsheet
    },
    HtmlService: {
      createTemplateFromFile: (name) => {
        const template = {
          name,
          evaluate: () => ({
            title: '',
            meta: [],
            xFrameMode: '',
            setTitle(title) {
              this.title = title;
              return this;
            },
            addMetaTag(name, value) {
              this.meta.push([name, value]);
              return this;
            },
            setXFrameOptionsMode(mode) {
              this.xFrameMode = mode;
              return this;
            }
          })
        };
        templates.push(template);
        return template;
      },
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' },
      createHtmlOutput: (html) => ({ html, setWidth() { return this; }, setHeight() { return this; } })
    },
    CacheService: {
      getScriptCache: () => ({
        put: (key, value, ttl) => {
          cache[key] = { value, ttl };
        }
      })
    }
  };
  vm.runInNewContext(`${codeJs}
globalThis.__editKeyApi = {
  EDIT_KEY_CONFIG_KEY: EDIT_KEY_CONFIG_KEY,
  WEB_APP_URL_CONFIG_KEY: WEB_APP_URL_CONFIG_KEY,
  EDIT_URL_CONFIG_KEY: EDIT_URL_CONFIG_KEY,
  EDIT_TOKEN_TTL_SECONDS: EDIT_TOKEN_TTL_SECONDS,
  generateEditKey_: generateEditKey_,
  normalizeWebAppUrl_: normalizeWebAppUrl_,
  ensureEditUrlConfig_: ensureEditUrlConfig_,
  refreshEditUrlConfig_: refreshEditUrlConfig_,
  buildEditUrl_: buildEditUrl_,
  buildSharedViewUrl_: buildSharedViewUrl_,
  onOpen: onOpen,
  doGet: doGet,
  setupSheet: setupSheet,
  promptWebAppUrl: promptWebAppUrl,
  regenerateEditKeyFromMenu: regenerateEditKeyFromMenu,
  refreshAndOpenEditUrl: refreshAndOpenEditUrl,
  showEditUrlDialog: showEditUrlDialog
};`, context);
  return { api: context.__editKeyApi, sheets, sheetOrder, menuItems, cache, templates, ui };
}

function rowByKey(sheet, key) {
  return sheet.rows.find((row) => row[0] === key);
}

function rowsByKey(sheet, key) {
  return sheet.rows.filter((row) => row[0] === key);
}

test('edit key config constants and URL normalizer are present', () => {
  const { api } = loadApi();

  assert.equal(api.EDIT_KEY_CONFIG_KEY, 'EDIT_KEY');
  assert.equal(api.WEB_APP_URL_CONFIG_KEY, 'WEB_APP_URL');
  assert.equal(api.EDIT_URL_CONFIG_KEY, 'EDIT_URL');
  assert.match(api.generateEditKey_(), /^ed_[A-Za-z0-9]+$/);
  assert.equal(
    api.normalizeWebAppUrl_(' https://script.google.com/macros/s/xxxxx/exec?x=1#hash/ '),
    'https://script.google.com/macros/s/xxxxx/exec'
  );
  assert.equal(
    api.normalizeWebAppUrl_('https://script.google.com/a/e.osakamanabi.jp/macros/s/xxxxx/exec/'),
    'https://script.google.com/a/macros/e.osakamanabi.jp/s/xxxxx/exec'
  );
  assert.ok(codeJs.includes('function getConfiguredWebAppUrl_()'));
});

test('setupSheet creates one EDIT_URL row while preserving an existing edit key', () => {
  let loaded = loadApi();

  loaded.api.setupSheet();
  assert.match(rowByKey(loaded.sheets.config, 'EDIT_KEY')[1], /^ed_[A-Za-z0-9]+$/);
  assert.equal(rowByKey(loaded.sheets.config, 'WEB_APP_URL')[1], '');
  assert.equal(rowsByKey(loaded.sheets.config, 'EDIT_URL').length, 1);
  assert.equal(rowByKey(loaded.sheets.config, 'EDIT_URL')[2], '編集用WebアプリURL。知っている人は編集できるため、共有範囲に注意してください。');

  loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'existing-key', '既存説明']
  ]);
  loaded.api.setupSheet();
  assert.equal(rowByKey(loaded.sheets.config, 'EDIT_KEY')[1], 'existing-key');
  assert.equal(rowByKey(loaded.sheets.config, 'WEB_APP_URL')[1], '');
  assert.equal(rowsByKey(loaded.sheets.config, 'EDIT_URL').length, 1);
  loaded.api.setupSheet();
  assert.equal(rowsByKey(loaded.sheets.config, 'EDIT_URL').length, 1);
});

test('app URLs use configured WEB_APP_URL and fall back to ScriptApp URL', () => {
  let loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'class-edit-key', ''],
    ['WEB_APP_URL', 'https://script.google.com/a/e.osakamanabi.jp/macros/s/xxxxx/exec?ignored=1#hash', '']
  ]);

  assert.equal(
    loaded.api.buildEditUrl_(),
    'https://script.google.com/a/macros/e.osakamanabi.jp/s/xxxxx/exec?mode=edit&editKey=class-edit-key'
  );
  assert.equal(
    loaded.api.buildSharedViewUrl_('share token'),
    'https://script.google.com/a/macros/e.osakamanabi.jp/s/xxxxx/exec?view=shared&token=share%20token'
  );
  loaded.api.doGet({ parameter: {} });
  assert.equal(
    loaded.templates[0].execUrl,
    'https://script.google.com/a/macros/e.osakamanabi.jp/s/xxxxx/exec'
  );

  loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'class-edit-key', ''],
    ['WEB_APP_URL', '', '']
  ]);

  assert.equal(
    loaded.api.buildEditUrl_(),
    'https://script.google.com/macros/s/deploy/exec?mode=edit&editKey=class-edit-key'
  );
  assert.equal(
    loaded.api.buildSharedViewUrl_('share token'),
    'https://script.google.com/macros/s/deploy/exec?view=shared&token=share%20token'
  );
  loaded.api.doGet({ parameter: {} });
  assert.equal(loaded.templates[0].execUrl, 'https://script.google.com/macros/s/deploy/exec');
});

test('refreshEditUrlConfig writes a normalized, encoded EDIT_URL and replaces a stale value', () => {
  const loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'class edit/key?', ''],
    ['WEB_APP_URL', 'https://script.google.com/a/e.osakamanabi.jp/macros/s/xxxxx/exec?ignored=1#hash', ''],
    ['EDIT_URL', 'stale-url', '']
  ]);

  const first = loaded.api.refreshEditUrlConfig_();
  assert.equal(first.url,
    'https://script.google.com/a/macros/e.osakamanabi.jp/s/xxxxx/exec?mode=edit&editKey=class%20edit%2Fkey%3F'
  );
  assert.equal(rowByKey(loaded.sheets.config, 'EDIT_URL')[1], first.url);
  assert.deepEqual(loaded.sheets.config.hyperlinkRows.at(-1), [4, 2]);

  rowByKey(loaded.sheets.config, 'WEB_APP_URL')[1] = 'https://script.google.com/macros/s/next/exec///?old=1';
  const second = loaded.api.refreshEditUrlConfig_();
  assert.equal(second.url,
    'https://script.google.com/macros/s/next/exec?mode=edit&editKey=class%20edit%2Fkey%3F'
  );
});

test('setupSheet clears EDIT_URL without failing when the service URL is unavailable', () => {
  const loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'class-edit-key', ''],
    ['WEB_APP_URL', '', ''],
    ['EDIT_URL', 'old-url', '']
  ], { serviceUrlError: true });

  assert.doesNotThrow(() => loaded.api.setupSheet());
  assert.equal(rowByKey(loaded.sheets.config, 'EDIT_URL')[1], '');
});

test('spreadsheet menu exposes edit URL operations', () => {
  const { api, menuItems } = loadApi();

  api.onOpen();

  assert.deepEqual(menuItems.map((item) => item.label), [
    '初期設定',
    '編集URLを更新・開く',
    '編集キーを再生成'
  ]);
});

test('setupSheet puts config first without disturbing existing data or other sheet order', () => {
  const loaded = loadApi([], {
    initialSheets: [
      { name: 'notes', rows: [['keep-notes']] },
      { name: 'config', rows: [
        ['設定項目', '値', '説明'],
        ['CUSTOM_FORMULA', '=1+1', '維持する説明'],
        ['EDIT_KEY', 'existing-key', '既存説明'],
        ['EDIT_URL', '=HYPERLINK("https://example.com","編集")', '式を維持する説明']
      ] },
      { name: 'archive', rows: [['keep-archive']] }
    ]
  });

  loaded.api.setupSheet();
  assert.equal(loaded.sheetOrder[0].getName(), 'config');
  assert.deepEqual(
    loaded.sheetOrder.filter((sheet) => ['notes', 'archive'].includes(sheet.getName())).map((sheet) => sheet.getName()),
    ['notes', 'archive']
  );
  assert.deepEqual(rowByKey(loaded.sheets.config, 'CUSTOM_FORMULA'), [
    'CUSTOM_FORMULA', '=1+1', '維持する説明'
  ]);
  assert.deepEqual(rowByKey(loaded.sheets.config, 'EDIT_URL'), [
    'EDIT_URL', '=HYPERLINK("https://example.com","編集")', '式を維持する説明'
  ]);

  loaded.api.setupSheet();
  assert.equal(loaded.sheetOrder.filter((sheet) => sheet.getName() === 'config').length, 1);
  assert.equal(loaded.sheetOrder.filter((sheet) => sheet.getName() === 'map_info').length, 1);
  assert.deepEqual(rowByKey(loaded.sheets.config, 'CUSTOM_FORMULA'), [
    'CUSTOM_FORMULA', '=1+1', '維持する説明'
  ]);
  assert.deepEqual(rowByKey(loaded.sheets.config, 'EDIT_URL'), [
    'EDIT_URL', '=HYPERLINK("https://example.com","編集")', '式を維持する説明'
  ]);
});

test('setupSheet safely renames a sole blank initial sheet to config before creating map_info', () => {
  const loaded = loadApi([], {
    initialSheets: [{ name: 'シート1', rows: [] }]
  });

  loaded.api.setupSheet();
  assert.equal(loaded.sheetOrder[0].getName(), 'config');
  assert.equal(loaded.sheetOrder.filter((sheet) => sheet.getName() === 'config').length, 1);
  assert.equal(loaded.sheetOrder[1].getName(), 'map_info');
  assert.deepEqual(loaded.sheetOrder[0].rows[0].slice(0, 3), ['設定項目', '値', '説明']);
});

test('menu edit URL operation and its compatibility wrapper select EDIT_URL without a modal', () => {
  const loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'class-edit-key', ''],
    ['WEB_APP_URL', 'https://script.google.com/macros/s/deploy/exec', '']
  ]);

  loaded.api.refreshAndOpenEditUrl();
  const editUrlRow = loaded.sheets.config.rows.findIndex((row) => row[0] === 'EDIT_URL') + 1;
  assert.equal(loaded.sheets.config.activated, true);
  assert.deepEqual(loaded.sheets.config.activatedRanges.at(-1), [editUrlRow, 2]);
  assert.equal(loaded.ui.alerts.length, 0);

  loaded.api.showEditUrlDialog();
  assert.deepEqual(loaded.sheets.config.activatedRanges.at(-1), [editUrlRow, 2]);
  assert.equal(codeJs.includes('showModalDialog'), false);
});

test('prompting for WEB_APP_URL and regenerating EDIT_KEY refresh EDIT_URL without retaining an old URL', () => {
  let loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'old-key', ''],
    ['WEB_APP_URL', 'https://script.google.com/macros/s/old/exec', '']
  ], { promptButton: 'OK', promptText: 'https://script.google.com/macros/s/new/exec?ignored=1' });
  loaded.api.refreshEditUrlConfig_();
  loaded.api.promptWebAppUrl();
  assert.equal(rowByKey(loaded.sheets.config, 'EDIT_URL')[1], 'https://script.google.com/macros/s/new/exec?mode=edit&editKey=old-key');

  loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'old-key', ''],
    ['WEB_APP_URL', 'https://script.google.com/macros/s/deploy/exec', '']
  ], { alertResult: 'YES' });
  const oldUrl = loaded.api.refreshEditUrlConfig_().url;
  loaded.api.regenerateEditKeyFromMenu();
  assert.notEqual(rowByKey(loaded.sheets.config, 'EDIT_URL')[1], oldUrl);
  assert.match(rowByKey(loaded.sheets.config, 'EDIT_URL')[1], /editKey=ed_/);
});

test('doGet issues edit tokens only for matching edit mode and key', () => {
  let loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'class-edit-key', ''],
    ['WEB_APP_URL', '', ''],
    ['EDIT_URL', 'https://example.invalid/?mode=edit&editKey=wrong-key', '']
  ]);
  loaded.api.doGet({ parameter: { mode: 'edit', editKey: 'class-edit-key' } });
  assert.match(loaded.templates[0].editToken, /^edt_[A-Za-z0-9]+$/);
  assert.equal(Object.values(loaded.cache)[0].ttl, loaded.api.EDIT_TOKEN_TTL_SECONDS);

  [
    {},
    { mode: 'public', editKey: 'class-edit-key' },
    { view: 'shared', mode: 'edit', editKey: 'class-edit-key' },
    { mode: 'edit', editKey: 'wrong-key' }
  ].forEach((parameter) => {
    loaded = loadApi([
      ['設定項目', '値', '説明'],
      ['EDIT_KEY', 'class-edit-key', ''],
      ['WEB_APP_URL', '', '']
    ]);
    loaded.api.doGet({ parameter });
    assert.equal(loaded.templates[0].editToken, '');
  });
});

test('EDIT_URL is not used for authorization or logged by the refresh helper', () => {
  const issueStart = codeJs.indexOf('function issueEditTokenFromRequest_');
  const issueEnd = codeJs.indexOf('function openDataSpreadsheet_', issueStart);
  const refreshStart = codeJs.indexOf('function refreshEditUrlConfig_');
  const refreshEnd = codeJs.indexOf('function promptWebAppUrl', refreshStart);
  assert.equal(codeJs.slice(issueStart, issueEnd).includes('EDIT_URL_CONFIG_KEY'), false);
  assert.equal(codeJs.slice(refreshStart, refreshEnd).includes('Logger.'), false);
});

test('index receives the edit token template value', () => {
  assert.ok(indexHtml.includes('window.__EDIT_TOKEN__ = <?!= JSON.stringify(editToken || \'\') ?>;'));
});

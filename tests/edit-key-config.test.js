const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');
const indexHtml = fs.readFileSync(path.join(root, 'index.html'), 'utf8');

function createSheet(rows = []) {
  return {
    rows: rows.map((row) => row.slice()),
    getLastRow() {
      return this.rows.length;
    },
    getRange(row, column, numRows = 1, numColumns = 1) {
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
        setFontWeight: () => range
      };
      return range;
    },
    appendRow(row) {
      this.rows.push(row.slice());
    },
    setFrozenRows() {},
    setColumnWidth() {}
  };
}

function loadApi(configRows = [['設定項目', '値', '説明']]) {
  const sheets = {
    config: createSheet(configRows)
  };
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
    alert: () => 'OK'
  };
  const context = {
    Logger: { log() {} },
    Utilities: { getUuid: () => '11111111-2222-4333-8444-555555555555' },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/deploy/exec' }) },
    SpreadsheetApp: {
      getUi: () => ui,
      getActiveSpreadsheet: () => ({
        getSheetByName: (name) => sheets[name] || null,
        insertSheet: (name) => {
          sheets[name] = createSheet();
          return sheets[name];
        }
      })
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
  EDIT_TOKEN_TTL_SECONDS: EDIT_TOKEN_TTL_SECONDS,
  generateEditKey_: generateEditKey_,
  normalizeWebAppUrl_: normalizeWebAppUrl_,
  ensureEditUrlConfig_: ensureEditUrlConfig_,
  buildEditUrl_: buildEditUrl_,
  buildSharedViewUrl_: buildSharedViewUrl_,
  onOpen: onOpen,
  doGet: doGet
};`, context);
  return { api: context.__editKeyApi, sheets, menuItems, cache, templates };
}

function rowByKey(sheet, key) {
  return sheet.rows.find((row) => row[0] === key);
}

test('edit key config constants and URL normalizer are present', () => {
  const { api } = loadApi();

  assert.equal(api.EDIT_KEY_CONFIG_KEY, 'EDIT_KEY');
  assert.equal(api.WEB_APP_URL_CONFIG_KEY, 'WEB_APP_URL');
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

test('edit config rows are created while preserving an existing edit key', () => {
  let loaded = loadApi();

  loaded.api.ensureEditUrlConfig_();
  assert.match(rowByKey(loaded.sheets.config, 'EDIT_KEY')[1], /^ed_[A-Za-z0-9]+$/);
  assert.equal(rowByKey(loaded.sheets.config, 'WEB_APP_URL')[1], '');

  loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'existing-key', '既存説明']
  ]);
  loaded.api.ensureEditUrlConfig_();
  assert.equal(rowByKey(loaded.sheets.config, 'EDIT_KEY')[1], 'existing-key');
  assert.equal(rowByKey(loaded.sheets.config, 'WEB_APP_URL')[1], '');
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

test('spreadsheet menu exposes edit URL operations', () => {
  const { api, menuItems } = loadApi();

  api.onOpen();

  assert.deepEqual(menuItems.map((item) => item.label), [
    '初期設定',
    'WebアプリURLを設定',
    '編集URLを表示',
    '編集キーを再生成'
  ]);
});

test('doGet issues edit tokens only for matching edit mode and key', () => {
  let loaded = loadApi([
    ['設定項目', '値', '説明'],
    ['EDIT_KEY', 'class-edit-key', ''],
    ['WEB_APP_URL', '', '']
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

test('index receives the edit token template value', () => {
  assert.ok(indexHtml.includes('window.__EDIT_TOKEN__ = <?!= JSON.stringify(editToken || \'\') ?>;'));
});

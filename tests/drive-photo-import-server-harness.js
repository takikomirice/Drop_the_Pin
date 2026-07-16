const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const codeJs = fs.readFileSync(path.join(root, 'Code.js'), 'utf8');

function iterator(values) {
  let index = 0;
  return {
    hasNext() { return index < values.length; },
    next() {
      if (index >= values.length) throw new Error('iterator exhausted');
      return values[index++];
    }
  };
}

function createHarness(options = {}) {
  const audit = {
    folderReads: [], fileReads: [], blobReads: 0, byteReads: 0, writes: [], sheetReads: []
  };
  const folders = new Map();
  const files = new Map();
  const rootId = options.rootId || 'root_ABCDEFGHIJKLMNO';

  function mutator(name) {
    return function forbiddenMutation() {
      audit.writes.push(name);
      throw new Error(`${name} must not be called`);
    };
  }

  function addFolder(id, name, parentIds = [], extra = {}) {
    const folder = {
      getId: () => id,
      getName: () => name,
      getParents: () => {
        if (extra.parentError) throw new Error('private parent failure ' + id);
        return iterator(parentIds.map((parentId) => folders.get(parentId)).filter(Boolean));
      },
      getFolders: () => iterator(
        Array.from(folders.values()).filter((candidate) => candidate.__parentIds.includes(id))
      ),
      getFiles: () => iterator(
        Array.from(files.values()).filter((candidate) => candidate.__parentIds.includes(id))
      ),
      isTrashed: () => {
        if (extra.trashError) throw new Error('private folder trash failure ' + id);
        return extra.trashed === true;
      },
      createFile: mutator('createFile'),
      createFolder: mutator('createFolder'),
      setName: mutator('setName'),
      setSharing: mutator('setSharing'),
      setTrashed: mutator('setTrashed'),
      moveTo: mutator('moveTo'),
      __parentIds: parentIds.slice()
    };
    folders.set(id, folder);
    return folder;
  }

  function addFile(id, name, mimeType, bytes, parentIds, extra = {}) {
    const values = Array.from(bytes || []);
    const file = {
      getId: () => id,
      getName: () => name,
      getMimeType: () => mimeType,
      getSize: () => extra.metadataSize == null ? values.length : extra.metadataSize,
      getLastUpdated: () => new Date(extra.modifiedAt || '2026-07-12T01:02:03.000Z'),
      getParents: () => {
        if (extra.parentError) throw new Error('private file parent failure ' + id);
        return iterator(parentIds.map((parentId) => folders.get(parentId)).filter(Boolean));
      },
      isTrashed: () => {
        if (extra.trashError) throw new Error('private file trash failure ' + id);
        return extra.trashed === true;
      },
      getBlob: () => {
        audit.blobReads += 1;
        if (extra.blobError) throw new Error('private blob failure ' + id);
        return {
          getBytes() {
            audit.byteReads += 1;
            return values.slice();
          }
        };
      },
      makeCopy: mutator('makeCopy'),
      moveTo: mutator('moveTo'),
      setName: mutator('setName'),
      setSharing: mutator('setSharing'),
      setTrashed: mutator('setTrashed'),
      __parentIds: parentIds.slice()
    };
    files.set(id, file);
    return file;
  }

  addFolder(rootId, 'Root');

  const configSheet = {
    getLastRow: () => options.rootMissing ? 1 : 2,
    getRange(row, column, rowCount, columnCount) {
      assert.equal(row, 2);
      assert.equal(column, 1);
      assert.equal(rowCount, 1);
      assert.equal(columnCount, 2);
      return {
        getValues: () => [[
          'IMAGE_DRIVE_URL',
          options.rootUrl || `https://drive.google.com/drive/folders/${rootId}`
        ]]
      };
    }
  };

  const readSheets = new Map(Object.entries(options.sheets || {}).map(([name, sourceRows]) => {
    const rows = sourceRows.map((row) => row.slice());
    const sheet = {
      getLastRow() { return rows.length; },
      getLastColumn() { return rows.reduce((max, row) => Math.max(max, row.length), 0); },
      getDataRange() {
        return sheet.getRange(1, 1, Math.max(1, sheet.getLastRow()), Math.max(1, sheet.getLastColumn()));
      },
      getRange(row, column, rowCount = 1, columnCount = 1) {
        return {
          getValues() {
            audit.sheetReads.push({ name, row, column, rowCount, columnCount });
            if (options.failSheetRead) throw new Error('private spreadsheet read detail');
            return Array.from({ length: rowCount }, (_unused, rowOffset) =>
              Array.from({ length: columnCount }, (_unusedColumn, columnOffset) =>
                (rows[row - 1 + rowOffset] || [])[column - 1 + columnOffset] ?? ''
              )
            );
          },
          setValues: mutator(`setValues:${name}`),
          setValue: mutator(`setValue:${name}`)
        };
      }
    };
    return [name, sheet];
  }));

  const context = {
    console,
    Buffer,
    Date,
    JSON,
    Math,
    Number,
    Object,
    String,
    Array,
    Error,
    RegExp,
    Set,
    Map,
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        return { getSheetByName(name) {
          if (name === 'config') return options.rootMissing ? null : configSheet;
          return readSheets.get(name) || null;
        } };
      }
    },
    CacheService: {
      getScriptCache() {
        return { get: (key) => key === 'EDIT_TOKEN_valid-token' ? '1' : null };
      }
    },
    DriveApp: {
      getFolderById(id) {
        audit.folderReads.push(id);
        const folder = folders.get(id);
        if (!folder) throw new Error('private folder not found ' + id);
        return folder;
      },
      getFileById(id) {
        audit.fileReads.push(id);
        const file = files.get(id);
        if (!file) throw new Error('private file not found ' + id);
        return file;
      }
    },
    Utilities: {
      base64Encode(bytes) {
        return Buffer.from(bytes.map((value) => value < 0 ? value + 256 : value)).toString('base64');
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(
    codeJs + '\n' + [
      'globalThis.__drivePhotoImportApi = {',
      '  listDrivePhotoImportFolder: typeof listDrivePhotoImportFolder === "function" ? listDrivePhotoImportFolder : null,',
      '  readDrivePhotoImportFile: typeof readDrivePhotoImportFile === "function" ? readDrivePhotoImportFile : null,',
      '  isDriveFolderWithinRoot_: typeof isDriveFolderWithinRoot_ === "function" ? isDriveFolderWithinRoot_ : null,',
      '  isDriveFileWithinRoot_: typeof isDriveFileWithinRoot_ === "function" ? isDriveFileWithinRoot_ : null',
      '};'
    ].join('\n'),
    context
  );

  return {
    api: context.__drivePhotoImportApi,
    audit,
    folders,
    files,
    rootId,
    addFolder,
    addFile,
    tokenPayload(value) { return { ...value, __editToken: 'valid-token' }; }
  };
}

module.exports = { createHarness };

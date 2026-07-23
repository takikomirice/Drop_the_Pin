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
    folderReads: [], fileReads: [], blobReads: 0, byteReads: 0, writes: [], sheetReads: [],
    locks: { attempts: 0, releases: 0, held: false, maxDepth: 0, queuedRuns: 0 }
  };
  const folders = new Map();
  const files = new Map();
  const rootId = options.rootId || 'root_ABCDEFGHIJKLMNO';
  let nextFolderId = 0;
  let nextFileId = 0;
  let nextFolderCreateHook = null;
  const concurrentQueue = [];

  function mutator(name) {
    return function forbiddenMutation() {
      audit.writes.push(name);
      throw new Error(`${name} must not be called`);
    };
  }

  function addFolder(id, name, parentIds = [], extra = {}) {
    const folder = {
      getId: () => id,
      getName: () => folder.__name,
      getParents: () => {
        if (extra.parentError) throw new Error('private parent failure ' + id);
        return iterator(folder.__parentIds.map((parentId) => folders.get(parentId)).filter(Boolean));
      },
      getFolders: () => iterator(
        Array.from(folders.values()).filter((candidate) => candidate.__parentIds.includes(id))
      ),
      getFiles: () => iterator(
        Array.from(files.values()).filter((candidate) => candidate.__parentIds.includes(id))
      ),
      getFilesByName: (value) => iterator(
        Array.from(files.values()).filter((candidate) => candidate.__parentIds.includes(id)
          && candidate.__name === String(value) && !candidate.__trashed)
      ),
      isTrashed: () => {
        if (extra.trashError) throw new Error('private folder trash failure ' + id);
        return extra.trashed === true;
      },
      createFile(blob) {
        const values = blob && typeof blob.getBytes === 'function'
          ? blob.getBytes() : Array.from(blob && blob.bytes || []);
        const fileName = blob && typeof blob.getName === 'function'
          ? blob.getName() : String(blob && blob.name || '');
        const mimeType = blob && typeof blob.getContentType === 'function'
          ? blob.getContentType() : String(blob && blob.mime || 'application/octet-stream');
        const fileId = `created_file_${String(++nextFileId).padStart(12, '0')}`;
        audit.writes.push({ method: 'createFile', parentId: id, name: fileName });
        return addFile(fileId, fileName, mimeType, values, [id]);
      },
      createFolder(value) {
        if (nextFolderCreateHook) {
          const hook = nextFolderCreateHook;
          nextFolderCreateHook = null;
          hook();
        }
        const childId = `created_folder_${String(++nextFolderId).padStart(12, '0')}`;
        audit.writes.push({ method: 'createFolder', parentId: id, name: String(value) });
        return addFolder(childId, String(value), [id]);
      },
      setName(value) {
        audit.writes.push({ method: 'setFolderName', id, name: String(value) });
        folder.__name = String(value);
        return folder;
      },
      setSharing: mutator('setSharing'),
      setTrashed: mutator('setTrashed'),
      moveTo: mutator('moveTo'),
      __name: String(name),
      __parentIds: parentIds.slice(),
      __trashed: extra.trashed === true
    };
    folders.set(id, folder);
    return folder;
  }

  function addFile(id, name, mimeType, bytes, parentIds, extra = {}) {
    const values = Array.from(bytes || []);
    const file = {
      getId: () => id,
      getName: () => file.__name,
      getMimeType: () => file.__mimeType,
      getSize: () => extra.metadataSize == null ? values.length : extra.metadataSize,
      getLastUpdated: () => new Date(extra.modifiedAt || '2026-07-12T01:02:03.000Z'),
      getParents: () => {
        if (extra.parentError) throw new Error('private file parent failure ' + id);
        return iterator(file.__parentIds.map((parentId) => folders.get(parentId)).filter(Boolean));
      },
      isTrashed: () => {
        if (extra.trashError) throw new Error('private file trash failure ' + id);
        return file.__trashed;
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
      moveTo(targetFolder) {
        const targetId = targetFolder && targetFolder.getId();
        audit.writes.push({ method: 'moveTo', id, parentId: targetId });
        file.__parentIds = [String(targetId)];
        return file;
      },
      setName(value) {
        audit.writes.push({ method: 'setFileName', id, name: String(value) });
        file.__name = String(value);
        return file;
      },
      setSharing: mutator('setSharing'),
      setTrashed(value) {
        audit.writes.push({ method: 'setTrashed', id, value: !!value });
        file.__trashed = !!value;
        return file;
      },
      __name: String(name),
      __mimeType: String(mimeType),
      __parentIds: parentIds.slice(),
      __trashed: extra.trashed === true,
      __bytes: values
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
    LockService: {
      getScriptLock() {
        return {
          tryLock() {
            audit.locks.attempts += 1;
            if (audit.locks.held) return false;
            audit.locks.held = true;
            audit.locks.maxDepth = Math.max(audit.locks.maxDepth, 1);
            return true;
          },
          releaseLock() {
            assert.equal(audit.locks.held, true, 'media structure lock must be held before release');
            audit.locks.releases += 1;
            audit.locks.held = false;
            const queued = concurrentQueue.shift();
            if (queued) queued();
          }
        };
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
      newBlob(bytes, mimeType, name) {
        const values = Array.from(bytes || []);
        return {
          bytes: values,
          mime: String(mimeType || ''),
          name: String(name || ''),
          getBytes: () => values.slice(),
          getContentType: () => String(mimeType || ''),
          getName: () => String(name || '')
        };
      },
      base64Encode(bytes) {
        return Buffer.from(bytes.map((value) => value < 0 ? value + 256 : value)).toString('base64');
      }
    }
  };
  vm.createContext(context);
  vm.runInContext(
    codeJs + '\n' + [
      'globalThis.__drivePhotoImportApi = {',
      '  ensureMediaDriveStructure: typeof ensureMediaDriveStructure === "function" ? ensureMediaDriveStructure : null,',
      '  ensureMediaDriveStructure_: typeof ensureMediaDriveStructure_ === "function" ? ensureMediaDriveStructure_ : null,',
      '  listDriveMediaInbox: typeof listDriveMediaInbox === "function" ? listDriveMediaInbox : null,',
      '  readDriveAudioImportFile: typeof readDriveAudioImportFile === "function" ? readDriveAudioImportFile : null,',
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
    directFolderNames(parentId) {
      return Array.from(folders.values())
        .filter((folder) => folder.__parentIds.includes(parentId) && !folder.__trashed)
        .map((folder) => folder.__name).sort();
    },
    directFileNames(parentId) {
      return Array.from(files.values())
        .filter((file) => file.__parentIds.includes(parentId) && !file.__trashed)
        .map((file) => file.__name).sort();
    },
    folderId(name, parentId = rootId) {
      const match = Array.from(folders.values()).find((folder) =>
        folder.__parentIds.includes(parentId) && !folder.__trashed && folder.__name === name);
      return match ? match.getId() : '';
    },
    fileBytes(name, parentId = rootId) {
      const match = Array.from(files.values()).find((file) =>
        file.__parentIds.includes(parentId) && !file.__trashed && file.__name === name);
      return match ? match.__bytes.slice() : null;
    },
    replaceFile(id, name, mimeType, bytes, parentIds, extra) {
      files.delete(id);
      return addFile(id, name, mimeType, bytes, parentIds, extra);
    },
    onNextFolderCreate(callback) { nextFolderCreateHook = callback; },
    runConcurrent(callback) {
      if (!audit.locks.held) return callback();
      audit.locks.queuedRuns += 1;
      concurrentQueue.push(callback);
      return undefined;
    },
    tokenPayload(value) { return { ...value, __editToken: 'valid-token' }; }
  };
}

module.exports = { createHarness };

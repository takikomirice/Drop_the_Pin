const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const target = require('../gas-project.json');
const sourceFiles = ['Code.js', 'index.html', 'shared.html'];
const remoteRoot = path.join(root, '.codex-remote');
const configPath = path.join(root, '.clasp.json');

function validateTarget(config, expected) {
  if (config.scriptId !== expected.scriptId) throw new Error('scriptId mismatch. Run gas:setup before continuing.');
  if (!['', '.'].includes(config.rootDir)) throw new Error('rootDir must be the repository root.');
}

function compareFiles(before, after) {
  return [...new Set([...Object.keys(before), ...Object.keys(after)])].sort().filter(name =>
    before[name]?.replace(/\r\n/g, '\n') !== after[name]?.replace(/\r\n/g, '\n'));
}

function assertUnchanged(before, after) {
  const changes = compareFiles(before, after);
  if (changes.length) throw new Error(`Remote changed since review: ${changes.join(', ')}. Take and review a new snapshot.`);
}

function preparePush(local, remote) {
  const expected = [...sourceFiles, 'appsscript.json'].sort();
  if (JSON.stringify(Object.keys(remote).sort()) !== JSON.stringify(expected)) {
    throw new Error('Unexpected remote files. Review manually; no files were pushed.');
  }
  for (const name of sourceFiles) {
    if (typeof local[name] !== 'string' || !local[name].trim()) throw new Error(`Missing local source: ${name}`);
  }
  return Object.fromEntries([...sourceFiles.map(name => [name, local[name]]), ['appsscript.json', remote['appsscript.json']]]);
}

function writeJson(file, value) {
  fs.writeFileSync(file, JSON.stringify(value, null, 2) + '\n');
}

function runNode(script, args) {
  const result = spawnSync(process.execPath, [script, ...args], {
    cwd: root, stdio: 'inherit', timeout: 180000,
    env: { ...process.env, NODE_USE_SYSTEM_CA: '1' }
  });
  if (result.error) throw result.error;
  if (result.status !== 0) throw new Error(`Command failed (exit ${result.status}).`);
}

function clasp(args) {
  runNode(path.join(root, 'node_modules/@google/clasp/build/src/index.js'), args);
}

function readSources(directory) {
  const files = {};
  function walk(folder, prefix = '') {
    for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
      if (entry.name.startsWith('.')) continue;
      const name = prefix + entry.name;
      if (entry.isDirectory()) walk(path.join(folder, entry.name), name + '/');
      else if (/\.(js|gs|html|json)$/.test(name)) files[name] = fs.readFileSync(path.join(folder, entry.name), 'utf8');
    }
  }
  walk(directory);
  return files;
}

function snapshot() {
  fs.mkdirSync(remoteRoot, { recursive: true });
  const directory = fs.mkdtempSync(path.join(remoteRoot, 'snapshot-'));
  const project = path.join(directory, '.clasp.json');
  writeJson(project, { scriptId: target.scriptId, rootDir: directory });
  clasp(['-P', project, 'pull']);
  console.log(`Snapshot: ${path.relative(root, directory)}`);
  return directory;
}

function reviewedSnapshot(argument) {
  if (!argument) throw new Error('Usage: gas:push -- --reviewed .codex-remote/snapshot-...');
  const directory = fs.realpathSync(path.resolve(root, argument));
  const relative = path.relative(fs.realpathSync(remoteRoot), directory);
  if (!relative || relative.startsWith('..') || path.isAbsolute(relative)) throw new Error('Snapshot must be inside .codex-remote.');
  const config = JSON.parse(fs.readFileSync(path.join(directory, '.clasp.json'), 'utf8'));
  if (config.scriptId !== target.scriptId) throw new Error('Snapshot scriptId mismatch.');
  return directory;
}

function main(args) {
  const command = args[0];
  if (command === 'setup') {
    fs.mkdirSync(remoteRoot, { recursive: true });
    if (fs.existsSync(configPath)) {
      const backup = fs.mkdtempSync(path.join(remoteRoot, 'previous-config-'));
      fs.copyFileSync(configPath, path.join(backup, '.clasp.json'));
      console.log(`Previous config saved: ${path.relative(root, backup)}`);
    }
    writeJson(configPath, { scriptId: target.scriptId, rootDir: '.' });
    console.log(`Connected to ${target.name} (${target.scriptId}).`);
    return;
  }
  if (command === 'links') {
    console.log(`GAS: https://script.google.com/home/projects/${target.scriptId}/edit`);
    console.log(`Sheet: https://docs.google.com/spreadsheets/d/${target.spreadsheetId}/edit`);
    console.log(`Project: https://drive.google.com/drive/folders/${target.projectFolderId}`);
    console.log(`Media: https://drive.google.com/drive/folders/${target.mediaFolderId}`);
    if (target.devUrl) console.log(`Development HEAD: ${target.devUrl}`);
    if (target.webAppUrl) console.log(`Versioned web app: ${target.webAppUrl}`);
    return;
  }
  validateTarget(JSON.parse(fs.readFileSync(configPath, 'utf8')), target);
  if (command === 'status') {
    clasp(['-P', configPath, 'show-file-status']);
  } else if (command === 'doctor') {
    console.log(`Node ${process.version}; target: ${target.name}`);
    for (const name of [...sourceFiles, 'appsscript.json']) fs.accessSync(path.join(root, name));
    clasp(['show-authorized-user']);
    clasp(['-P', configPath, 'list-deployments', target.scriptId]);
  } else if (command === 'snapshot') {
    const directory = snapshot();
    const remote = readSources(directory);
    const local = Object.fromEntries([...sourceFiles, 'appsscript.json'].map(name => [name, fs.readFileSync(path.join(root, name), 'utf8')]));
    console.log(`Local/remote differences: ${compareFiles(remote, local).join(', ') || 'none'}`);
  } else if (command === 'deployments') {
    clasp(['-P', configPath, 'list-deployments', target.scriptId]);
  } else if (command === 'push') {
    if (args.length !== 3 || args[1] !== '--reviewed') throw new Error('Usage: gas:push -- --reviewed .codex-remote/snapshot-...');
    const reviewed = readSources(reviewedSnapshot(args[2]));
    // Run validation before fetching the latest remote to minimize the race window.
    runNode(path.join(root, 'node_modules/npm/bin/npm-cli.js'), ['run', 'validate']);
    const latestDirectory = snapshot();
    const latest = readSources(latestDirectory);
    assertUnchanged(reviewed, latest);
    const local = Object.fromEntries(sourceFiles.map(name => [name, fs.readFileSync(path.join(root, name), 'utf8')]));
    const filesToPush = preparePush(local, latest);
    const staging = fs.mkdtempSync(path.join(remoteRoot, 'push-'));
    // Deployment permissions, timezone and scopes belong to the remote environment.
    for (const [name, content] of Object.entries(filesToPush)) fs.writeFileSync(path.join(staging, name), content);
    const project = path.join(staging, '.clasp.json');
    writeJson(project, { scriptId: target.scriptId, rootDir: staging });
    const ignore = path.join(staging, '.claspignore');
    fs.writeFileSync(ignore, '**/**\n!Code.js\n!index.html\n!shared.html\n!appsscript.json\n');
    const staged = readSources(staging);
    clasp(['-P', project, '-I', ignore, 'push']);
    const verified = readSources(snapshot());
    assertUnchanged(staged, verified);
    console.log('GAS HEAD updated and verified. Versioned web app deployments were not changed.');
  } else {
    throw new Error('Commands: setup, links, status, doctor, snapshot, deployments, push --reviewed <snapshot>');
  }
}

module.exports = { validateTarget, compareFiles, assertUnchanged, preparePush };
if (require.main === module) {
  try { main(process.argv.slice(2)); }
  catch (error) { console.error(error.message); process.exitCode = 1; }
}

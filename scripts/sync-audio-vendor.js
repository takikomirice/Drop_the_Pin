const esbuild = require('esbuild');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const checkOnly = process.argv.slice(2).includes('--check');
const vendorStartSentinel = 'AUDIO_VENDOR_BUNDLE_START';
const vendorEndSentinel = 'AUDIO_VENDOR_BUNDLE_END';
const audioVendorEntrySource = [
  "import { canEncodeAudio, BufferTarget, Output, Mp3OutputFormat, AudioBufferSource } from 'mediabunny';",
  "import { registerMp3Encoder } from '@mediabunny/mp3-encoder';",
  'globalThis.Mediabunny = { canEncodeAudio, BufferTarget, Output, Mp3OutputFormat, AudioBufferSource };',
  'globalThis.MediabunnyMp3Encoder = { registerMp3Encoder };'
].join('\n');

function findPackageDirectory(packageName) {
  let directory = path.dirname(require.resolve(packageName));

  while (directory !== path.dirname(directory)) {
    const packageJsonPath = path.join(directory, 'package.json');
    if (fs.existsSync(packageJsonPath)) {
      const pkg = JSON.parse(fs.readFileSync(packageJsonPath, 'utf8'));
      if (pkg.name === packageName) return directory;
    }
    directory = path.dirname(directory);
  }

  throw new Error(`Could not find installed package directory for ${packageName}.`);
}

function buildVendorRegion() {
  const tempDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'drop-the-pin-audio-vendor-'));
  const entryPath = path.join(tempDirectory, 'audio-vendor-entry.js');

  try {
    fs.writeFileSync(entryPath, audioVendorEntrySource, 'utf8');
    const result = esbuild.buildSync({
      bundle: true,
      entryPoints: [entryPath],
      format: 'iife',
      minify: true,
      nodePaths: [path.join(root, 'node_modules')],
      platform: 'browser',
      target: 'es2022',
      write: false
    });
    const bundle = result.outputFiles[0].text.replace(/<\/script/gi, '<\\/script').trimEnd();
    return `${vendorStartSentinel}\n<script>\n${bundle}\n</script>\n${vendorEndSentinel}`;
  } finally {
    fs.rmSync(tempDirectory, { force: true, recursive: true });
  }
}

function countOccurrences(source, needle) {
  return source.split(needle).length - 1;
}

function replaceVendorRegion(indexSource, vendorRegion) {
  if (countOccurrences(indexSource, vendorStartSentinel) !== 1
      || countOccurrences(indexSource, vendorEndSentinel) !== 1) {
    throw new Error('index.html must contain exactly one audio vendor marker pair.');
  }
  const startIndex = indexSource.indexOf(vendorStartSentinel);
  const endIndex = indexSource.indexOf(vendorEndSentinel);
  if (endIndex <= startIndex) {
    throw new Error('index.html audio vendor markers are out of order.');
  }
  if (indexSource.slice(0, startIndex).trim()) {
    throw new Error('index.html audio vendor region must be at the raw prefix.');
  }

  const currentRegion = indexSource.slice(
    startIndex + vendorStartSentinel.length,
    endIndex
  );
  const scriptMatch = /^(?:\r\n|\n)<script>(?:\r\n|\n)([\s\S]*?)(?:\r\n|\n)<\/script>(?:\r\n|\n)$/.exec(currentRegion);
  if (!scriptMatch) {
    throw new Error('index.html audio vendor region must use a plain script wrapper.');
  }
  if (!scriptMatch[1].trim()) {
    throw new Error('index.html audio vendor source must be nonempty.');
  }

  const afterEndIndex = endIndex + vendorEndSentinel.length;
  if (!/^(?:\r\n|\n)<!DOCTYPE html>/.test(indexSource.slice(afterEndIndex))) {
    throw new Error('index.html audio vendor region must be adjacent to <!DOCTYPE html>.');
  }
  return indexSource.slice(0, startIndex)
    + vendorRegion
    + indexSource.slice(afterEndIndex);
}

function desiredArtifacts() {
  const indexPath = path.join(root, 'index.html');
  const indexSource = fs.readFileSync(indexPath, 'utf8');
  return new Map([
    [
      indexPath,
      replaceVendorRegion(indexSource, buildVendorRegion())
    ],
    [
      path.join(root, 'vendor', 'mediabunny-LICENSE.txt'),
      fs.readFileSync(path.join(findPackageDirectory('mediabunny'), 'LICENSE'), 'utf8')
    ],
    [
      path.join(root, 'vendor', 'mediabunny-mp3-encoder-LICENSE.txt'),
      fs.readFileSync(path.join(findPackageDirectory('@mediabunny/mp3-encoder'), 'LICENSE'), 'utf8')
    ]
  ]);
}

function fileMatches(filePath, content) {
  return fs.existsSync(filePath) && fs.readFileSync(filePath, 'utf8') === content;
}

function syncArtifacts() {
  const mismatches = [];

  for (const [filePath, content] of desiredArtifacts()) {
    if (fileMatches(filePath, content)) continue;
    mismatches.push(path.relative(root, filePath));
    if (!checkOnly) {
      fs.mkdirSync(path.dirname(filePath), { recursive: true });
      fs.writeFileSync(filePath, content, 'utf8');
    }
  }

  if (checkOnly && mismatches.length > 0) {
    console.error(`Audio vendor artifacts are out of date: ${mismatches.join(', ')}`);
    process.exitCode = 1;
  }
}

module.exports = { replaceVendorRegion };

if (require.main === module) {
  syncArtifacts();
}

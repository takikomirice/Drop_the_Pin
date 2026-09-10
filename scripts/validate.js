const { spawnSync } = require('node:child_process');
const path = require('node:path');
const fs = require('node:fs');
const root = path.resolve(__dirname, '..');
const tests = fs.readdirSync(path.join(root, 'tests')).filter(name => name.endsWith('.test.js')).map(name => `tests/${name}`);
const checks = [
  ['scripts/sync-audio-vendor.js', '--check'],
  ['scripts/sync-map-experience.js', '--check'],
  ['--check', 'Code.js'],
  ['--test', ...tests],
  ['node_modules/@playwright/test/cli.js', 'test']
];
function runChecks(commands) {
  for (const args of commands) {
    const result = spawnSync(process.execPath, args, { cwd: root, stdio: 'inherit' });
    if (result.error) console.error(result.error.message);
    if (result.error || result.status !== 0) return result.status || 1;
  }
  return 0;
}
module.exports = { runChecks };
if (require.main === module) process.exitCode = runChecks(checks);

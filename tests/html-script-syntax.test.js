const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function inlineScripts(fileName) {
  const html = fs.readFileSync(path.join(root, fileName), 'utf8');
  const scripts = [];
  const scriptRe = /<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = scriptRe.exec(html)) !== null) {
    const source = match[1].trim();
    if (source) scripts.push(source);
  }
  return scripts;
}

function replaceAppsScriptTemplates(source) {
  return source
    .replace(/<\?!=[\s\S]*?\?>/g, "''")
    .replace(/<\?[\s\S]*?\?>/g, '');
}

['index.html', 'shared.html'].forEach((fileName) => {
  test(`${fileName} inline scripts parse`, () => {
    inlineScripts(fileName).forEach((source, index) => {
      new vm.Script(replaceAppsScriptTemplates(source), { filename: `${fileName}#script${index + 1}` });
    });
  });
});

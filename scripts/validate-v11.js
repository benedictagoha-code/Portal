'use strict';
const fs = require('node:fs');
const vm = require('node:vm');

const read = (name) => fs.readFileSync(name, 'utf8');
const fail = [];
const code = read('Code.gs');
const bridge = read('Bridge.html');
const app = read('index.html');

function checkJavaScript(label, source) {
  try { new vm.Script(source, { filename: label }); }
  catch (error) { fail.push(label + ': ' + error.message); }
}
function firstScript(label, html) {
  const match = html.match(/<script\b[^>]*>([\s\S]*?)<\/script>/i);
  if (!match) { fail.push(label + ': no script block found'); return; }
  checkJavaScript(label, match[1]);
}
checkJavaScript('Code.gs', code);
firstScript('Bridge.html', bridge);
firstScript('index.html', app);

const routeBlock = code.match(/const ROUTES\s*=\s*\{([\s\S]*?)\n    \};/);
const allowBlock = bridge.match(/const ALLOWED_METHODS\s*=\s*new Set\(\[([\s\S]*?)\]\);/);
if (!routeBlock) fail.push('Code.gs: ROUTES map not found');
if (!allowBlock) fail.push('Bridge.html: ALLOWED_METHODS not found');
const routes = routeBlock ? [...routeBlock[1].matchAll(/^\s*([A-Za-z0-9_]+):/gm)].map(x => x[1]) : [];
const allowed = allowBlock ? [...allowBlock[1].matchAll(/'([A-Za-z0-9_]+)'/g)].map(x => x[1]) : [];
const calls = [...app.matchAll(/api\(['"]([^'"]+)/g)].map(x => x[1]).filter((v,i,a) => a.indexOf(v) === i);
for (const name of calls) if (!routes.includes(name)) fail.push('Frontend API has no backend route: ' + name);
for (const name of routes) if (!allowed.includes(name)) fail.push('Backend route missing from bridge allowlist: ' + name);
for (const name of allowed) if (!routes.includes(name)) fail.push('Bridge exposes unknown backend route: ' + name);

const ids = [...app.matchAll(/\bid=["']([^"']+)["']/g)].map(x => x[1]);
const seen = new Set();
for (const id of ids) {
  if (seen.has(id)) fail.push('index.html: duplicate element id ' + id);
  seen.add(id);
}
// Regression tests for lecturer display-name correction.
const nameHelper = code.match(/function normaliseDisplayName_\(name\) \{[\s\\S]*?\n\}/);
if (!nameHelper) fail.push('Code.gs: lecturer display-name normaliser missing');
else {
  const helperSource = nameHelper[0].replace(/^function normaliseDisplayName_/, 'function');
  const normaliseDisplayName = vm.runInNewContext('(' + helperSource + ')');
  const cases = [
    ['Dr Agoha Benedict', 'Professor Benedict Agoha'],
    ['Dr. Agoha Benedict', 'Professor Benedict Agoha'],
    ['Agoha Benedict', 'Professor Benedict Agoha'],
    ['Professor Other Lecturer', 'Professor Other Lecturer']
  ];
  for (const [input, expected] of cases) {
    if (normaliseDisplayName(input) !== expected) fail.push('Display-name regression: ' + input);
  }
}
if (!code.includes('name: normaliseDisplayName_(match.name)')) fail.push('Login does not apply the corrected lecturer display name.');
for (const method of ['createClassroomPoll','getClassroomPolls','voteClassroomPoll','closeClassroomPoll']) {
  if (!routes.includes(method) || !allowed.includes(method) || !code.includes('function ' + method + '(')) fail.push('Classroom poll wiring regression: ' + method);
}
if (!app.includes('onCreateClassroomPoll') || !app.includes('onVoteClassroomPoll')) fail.push('Classroom poll UI handlers missing.');

if (fail.length) {
  console.error('V11 validation failed:\n- ' + fail.join('\n- '));
  process.exit(1);
}
console.log('V11 validation passed.');
console.log('JavaScript syntax: Code.gs, Bridge.html, index.html');
console.log('API wiring: ' + calls.length + ' frontend calls / ' + routes.length + ' backend routes / ' + allowed.length + ' bridge permissions');

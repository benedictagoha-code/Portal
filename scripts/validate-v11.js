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
  const match = html.match(/<script\\b[^>]*>([\\s\\S]*?)<\\/script>/i);
  if (!match) { fail.push(label + ': no script block found'); return; }
  checkJavaScript(label, match[1]);
}
checkJavaScript('Code.gs', code);
firstScript('Bridge.html', bridge);
firstScript('index.html', app);

const routeBlock = code.match(/const ROUTES\\s*=\\s*\\{([\\s\\S]*?)\\n    \\};/);
const allowBlock = bridge.match(/const ALLOWED_METHODS\\s*=\\s*new Set\\(\\[([\\s\\S]*?)\\]\\);/);
if (!routeBlock) fail.push('Code.gs: ROUTES map not found');
if (!allowBlock) fail.push('Bridge.html: ALLOWED_METHODS not found');
const routes = routeBlock ? [...routeBlock[1].matchAll(/^\\s*([A-Za-z0-9_]+):/gm)].map(x => x[1]) : [];
const allowed = allowBlock ? [...allowBlock[1].matchAll(/'([A-Za-z0-9_]+)'/g)].map(x => x[1]) : [];
const calls = [...app.matchAll(/api\\(['"]([^'"]+)/g)].map(x => x[1]).filter((v,i,a) => a.indexOf(v) === i);
for (const name of calls) if (!routes.includes(name)) fail.push('Frontend API has no backend route: ' + name);
for (const name of routes) if (!allowed.includes(name)) fail.push('Backend route missing from bridge allowlist: ' + name);
for (const name of allowed) if (!routes.includes(name)) fail.push('Bridge exposes unknown backend route: ' + name);

const ids = [...app.matchAll(/\\bid=["']([^"']+)["']/g)].map(x => x[1]);
const seen = new Set();
for (const id of ids) {
  if (seen.has(id)) fail.push('index.html: duplicate element id ' + id);
  seen.add(id);
}
if (fail.length) {
  console.error('V11 validation failed:\\n- ' + fail.join('\\n- '));
  process.exit(1);
}
console.log('V11 validation passed.');
console.log('JavaScript syntax: Code.gs, Bridge.html, index.html');
console.log('API wiring: ' + calls.length + ' frontend calls / ' + routes.length + ' backend routes / ' + allowed.length + ' bridge permissions');

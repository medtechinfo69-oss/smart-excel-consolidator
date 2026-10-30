/* ============================================================
 * Build the local French commune database used by the application.
 *
 * Source (official, up to date): API Découpage administratif de l'État
 *   https://geo.api.gouv.fr/communes?fields=nom,code,codesPostaux
 *   (data.gouv.fr / IGN, Licence Ouverte 2.0)
 * It carries the official INSEE code of every commune (métropole, DOM, COM)
 * and the list of postal codes attached to each commune.
 *
 * Output: js/data/communes.js  ->  SEC.COMMUNE_DB = "name\tinsee\tcp1,cp2…\n…"
 * The application stays 100% local: it only reads that generated file.
 *
 * Run: node test/tools/build-communes.js  (needs internet, once)
 * ============================================================ */
'use strict';
const fs = require('fs');
const path = require('path');

const API = 'https://geo.api.gouv.fr/communes?fields=nom,code,codesPostaux&format=json';
const ROOT = path.join(__dirname, '..', '..');
const OUT = path.join(ROOT, 'js', 'data', 'communes.js');
const CACHE = path.join(ROOT, '.cache', 'communes.json');

function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

function sourcePath() {
  if (process.argv[2]) return process.argv[2];
  if (fs.existsSync(CACHE)) return CACHE;
  return null;
}

(async function main() {
  let p = sourcePath();
  if (!p) {
    console.log('Downloading ' + API);
    const res = await fetch(API);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    fs.mkdirSync(path.dirname(CACHE), { recursive: true });
    fs.writeFileSync(CACHE, await res.text());
    p = CACHE;
  }
  const list = JSON.parse(fs.readFileSync(p, 'utf8'));
  const lines = [];
  let cpCount = 0;
  list.forEach(function (c) {
    const nom = String(c.nom || '').replace(/[\t\r\n]/g, ' ').trim();
    const code = String(c.code || '').trim();
    if (!nom || !/^\d[0-9AB]{4}$/.test(code)) return;
    const cps = Array.from(new Set((c.codesPostaux || []).filter(function (cp) {
      return /^\d{5}$/.test(cp);
    }))).sort();
    cpCount += cps.length;
    lines.push(nom + '\t' + code + '\t' + cps.join(','));
  });
  lines.sort(function (a, b) {
    return a.split('\t')[0].localeCompare(b.split('\t')[0], 'fr');
  });
  const header = '/* AUTO-GENERATED - do not edit. Source: geo.api.gouv.fr (IGN / data.gouv.fr),\n' +
    ' * Licence Ouverte 2.0, ' + lines.length + ' communes, generated ' + new Date().toISOString().slice(0, 10) + '.\n' +
    ' * Format: "Nom<TAB>code INSEE<TAB>codes postaux" per line. */\n';
  fs.mkdirSync(path.dirname(OUT), { recursive: true });
  fs.writeFileSync(OUT,
    header +
    '(function (root) {\n' +
    '  "use strict";\n' +
    '  var SEC = (root.SEC = root.SEC || {});\n' +
    '  SEC.COMMUNE_DB_VERSION = ' + JSON.stringify(new Date().toISOString().slice(0, 10)) + ';\n' +
    '  /* One chunk per 1000 communes keeps the file fast to parse. */\n' +
    '  SEC.COMMUNE_DB = [\n' +
    chunk(lines, 1000).map(function (c) { return '    ' + JSON.stringify(c.join('\n')); }).join(',\n') +
    '\n  ].join("\\n");\n' +
    '})(typeof self !== \'undefined\' ? self : globalThis);\n');
  console.log('Wrote ' + OUT + ': ' + lines.length + ' communes, ' + cpCount + ' (commune, postal code) pairs');
})().catch(function (e) { console.error('FAILED: ' + e.message); process.exit(1); });

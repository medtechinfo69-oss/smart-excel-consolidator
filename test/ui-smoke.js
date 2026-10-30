/* ============================================================
 * UI smoke test - loads the real index.html scripts inside a
 * simulated DOM (jsdom) and drives the full user workflow.
 * Run: node test/ui-smoke.js
 * ============================================================ */
'use strict';
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');
const { TextDecoder, TextEncoder } = require('util');

const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
const dom = new JSDOM(html, {
  runScripts: 'outside-only',
  pretendToBeVisual: true,
  url: 'http://localhost/'
});
const { window } = dom;
/* jsdom window lacks TextDecoder/TextEncoder; every real browser provides them. */
if (!window.TextDecoder) window.TextDecoder = TextDecoder;
if (!window.TextEncoder) window.TextEncoder = TextEncoder;

/* ---- Browser API shims (jsdom lacks Blob.arrayBuffer/text; real browsers have them) ---- */
window.CSS = window.CSS || {};
window.CSS.escape = s => String(s).replace(/[^a-zA-Z0-9_\-]/g, m => '\\' + m);
const blobBytes = new window.WeakMap();
const NativeBlob = window.Blob;
function toNodeBuffer(p) {
  if (p === null || p === undefined) return Buffer.alloc(0);
  if (typeof p === 'string') return Buffer.from(p);
  if (p.buffer && typeof p.byteLength === 'number') {
    return Buffer.from(p.buffer, p.byteOffset || 0, p.byteLength); // typed array / DataView (any realm)
  }
  if (typeof p.byteLength === 'number') return Buffer.from(p); // ArrayBuffer (any realm)
  return Buffer.from(String(p));
}
function BlobShim(parts, options) {
  const b = new NativeBlob(parts, options);
  blobBytes.set(b, Buffer.concat((parts || []).map(toNodeBuffer)));
  return b;
}
BlobShim.prototype = NativeBlob.prototype;
window.Blob = BlobShim;
NativeBlob.prototype.arrayBuffer = async function () {
  const buf = blobBytes.get(this);
  if (!buf) throw new Error('shim: untracked blob');
  /* Return a same-realm ArrayBuffer: cross-realm buffers break instance-of
   * checks inside SheetJS when running under jsdom (real browsers are fine). */
  const ab = new window.ArrayBuffer(buf.length);
  new window.Uint8Array(ab).set(buf);
  return ab;
};
NativeBlob.prototype.text = async function () {
  const buf = blobBytes.get(this);
  return buf ? buf.toString('utf8') : '';
};
const blobStore = {};
let blobSeq = 0;
window.URL.createObjectURL = blob => { const u = 'blob:' + (++blobSeq); blobStore[u] = blob; return u; };
window.URL.revokeObjectURL = () => {};
window.HTMLAnchorElement.prototype.click = function () { this._clicked = true; };
const errors = [];
window.addEventListener('error', e => errors.push(e.message));

/* Load the exact scripts referenced by index.html, in order. */
const scriptSrcs = [...html.matchAll(/<script src="([^"]+)"><\/script>/g)].map(m => m[1]);
for (const src of scriptSrcs) {
  const code = fs.readFileSync(path.join(ROOT, src), 'utf8');
  window.eval(code);
}
const document = window.document;
const $ = sel => document.querySelector(sel);
const sleep = ms => new Promise(r => setTimeout(r, ms));
async function waitFor(fn, timeout = 15000, label = 'condition') {
  const start = Date.now();
  while (Date.now() - start < timeout) {
    if (fn()) return;
    await sleep(50);
  }
  throw new Error('timeout waiting for: ' + label);
}

let passed = 0, failed = 0;
function check(name, cond) {
  if (cond) { passed++; console.log('  PASS [UI] ' + name); }
  else { failed++; console.log('  FAIL [UI] ' + name); }
}

function makeFile(fixtureName, asName) {
  const buffer = fs.readFileSync(path.join(__dirname, 'fixtures', fixtureName));
  const file = new window.File([buffer], asName || fixtureName, {
    type: 'application/octet-stream', lastModified: Date.now()
  });
  blobBytes.set(file, buffer); // File builds through native Blob; register bytes explicitly
  return file;
}
function setInputFiles(input, files) {
  Object.defineProperty(input, 'files', { value: files, configurable: true, writable: false });
  input.dispatchEvent(new window.Event('change', { bubbles: true }));
}

(async function () {
  console.log('\n=== UI smoke test ===\n');

  /* Initial render */
  check('11 official headers rendered', document.querySelectorAll('#headerList li').length === 11);
  check('Start disabled with no files', $('#btnStart').disabled === true);
  check('Export disabled before processing', $('#btnExportExcel').disabled === true);
  check('vendor SheetJS loaded', typeof window.XLSX === 'object' && !!window.XLSX.read);

  /* Scenario 1 through the real handler (manual dispatch: disabled buttons
   * do not fire activation clicks per the HTML spec, even in real browsers). */
  $('#btnStart').dispatchEvent(new window.MouseEvent('click', { bubbles: true }));
  await sleep(50);
  check('Error toast when starting without files', document.querySelectorAll('.toast.error').length > 0);

  /* Select 4 mixed files (2 review-heavy) */
  const input = $('#fileInput');
  setInputFiles(input, [
    makeFile('01_clients_semicolon.csv'),
    makeFile('02_clients_comma.csv'),
    makeFile('06_ambiguous.xlsx'),
    makeFile('07_edge_cases.xlsx')
  ]);
  await sleep(50);
  /* 4 files + the totals row. */
  check('4 files listed in both tables', document.querySelectorAll('#fileTableBody tr').length === 5);
  check('file counter shows 4', $('#fileCounter').textContent === '4');
  check('Start enabled with files', $('#btnStart').disabled === false);

  /* Duplicate detection */
  setInputFiles(input, [makeFile('01_clients_semicolon.csv', '01_clients_semicolon.csv')]);
  await sleep(50);
  check('duplicate flagged in table', document.body.innerHTML.indexOf('Doublon') !== -1);

  /* Run the batch through the real Start button */
  $('#btnStart').click();
  check('Start locked while running', $('#btnStart').disabled === true);
  check('Cancel enabled while running', $('#btnCancel').disabled === false);
  await waitFor(
    () => $('#btnCancel').disabled === true && Number($('#sumOkRows').textContent) > 0,
    20000, 'batch finish (cancel released, rows extracted)'
  );
  check('Cancel disabled after batch', $('#btnCancel').disabled === true);
  check('lines added card > 0', Number($('#sumOkRows').textContent) > 0);
  check('no failed files', Number($('#sumFailed').textContent) === 0);
  check('success files card > 0', Number($('#sumSuccess').textContent) > 0);
  /* No manual validation UI: columns are resolved automatically. */
  check('no manual review section', !document.querySelector('#reviewSection'));
  check('no Correction button', !document.querySelector('#btnAutoFix'));
  check('no confidence threshold inputs', !document.querySelector('#thHigh') && !document.querySelector('#thMedium'));
  check('no manual destination selects', document.querySelectorAll('[data-review-select]').length === 0);

  await waitFor(() => $('#btnExportExcel').disabled === false, 15000, 'export enabled');
  await sleep(200);
  check('preview rendered', $('#previewSection').hidden === false && document.querySelectorAll('#previewTable tbody tr').length > 0);
  check('Export enabled after automatic processing', $('#btnExportExcel').disabled === false);

  $('#btnExportExcel').click();
  await sleep(300);
  const valItems = document.querySelectorAll('#validationList li');
  check('validation checklist rendered', valItems.length >= 6);
  check('all validation checks pass', [...valItems].every(li => li.classList.contains('pass')));
  check('workbook download triggered', Object.keys(blobStore).length >= 1);

  /* Header configuration is manual only (no template / JSON buttons). */
  check('template, JSON load and rules export buttons removed',
    !document.querySelector('#btnLoadTemplate') && !document.querySelector('#btnLoadJson') &&
    !document.querySelector('#btnExportHeaders'));
  check('manual header inputs available', document.querySelectorAll('#manualInputs [data-manual]').length === 11);

  /* Verify the downloaded workbook really opens and has the right shape */
  const wbBlob = Object.values(blobStore)[0];
  const wbBuf = await wbBlob.arrayBuffer();
  const wb = window.XLSX.read(wbBuf, { type: 'array' });
  const aoa = window.XLSX.utils.sheet_to_json(wb.Sheets['tableau final'], { header: 1, defval: '' });
  check('downloaded workbook: exact 11 headers', JSON.stringify(aoa[0]) === JSON.stringify(window.SEC.OFFICIAL_HEADERS));
  check('downloaded workbook: data rows present', aoa.length > 5);
  check('downloaded workbook: report sheet included', wb.SheetNames.indexOf('Rapport') !== -1);

  /* Every previewed value is found in the SAME column of the workbook. The
   * preview shows the first 100 rows of the session (not the whole batch), so
   * compare against the workbook's columns only (any row, same column). */
  const previewRows = [...document.querySelectorAll('#previewTable tbody tr')]
    .map(tr => [...tr.querySelectorAll('td')].map(td => td.textContent.trim()));
  const wbData = aoa.slice(1);
  /* Formula-injection guard prefixes EVERY dangerous segment with an
   * apostrophe ("a | =1+1" -> "a | '=1+1"): compare with that applied. */
  const guard = v => v.split(' | ').map(p => /^[=+\-@]/.test(p) ? "'" + p : p).join(' | ');
  const sameValue = (v, got) => v === got || (/^\d+$/.test(v) && /^\d+$/.test(got) && String(Number(v)) === String(Number(got))) ||
    got === "'" + v || got === guard(v);
  const missing = [];
  previewRows.forEach(pr => pr.forEach((v, c) => {
    if (v === '') return;
    const hit = wbData.some(wr => sameValue(v, wr[c] == null ? '' : String(wr[c])));
    if (!hit) missing.push(c + ':' + v);
  }));
  if (missing.length) console.log('  missing preview cells:', missing.slice(0, 10).join(', '));
  check('preview and workbook agree column by column',
    previewRows.length > 0 && missing.length === 0, missing.slice(0, 4).join(', '));

  /* An unknown column is resolved automatically: its data must be exported
   * (kept in « Commentaire ») without any manual destination choice. */
  $('#btnRemoveAll').click();
  await sleep(300);
  const manualText = 'Nom;Truc bizarre;Mail\nDupont;a@b.fr;c@d.fr\nMartin;e@f.fr;g@h.fr\n';
  const manualCsv = new window.File([Buffer.from(manualText, 'utf8')], 'manuel.csv', { type: 'text/csv' });
  blobBytes.set(manualCsv, Buffer.from(manualText, 'utf8'));
  setInputFiles(input, [manualCsv]);
  await sleep(50);
  $('#btnStart').click();
  await waitFor(() => $('#btnCancel').disabled === true, 15000, 'manual file processed');
  await waitFor(() => {
    const row = document.querySelectorAll('#previewTable tbody tr')[0];
    return row && row.textContent.indexOf('a@b.fr') !== -1;
  }, 15000, 'values in preview');
  check('no destination select needed', document.querySelectorAll('[data-review-select]').length === 0);
  check('e-mail value exported', $('#previewTable tbody').textContent.indexOf('a@b.fr') !== -1);
  check('unknown column value preserved', $('#previewTable tbody').textContent.indexOf('a@b.fr') !== -1 &&
    $('#previewTable tbody').textContent.indexOf('e@f.fr') !== -1);

  /* Remove all resets the UI */
  $('#btnRemoveAll').click();
  await sleep(300);
  check('Remove all empties file list',
    document.querySelectorAll('#fileTableBody tr:not(.totals-row)').length === 0);
  check('Start disabled again after clearing', $('#btnStart').disabled === true);

  /* Manual header edit round trip */
  $('#btnManualEdit').click();
  check('manual editor is always visible', $('#manualPanel').hidden === false);
  const inputs = document.querySelectorAll('[data-manual]');
  check('11 manual inputs', inputs.length === 11);
  inputs[0].value = 'Nom de famille';
  $('#btnSaveManual').click();
  await sleep(50);
  check('edited header text applied', document.querySelector('#headerList li').textContent.indexOf('Nom de famille') !== -1);
  $('#btnResetManual').click();
  check('reset restores official headers', document.querySelectorAll('#headerList li')[4].textContent.indexOf('Adresse') !== -1);

  check('no uncaught window errors', errors.length === 0);
  if (errors.length) console.log('  errors:', errors);

  console.log('\n=== UI results: ' + passed + ' passed, ' + failed + ' failed ===\n');
  fs.writeFileSync(path.join(__dirname, 'ui-results.json'), JSON.stringify({ passed, failed, errors }, null, 2));
  process.exit(failed ? 1 : 0);
})().catch(err => { console.error('UI HARNESS ERROR', err); process.exit(2); });

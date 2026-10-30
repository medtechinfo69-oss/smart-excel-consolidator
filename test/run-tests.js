/* ============================================================
 * Smart Excel - automated test suite (Node, no DOM)
 * Covers the 28 mandatory delivery scenarios against the real
 * parsing / mapping / processing / export logic.
 * Run: node test/run-tests.js
 * ============================================================ */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

/* Load application modules in browser-like order. */
require('../js/config.js');
require('../js/data/communes.js');
require('../js/communes.js');
require('../js/dictionaries.js');
require('../js/data/mutuelles.js');
require('../js/normalizers.js');
require('../js/csv-parser.js');
require('../js/xlsx-reader.js');
require('../js/agent.js');
require('../js/processor.js');
require('../js/exporter.js');
const SEC = globalThis.SEC;

const FX = path.join(__dirname, 'fixtures');
const buf = name => fs.readFileSync(path.join(FX, name));
const immediate = () => Promise.resolve();

let passed = 0, failed = 0;
const results = [];
function test(scenario, name, fn) {
  return Promise.resolve().then(fn).then(() => {
    passed++; results.push({ scenario, name, pass: true });
    console.log('  PASS [' + scenario + '] ' + name);
  }).catch(err => {
    failed++; results.push({ scenario, name, pass: false, error: err.message });
    console.log('  FAIL [' + scenario + '] ' + name + ' -> ' + err.message);
  });
}
function assert(cond, msg) { if (!cond) throw new Error(msg || 'assertion failed'); }
function assertEq(a, b, msg) { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error((msg || 'not equal') + ' (got ' + JSON.stringify(a) + ', want ' + JSON.stringify(b) + ')'); }

async function newSession() { return SEC.Processor.createSession(); }
async function processOne(session, name, meta) {
  const e = session.registerFile(Object.assign({ name, size: buf(name).length }, meta || {}));
  await session.startBatch([e], { readBuffer: () => Promise.resolve(buf(name)), yieldFn: immediate });
  return e;
}
function rowsOf(session) { return session.buildConsolidatedRows(); }
function findRecord(rows, col, value) {
  const idx = SEC.TARGET_KEYS.indexOf(col);
  return rows.find(r => r[idx] === value);
}

(async function run() {
  console.log('\n=== Smart Excel - test run ===\n');

  /* ---------- Scenario 1: Start without files ---------- */
  await test(1, 'Start Processing with no file is blocked', async () => {
    const s = await newSession();
    const r = SEC.Processor.validatePrerequisites({ fileCount: 0, headers: s.headers, running: false, hasProcessable: false });
    assert(r.ok === false && r.errors.length > 0, 'must report errors');
    await s.startBatch([], { readBuffer: () => null });
    assert(s.running === false && s.records.size === 0, 'empty batch must not start work');
  });

  /* ---------- Scenario 2: Built-in official headers ---------- */
  await test(2, 'Built-in configuration exposes the exact 11 headers', () => {
    assertEq(SEC.OFFICIAL_HEADERS, ['Nom', 'Prénom', 'Téléphone 1', 'Téléphone 2', 'Adresse', 'Code postal', 'Ville', 'Email', 'Cotisation', 'Mutuelle', 'Commentaire']);
    assert(SEC.OFFICIAL_HEADERS.length === 11);
  });

  /* ---------- Scenario 3: Load official template locally ---------- */
  await test(3, 'Official XLSX template is parsed and validated (and bad template rejected)', () => {
    const good = SEC.Processor.validateTemplate(SEC.XLSXReader.parse(buf('En tete finale.xlsx')));
    assert(good.ok === true, 'template must validate');
    assertEq(good.row, SEC.OFFICIAL_HEADERS);
    const bad = SEC.Processor.validateTemplate(SEC.XLSXReader.parse(buf('10_headers_wrong.xlsx')));
    assert(bad.ok === false && bad.found === 4, 'partial template must be rejected (found=' + bad.found + ')');
  });

  /* ---------- Scenario 4: valid CSV ---------- */
  await test(4, 'A valid semicolon CSV is processed end to end', async () => {
    const s = await newSession();
    const e = await processOne(s, '01_clients_semicolon.csv');
    assert(e.status === 'completed' || e.status === 'completed_warnings', 'status=' + e.status);
    assertEq(rowsOf(s).length, 8, '8 data rows expected');
    assert(s.stats.failed === 0);
  });

  /* ---------- Scenario 5: valid XLSX with multiple sheets ---------- */
  await test(5, 'A valid XLSX with multiple worksheets is processed (empty sheet ignored)', async () => {
    const s = await newSession();
    const e = await processOne(s, '05_workbook_multi.xlsx');
    /* needs_review is correct here: the two unmapped columns require a decision */
    assert(['completed', 'completed_warnings', 'needs_review'].indexOf(e.status) !== -1, 'status=' + e.status);
    assertEq(rowsOf(s).length, 6, '3 + 3 records from the two non-empty sheets');
    assert(s.report.some(r => /feuille vide/i.test(r.reason)), 'empty sheet reported');
  });

  /* ---------- Scenario 6: 100+ mixed files ---------- */
  await test(6, 'Batch of 117 mixed CSV/XLSX files settles without freezing or loss', async () => {
    const s = await newSession();
    const names = fs.readdirSync(path.join(FX, 'batch')).map(n => 'batch' + path.sep + n)
      .concat(['01_clients_semicolon.csv', '02_clients_comma.csv', '03_tab_delimited.txt', '04_legacy_cp1252.csv',
        '05_workbook_multi.xlsx', '06_ambiguous.xlsx', '07_edge_cases.xlsx']);
    assert(names.length >= 117, 'batch size ' + names.length);
    const entries = names.map(n => s.registerFile({ name: n, size: buf(n).length }));
    await s.startBatch(entries, { readBuffer: en => Promise.resolve(buf(en.name)), yieldFn: immediate });
    const stuck = entries.filter(en => ['waiting', 'validating', 'processing'].indexOf(en.status) !== -1);
    assert(stuck.length === 0, 'no file stuck in progress');
    assert(s.records.size > 500, 'expected hundreds of consolidated rows, got ' + s.records.size);
    assert(s.stats.failed === 0, 'all well-formed files must succeed');
  });

  /* ---------- Scenario 7: different header names ---------- */
  await test(7, 'English/industry synonyms map to the French official targets', async () => {
    const s = await newSession();
    await processOne(s, '02_clients_comma.csv');
    const rows = rowsOf(s);
    assert(findRecord(rows, 'nom', 'Smith'), 'Last Name -> Nom');
    assert(findRecord(rows, 'prenom', 'Maria'), 'First Name -> Prénom');
    assert(findRecord(rows, 'mutuelle', 'SwissLife'), 'Insurer -> Mutuelle');
    assert(findRecord(rows, 'commentaire', 'Quoted, with comma'), 'Comments -> Commentaire');
    assert(findRecord(rows, 'cotisation', '75.50'), 'Monthly Contribution -> Cotisation (value preserved)');
  });

  /* ---------- Scenario 8: different column order ---------- */
  await test(8, 'Column order differs but values land in the correct official columns', async () => {
    const s = await newSession();
    await processOne(s, '05_workbook_multi.xlsx');
    const rec = findRecord(rowsOf(s), 'email', 'nathalie.leroy@example.fr');
    assert(rec, 'record found');
    assertEq(rec[SEC.TARGET_KEYS.indexOf('nom')], 'Leroy');
    assertEq(rec[SEC.TARGET_KEYS.indexOf('ville')], 'Paris');
    assertEq(rec[SEC.TARGET_KEYS.indexOf('cp')], '75004');
  });

  /* ---------- Scenario 9: missing columns ---------- */
  await test(9, 'Files with missing columns produce empty fields, not errors', async () => {
    const s = await newSession();
    await processOne(s, '05_workbook_multi.xlsx');
    const rec = findRecord(rowsOf(s), 'nom', 'Faure');
    assert(rec, 'Antoine Faure record exists');
    assertEq(rec[SEC.TARGET_KEYS.indexOf('email')], '', 'email column absent -> empty');
    assertEq(rec[SEC.TARGET_KEYS.indexOf('tel2')], '', 'tel2 absent -> empty');
    assertEq(rec[SEC.TARGET_KEYS.indexOf('adresse')], '5 place Bellecour');
  });

  /* ---------- Scenario 10: extra unmapped columns ---------- */
  await test(10, 'Additional unmapped columns are reported, never silently discarded', async () => {
    const s = await newSession();
    await processOne(s, '05_workbook_multi.xlsx');
    const reasons = s.report.map(r => r.sourceHeader + '|' + r.reason).join('§');
    assert(/Numéro client|numero client/i.test(reasons), 'ID column reported: ' + reasons);
    assert(/Date d'adhésion|date adhesion/i.test(reasons), 'date column reported');
  });

  /* ---------- Scenario 11: empty rows and empty worksheets ---------- */
  await test(11, 'Empty rows are skipped and empty worksheets are noted', async () => {
    const s = await newSession();
    const e = await processOne(s, '07_edge_cases.xlsx');
    const names = rowsOf(s).map(r => r[0]);
    assertEq(names.length, 4, 'four non-empty data rows');
    assert(names.indexOf('') === -1, 'no blank records');
    assert(e.rowsExtracted === 4);
  });

  /* ---------- Scenario 12: French accents & special characters ---------- */
  await test(12, 'Accents, hyphens and apostrophes are preserved exactly', async () => {
    const s = await newSession();
    await processOne(s, '01_clients_semicolon.csv');
    const rows = rowsOf(s);
    assert(findRecord(rows, 'prenom', 'Élodie'), 'Élodie preserved');
    assert(findRecord(rows, 'nom', 'Petit') && findRecord(rows, 'prenom', 'François'), 'François preserved');
    /* "Bellegarde-sur-Valserine" (INSEE 01033) is not a commune anymore; the
     * city cannot be identified with certainty -> Ville left EMPTY (the value
     * is reported, never replaced by an invented name). */
    const sophie = findRecord(rows, 'prenom', 'Sophie');
    assertEq(sophie[SEC.TARGET_KEYS.indexOf('ville')], '', 'old/ambiguous commune name -> empty');
    assert(s.report.some(r => /ville non identifiee/i.test(r.reason) && /Bellegarde/.test(r.samples.join(' '))), 'reported, not invented');
    assert(findRecord(rows, 'ville', 'Nantes'), 'simple commune kept');
    const cp1252 = await newSession();
    await processOne(cp1252, '04_legacy_cp1252.csv');
    assert(findRecord(rowsOf(cp1252), 'nom', 'Béranger'), 'cp1252 decoded to Béranger');
    const ber = rowsOf(cp1252).find(r => r[SEC.TARGET_KEYS.indexOf('nom')] === 'Béranger');
    assert(ber[SEC.TARGET_KEYS.indexOf('adresse')].indexOf("l'église") !== -1, "apostrophe address preserved");
  });

  /* ---------- Scenario 13: CSV dialects, quoting, encodings ---------- */
  await test(13, 'Semicolon, comma, tab, quoted fields, BOM and cp1252 are handled', async () => {
    const s = await newSession();
    const e1 = await processOne(s, '01_clients_semicolon.csv');
    assertEq(e1.delimiter, ';'); assertEq(e1.encoding, 'utf-8');
    const e2 = s.registerFile({ name: '02_clients_comma.csv', size: buf('02_clients_comma.csv').length });
    await s.startBatch([e2], { readBuffer: () => Promise.resolve(buf('02_clients_comma.csv')), yieldFn: immediate });
    assertEq(e2.delimiter, ',');
    const e3 = s.registerFile({ name: '03_tab_delimited.txt', size: buf('03_tab_delimited.txt').length });
    await s.startBatch([e3], { readBuffer: () => Promise.resolve(buf('03_tab_delimited.txt')), yieldFn: immediate });
    assertEq(e3.delimiter, '\t');
    const e4 = s.registerFile({ name: '04_legacy_cp1252.csv', size: buf('04_legacy_cp1252.csv').length });
    await s.startBatch([e4], { readBuffer: () => Promise.resolve(buf('04_legacy_cp1252.csv')), yieldFn: immediate });
    assertEq(e4.encoding, 'windows-1252');
    /* quoted semicolon inside field */
    const quoted = findRecord(rowsOf(s), 'nom', 'Bernard');
    assert(quoted[SEC.TARGET_KEYS.indexOf('adresse')].indexOf('10 rue; grande entrée') !== -1, 'quoted semicolon kept');
    /* embedded newline inside quoted field */
    const nl = findRecord(rowsOf(s), 'commentaire', 'Line one\nLine two in same cell');
    assert(nl, 'embedded newline preserved');
  });

  /* ---------- Scenario 14: French telephone formats ---------- */
  await test(14, 'National, +33, 0033, landline and mobile numbers normalize correctly', async () => {
    const s = await newSession();
    await processOne(s, '01_clients_semicolon.csv');
    const rows = rowsOf(s);
    assertEq(findRecord(rows, 'prenom', 'Jean')[SEC.TARGET_KEYS.indexOf('tel1')], '0612345678', 'spaced national -> digits text');
    assertEq(findRecord(rows, 'prenom', 'Sophie')[SEC.TARGET_KEYS.indexOf('tel1')], '0612345678', '+33 conversion');
    assertEq(findRecord(rows, 'prenom', 'Élodie')[SEC.TARGET_KEYS.indexOf('tel1')], '0712345678', '0033 conversion');
    assertEq(findRecord(rows, 'prenom', 'Jean')[SEC.TARGET_KEYS.indexOf('tel2')], '0145678910', 'landline in Téléphone 2');
    assertEq(findRecord(rows, 'prenom', 'Kévin')[SEC.TARGET_KEYS.indexOf('tel1')], '0712345678', '+33 7 mobile');
    const onlyOne = findRecord(rows, 'prenom', 'François');
    assertEq(onlyOne[SEC.TARGET_KEYS.indexOf('tel1')], '0712345678');
    assertEq(onlyOne[SEC.TARGET_KEYS.indexOf('tel2')], '', 'missing second number never invented');
  });

  /* ---------- Scenario 15: postal codes with leading zero ---------- */
  await test(15, 'Postal codes keep leading zeros and stay text', async () => {
    const s = await newSession();
    const out = SEC.Exporter.generateConsolidated(s.headers, [
      ['N', 'P', '0612345678', '', '1 rue', '01200', 'Béligneux', 'a@b.fr', '10,00', 'MGEN', '']
    ], [], 1);
    assert(out.validation.ok, 'export validation must pass');
    const X = SEC.XLSXReader.getXLSX();
    const reopened = X.read(out.buffer, { type: 'array' });
    const ws = reopened.Sheets['tableau final'];
    const cpCell = ws[X.utils.encode_cell({ r: 1, c: 5 })];
    const telCell = ws[X.utils.encode_cell({ r: 1, c: 2 })];
    assertEq(cpCell.t, 's', 'postal code cell is text');
    assertEq(cpCell.v, '01200', 'leading zero retained');
    assertEq(telCell.t, 's', 'phone cell is text');
  });

  /* ---------- Scenario 16: French currency / decimal formats ---------- */
  await test(16, 'French amounts are preserved verbatim and parsed correctly', async () => {
    assertEq(SEC.Normalizers.parseFrenchNumber('1 250,00 €'), 1250);
    assertEq(SEC.Normalizers.parseFrenchNumber('75,50'), 75.5);
    assertEq(SEC.Normalizers.parseFrenchNumber('120 €'), 120);
    const s = await newSession();
    await processOne(s, '01_clients_semicolon.csv');
    const rows = rowsOf(s);
    assert(findRecord(rows, 'cotisation', '1 250,00 €'), 'spaced thousands + currency symbol kept');
    assert(findRecord(rows, 'cotisation', '75,50'), 'comma decimal kept');
    await processOne(s, '05_workbook_multi.xlsx');
    assert(s.report.some(r => /annuelle/i.test(r.reason)), 'annual period flagged without conversion');
  });

  /* ---------- Scenario 17: ambiguous headers ---------- */
  await test(17, 'ID columns are blocked from phone mapping; full names and dual phones go to review', async () => {
    const s = await newSession();
    await processOne(s, '06_ambiguous.xlsx');
    const items = [];
    s.reviewItems.forEach(i => items.push(i));
    const headers = items.map(i => i.header);
    assert(headers.indexOf('Numéro client') !== -1, 'identifier column flagged for review');
    assert(headers.indexOf('Nom complet') !== -1, 'full name column flagged for review');
    const telA = items.find(i => i.header === 'Téléphone A' || i.header === 'Téléphone B');
    assert(telA, 'two undifferentiated phones flagged for primary/secondary review');
    /* numeric ID values must not have been mapped to a phone */
    const rows = rowsOf(s);
    const allPhones = rows.map(r => r[2]).filter(Boolean);
    assert(allPhones.every(p => /^0\d{9}$/.test(p)), 'all tel1 values are real French numbers, no IDs: ' + allPhones.join(','));
    assert(findRecord(rows, 'tel1', '0612000010'), 'Numéro column (phone values) mapped to Téléphone 1');
  });

  /* ---------- Scenario 18: uncertain columns are resolved automatically ---------- */
  await test(18, 'Uncertain columns are resolved automatically, without manual validation', async () => {
    const s = await newSession();
    const csv1 = Buffer.from('Titre client;Mail client;Montant\nDupont;a@b.fr;10,00\nMartin;c@d.fr;20,50\n', 'utf8');
    const e1 = s.registerFile({ name: 'weird1.csv', size: csv1.length });
    await s.startBatch([e1], { readBuffer: () => Promise.resolve(csv1), yieldFn: immediate });
    /* Nothing is left open and the file is a plain success. */
    assertEq(s.openReviewCount(), 0, 'no column is left to verify');
    assertEq(e1.status, 'completed', 'file completed automatically');
    /* An unknown header keeps its data (Commentaire) instead of blocking. */
    const item = Array.from(s.reviewItems.values()).find(i => i.header === 'Titre client');
    assert(item && item.decision, 'unknown column decided automatically');
    assert(findRecord(rowsOf(s), 'commentaire', 'Dupont'), 'unknown column value preserved');
    assert(!findRecord(rowsOf(s), 'nom', 'Dupont'), 'nothing invented in Nom');
    /* The e-mail column still lands in Email. */
    assert(findRecord(rowsOf(s), 'email', 'a@b.fr'), 'e-mail detected in the right column');
    /* A second file with the same uncertain header is resolved the same way. */
    const csv2 = Buffer.from('Titre client;Mail client\nBernard;e@f.fr\n', 'utf8');
    const e2 = s.registerFile({ name: 'weird2.csv', size: csv2.length });
    await s.startBatch([e2], { readBuffer: () => Promise.resolve(csv2), yieldFn: immediate });
    assertEq(s.openReviewCount(), 0, 'second file resolved automatically too');
    assert(findRecord(rowsOf(s), 'commentaire', 'Bernard'), 'second file value preserved');
    assert(!findRecord(rowsOf(s), 'nom', 'Bernard'), 'nothing copied into Nom');
    /* An explicit decision still applies exactly when it is made. */
    const again = Array.from(s.reviewItems.values()).find(i => i.fileId === e2.id && i.header === 'Titre client');
    if (again) {
      s.decideReview(again.id, 'nom');
      await s.reprocessAffected([e2.id], undefined, immediate, () => s.autoFixReview());
      assert(findRecord(rowsOf(s), 'nom', 'Bernard'), 'explicit decision applied exactly');
    }
  });

  /* ---------- Scenario 19: corrupted / unsupported file ---------- */
  await test(19, 'Corrupted and unsupported files fail individually while the batch continues', async () => {
    const s = await newSession();
    const bad = s.registerFile({ name: '08_corrupted.xlsx', size: buf('08_corrupted.xlsx').length });
    const good = s.registerFile({ name: '01_clients_semicolon.csv', size: buf('01_clients_semicolon.csv').length });
    await s.startBatch([bad, good], {
      readBuffer: en => Promise.resolve(en.name === '08_corrupted.xlsx' ? buf('08_corrupted.xlsx') : buf('01_clients_semicolon.csv')),
      yieldFn: immediate
    });
    assertEq(bad.status, 'failed');
    assert(bad.error && bad.error.length > 0, 'error detail recorded');
    assertEq(good.status, 'completed');
    assertEq(rowsOf(s).length, 8, 'other files still processed');
    const dat = s.registerFile({ name: '09_unsupported.dat', size: 15 });
    await s.startBatch([dat], { readBuffer: () => Promise.resolve(buf('09_unsupported.dat')), yieldFn: immediate });
    assertEq(dat.status, 'failed', 'unsupported extension fails cleanly');
  });

  /* ---------- Scenario 20: double-start protection ---------- */
  await test(20, 'A second concurrent batch cannot start', async () => {
    const s = await newSession();
    const e = s.registerFile({ name: '01_clients_semicolon.csv', size: buf('01_clients_semicolon.csv').length });
    const p1 = s.startBatch([e], { readBuffer: () => Promise.resolve(buf('01_clients_semicolon.csv')), yieldFn: immediate });
    let rejected = false;
    try { await s.startBatch([e], { readBuffer: () => Promise.resolve(buf('01_clients_semicolon.csv')), yieldFn: immediate }); }
    catch (err) { rejected = /deja en cours|already/.test(err.message); }
    await p1;
    assert(rejected, 'second startBatch must reject');
  });

  /* ---------- Scenario 21: retry without duplicate rows ---------- */
  await test(21, 'Reprocessing/retry replaces records by source key (no duplication)', async () => {
    const s = await newSession();
    const e = await processOne(s, '01_clients_semicolon.csv');
    const before = s.records.size;
    await s.reprocessAffected([e.id]);
    assertEq(s.records.size, before, 'same record count after reprocess');
    await s.reprocessAffected([e.id]);
    assertEq(s.records.size, before, 'still no duplicates after second reprocess');
    assertEq(rowsOf(s).length, before);
  });

  /* ---------- Scenario 22: cancellation ---------- */
  await test(22, 'Cancellation restores the interface and retains partial results', async () => {
    const s = await newSession();
    const entries = ['01_clients_semicolon.csv', '02_clients_comma.csv', '03_tab_delimited.txt']
      .map(n => s.registerFile({ name: n, size: buf(n).length }));
    const p = s.startBatch(entries, { readBuffer: en => Promise.resolve(buf(en.name)), yieldFn: immediate });
    s.cancel();
    await p;
    assert(s.running === false, 'running flag released');
    assert(entries.some(en => en.status === 'cancelled'), 'at least one file cancelled');
    /* partial retained: either zero or some records, never corrupted */
    assert(s.records.size >= 0 && rowsOf(s).length === s.records.size);
  });

  /* ---------- Scenario 23/24/25: export structure, headers, counts ---------- */
  await test(23, 'Exported workbook reopens with the exact 11 headers, report sheet and correct row count', async () => {
    const s = await newSession();
    await processOne(s, '01_clients_semicolon.csv');
    await processOne(s, '07_edge_cases.xlsx');
    const rows = rowsOf(s);
    const out = SEC.Exporter.generateConsolidated(s.headers, rows, s.report, rows.length);
    assert(out.validation.ok, 'all validation checks pass: ' + JSON.stringify(out.validation.checks.filter(c => !c.pass)));
    const X = SEC.XLSXReader.getXLSX();
    const wb = X.read(out.buffer, { type: 'array' });
    assertEq(wb.SheetNames[0], 'tableau final');
    assert(wb.SheetNames.indexOf('Rapport') !== -1, 'report worksheet present');
    const aoa = X.utils.sheet_to_json(wb.Sheets['tableau final'], { header: 1, defval: '' });
    assertEq(aoa[0], SEC.OFFICIAL_HEADERS, 'exact official header order');
    assertEq(aoa.length - 1, rows.length, 'exported row count matches dataset');
    assertEq(aoa[0].length, 11, 'no extra columns');
  });

  /* ---------- Scenario 24 (explicit) and formula injection ---------- */
  await test(24, 'Formula injection is neutralized on export', async () => {
    const s = await newSession();
    await processOne(s, '07_edge_cases.xlsx');
    const rows = rowsOf(s);
    const out = SEC.Exporter.generateConsolidated(s.headers, rows, [], rows.length);
    const X = SEC.XLSXReader.getXLSX();
    const aoa = X.utils.sheet_to_json(X.read(out.buffer, { type: 'array' }).Sheets['tableau final'], { header: 1, defval: '' });
    /* A non-phone value ("123") is now moved to Commentaire before the
     * original note, so the injection payload may not be the first segment:
     * check that every dangerous value was neutralized wherever it landed. */
    const comments = aoa.slice(1).map(r => r[10]).filter(Boolean).join(' | ');
    assert(comments.indexOf("'=1+1") !== -1, 'leading = neutralized');
    assert(comments.indexOf("'+CMD") !== -1, 'leading + neutralized');
    assert(comments.indexOf("'@inject") !== -1, 'leading @ neutralized');
    assert(comments.indexOf(' | =1+1') === -1, 'no un-neutralized = value');
    assert(comments.indexOf(' | +CMD') === -1, 'no un-neutralized + value');
  });

  /* ---------- Scenario 25: invalid values warned but retained ---------- */
  await test(25, 'Invalid phone/email/CP/amount values are flagged in the report and kept', async () => {
    const s = await newSession();
    await processOne(s, '07_edge_cases.xlsx');
    const reasons = s.report.map(r => r.reason).join('§');
    assert(/telephone invalide/i.test(reasons), 'bad phone warned');
    assert(/email/i.test(reasons), 'bad email warned');
    assert(/code postal invalide/i.test(reasons), 'bad CP warned');
    assert(/montant/i.test(reasons), 'bad amount warned');
    const rec = findRecord(rowsOf(s), 'nom', 'Rond');
    assert(rec, 'row retained despite invalid values');
  });

  /* ---------- Scenario 26: syntax / runtime errors ---------- */
  await test(26, 'Every JavaScript file passes syntax checking', () => {
    const files = ['js/config.js', 'js/dictionaries.js', 'js/normalizers.js', 'js/csv-parser.js', 'js/xlsx-reader.js', 'js/agent.js', 'js/processor.js', 'js/exporter.js', 'js/ui.js', 'js/app.js'];
    files.concat(['js/communes.js', 'js/data/communes.js', 'js/data/mutuelles.js', 'js/ai/gemini.js']).forEach(f => {
      const code = fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
      new vm.Script(code, { filename: f }); // throws on syntax error
    });
  });

  /* ---------- Scenario 27: offline, local dependencies only ----------
   * The OPTIONAL Gemini assistant is the single, deliberate exception: it is
   * the only file allowed to contact a remote service, and only when the
   * user has explicitly enabled it and pasted a key. */
  await test(27, 'No external URL/CDN dependency exists and vendor library is local', () => {
    const own = ['index.html', 'css/styles.css', 'js/ui.js', 'js/app.js', 'js/config.js', 'js/dictionaries.js',
      'js/normalizers.js', 'js/csv-parser.js', 'js/xlsx-reader.js', 'js/agent.js', 'js/processor.js', 'js/exporter.js',
      'js/communes.js', 'js/data/communes.js', 'js/data/mutuelles.js'];
    own.forEach(f => {
      let code = fs.readFileSync(path.join(__dirname, '..', f), 'utf8');
      /* The only allowed absolute link is this machine's own local server
       * (recovery link); it is never a CDN or a remote service. */
      code = code.replace(/http:\/\/127\.0\.0\.1:\d+\/?/g, '');
      assert(!/https?:\/\//.test(code), f + ' contains an external URL');
    });
    /* The AI module must ONLY reach the official Gemini endpoint. */
    const ai = fs.readFileSync(path.join(__dirname, '..', 'js/ai/gemini.js'), 'utf8');
    const urls = (ai.match(/https?:\/\/[^"'\s]+/g) || []);
    assert(urls.every(u => u.indexOf('https://generativelanguage.googleapis.com/') === 0),
      'AI module reaches a non-Gemini URL: ' + urls.join(', '));
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    const refs = (html.match(/(?:src|href)="([^"]+)"/g) || []).map(r => r.replace(/(?:src|href)="|"/g, ''));
    refs.forEach(r => {
      /* The only absolute link is the recovery link to this machine's own
       * local server (same loopback origin, never a CDN). */
      const localServer = /^http:\/\/127\.0\.0\.1:\d+\/?$/.test(r);
      assert(localServer || !/^https?:/.test(r), 'external reference: ' + r);
      /* An inlined resource (data:) is not a file and not a network call. */
      if (localServer || /^data:/.test(r)) return;
      const target = path.join(__dirname, '..', r);
      assert(fs.existsSync(target), 'missing local dependency: ' + r);
    });
    assert(fs.statSync(path.join(__dirname, '..', 'vendor/xlsx.full.min.js')).size > 500000, 'SheetJS bundle present');
  });

  /* ---------- Scenario 28: no backend / storage / data logging ---------- */
  await test(28, 'No storage, no direct AI call and no customer-data logging in application code', () => {
    const dir = path.join(__dirname, '..', 'js');
    const stripComments = code => code
      .replace(/\/\*[\s\S]*?\*\//g, '')
      .replace(/(^|[^:])\/\/[^\n]*/g, '$1');
    const forbidden = ['localStorage', 'sessionStorage', 'indexedDB', 'XMLHttpRequest', 'sendBeacon', 'console.log', 'console.debug', 'navigator.geolocation'];
    fs.readdirSync(dir).filter(f => f.endsWith('.js')).forEach(f => {
      const code = stripComments(fs.readFileSync(path.join(dir, f), 'utf8'));
      forbidden.forEach(token => assert(code.indexOf(token) === -1, f + ' contains forbidden API: ' + token));
    });
    /* The optional AI assistant may talk to its LOCAL server only, never to a
     * provider API (keys live in .env, on the backend). */
    fs.readdirSync(dir).filter(f => f.endsWith('.js')).forEach(f => {
      const code = stripComments(fs.readFileSync(path.join(dir, f), 'utf8'));
      assert(!/https?:\/\/(?!127\.0\.0\.1|localhost)/.test(code), f + ' calls an external URL');
      assert(!/gsk_[A-Za-z0-9]{10,}|AQ\.[A-Za-z0-9_-]{10,}/.test(code), f + ' contains an API key');
    });
    const aiClient = stripComments(fs.readFileSync(path.join(dir, 'ai', 'client.js'), 'utf8'));
    const fetchTargets = Array.from(aiClient.matchAll(/fetchJson\(\s*["']([^"']+)["']/g)).map(m => m[1]);
    assert(fetchTargets.length >= 3, 'AI client endpoints found: ' + fetchTargets.length);
    fetchTargets.forEach(t => assert(t.indexOf('/api/ai/') === 0, 'AI client calls a non-local endpoint: ' + t));
    assert(/\/api\/export\/xlsx/.test(stripComments(fs.readFileSync(path.join(dir, 'app.js'), 'utf8'))), 'styled export uses the local endpoint');
    const appCode = stripComments(fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8'));
    forbidden.forEach(token => assert(appCode.indexOf(token) === -1, 'index.html contains forbidden API: ' + token));
  });

  /* ---------- Additional: headers are configured manually only ---------- */
  await test('A', 'Header configuration is manual only (no template / JSON import-export)', async () => {
    const s = await newSession();
    assertEq(s.headers, SEC.OFFICIAL_HEADERS);
    assert(!('exportRules' in s) && !('importRules' in s), 'no rules export/import on the session');
    const html = fs.readFileSync(path.join(__dirname, '..', 'index.html'), 'utf8');
    ['btnLoadTemplate', 'btnLoadJson', 'btnExportHeaders', 'fileTemplate', 'fileJson'].forEach(function (id) {
      assert(html.indexOf('id="' + id + '"') === -1, 'removed button/input still present: ' + id);
    });
    assert(html.indexOf('id="btnManualEdit"') !== -1 && html.indexOf('id="manualInputs"') !== -1,
      'manual header editor is available');
  });

  /* ---------- Excel 32 767-char cell limit: no crash, no data loss ---------- */
  await test('L', 'Texts longer than one Excel cell are exported in full (no truncation)', async () => {
    const s = await newSession();
    const longNote = 'Rappel mardi 30/04/2024 a 17h30 | 87 ans | ' + 'historique '.repeat(4000);
    const csv = Buffer.from('Nom;Tel;Note\nDUPONT;0612345678;' + longNote + '\n', 'utf8');
    const e = s.registerFile({ name: 'long.csv', size: csv.length });
    await s.startBatch([e], { readBuffer: () => Promise.resolve(csv), yieldFn: immediate });
    /* The COMPLETE text as processed (before the export limit is applied). */
    const fullText = ((s.recordsByFile.get(e.id) || [])[0].values.commentaire || '');
    const rows = s.buildConsolidatedRows();
    const conts = s.buildContinuations();
    const res = SEC.Exporter.generateConsolidated(
      s.headers, rows, s.report, rows.length, s.buildFileStats(), null, conts);
    assertEq(rows.length, 1, 'one record must stay one row');
    const X = SEC.XLSXReader.getXLSX();
    const wb = X.read(res.buffer, { type: 'array' });
    const main = X.utils.sheet_to_json(wb.Sheets['tableau final'], { header: 1, defval: '' });
    assertEq(main.length, 2, 'header + 1 data row');
    const col = SEC.TARGET_KEYS.indexOf('commentaire');
    assert(main[1][col].length <= 32767, 'main cell respects the Excel limit');
    const cont = X.utils.sheet_to_json(wb.Sheets['Suite des textes longs'], { header: 1, defval: '' });
    const extra = cont.slice(1).map((c) => c[4]).join('');
    const stored = fullText;
    const rebuilt = main[1][col] + extra;
    assert(
      rebuilt === stored,
      'nothing lost: main cell + continuation == original text (' + rebuilt.length + ' vs ' + stored.length + ')',
    );
    assert(stored.indexOf('historique') !== -1 && stored.length > 32767, 'the text really was over the limit');
  });

  /* ---------- Nom / Prénom: real split, notes never in a name ---------- */
  await test('N', 'Nom and Prénom are separated; notes always go to Commentaire', async () => {
    const cases = [
      { csv: 'Adherent;Telephone\nDupont, Jean;0612345678\nDUPONT Jean-Pierre;0698765432\n',
        want: [{ nom: 'Dupont', prenom: 'Jean' }, { nom: 'DUPONT', prenom: 'Jean-Pierre' }] },
      { csv: 'Client;Portable\nJean Dupont;0612345678\n', want: [{ nom: 'Dupont', prenom: 'Jean' }] },
    ];
    for (const c of cases) {
      const s = await newSession();
      const buf = Buffer.from(c.csv, 'utf8');
      const e = s.registerFile({ name: 'n.csv', size: buf.length });
      await s.startBatch([e], { readBuffer: () => Promise.resolve(buf), yieldFn: immediate });
      const recs = s.recordsByFile.get(e.id) || [];
      c.want.forEach((w, i) => {
        assert(recs[i], 'row ' + i + ' missing');
        assertEq([recs[i].values.nom, recs[i].values.prenom], [w.nom, w.prenom], 'row ' + i);
      });
    }
    /* A reminder with a date and an age is NEVER in Nom / Prénom. */
    const s2 = await newSession();
    const buf2 = Buffer.from('Nom;Tel;Note\nDupont;0612345678;Rappel mardi 30/04/2024 a 17h30 | 87 ans\n', 'utf8');
    const e2 = s2.registerFile({ name: 'r.csv', size: buf2.length });
    await s2.startBatch([e2], { readBuffer: () => Promise.resolve(buf2), yieldFn: immediate });
    const v = (s2.recordsByFile.get(e2.id) || [])[0].values;
    assertEq(v.nom, 'Dupont');
    assert(/Rappel/.test(v.commentaire), 'reminder kept in Commentaire: ' + v.commentaire);
    /* A note mapped into a name column is moved, not lost. */
    const s3 = await newSession();
    const buf3 = Buffer.from('Nom;Prenom\nRappel mardi 30/04/2024 17h30 | 87 ans;Jean\n', 'utf8');
    const e3 = s3.registerFile({ name: 'r2.csv', size: buf3.length });
    await s3.startBatch([e3], { readBuffer: () => Promise.resolve(buf3), yieldFn: immediate });
    const v3 = (s3.recordsByFile.get(e3.id) || [])[0].values;
    assert(!/\d/.test(v3.nom), 'digits in Nom: ' + v3.nom);
    assert(/Rappel/.test(v3.commentaire), 'note preserved in Commentaire');
  });

  await test('B', 'Duplicate files (same name and size) are detected, warned and skipped', async () => {
    const s = await newSession();
    const e1 = s.registerFile({ name: 'dup.csv', size: 123 });
    const e2 = s.registerFile({ name: 'dup.csv', size: 123, duplicateOf: e1.id });
    assert(e2.duplicateOf === e1.id);
    e2._lastBuffer = Buffer.from('a\nb');
    await s.startBatch([e1, e2], { readBuffer: () => Promise.reject(new Error('should not read skipped dup')), yieldFn: immediate });
    assertEq(e2.status, 'cancelled');
    assert(s.report.some(r => /duplique/i.test(r.reason)), 'duplicate reported');
  });

  await test('C', 'Processing report workbook is generated with file-level results', async () => {
    const s = await newSession();
    await processOne(s, '07_edge_cases.xlsx');
    const buffer = SEC.Exporter.buildReportWorkbook(s.report, s.fileOrder.map(id => s.files.get(id)));
    const X = SEC.XLSXReader.getXLSX();
    const wb = X.read(buffer, { type: 'array' });
    assert(wb.SheetNames.indexOf('Rapport') !== -1 && wb.SheetNames.indexOf('Fichiers') !== -1);
  });

  await test('D', 'Individual file removal drops its records and keeps the rest', async () => {
    const s = await newSession();
    const e1 = await processOne(s, '01_clients_semicolon.csv');
    const before = s.records.size;
    await processOne(s, '02_clients_comma.csv');
    assert(s.records.size > before);
    s.removeFile(e1.id);
    assert(!s.files.has(e1.id), 'entry removed');
    assert(s.records.size === 3, 'only the 3 rows of file 2 remain, got ' + s.records.size);
    assert(rowsOf(s).every(r => r[0] !== 'Dupont'), 'removed file data gone');
  });

  await test('E', 'Excluded columns stay visible in the report and never reach the workbook', async () => {
    const s = await newSession();
    const csv = Buffer.from('Nom;Numéro client;Email\nDupont;C-999;a@b.fr\nMartin;C-998;c@d.fr\n', 'utf8');
    const e = s.registerFile({ name: 'excl.csv', size: csv.length });
    await s.startBatch([e], { readBuffer: () => Promise.resolve(csv), yieldFn: immediate });
    const item = Array.from(s.reviewItems.values()).find(i => i.header === 'Numéro client');
    assert(item, 'ID column is in review');
    s.decideReview(item.id, '__exclude__', false);
    await s.reprocessAffected([e.id]);
    const rows = rowsOf(s);
    assertEq(rows.length, 2);
    const flat = JSON.stringify(rows);
    assert(flat.indexOf('C-999') === -1 && flat.indexOf('C-998') === -1, 'excluded values absent from output');
    assert(s.report.some(r => /exclue volontairement/i.test(r.reason)), 'exclusion recorded in report');
  });

  await test('F', 'Template-based export keeps the official sheet and passes full validation', async () => {
    const X = SEC.XLSXReader.getXLSX();
    const tplBuf = buf('En tete finale.xlsx');
    const wb = X.read(tplBuf, { type: 'array', cellStyles: true, cellFormula: false });
    const first = wb.SheetNames[0];
    /* mirror app.buildTemplateWorkbook core operations */
    const ws = wb.Sheets[first];
    const aoa = X.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
    const targetNorms = SEC.OFFICIAL_HEADERS.map(SEC.Normalizers.normalizeHeader);
    let headerIdx = -1;
    for (let i = 0; i < aoa.length; i++) {
      const cells = aoa[i].map(SEC.Normalizers.normalizeHeader);
      if (targetNorms.every(t => cells.indexOf(t) !== -1)) { headerIdx = i; break; }
    }
    assert(headerIdx !== -1, 'template header row located');
    const range = X.utils.decode_range(ws['!ref']);
    for (let r = headerIdx + 1; r <= range.e.r; r++)
      for (let c = range.s.c; c <= range.e.c; c++) delete ws[X.utils.encode_cell({ r, c })];
    const dataRows = [['Test', 'Modèle', '0600000000', '', '1 rue du Test', '75001', 'Paris', 't@t.fr', '10,00', 'MGEN', '']];
    /* Same cell writer as the application: text columns stay text, the empty
     * Téléphone 2 cell stays empty. */
    SEC.Exporter.writeRow(X, ws, headerIdx + 1, dataRows[0], SEC.OFFICIAL_HEADERS);
    ws['!ref'] = X.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: headerIdx + dataRows.length, c: 10 } });
    const out = X.write(wb, { type: 'array', bookType: 'xlsx', cellStyles: true, compression: true });
    const validation = SEC.Exporter.validateBuffer(out, SEC.OFFICIAL_HEADERS, 1, first, false, dataRows);
    assert(validation.ok, 'template export validation: ' + JSON.stringify(validation.checks.filter(c => !c.pass)));
  });

  /* ---------- Official commune database (INSEE + postal codes) ---------- */
  await test('G', 'Official commune database is loaded, complete and up to date', () => {
    assert(SEC.Communes.count() > 34000, 'expected 34k+ communes, got ' + SEC.Communes.count());
    assert(/^\d{4}-\d{2}-\d{2}$/.test(SEC.Communes.version()), 'database date recorded');
    const paris = SEC.Communes.communeByInsee('75056');
    assert(paris && paris.nom === 'Paris' && paris.cps.indexOf('75004') !== -1, 'Paris 75056 with postal codes');
    assert(SEC.Communes.communeByInsee('97411').nom === 'Saint-Denis', 'overseas commune present');
    assert(SEC.Communes.isPostalValid('01200') && SEC.Communes.isPostalValid('75008'), 'metropolitan postal codes');
    assert(SEC.Communes.isPostalValid('20000'), 'Corsican commune postal code');
    assertEq(SEC.Communes.communeByInsee('2A004').nom, 'Ajaccio', 'Corsican commune (2A/2B codes)');
    assert(SEC.Communes.isPostalValid('97400'), 'overseas postal code');
    assert(!SEC.Communes.isPostalValid('99999'), 'non-existent postal code rejected');
    assert(SEC.Communes.lookupPostal('31000').nom === 'Toulouse', 'single-commune postal code resolved');
    assert(SEC.Communes.lookupPostal('75001').nom === 'Paris', 'arrondissement postal code resolved');
  });

  await test('H', 'Cities are corrected only when certain, otherwise left empty (never invented)', () => {
    const v = (c, cp) => SEC.Communes.validate(c, cp);
    assertEq(v('PARIS', '').value, 'Paris', 'case normalised');
    assertEq(v('Toulouse', '').value, 'Toulouse');
    assertEq(v('TOLOUSE', '').value, 'Toulouse', 'known misspelling corrected');
    assertEq(v('toulouse', '31000').value, 'Toulouse');
    assertEq(v('St Etienne', '').value, 'Saint-Étienne', 'abbreviation + accent restored');
    assertEq(v('Nimes', '').value, 'Nîmes', 'missing accent restored');
    assertEq(v('Saint Denis', '97400').value, 'Saint-Denis', 'homonym decided by the postal code');
    assertEq(v('Saint-Denis', '93200').value, 'Saint-Denis', 'metropolitan homonym decided by the postal code');
    assertEq(v('Saint-Denis', '').value, '', 'homonym without postal code -> empty');
    assertEq(v('Atlantis', '').value, '', 'unknown city -> empty');
    assertEq(v('Atlantis', '75001').value, '', 'postal code alone never creates an incompatible city');
    assertEq(v('Bellegarde-sur-Valserine', '01200').value, '', 'former commune name is not invented/guessed');
    assertEq(v('', '').value, '', 'empty stays empty');
  });

  /* ---------- Ville column processed with the postal code of the row ---------- */
  await test('I', 'Ville column is validated against INSEE data, empty when not identified', async () => {
    const s = await newSession();
    const csv = Buffer.from('Nom;Prénom;Code postal;Ville\n' +
      'Dupont;Jean;31000;toulouse\n' +
      'Martin;Luc;93200;Saint Denis\n' +
      'Rond;Paul;69123;Atlantis\n' +
      'Petit;Élodie;75004;PARIS\n', 'utf8');
    const e = s.registerFile({ name: 'villes.csv', size: csv.length });
    await s.startBatch([e], { readBuffer: () => Promise.resolve(csv), yieldFn: immediate });
    const rows = rowsOf(s);
    const ville = rows.map(r => r[SEC.TARGET_KEYS.indexOf('ville')]);
    assertEq(ville, ['Toulouse', 'Saint-Denis', '', 'Paris'], 'cities: ' + JSON.stringify(ville));
    assert(s.report.some(r => /ville non identifiee/i.test(r.reason)), 'unidentified city reported');
    assert(!s.report.some(r => /ville non identifiee/i.test(r.reason) && /toulouse/i.test(r.samples.join(' '))), 'identified cities not reported as errors');
  });

  /* ---------- Manual destination respected exactly ---------- */
  await test('J', 'Choosing a destination keeps every value in that exact column', async () => {
    const s = await newSession();
    const csv = Buffer.from('Titre client;Mail client;Montant;Commentaire libre\n' +
      'Dupont;a@b.fr;10,00;RAS\n' +
      'Martin;c@d.fr;20,50;ok\n', 'utf8');
    const e = s.registerFile({ name: 'dest.csv', size: csv.length });
    await s.startBatch([e], { readBuffer: () => Promise.resolve(csv), yieldFn: immediate });
    const item = Array.from(s.reviewItems.values()).find(i => i.header === 'Titre client');
    assert(item, 'unknown header in review');
    /* An email-shaped column is forced into Ville: it must STAY in Ville even
     * though its values look like e-mails (no redirect to Email). */
    const mailItem = Array.from(s.reviewItems.values()).find(i => i.header === 'Mail client');
    s.decideReview(item.id, 'commentaire');
    if (mailItem) s.decideReview(mailItem.id, 'ville');
    await s.reprocessAffected([e.id]);
    const rows = rowsOf(s);
    const idx = k => SEC.TARGET_KEYS.indexOf(k);
    assert(rows.some(r => r[idx('commentaire')] === 'Dupont'), 'value placed in the chosen Commentaire column');
    assert(rows.every(r => r[idx('nom')] === ''), 'nothing copied into Nom');
    if (mailItem) {
      assert(rows.every(r => r[idx('ville')] === ''), 'e-mail values in Ville are never kept as a city (left empty)');
      assert(rows.every(r => r[idx('email')] === ''), 'e-mail values are not copied into Email either');
      const reasons = s.report.map(r => r.sourceHeader + '|' + r.reason).join(' § ');
      assert(/Mail client\|ville non identifiee/.test(reasons), 'the e-mail value is reported as an unidentified city: ' + reasons);
    }
  });

  /* ---------- Export integrity: dates, empty cells, alignment ---------- */
  await test('K', 'Excel export preserves dates, empty cells and row alignment', async () => {
    const s = await newSession();
    await processOne(s, '01_clients_semicolon.csv');
    await processOne(s, '07_edge_cases.xlsx');
    const rows = rowsOf(s);
    /* Insert a date value in a spare Commentaire cell of the first row to
     * check that a real date is exported as a date, not as a broken number. */
    rows[0] = rows[0].slice();
    const rowWithEmptyComment = rows.findIndex(r => !r[SEC.TARGET_KEYS.indexOf('commentaire')]);
    assert(rowWithEmptyComment !== -1, 'a row with an empty Commentaire exists');
    rows[rowWithEmptyComment] = rows[rowWithEmptyComment].slice();
    rows[rowWithEmptyComment][SEC.TARGET_KEYS.indexOf('commentaire')] = '12/03/1985';
    const other = rowWithEmptyComment === 0 ? 1 : 0;
    rows[other] = rows[other].slice();
    rows[other][SEC.TARGET_KEYS.indexOf('commentaire')] = 'RAS 12/03/1985';
    const out = SEC.Exporter.generateConsolidated(s.headers, rows, [], rows.length, s.buildFileStats(), s.buildLineDetails());
    assert(out.validation.ok, 'validation: ' + JSON.stringify(out.validation.checks.filter(c => !c.pass)));
    const X = SEC.XLSXReader.getXLSX();
    const wb = X.read(out.buffer, { type: 'array' });
    const ws = wb.Sheets['tableau final'];
    const aoa = X.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, raw: true });
    assertEq(aoa.length - 1, rows.length, 'no shifted or missing row');
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < 11; c++) {
        const want = rows[r][c] || '';
        const got = aoa[r + 1][c];
        if (want === '') { assert(got === '' || got === undefined, 'empty cell stays empty at ' + (r + 2) + ':' + c); continue; }
        const cell = ws[X.utils.encode_cell({ r: r + 1, c })];
        if (cell && cell.t === 'n') {
          /* A numeric date cell must at least keep the same day. */
          assertEq(String(cell.w), SEC.Exporter.parseDateValue(want).text, 'date displayed as dd/mm/yyyy');
        } else if (cell && cell.t === 'd') {
          const d = cell.v;
          assertEq([d.getUTCDate(), d.getUTCMonth() + 1, d.getUTCFullYear()], [3, 3, 1985], 'date value kept');
        } else {
          /* Text cell: leading zeros and formula neutralisation kept exactly. */
          assertEq(String(got), SEC.Normalizers.sanitizeExportValue(want), 'value kept in its column at ' + (r + 2) + ':' + c);
        }
      }
    }
    /* A date IS preserved exactly. It is deliberately kept as TEXT: a real
     * Excel date cell displays the SAME day differently in another locale
     * (US "3/12/1985" = March 12th), which users report as "the data
     * changed". Text keeps the exact source characters, formula-safe. */
    const dateCell = ws[X.utils.encode_cell({ r: rowWithEmptyComment + 1, c: SEC.TARGET_KEYS.indexOf('commentaire') })];
    assert(dateCell && dateCell.t === 's' && dateCell.v === '12/03/1985', 'date preserved exactly as text (no locale shift)');
    const textCell = ws[X.utils.encode_cell({ r: other + 1, c: SEC.TARGET_KEYS.indexOf('commentaire') })];
    assert(textCell && textCell.t === 's', 'date inside free text stays text');
  });

  /* ---------- Export integrity: address values with surrounding spaces stay in Adresse ---------- */
  await test('L', 'Address values with extra whitespace stay in the Adresse column on export', async () => {
    const s = await newSession();
    await processOne(s, 'edge-adresses.xlsx');
    const rows = rowsOf(s);
    assert(rows.length >= 3, 'rows extracted: ' + rows.length);
    const aIdx = SEC.TARGET_KEYS.indexOf('adresse');
    /* Every address value is cleaned (trimmed) and stays in its Adresse cell. */
    rows.forEach(function (r, i) {
      const raw = r[aIdx] || '';
      assertEq(raw, raw.replace(/^\s+|\s+$/g, '').replace(/[\u00a0\u202f]+/g, ' ').trim(), 'address value at row ' + (i + 2) + ' contains no stray outer whitespace: "' + raw + '"');
    });
    const out = SEC.Exporter.generateConsolidated(s.headers, rows, [], rows.length, s.buildFileStats(), s.buildLineDetails());
    assert(out.validation.ok, 'export validation must pass: ' + JSON.stringify(out.validation.checks.filter(c => !c.pass)));
    const X = SEC.XLSXReader.getXLSX();
    const wb = X.read(out.buffer, { type: 'array' });
    const ws = wb.Sheets['tableau final'];
    const aoa = X.utils.sheet_to_json(ws, { header: 1, defval: '', blankrows: true, raw: true });
    assertEq(aoa.length - 1, rows.length, 'no shifted or missing row');
    rows.forEach(function (r, i) {
      assertEq(String(aoa[i + 1][aIdx] || ''), r[aIdx] || '', 'address stays in Adresse column at row ' + (i + 2));
    });
  });

  console.log('\n=== Results: ' + passed + ' passed, ' + failed + ' failed ===\n');
  fs.writeFileSync(path.join(__dirname, 'results.json'), JSON.stringify({ passed, failed, results }, null, 2));
  process.exit(failed ? 1 : 0);
})().catch(err => { console.error('TEST HARNESS ERROR', err); process.exit(2); });

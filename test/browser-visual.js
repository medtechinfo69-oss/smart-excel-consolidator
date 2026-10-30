/* Real-Chromium end-to-end verification via Playwright.
 * Run: node test/browser-visual.js
 * Screenshots -> test/screenshots, downloaded workbook -> test/downloads. */
'use strict';
const path = require('path');
const fs = require('fs');
const { chromium } = require('playwright-core');

const ROOT = path.join(__dirname, '..');
const FX = path.join(__dirname, 'fixtures');
const CHROME = require('child_process').execSync(
  'ls -d ~/.cache/ms-playwright/chromium-*/chrome-linux/chrome | head -1', { shell: '/bin/bash' }
).toString().trim();
fs.mkdirSync(path.join(__dirname, 'screenshots'), { recursive: true });
fs.mkdirSync(path.join(__dirname, 'downloads'), { recursive: true });
const files = [
  '01_clients_semicolon.csv', '02_clients_comma.csv', '03_tab_delimited.txt',
  '04_legacy_cp1252.csv', '05_workbook_multi.xlsx', '06_ambiguous.xlsx', '07_edge_cases.xlsx'
].map(f => path.join(FX, f));

let passed = 0, failed = 0;
function check(name, cond) {
  if (cond) { passed++; console.log('  PASS [browser] ' + name); }
  else { failed++; console.log('  FAIL [browser] ' + name); }
}

(async () => {
  const browser = await chromium.launch({ executablePath: CHROME, headless: true, args: ['--no-sandbox'] });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  const consoleErrors = [];
  page.on('console', m => { if (m.type() === 'error') consoleErrors.push(m.text()); });
  page.on('pageerror', e => consoleErrors.push(String(e)));

  await page.goto('file://' + path.join(ROOT, 'index.html'));
  await page.waitForSelector('#headerList li', { timeout: 10000 });
  check('11 headers rendered', (await page.$$('#headerList li')).length === 11);
  check('start initially disabled', await page.$eval('#btnStart', el => el.disabled));
  await page.screenshot({ path: path.join(__dirname, 'screenshots/01-initial.png'), fullPage: true });

  /* Upload through the real hidden file input */
  await page.setInputFiles('#fileInput', files);
  await page.waitForFunction(() => document.querySelectorAll('#fileTableBody tr').length === 7);
  check('7 files listed', (await page.$$('#fileTableBody tr')).length === 7);
  check('start enabled after selection', await page.$eval('#btnStart', el => !el.disabled));
  await page.screenshot({ path: path.join(__dirname, 'screenshots/02-files-loaded.png'), fullPage: true });

  /* Start and wait for completion */
  await page.click('#btnStart');
  check('start locked while running', await page.$eval('#btnStart', el => el.disabled));
  await page.waitForFunction(() => document.querySelector('#sumRows').textContent !== '0', { timeout: 30000 });
  await page.waitForFunction(() => document.querySelector('#btnCancel').disabled === true, { timeout: 30000 });
  const rows = Number(await page.$eval('#sumRows', el => el.textContent));
  check('rows extracted (' + rows + ')', rows > 20);
  check('zero failed files', (await page.$eval('#sumFailed', el => el.textContent)) === '0');
  check('no manual review section', (await page.$$('#reviewSection')).length === 0);
  await page.screenshot({ path: path.join(__dirname, 'screenshots/03-processing-review.png'), fullPage: true });

  /* Details modal: worksheet selection + header-row correction on the multi-sheet workbook */
  await page.$$eval('#processTableBody tr', rows => {
    const tr = rows.find(r => r.textContent.includes('05_workbook_multi.xlsx'));
    tr.querySelector('[data-action="details"]').click();
  });
  await page.waitForSelector('#modalBackdrop:not([hidden])', { timeout: 10000 });
  const sheetChecks = await page.$$eval('#modalBody [data-sheet-name]', els => els.map(e => ({ name: e.getAttribute('data-sheet-name'), checked: e.checked, disabled: e.disabled })));
  check('modal lists all worksheets incl. the empty one', sheetChecks.length === 3);
  check('empty worksheet checkbox is disabled', sheetChecks.some(s => s.name === 'Feuille vide' && s.disabled));
  check('two non-empty worksheets selected', sheetChecks.filter(s => s.checked).length === 2);
  check('header-row input present (empty rows are compacted)', await page.$eval('#modalBody input[type="number"]', el => el.value === '3'));
  await page.click('#btnCloseModal');
  await page.waitForFunction(() => document.querySelector('#modalBackdrop').hidden === true, { timeout: 5000 });
  check('modal closes and stops intercepting clicks', await page.$eval('#modalBackdrop', el => el.hidden === true));

  /* Columns were resolved automatically during processing */
  await page.waitForFunction(() => document.querySelector('#btnExportExcel').disabled === false, { timeout: 30000 });
  check('export enabled after automatic processing', await page.$eval('#btnExportExcel', el => !el.disabled));
  check('preview populated', (await page.$$('#previewTable tbody tr')).length > 0);

  /* Export and validate the real downloaded workbook */
  const downloadPromise = page.waitForEvent('download', { timeout: 20000 });
  await page.click('#btnExportExcel');
  const download = await downloadPromise;
  const outPath = path.join(__dirname, 'downloads', 'fichier_final.xlsx');
  await download.saveAs(outPath);
  check('workbook downloaded', fs.existsSync(outPath) && fs.statSync(outPath).size > 5000);
  const allPass = await page.$$eval('#validationList li', lis => lis.length > 5 && lis.every(li => li.classList.contains('pass')));
  check('all export validation checks pass', allPass);

  /* Header configuration is manual only (no template / JSON buttons). */
  check('header template/JSON buttons removed',
    !(await page.$('#btnLoadTemplate')) && !(await page.$('#btnLoadJson')) && !(await page.$('#btnExportHeaders')));

  await page.screenshot({ path: path.join(__dirname, 'screenshots/04-final-export.png'), fullPage: true });

  /* Mobile viewport sanity */
  await page.setViewportSize({ width: 390, height: 844 });
  await page.waitForTimeout(300);
  await page.screenshot({ path: path.join(__dirname, 'screenshots/05-mobile.png'), fullPage: false });
  const noHorizontalScroll = await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1);
  check('mobile layout has no horizontal scroll', noHorizontalScroll);

  check('zero browser console errors', consoleErrors.length === 0);
  if (consoleErrors.length) console.log('   console:', consoleErrors.slice(0, 5));

  await browser.close();
  const XLSX = require('../vendor/xlsx.full.min.js');
  const expected = ['Nom', 'Prénom', 'Téléphone 1', 'Téléphone 2', 'Adresse', 'Code postal', 'Ville', 'Email', 'Cotisation', 'Mutuelle', 'Commentaire'];
  const wb = XLSX.read(fs.readFileSync(outPath), { type: 'buffer' });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets['tableau final'], { header: 1, defval: '' });
  check('downloaded workbook: exact 11 headers', JSON.stringify(aoa[0]) === JSON.stringify(expected));
  check('downloaded workbook: data rows = ' + (aoa.length - 1), aoa.length - 1 === rows);
  check('downloaded workbook: Rapport sheet', wb.SheetNames.indexOf('Rapport') !== -1);

  console.log('\n=== Browser results: ' + passed + ' passed, ' + failed + ' failed ===\n');
  fs.writeFileSync(path.join(__dirname, 'browser-results.json'), JSON.stringify({ passed, failed, consoleErrors }, null, 2));
  process.exit(failed ? 1 : 0);
})().catch(e => { console.error('BROWSER HARNESS ERROR', e); process.exit(2); });

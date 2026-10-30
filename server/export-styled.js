/* ============================================================
 * Smart Excel - optional styled .xlsx builder
 * ------------------------------------------------------------
 * Reached ONLY through POST /api/export/xlsx when a spreadsheet
 * library with style-write support is installed (exceljs is the
 * recommended one: `npm i exceljs`). It is feature-detected: when
 * the module is absent the endpoint answers { ok:false, code:'no-lib' }
 * and the browser keeps its own 100% local export.
 *
 * Input (JSON, built in the browser):
 *   { headers: [11], rows: [[...]],
 *     template?: { base64, sheet, headerRowIndex },
 *     report?: [[...]], extraSheets?: [{name, rows}] }
 * Output: Buffer of the .xlsx file.
 * ============================================================ */
"use strict";

const HEADER_FILL = "$FF1D4ED8";
const LEADING_ZERO_RE = /^0\d+$/;

function textColumn(index) {
  /* Téléphone 1, Téléphone 2, Adresse, Code postal, Cotisation */
  return [2, 3, 4, 5, 8].indexOf(index) !== -1;
}

async function buildStyledWorkbook(libName, body) {
  let ExcelJS;
  try {
    ExcelJS = require(
      require.resolve(libName, {
        paths: [__dirname + "/..", __dirname],
      }),
    );
  } catch (e) {
    throw new Error("Bibliothèque d'export introuvable: " + libName);
  }

  const headers = Array.isArray(body.headers) ? body.headers.slice(0, 11) : [];
  const rows = Array.isArray(body.rows) ? body.rows : [];
  const wb = new ExcelJS.Workbook();
  wb.creator = "Smart Excel Consolidator";
  wb.created = new Date();

  let ws;
  let headerRowIndex = 0;

  if (body.template && body.template.base64) {
    const templateBuffer = Buffer.from(body.template.base64, "base64");
    await wb.xlsx.load(templateBuffer);
    ws =
      wb.getWorksheet(body.template.sheet) ||
      wb.worksheets[0] ||
      wb.addWorksheet("tableau final");
    headerRowIndex = Number.isInteger(body.template.headerRowIndex)
      ? body.template.headerRowIndex
      : 0;
  } else {
    ws = wb.addWorksheet("tableau final");
  }

  /* Header row: keep the template's own styling when possible. */
  const headerRow = ws.getRow(headerRowIndex + 1);
  headers.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = String(h);
    cell.font = Object.assign(
      { bold: true, color: { argb: "FFFFFFFF" } },
      cell.font || {},
    );
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: HEADER_FILL },
    };
    cell.alignment = { vertical: "middle" };
  });

  rows.forEach((r, ri) => {
    const row = ws.getRow(headerRowIndex + 2 + ri);
    for (let c = 0; c < headers.length; c++) {
      const value = r[c] == null ? "" : String(r[c]);
      if (value === "") {
        row.getCell(c + 1).value = null;
        continue;
      }
      const cell = row.getCell(c + 1);
      if (textColumn(c)) {
        cell.value = value;
        cell.numFmt = "@"; // keep "0612345678" / "01200" exactly as text
      } else if (
        /^-?\d+(?:\.\d+)?$/.test(value) &&
        !LEADING_ZERO_RE.test(value)
      ) {
        cell.value = Number(value);
      } else {
        cell.value = value;
      }
    }
  });

  const last = headerRowIndex + 1 + Math.max(rows.length, 1);
  for (let r = headerRowIndex + 2; r <= last; r++) {
    const row = ws.getRow(r);
    for (let c = 0; c < headers.length; c++) {
      row.getCell(c + 1).border = {
        top: { style: "thin", color: { argb: "FFE2E8F0" } },
        bottom: { style: "thin", color: { argb: "FFE2E8F0" } },
      };
      row.getCell(c + 1).alignment = { vertical: "middle", wrapText: false };
    }
  }
  ws.views = [{ state: "frozen", ySplit: headerRowIndex + 1 }];
  ws.autoFilter = {
    from: { row: headerRowIndex + 1, column: 1 },
    to: { row: last, column: Math.max(headers.length, 1) },
  };

  if (Array.isArray(body.report) && body.report.length) {
    const rep = wb.addWorksheet("Rapport");
    body.report.forEach((r, ri) => {
      rep.addRow(r.map((v) => (v == null ? "" : String(v))));
      if (ri === 0) rep.getRow(1).font = { bold: true };
    });
    rep.columns.forEach((col) => {
      col.width = 24;
    });
  }
  (body.extraSheets || []).forEach((extra) => {
    if (!extra || !Array.isArray(extra.rows) || !extra.rows.length) return;
    const sheet = wb.addWorksheet(String(extra.name || "Feuille").slice(0, 31));
    extra.rows.forEach((r, ri) => {
      sheet.addRow(r.map((v) => (v == null ? "" : String(v))));
      if (ri === 0) sheet.getRow(1).font = { bold: true };
    });
  });

  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out);
}

module.exports = { buildStyledWorkbook };

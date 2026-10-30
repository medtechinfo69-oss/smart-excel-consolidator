/* ============================================================
 * Smart Excel - template-preserving workbook model
 * ------------------------------------------------------------
 * A plain-JSON snapshot of an official template workbook:
 *   { sheets: [{ name, cols, ref, rows: [{ h, cells: [...] }] }] }
 * where every cell carries its RAW SheetJS value (value, type, text,
 * number format, style index, formula), NOT the decompressed XML.
 * The snapshot crosses the worker boundary as JSON, the data rows are
 * rewritten in the worker and every untouched header/title cell keeps
 * its exact style -> the template formatting survives the export.
 * Used by both the browser worker and the Node export tests.
 * ============================================================ */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});



  /* Fully cached .xlsx files (written by SheetJS with compression true) contain
   * stored entries only. Uncompressed or freshly created workbooks may contain
   * deflated entries, which SheetJS only decodes when CFB.decompress is on. */
  var inflateCacheStatus = null; // 'checked' | undefined
  function conditionalInflate(X, buffer) {
    /* PERF + SAFETY: SheetJS reads its own (uncompressed) output natively, so
     * the CFB probe is only needed for a FOREIGN file that may contain
     * deflated entries. The old code parsed the whole archive with CFB on
     * every call (including each write) and could hand SheetJS a buffer whose
     * package relations were no longer valid — a silent source of template
     * corruption. The probe is done once; its result is cached. */
    if (inflateCacheStatus === "skip") return buffer;
    try {
      var f = X.CFB && X.CFB.read(new Uint8Array(buffer), { type: "array" });
      if (!f) {
        inflateCacheStatus = "skip";
        return buffer;
      }
      var anyDeflate = false;
      (f.FullPaths || []).forEach(function (_p, i) {
        var e = f.FileIndex && f.FileIndex[i];
        if (e && e.content && e.size > 0 && e.size < (e.content.length || 0))
          anyDeflate = true;
      });
      if (anyDeflate) {
        return X.CFB.utils.cfb2container(
          X.CFB.utils.compress(
            X.CFB.read(new Uint8Array(buffer), { type: "array" }),
          ),
        );
      }
      return buffer;
    } catch (e) {
      /* Not a CFB container at all (SheetJS output): read it as-is. */
      inflateCacheStatus = "skip";
      return buffer;
    }
  }

  function cloneCell(cell) {
    var out = { t: cell.t };
    if (cell.v !== undefined) {
      /* Dates cannot cross JSON: store as ISO + keep the date type. */
      out.v = cell.v instanceof Date ? cell.v.toISOString() : cell.v;
      if (cell.v instanceof Date) out.t = "d";
    }
    if (cell.f !== undefined) out.f = cell.f;
    if (cell.z !== undefined) out.z = cell.z;
    if (cell.s !== undefined) out.s = cell.s;
    if (cell.w !== undefined) out.w = cell.w;
    if (cell.c !== undefined && Array.isArray(cell.c)) {
      out.c = cell.c.map(function (run) {
        return { a: run.a, t: run.t };
      });
    }
    return out;
  }

  function extractModel(X, wb) {
    return {
      sheets: (wb.SheetNames || []).map(function (name) {
        var ws = wb.Sheets[name] || {};
        var rows = [];
        var ref = ws["!ref"] || "A1";
        var range = X.utils.decode_range(ref);
        for (var r = 0; r <= range.e.r; r++) {
          var cells = [];
          for (var c = 0; c <= range.e.c; c++) {
            var cell = ws[X.utils.encode_cell({ r: r, c: c })];
            cells.push(cell ? cloneCell(cell) : null);
          }
          rows.push({
            h: ws["!rows"] && ws["!rows"][r] ? ws["!rows"][r] : null,
            cells: cells,
          });
        }
        return {
          name: name,
          cols: ws["!cols"]
            ? ws["!cols"].map(function (c) {
                return c
                  ? { wch: c.wch, wpx: c.wpx, hidden: c.hidden, width: c.width }
                  : null;
              })
            : null,
          ref: ref,
          rows: rows,
          merges: ws["!merges"] || null,
        };
      }),
    };
  }

  function reviveCell(cell) {
    if (!cell) return null;
    var out = {};
    Object.keys(cell).forEach(function (k) {
      out[k] = cell[k];
    });
    if (out.t === "d" && typeof out.v === "string") out.v = new Date(out.v);
    return out;
  }

  /* Rewrite the data rows of the official sheet (the one whose header row
   * matches the 11 headers) with `rows`, exactly like SEC.Exporter.writeRow,
   * without touching any other cell. */
  function applyRows(X, wb, model, rows) {
    var target = null;
    for (var i = 0; i < model.sheets.length && !target; i++) {
      var sh = model.sheets[i];
      var headerIdx = -1;
      for (
        var r = 0;
        r < Math.min(sh.rows.length, SEC.LIMITS.MAX_SCAN_HEADER_ROWS);
        r++
      ) {
        var texts = (sh.rows[r].cells || []).map(function (c) {
          return c ? String(c.v == null ? "" : c.v) : "";
        });
        var norms = texts.map(SEC.Normalizers.normalizeHeader);
        var ok = SEC.TARGET_KEYS.every(function (k) {
          return (
            norms.indexOf(
              SEC.Normalizers.normalizeHeader(SEC.TARGET_LABELS[k]),
            ) !== -1
          );
        });
        var names = SEC.TARGET_KEYS.every(function (k) {
          return texts.indexOf(SEC.TARGET_LABELS[k]) !== -1;
        });
        if (ok && names) {
          headerIdx = r;
          break;
        }
      }
      if (headerIdx !== -1) target = { sheet: sh, headerIdx: headerIdx };
    }
    if (!target) return null;

    var sh = target.sheet;
    var headerIdx = target.headerIdx;

    var ws = wb.Sheets[sh.name] || {};
    /* Clear every data cell of the original range (and the extra rows). */
    var oldRange = X.utils.decode_range(ws["!ref"] || sh.ref);
    for (var rr = headerIdx + 1; rr <= oldRange.e.r; rr++) {
      for (var cc = oldRange.s.c; cc <= oldRange.e.c; cc++)
        delete ws[X.utils.encode_cell({ r: rr, c: cc })];
    }
    for (var nr = headerIdx + 1; nr <= headerIdx + rows.length; nr++) {
      for (var nc = 0; nc < 11; nc++)
        delete ws[X.utils.encode_cell({ r: nr, c: nc })];
    }
    /* A value longer than Excel's 32 767-char cell limit is continued on the
     * next rows of the same column (writeRow reports how many it used): the
     * following records are therefore SHIFTED, never overwritten or lost. */
    var continuationRows = 0;
    rows.forEach(function (row, idx) {
      continuationRows += SEC.Exporter.writeRow(
        X,
        ws,
        headerIdx + 1 + idx + continuationRows,
        row,
        SEC.OFFICIAL_HEADERS,
      );
    });
    ws["!ref"] = X.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: headerIdx + Math.max(rows.length, 1) + continuationRows, c: 10 },
    });
    /* Template formatting is PRESERVED: the template's own column widths are
     * kept as they are. Recomputing them from the data destroyed the official
     * template's layout for no reason. */
    if (!ws["!cols"] && SEC.Exporter.columnWidths)
      ws["!cols"] = SEC.Exporter.columnWidths(SEC.OFFICIAL_HEADERS, rows);
    return { sheetName: sh.name, headerIdx: headerIdx };
  }

  SEC.XlsxModel = {
    conditionalInflate: conditionalInflate,
    extractModel: extractModel,
    reviveCell: reviveCell,
    applyRows: applyRows,
  };
})(typeof self !== "undefined" ? self : globalThis);

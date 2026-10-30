/* ============================================================
 * Smart Excel - parsing Web Worker
 * ------------------------------------------------------------
 * SheetJS (and the CSV parser) run OFF the interface thread, so
 * large workbooks never freeze the page. The worker does:
 *   parse  : ArrayBuffer -> [{ name, rows, isEmpty }]
 *   read   : conditionally decompress a .xlsx and read raw cell
 *            values (used by the template-preserving export)
 *   write  : rebuild a .xlsx from raw cells (style kept)
 * Works both as a classic worker (importScripts) and as a module
 * (import 'xlsx.mjs'), so it also runs inside Node tests.
 * ============================================================ */
/* eslint-disable no-restricted-globals */
(function () {
  "use strict";
  var inWorker = typeof importScripts === "function";
  var XLSXLib = null;
  var loaded = null;

  /* Model/rows dispatcher shared with the main thread (single source of
   * truth: js/xlsx-model.js). */
  var model = null;

  function isXlsxLib(lib) {
    return !!(
      lib &&
      typeof lib.read === "function" &&
      lib.utils &&
      lib.utils.sheet_to_json
    );
  }

  function loadLibs() {
    if (loaded) return loaded;
    if (inWorker) {
      if (!isXlsxLib(XLSXLib)) {
        importScripts("../vendor/xlsx.full.min.js");
        XLSXLib = self.XLSX;
      }
      importScripts(
        "../js/encoding.js",
        "../js/csv-parser.js",
        "../js/xlsx-reader.js",
        "../js/xlsx-model.js",
        "../js/exporter.js",
      );
      model = self.SEC.XlsxModel;
      loaded = Promise.resolve(self.SEC);
      return loaded;
    }
    /* Node / jsdom test environment */
    loaded = Promise.all([
      import("xlsx/xlsx.mjs"),
      import("../js/encoding.js"),
      import("../js/csv-parser.js"),
      import("../js/xlsx-reader.js"),
      import("../js/xlsx-model.js"),
      import("../js/exporter.js"),
    ]).then(function (mods) {
      XLSXLib = mods[0];
      model = globalThis.SEC.XlsxModel;
      return globalThis.SEC;
    });
    return loaded;
  }

  function post(msg, transfer) {
    self.postMessage(msg, transfer || []);
  }

  function withProgress(id) {
    return function (pct, stage) {
      post({ type: "progress", id: id, pct: pct, stage: stage });
    };
  }

  var handlers = {
    parse: function (payload) {
      var progress = withProgress(payload._id);
      return loadLibs().then(function (SEC) {
        SEC.XLSXReader.setLib(XLSXLib);
        var ext = String(payload.ext || "").toLowerCase();
        if (ext === "csv" || ext === "txt") {
          progress(35, "Lecture du contenu");
          var parsed = SEC.CSV.parse(payload.buffer);
          return {
            sheets: [
              {
                name: "(CSV)",
                rows: parsed.rows,
                isEmpty: parsed.rows.length === 0,
              },
            ],
            encoding: parsed.encoding,
            delimiter: parsed.delimiter,
          };
        }
        progress(40, "Lecture du classeur");
        var sheets = SEC.XLSXReader.parseFromLib(XLSXLib, payload.buffer);
        progress(70, "Mise en forme des feuilles");
        return { sheets: sheets, encoding: "", delimiter: "" };
      });
    },
    read: function (payload) {
      return loadLibs().then(function () {
        var bytes = model.conditionalInflate(XLSXLib, payload.buffer);
        var wb = XLSXLib.read(bytes, {
          type: "array",
          cellStyles: true,
          cellFormula: false,
          bookVBA: false,
        });
        return model.extractModel(XLSXLib, wb);
      });
    },
    write: function (payload) {
      return loadLibs().then(function () {
        var wb = XLSXLib.read(
          model.conditionalInflate(XLSXLib, payload.buffer),
          {
            type: "array",
            cellStyles: true,
            cellFormula: false,
            bookVBA: false,
          },
        );
        model.applyRows(XLSXLib, wb, payload.model, payload.rows);
        var out = XLSXLib.write(wb, {
          type: "array",
          bookType: "xlsx",
          cellStyles: true,
          compression: true,
        });
        return { buffer: out };
      });
    },
    /* Build the standard 11-column workbook + report sheets OFF the interface
     * thread: the Export button used to freeze the page on big batches. */
    export: function (payload, id) {
      var progress = withProgress(id);
      return loadLibs().then(function (SEC) {
        SEC.XLSXReader.setLib(XLSXLib);
        progress(20, "Ecriture du classeur Excel");
        var result = SEC.Exporter.generateConsolidated(
          payload.headers,
          payload.rows,
          payload.report,
          payload.expectedCount,
          payload.fileStats,
          payload.lineDetails,
          payload.continuations,
        );
        progress(80, "Verification du classeur");
        return {
          buffer: result.buffer,
          validation: result.validation,
          sheets: (result.sheetNames || []).slice(),
        };
      });
    },
  };

  self.onmessage = function (e) {
    var msg = e.data || {};
    var handler = handlers[msg.type];
    if (!handler) {
      post({
        type: "error",
        id: msg.id,
        error: "Type de message inconnu : " + msg.type,
      });
      return;
    }
    Promise.resolve()
      .then(function () {
        return handler(msg.payload || {}, msg.id);
      })
      .then(function (result) {
        var transfer = [];
        if (result && result.buffer && result.buffer.byteLength)
          transfer.push(result.buffer);
        post({ type: "done", id: msg.id, result: result }, transfer);
      })
      .catch(function (err) {
        post({
          type: "error",
          id: msg.id,
          error: (err && err.message) || String(err),
        });
      });
  };
})();

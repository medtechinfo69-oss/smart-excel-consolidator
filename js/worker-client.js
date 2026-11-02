/* ============================================================
 * Smart Excel - parsing / export engine (main thread)
 * ------------------------------------------------------------
 * Heavy parsing (SheetJS) and the template export run directly
 * on the main thread: no Web Worker, no local server — the
 * application stays fully self-contained everywhere (double-
 * click, USB key, static hosting).
 * ============================================================ */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});

  var state = {
    listeners: [],
    reason: "",
  };

  function notify(kind, payload) {
    state.listeners.forEach(function (fn) {
      try {
        fn(kind, payload);
      } catch (e) {
        /* never break the app */
      }
    });
  }

  /* The four operations of the engine, run inline. */
  function runInline(type, payload, onProgress) {
    /* One macrotask first: lets the browser PAINT the busy overlay /
     * progress BEFORE the synchronous SheetJS work blocks the thread —
     * otherwise the overlay was set and consumed in the same task and the
     * user saw a dead page for the whole parse. */
    return new Promise(function (resolve) { setTimeout(resolve, 0); }).then(function () {
      var X = SEC.XLSXReader.getXLSX();
      var model = SEC.XlsxModel;
      if (type === "parse") {
        if (onProgress) onProgress(40, "Lecture du fichier");
        var ext = String(payload.ext || "").toLowerCase();
        if (ext === "csv" || ext === "txt") {
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
        var sheets = SEC.XLSXReader.parseFromLib(X, payload.buffer, ext);
        if (onProgress) onProgress(70, "Mise en forme des feuilles");
        return { sheets: sheets, encoding: "", delimiter: "" };
      }
      if (type === "read") {
        var wb = X.read(model.conditionalInflate(X, payload.buffer), {
          type: "array",
          cellStyles: true,
          cellFormula: false,
          bookVBA: false,
        });
        return model.extractModel(X, wb);
      }
      if (type === "export") {
        if (onProgress) onProgress(20, "Ecriture du classeur Excel");
        var result = SEC.Exporter.generateConsolidated(
          payload.headers,
          payload.rows,
          payload.report,
          payload.expectedCount,
          payload.fileStats,
          payload.lineDetails,
          payload.continuations,
        );
        return {
          buffer: result.buffer,
          validation: result.validation,
          sheets: (result.sheetNames || []).slice(),
        };
      }
      if (type === "write") {
        var out = X.read(model.conditionalInflate(X, payload.buffer), {
          type: "array",
          cellStyles: true,
          cellFormula: false,
          bookVBA: false,
        });
        model.applyRows(X, out, payload.model, payload.rows);
        return {
          buffer: X.write(out, {
            type: "array",
            bookType: "xlsx",
            cellStyles: true,
            compression: true,
          }),
        };
      }
      return Promise.reject(new Error("Type inconnu : " + type));
    });
  }

  SEC.WorkerClient = {
    parse: function (buffer, ext, onProgress) {
      return runInline("parse", { buffer: buffer, ext: ext }, onProgress);
    },
    readModel: function (buffer) {
      return runInline("read", { buffer: buffer });
    },
    exportXlsx: function (payload, onProgress) {
      return runInline("export", payload, onProgress);
    },
    writeModel: function (templateBuffer, model, rows, onProgress) {
      return runInline(
        "write",
        { buffer: templateBuffer, model: model, rows: rows },
        onProgress,
      );
    },
    /* Nothing to pre-warm: everything runs on this thread. */
    prewarm: function () {},
    mode: function () {
      return "main";
    },
    reason: function () {
      return state.reason;
    },
    on: function (fn) {
      state.listeners.push(fn);
    },
    _inline: runInline,
  };
  notify("ready", { mode: "main" });
})(typeof self !== "undefined" ? self : globalThis);

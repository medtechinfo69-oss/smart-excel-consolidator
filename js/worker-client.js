/* ============================================================
 * Smart Excel - worker client with inline fallback
 * ------------------------------------------------------------
 * Heavy parsing (SheetJS) and the template export run in a Web
 * Worker so the interface never freezes. When no worker can be
 * created (file:// restrictions, jsdom tests, old browser), the
 * very same handlers run on the main thread before the next paint:
 * the application stays fully functional everywhere.
 * ============================================================ */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});

  var state = {
    worker: null,
    mode: "idle", // 'worker' | 'main'
    seq: 0,
    pending: new Map(),
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



  /* Same three operations as workers/parse.worker.js, main-thread version. */
  function runInline(type, payload, onProgress) {
    return Promise.resolve().then(function () {
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

  function workerUrl() {
    try {
      return new URL("../workers/parse.worker.js", document.baseURI).href;
    } catch (e) {
      return "../workers/parse.worker.js";
    }
  }

  function spawn() {
    if (state.worker || state.mode === "main" || state.mode === "starting")
      return state.worker;
    if (typeof Worker !== "function") {
      state.mode = "main";
      state.reason = "Web Worker indisponible";
      return null;
    }
    state.mode = "starting";
    try {
      state.worker = new Worker(workerUrl());
    } catch (err) {
      state.worker = null;
      state.mode = "main";
      state.reason =
        "Web Worker bloqué : " + ((err && err.message) || "inconnu");
      return null;
    }
    state.worker.onmessage = function (e) {
      var msg = e.data || {};
      if (msg.type === "progress") {
        var entry = state.pending.get(msg.id);
        if (entry && entry.onProgress) entry.onProgress(msg.pct, msg.stage);
        return;
      }
      var job = state.pending.get(msg.id);
      if (!job) return;
      state.pending.delete(msg.id);
      if (msg.type === "error") job.reject(new Error(msg.error));
      else setTimeout(function () { job.resolve(msg.result); }, 0);
    };
    state.worker.onerror = function (e) {
      var err = new Error((e && e.message) || "Erreur du Web Worker");
      state.pending.forEach(function (job) {
        job.reject(err);
      });
      state.pending.clear();
      /* Any import failure (e.g. file:// security) -> transparent fallback. */
      try {
        state.worker.terminate();
      } catch (x) {
        /* ignore */
      }
      state.worker = null;
      state.mode = "main";
      state.reason = err.message;
      notify("fallback", { reason: err.message });
    };
    state.mode = "worker";
    return state.worker;
  }

  /* Resolves with the handler result. Never rejects because of the worker
   * itself: a worker failure retries the same job on the main thread. */
  function run(type, payload, onProgress) {
    var worker = spawn();
    if (!worker) return runInline(type, payload, onProgress);
    var id = ++state.seq;
    return new Promise(function (resolve, reject) {
      state.pending.set(id, {
        resolve: resolve,
        reject: reject,
        onProgress: onProgress,
      });
      /* Move the CSV/XLSX bytes to the worker (no copy): reading stays fast
       * on big files and the main thread keeps no second copy. The File is
       * re-read if a retry needs the bytes again. */
      var transfer = [];
      if (
        payload &&
        payload.buffer &&
        payload.buffer.byteLength &&
        (type === "write" || type === "parse") &&
        payload.buffer instanceof ArrayBuffer
      )
        transfer.push(payload.buffer);
      try {
        worker.postMessage(
          Object.assign({ type: type, id: id }, { payload: payload }),
          transfer,
        );
      } catch (err) {
        state.pending.delete(id);
        runInline(type, payload, onProgress).then(resolve, reject);
      }
    }).catch(function (err) {
      if (state.mode === "main") throw err;
      /* Worker crashed on this file (e.g. unsupported encoding): retry locally. */
      return runInline(type, payload, onProgress);
    });
  }

  function mode() {
    return state.mode || "undetermined";
  }

  SEC.WorkerClient = {
    parse: function (buffer, ext, onProgress) {
      return run("parse", { buffer: buffer, ext: ext }, onProgress);
    },
    readModel: function (buffer) {
      return run("read", { buffer: buffer });
    },
    /* Build the final workbook. Runs in the worker when available: a big
     * export never blocks the page (no freeze, no stuck button). */
    exportXlsx: function (payload, onProgress) {
      return run("export", payload, onProgress);
    },
    writeModel: function (templateBuffer, model, rows, onProgress) {
      return run(
        "write",
        { buffer: templateBuffer, model: model, rows: rows },
        onProgress,
      );
    },
    prewarm: function () {
      spawn();
    },
    mode: mode,
    reason: function () {
      return state.reason;
    },
    on: function (fn) {
      state.listeners.push(fn);
    },
    _inline: runInline,
  };
})(typeof self !== "undefined" ? self : globalThis);

/* ============================================================
 * Smart Excel - Google Gemini client (browser side)
 * ------------------------------------------------------------
 * The assistant is Google Gemini, and the API key NEVER reaches the
 * browser: every call goes through the local server (/api/ai/*),
 * which reads GEMINI_API_KEY from .env.
 *
 * Guarantees:
 *  - FULL-SHEET mode (default when the assistant is ticked): Gemini receives
 *    the text of EVERY row of EVERY sheet and places each piece of
 *    information in its official column, splits Nom / Prénom and sends what
 *    it cannot identify to Commentaire. Batches of 300 rows keep the flow
 *    memory-bounded; the same rows are never sent twice.
 *  - Column mode (headers + a few samples of the ambiguous columns) is kept
 *    for the connection test and as a column-level check.
 *  - The AI is never silently skipped: a failure comes back as
 *    { __error, code, message } so the caller names it, and the local engine
 *    finishes processing and export. "Enabled but not usable yet" (no key,
 *    server down) stops the batch with an explicit reason instead of quietly
 *    running 100 % locally.
 * ============================================================ */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});

  var DEFAULT_PROVIDER = "gemini";
  /* The two providers the assistant can use. The choice is made in the
   * interface and travels with every request, so the provider that is
   * SELECTED is always the one that really processes the data. */
  var PROVIDER_LABELS = { gemini: "Google Gemini", groq: "Groq (Llama)" };
  var MAX_SAMPLES = 6;
  var SAMPLE_LEN = 60;
  var MAX_COLUMNS = 60; // ambiguous columns per request (cost control)
  /* Hard deadline for one mapping request. Kept under the processing watchdog
   * so the local engine always regains control (never hangs a batch). */
  var MAP_TIMEOUT_MS = 18000;
  var STATUS_TIMEOUT_MS = 3000;
  var TEST_TIMEOUT_MS = 30000; // a real generation (cold model) can be slow

  var state = {
    enabled: false,
    /* "IA obligatoire" : when the provider cannot classify a line, the line is
     * NOT quietly placed by the local engine — it is sent to « À valider ».
     * The user explicitly asked for this; the local engine stays the fallback
     * only when the assistant is switched off. */
    strict: false,
    provider: DEFAULT_PROVIDER,
    model: "auto",
    available: {}, // state.provider -> true when the backend has a key
    modelsByProvider: {}, // state.provider -> [model ids] (from /api/ai/status)
    checked: false,
    reachable: null, // null unknown / true / false
    stopped: false, // no more calls this session (quota/auth/offline)
    lastError: "",
    lastModel: "",
    lastModelRequested: "auto",
    requests: 0,
    mappings: 0,
    failures: 0,
    skip: {}, // headerNorm -> true (quota: never retried this session)
    /* Full-sheet mode counters (shown in the report and the status line). */
    sheetRows: 0, // rows actually sent to Gemini
    sheetPlaced: 0, // rows Gemini classified
    sheetMissing: 0, // rows Gemini could not answer (kept by the local engine)
    sheetCalls: 0, // HTTP requests made in full-sheet mode
    lastSheetError: "",
  };
  var cache = new Map(); // headerNorm -> target key
  var cacheParts = {}; // headerNorm -> { nom, prenom } (Gemini name split)
  /* Test hook (Node suite only): forces LIBSIM in the sandboxed context. */
  var SIM = (function () {
    try {
      return LIBSIM || null;
    } catch (e) {
      return null;
    }
  })();

  function isSimulating() {
    return !!(SIM && SIM.aiState);
  }
  function testState() {
    return (SIM && SIM.aiState) || {};
  }

  /* Is the selected provider usable? (backend reachable + key present in .env) */
  function backendAvailable() {
    if (isSimulating()) {
      return (
        !!testState().available[state.provider] && testState().reachable !== false
      );
    }
    if (!state.checked) return true; // optimistic until the server answers
    return !!state.available[state.provider];
  }

  function enabled() {
    if (!state.enabled || state.stopped) return false;
    return backendAvailable();
  }

  /* Only "enabled" and "model" are configurable: there is one state.provider. */
  function configure(opts) {
    opts = opts || {};
    if (opts.enabled !== undefined) state.enabled = !!opts.enabled;
    if (opts.model !== undefined) state.model = opts.model || "auto";
    if (opts.strict !== undefined) state.strict = !!opts.strict;
    if (opts.provider !== undefined) {
      var p = String(opts.provider || "gemini").toLowerCase();
      if (p === "gemini" || p === "groq") state.provider = p;
    }
  }

  /* OFFICIAL status shown in the interface: enabled = ticked AND no earlier
   * failure. `usable` uses the same rule so the two never contradict each
   * other (the tick used to say "active" while `enabled()` said false). */
  function usable() {
    return state.enabled && !state.stopped && backendAvailable();
  }

  function status() {
    return {
      enabled: usable(),
      ticked: state.enabled,
      strict: state.strict,
      usable: usable(),
      provider: state.provider,
      label: PROVIDER_LABELS[state.provider] || PROVIDER_LABELS.gemini,
      model: state.model,
      /* The model that really answered the last request (auto mode). */
      usedModel: state.lastModel,
      usedModelRequested: state.lastModelRequested || "auto",
      models: state.modelsByProvider[state.provider] || [],
      available: isSimulating() ? testState().available : state.available,
      checked: isSimulating() ? true : state.checked,
      reachable: isSimulating() ? testState().reachable : state.reachable,
      stopped: state.stopped,
      lastError: state.lastError,
      lastModel: state.lastModel,
      requests: state.requests,
      mappings: state.mappings,
      failures: state.failures,
      sheetRows: state.sheetRows,
      sheetPlaced: state.sheetPlaced,
      sheetMissing: state.sheetMissing,
      sheetCalls: state.sheetCalls,
      lastSheetError: state.lastSheetError,
      cached: cache.size,
      /* A page opened from disk can never reach the local server. */
      filePage: isFilePage(),
    };
  }

  function withTimeout(ms) {
    var controller =
      typeof AbortController === "function" ? new AbortController() : null;
    var timer = controller
      ? setTimeout(function () {
          controller.abort();
        }, ms)
      : null;
    return {
      signal: controller ? controller.signal : undefined,
      done: function () {
        if (timer) clearTimeout(timer);
      },
    };
  }

  /* The page was opened directly from disk: a browser NEVER lets a file://
   * page call the local server (CORS), whatever the assistant is. Saying so
   * clearly beats a raw "Failed to fetch". */
  function nodeServerError() {
    return new Error(
      "page ouverte directement (file://) : ouvrez l'application avec « Ouvrir l'application.bat » (http://127.0.0.1:3000) pour activer Gemini",
    );
  }
  /* Base URL of the local server. On http(s) the page IS served by it, so a
   * RELATIVE url reaches the right instance whatever the host/port used
   * (localhost, 127.0.0.1, another port). A page opened from disk must target
   * the classic port. */
  function apiUrl(url) {
    if (!isFilePage()) return url;
    return "http://127.0.0.1:3000" + url;
  }

  var filePageOverride = null; // test hook only
  function isFilePage() {
    if (filePageOverride !== null) return filePageOverride;
    try {
      return !!(root.location && root.location.protocol === "file:");
    } catch (e) {
      return false;
    }
  }

  function fetchJson(url, options, timeoutMs, allowFromFile) {
    if (typeof fetch !== "function")
      return Promise.reject(new Error("fetch indisponible"));
    if (isFilePage() && !isSimulating() && !allowFromFile)
      return Promise.reject(nodeServerError());
    var guard = withTimeout(timeoutMs || MAP_TIMEOUT_MS);
    var opts = Object.assign(
      { headers: { "Content-Type": "application/json" } },
      options || {},
    );
    if (guard.signal) opts.signal = guard.signal;
    return fetch(apiUrl(url), opts)
      .then(function (res) {
        var ctype =
          res.headers && res.headers.get
            ? res.headers.get("content-type") || ""
            : "";
        if (res.status === 404 && ctype.indexOf("json") === -1) {
          /* The port answers but not with our API: an older copy of the app is
           * still running there. Nothing to do with the key. */
          throw new Error(
            "une ancienne version du serveur local tourne encore — fermez la fenetre noire puis relancez « Ouvrir l'application.bat »",
          );
        }
        if (ctype.indexOf("html") !== -1) {
          throw new Error("Serveur local injoignable (lancez « Ouvrir l'application.bat »)");
        }

        return res.text().then(function (raw) {
          var json = null;
          try {
            json = JSON.parse(raw);
          } catch (e) {
            /* server down / file:// */
          }
          if (!json)
            throw new Error("Serveur local injoignable (lancez « npm start »)");
          return json;
        });
      })
      .catch(function (err) {
        if (err && err.name === "AbortError")
          throw new Error("délai dépassé (assistant IA)");
        /* "Failed to fetch" means the local server is not running (or the page
         * is not served by it): explain it instead of the browser's wording. */
        if (err && /failed to fetch|load failed|networkerror/i.test(err.message || "")) {
          throw new Error(
            isFilePage()
              ? nodeServerError().message
              : "serveur local introuvable — lancez « Ouvrir l'application.bat » (ou npm start)",
          );
        }
        throw err;
      })
      .then(
        function (json) {
          guard.done();
          return json;
        },
        function (err) {
          guard.done();
          throw err;
        },
      );
  }

  /* Ask the backend whether the Gemini key is present (and its models). */
  function refreshStatus() {
    if (isSimulating() && testState().reachable === false)
      return Promise.resolve(status());
    return fetchJson("/api/ai/status", { method: "GET" }, STATUS_TIMEOUT_MS)
      .then(function (json) {
        /* Only the Gemini entry matters: read it by its exact key (iterating
         * the object left `state.available` empty, so the assistant disabled
         * itself with no error at all). */
        var info = (json.providers || {})[state.provider];
        state.available = {};
        if (info) info = Object.assign({}, info);
        state.available[state.provider] = !!(info && info.configured);
        if (info && Array.isArray(info.models) && info.models.length)
          state.modelsByProvider[state.provider] = info.models;
        else if (!state.modelsByProvider[state.provider])
          state.modelsByProvider[state.provider] = [];
        state.checked = true;
        state.reachable = true;
        return status();
      })
      .catch(function (err) {
        state.checked = true;
        state.reachable = false;
        state.lastError = err.message;
        return status();
      });
  }

  /* Real connection test: lists the free-tier models and generates a tiny
   * mapping. Reports the exact model used, or the exact error. */
  function test() {
    return refreshStatus()
      .then(function (st) {
        if (!st.reachable)
          return { ok: false, code: "network", error: st.lastError };
        if (!st.available[state.provider]) {
          return {
            ok: false,
            code: "no-key",
            error: "Aucune clé GEMINI_API_KEY dans .env",
          };
        }
        return fetchJson(
          "/api/ai/test",
          {
            method: "POST",
            body: JSON.stringify({ provider: state.provider, model: state.model }),
          },
          TEST_TIMEOUT_MS,
        ).then(function (res) {
          if (res.ok) {
            state.lastModel = res.model || "";
            if (Array.isArray(res.models) && res.models.length)
              state.modelsByProvider[state.provider] = res.models;
            return { ok: true, model: res.model, models: res.models || [] };
          }
          state.lastError = res.error || res.code || "erreur IA";
          return {
            ok: false,
            code: res.code,
            error: res.error,
            attempts: res.attempts,
          };
        });
      })
      .catch(function (err) {
        return { ok: false, code: "network", error: err.message };
      });
  }

  /* Paste a key: the server validates it with Google, stores it and writes it
   * to .env. Works from a page opened from disk too (its "null" origin can
   * only POST a key to this machine, never read another site's data). */
  function setKey(rawKey) {
    var key = String(rawKey || "").trim();
    if (!key) return Promise.resolve({ ok: false, code: "bad-request", error: "Clé vide" });
    if (typeof fetch !== "function")
      return Promise.resolve({ ok: false, code: "network", error: "fetch indisponible" });
    /* This call is deliberately allowed from file:// too (see filePage below):
     * fetchJson refuses other calls there, so it is bypassed on purpose. */
    return fetchJson(
      "/api/ai/key",
      { method: "POST", body: JSON.stringify({ key: key, provider: state.provider }) },
      TEST_TIMEOUT_MS,
    )
      .then(function (res) {
        if (!res.ok) {
          state.lastError = res.error || res.code || "clé refusée";
          return { ok: false, code: res.code, error: res.error };
        }
        /* Proven: the assistant restarts cleanly with this key. */
        state.stopped = false;
        state.skip = {};
        state.lastError = "";
        state.available[state.provider] = true;
        state.checked = true;
        state.reachable = true;
        state.lastModel = res.model || state.lastModel;
        if (Array.isArray(res.models) && res.models.length)
          state.modelsByProvider[state.provider] = res.models;
        cache.clear(); // a new key may answer differently: nothing stale
        return { ok: true, model: res.model, models: res.models || [], masked: res.masked };
      })
      .catch(function (err) {
        return { ok: false, code: "network", error: err.message };
      });
  }

  function models() {
    return fetchJson(
      "/api/ai/models?provider=" + state.provider,
      { method: "GET" },
      MAP_TIMEOUT_MS,
    ).catch(function (err) {
      return { ok: false, code: "network", error: err.message, models: [] };
    });
  }

  function columnPayload(col) {
    var seen = {};
    var samples = [];
    var list = (col.profile && col.profile.samples) || [];
    list.forEach(function (v) {
      var s = String(v == null ? "" : v)
        .trim()
        .slice(0, SAMPLE_LEN);
      if (!s || seen[s]) return;
      seen[s] = true;
      if (samples.length < MAX_SAMPLES) samples.push(s);
    });
    return { header: String(col.header || ""), samples: samples };
  }

  /* Value types that are certain on their own (the content IS the format:
   * nobody writes an e-mail address in a city column). A column made of those
   * at 90%+ is not ambiguous and is never sent online. */
  var DECISIVE_TYPES = ["email", "phone", "postal", "amount", "mutuelle"];

  /* A column is "ambiguous" when the local engine could not decide it:
   *  - the header text itself was recognised (dictionary/token match) and the
   *    values agree → certain, never sent;
   *  - the content IS the format (real e-mails / phones / postal codes /
   *    amounts / insurer names at 90%+) → certain, never sent;
   *  - anything else ("Titre client", "Ref dossier", a weak guess…) → sent.
   * The AI never sees a whole sheet: only these columns, header + samples. */
  function isAmbiguous(col) {
    if (!col || col.blocked || col.fullName || col.excluded || col.learned)
      return false;
    /* Columns already decided (merge) are final. */
    if (col.mergeInto) return false;
    var v = col.validation && col.validation.status;
    var profile = col.profile || {};
    /* The local engine could not choose ANY destination: this column was not
     * even PROFILED (no values read, e.g. the light pre-analysis pass): that
     * is a contract mistake, never a reason to upload a column blind. */
    if (!col.best)
      return !!(profile.ratios && Object.keys(profile.ratios).length);
    /* A real header match that is CONSISTENT with the values is certain
     * ("Titre client" full of names scores 68 and is not: nothing proves it).
     * The 85+ score is already the proof — the old margin>=20 test compared
     * two proposals derived from the SAME evidence and made almost every
     * column "certain" (0 destination), so Gemini was never consulted. */
    var margin = typeof col.margin === "number" ? col.margin : 100;
    if (col.headerMatched === true && col.score >= 85 && v !== "warn") return false;
    /* A value-proven destination with a clear lead is decided locally.
     * Otherwise the column keeps its data but needs a real decision (Gemini). */
    if (col.score >= 72 && margin >= 25 && v !== "warn") return false;
    var r = (col.profile && col.profile.ratios) || {};
    var decisive = DECISIVE_TYPES.some(function (k) {
      return (r[k] || 0) >= 0.9;
    });
    /* A whole column of official commune names is a Ville column even when
     * the header is unusual ("Commune de résidence", "Implantation"). */
    if (
      !decisive &&
      SEC.Agent &&
      SEC.Agent.looksLikeCommuneColumn &&
      SEC.Agent.looksLikeCommuneColumn(col.profile) &&
      (!col.best || col.best === "ville")
    ) {
      return false;
    }
    /* The content IS the format: certain whatever the header says. */
    if (decisive && v !== "warn" && col.best && col.score >= 60) return false;
    return true;
  }

  /* Consult Gemini for the ambiguous columns of one worksheet.
   * Resolves { headerNorm -> target } (possibly empty) — never rejects. */
  function suggestSheet(fileName, sheetName, columns) {
    if (!enabled()) return Promise.resolve(null);
    var pending = [];
    var seen = {};
    var hints = {};
    columns.forEach(function (col) {
      if (!isAmbiguous(col)) return;
      var k = col.headerNorm;
      if (!k || seen[k]) return;
      seen[k] = true;
      /* Already answered earlier in this session (any file): reuse it.
       * A FIXED model chosen by the user must answer with THAT model, so only
       * automatic mode reuses answers produced by another model. */
      var cacheKey = (state.model || "auto") + "|" + k;
      if (cache.has(cacheKey)) {
        hints[k] = cache.get(cacheKey);
        return;
      }
      /* Quota exhausted earlier for this header: never retried. */
      if (state.skip[k]) return;
      pending.push(col);
    });
    if (!pending.length) {
      /* An empty answer still counts as "asked": the report stays honest. */
      return Promise.resolve(hints);
    }
    /* Nothing online can answer this sheet: tell the caller, so the local
     * engine keeps its best evidence instead of demoting it. */
    var unavailable = function () {
      hints.__unavailable = true;
      return hints;
    };
    var payload = pending.slice(0, MAX_COLUMNS).map(columnPayload);
    state.requests++;
    return fetchJson(
      "/api/ai/map",
      {
        method: "POST",
        body: JSON.stringify({
          provider: state.provider,
          model: state.model,
          columns: payload,
        }),
      },
      MAP_TIMEOUT_MS,
    )
      .then(function (res) {
        if (!res.ok) {
          state.failures++;
          state.lastError = res.error || res.code || "erreur IA";
          /* Quota exhausted, key refused, backend gone: stop calling for the
           * rest of the session — the local engine finishes the job. */
          if (
            [
              "quota",
              "auth",
              "no-key",
              "network",
              "bad-request",
              "empty",
            ].indexOf(res.code) !== -1
          ) {
            state.stopped = true;
            payload.forEach(function (p) {
              var k = SEC.Normalizers
                ? SEC.Normalizers.normalizeHeader(p.header)
                : String(p.header || "").toLowerCase();
              if (k) state.skip[k] = true;
            });
            if (res.code === "network") state.reachable = false;
            return unavailable();
          }
          return {};
        }
        state.lastModel = res.model || state.lastModel;
        state.lastModelRequested = state.model || "auto";
        var out = {};
        Object.keys(res.mappings || {}).forEach(function (header) {
          var k = SEC.Normalizers
            ? SEC.Normalizers.normalizeHeader(header)
            : String(header || "").toLowerCase();
          var target = res.mappings[header];
          if (!k || !target) return;
          cache.set((state.model || "auto") + "|" + k, target);
          out[k] = target;
          state.mappings++;
        });
        /* Gemini also tells WHERE the surname and the first name are ("name"
         * columns). Kept apart from the destination hints. */
        Object.keys(res.fullName || {}).forEach(function (header) {
          var k = SEC.Normalizers
            ? SEC.Normalizers.normalizeHeader(header)
            : String(header || "").toLowerCase();
          var v = res.fullName[header];
          if (!k || !v || (!v.nom && !v.prenom)) return;
          cacheParts[k] = { nom: String(v.nom || ""), prenom: String(v.prenom || "") };
        });
        /* Every asked header was answered: merge with the cached ones. */
        Object.keys(hints).forEach(function (k) {
          if (k !== "__unavailable" && !out[k]) out[k] = hints[k];
        });
        return out;
      })
      .catch(function (err) {
        state.failures++;
        state.lastError = err.message;
        state.reachable = false;
        state.stopped = true;
        return unavailable();
      });
  }

  /* ---------------- FULL-SHEET mode ----------------
   * Gemini reads the REAL text of every row and returns, per line, which
   * official column each piece belongs to (nom / prenom separated here, not
   * by the header).
   *   headers: array of source header texts (position = source column)
   *   rows:    array of source rows (arrays of raw cells, position matter)
   *   baseIndex: source row number of rows[0] (returned as-is)
   * Resolves { rows: [{ index, values }], placed, missing, failed }
   *           or { __error: true, code, message } — never rejects. */

  /* 180 rows per request: big enough to stay fast, small enough that the JSON
   * answer is never truncated (a truncated answer costs a retry). */
  /* 60 rows per request: answers in a few seconds and is never truncated. */
  var SHEET_ROWS_PER_CALL = 60;
  var SHEET_MAX_CELLS = 40; // wider sheets are cut: 11 official columns exist
  var SHEET_CELL_LEN = 300;
  var SHEET_TIMEOUT_MS = 120000;
  /* The free tier allows 15 requests/MINUTE. Two requests in flight already
   * finish faster than the limiter: ONE worker, and a minimum delay between
   * two requests, keeps the batch under the limit instead of triggering
   * 429 "retry in 4s" on every sheet. */
  /* Gemini's FREE tier allows 15 requests/minute: one request every 4.3 s.
   * Groq allows far more, so it is not throttled at all (only 2 in flight). */
  var SHEET_MIN_GAP_MS = 4300;
  var GROQ_MIN_GAP_MS = 0;
  var sheetNextSlot = 0;

  function emptySheetResult() {
    return { rows: [], placed: 0, missing: 0, failed: 0, calls: 0 };
  }

  function sleepMs(ms) {
    return new Promise(function (r) { setTimeout(r, ms); });
  }

  function sheetRowsPayload(rows, start, headers) {
    var out = [];
    for (var i = start; i < rows.length; i++) {
      var src = rows[i] || [];
      var cells = [];
      for (var c = 0; c < headers.length && c < SHEET_MAX_CELLS; c++) {
        var v = src[c];
        cells.push(v == null ? "" : String(v).slice(0, SHEET_CELL_LEN));
      }
      /* A fully empty line is not sent (it is skipped by the extraction). */
      var any = false;
      for (var k = 0; k < cells.length; k++) {
        if (cells[k] && cells[k].replace(/\s+/g, "")) { any = true; break; }
      }
      if (any) out.push({ index: i, cells: cells });
    }
    return out;
  }

  function mapSheet(fileName, sheetName, headers, rows, baseIndex, onProgress) {
    /* The assistant is ticked but cannot be used: report WHY, never pretend
     * the local engine is the assistant. */
    if (!usable()) {
      return Promise.resolve({
        __error: true,
        code: state.enabled ? (state.reachable === false ? "network" : "unavailable") : "disabled",
        message: !state.enabled
          ? "Assistant IA désactivé"
          : state.lastError ||
            (state.reachable === false
              ? "serveur local injoignable (ouvrez « Ouvrir l'application.bat »)"
              : "clé Gemini absente ou refusée"),
      });
    }
    var list = Array.isArray(rows) ? rows : [];
    var head = (Array.isArray(headers) ? headers : []).map(function (h) {
      return h == null ? "" : String(h);
    });
    if (!list.length || !head.length) return Promise.resolve(emptySheetResult());

    var pending = [];
    for (var start = 0; start < list.length; start += SHEET_ROWS_PER_CALL) {
      pending.push(sheetRowsPayload(list, start, head));
    }
    pending = pending.filter(function (b) { return b.length; });
    if (!pending.length) return Promise.resolve(emptySheetResult());

    var result = emptySheetResult();
    result.rows = [];
    var totalRowsToDo = 0;
    pending.forEach(function (b) { totalRowsToDo += b.length; });
    var doneRows = 0;
    function reportProgress() {
      doneRows = result.placed + result.failed;
      if (typeof onProgress === 'function') {
        try { onProgress(doneRows, totalRowsToDo); } catch (e) { /* UI must never break */ }
      }
    }
    var firstError = null;
    var stop = false;

    function send(batch) {
      if (stop) return Promise.resolve();
      var payload = {
        provider: state.provider,
        model: state.model,
        fileName: String(fileName || ""),
        sheet: String(sheetName || ""),
        headers: head,
        rows: batch.map(function (r) {
          return {
            i: (typeof baseIndex === "number" ? baseIndex : 0) + r.index,
            c: r.cells,
          };
        }),
      };
      state.requests++;
      state.sheetCalls++;
      state.sheetRows += payload.rows.length;
      result.calls++;
      /* Rate pacing: never start a request before its slot (shared across all
       * sheets and files of the run). */
      var gap = state.provider === "groq" ? GROQ_MIN_GAP_MS : SHEET_MIN_GAP_MS;
      var now = Date.now();
      var wait = gap ? Math.max(0, sheetNextSlot - now) : 0;
      if (gap) sheetNextSlot = Math.max(now, sheetNextSlot) + gap;
      var go = wait ? sleepMs(wait) : Promise.resolve();
      return go
        .then(function () {
          return fetchJson(
            "/api/ai/sheet",
            { method: "POST", body: JSON.stringify(payload) },
            SHEET_TIMEOUT_MS,
          );
        })
        .then(function (res) {
          if (!res || !res.ok) {
            var code = (res && res.code) || "network";
            var msg = (res && res.error) || "erreur IA";
            state.failures++;
            state.lastSheetError = msg;
            state.lastError = msg;
            if (!firstError) firstError = { code: code, message: msg };
            /* A PER-MINUTE rate limit is transient: the server already waited
             * for the delay Gemini asked for, so this batch is retried instead
             * of giving up on the sheet (it is what made the last file fall
             * back to the local engine after a "retry in 4.3s"). */
            if (code === "rate" && (result.rateRetries = (result.rateRetries || 0) + 1) <= 2) {
              state.failures--; // not a real failure yet
              return sleepMs(1500 * result.rateRetries).then(function () { return send(batch); });
            }
            /* NETWORK / timeout: retry the same batch a few times. A single
             * aborted request must never mark hundreds of lines as
             * "unclassified" while the next request would have worked. */
            if (["network", "empty", "truncated", "busy"].indexOf(code) !== -1 &&
                (result.netRetries = (result.netRetries || 0) + 1) <= 3) {
              state.failures--;
              state.reachable = true;
              return sleepMs(1200 * result.netRetries).then(function () { return send(batch); });
            }
            result.failed += payload.rows.length;
            reportProgress();
            /* Daily quota / key / server: every further call is pointless. */
            if (
              ["quota", "auth", "no-key", "network", "bad-request", "forbidden"].indexOf(code) !== -1
            ) {
              if (code !== "bad-request") state.stopped = true;
              if (code === "network") state.reachable = false;
            }
            return;
          }
          state.lastModel = res.model || state.lastModel;
          state.lastModelRequested = state.model || "auto";
          var answered = res.rows || {};
          var placed = 0;
          var unplaced = [];
          payload.rows.forEach(function (r) {
            var values = answered[r.i];
            if (values && typeof values === "object") {
              result.rows.push({ index: r.i, values: values });
              placed++;
            } else unplaced.push(r);
          });
          result.placed += placed;
          result.missing += payload.rows.length - placed;
          if (res.truncated) result.truncated = (result.truncated || 0) + res.truncated;
          reportProgress();
          /* Lines the model did not answer (truncated answer, partial reply)
           * are re-sent, in smaller batches, instead of being abandoned to
           * the local engine. */
          if (unplaced.length) {
            result.resent = (result.resent || 0) + unplaced.length;
            var half = Math.ceil(unplaced.length / 2);
            var parts = unplaced.length > 60
              ? [unplaced.slice(0, half), unplaced.slice(half)]
              : [unplaced];
            return parts.reduce(function (chain, part) {
              return chain.then(function () {
                if (!part.length) return null;
                if ((result.resentRounds = (result.resentRounds || 0)) > 400) return null;
                result.resentRounds++;
                result.missing -= part.length;
                return send(part);
              });
            }, Promise.resolve());
          }
        })
        .catch(function (err) {
          state.failures++;
          state.reachable = false;
          state.stopped = true;
          state.lastSheetError = err.message;
          state.lastError = err.message;
          if (!firstError) firstError = { code: "network", message: err.message };
          result.failed += payload.rows.length;
        });
    }

    /* Batches are sent a few at a time; the caller keeps the UI alive. */
    var next = 0;
    function worker() {
      if (next >= pending.length) return Promise.resolve();
      var batch = pending[next++];
      return send(batch).then(worker);
    }
    var workers = [];
    var maxWorkers = state.provider === "groq" ? 2 : 1;
    for (var w = 0; w < Math.min(maxWorkers, pending.length); w++)
      workers.push(worker());

    return Promise.all(workers).then(function () {
      if (!result.rows.length && firstError) {
        return {
          __error: true,
          code: firstError.code,
          message: firstError.message,
          failed: result.failed,
          calls: result.calls,
        };
      }
      /* Some rows unplaced: always reported, never silently replaced. */
      result.error = firstError;
      return result;
    });
  }

  SEC.AI = {
    PROVIDER_LABELS: PROVIDER_LABELS,
    configure: configure,
    mapSheet: mapSheet,
    enabled: enabled,
    status: status,
    refreshStatus: refreshStatus,
    test: test,
    setKey: setKey,
    models: models,
    providers: PROVIDER_LABELS,
    PROVIDERS_LABELS: PROVIDER_LABELS,
    suggestSheet: suggestSheet,
    isAmbiguous: isAmbiguous,
    /* Name split already computed for this column (any file, this session). */
    nameParts: function (headerNorm) {
      return cacheParts[headerNorm] || null;
    },
    lastError: function () {
      return state.lastError;
    },
    model: function () {
      return state.lastModel;
    },
    stop: function (reason) {
      state.stopped = true;
      state.lastError = reason || state.lastError;
    },
    reset: function () {
      state.stopped = false;
      state.lastError = "";
      state.skip = {};
    },
    /* A new batch starts from scratch: counters describe THIS run only. */
    startRun: function () {
      state.stopped = false;
      state.skip = {};
      state.lastSheetError = "";
      state.sheetRows = 0;
      state.sheetPlaced = 0;
      state.sheetMissing = 0;
      state.sheetCalls = 0;
    },
    _clearCache: function () {
      cache.clear();
    },
    /* Test hook: simulate a page opened directly from disk. */
    _setFilePageForTest: function (v) {
      filePageOverride = v === undefined ? null : !!v;
    },
  };
})(typeof self !== "undefined" ? self : globalThis);

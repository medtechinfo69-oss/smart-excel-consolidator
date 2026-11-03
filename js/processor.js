/* ============================================================
 * Smart Excel - Processing engine & session state
 * Sequential, cancellable, memory-conscious, retry-safe.
 * Customer data lives only in memory for the current session.
 * ============================================================ */
(function (root) {
  'use strict';
  var SEC = (root.SEC = root.SEC || {});
  var N = SEC.Normalizers;
  var S = SEC.STATUS;

  /* REAL event-loop yield. queueMicrotask keeps draining the microtask queue
   * inside the same task: the browser never paints and never handles input,
   * which is exactly the "page freezes during processing" symptom. A 0 ms
   * timeout (or setImmediate in Node) hands control back to the browser so
   * progress bars, the Cancel button and scrolling stay alive. */
  function nextTick() {
    return new Promise(function (res) {
      if (typeof setImmediate === 'function') setImmediate(res);
      else setTimeout(res, 0);
    });
  }

  /* Optional online assistant (OpenRouter), called directly from the page.
   * The AI client only exists when a key is stored in this browser and the
   * assistant is enabled: without it, everything stays local. */
  function aiAssistant() {
    return SEC.AI && typeof SEC.AI.suggestSheet === 'function' ? SEC.AI : null;
  }

  /* Apply the (validated) Gemini destination hints to the uncertain columns
   * of a worksheet BEFORE extraction. With Gemini's full-scan mode (default),
   * every column below 95 local confidence is corrected by the IA answer;
   * a hint NEVER overrides a decision the user already made for that column.
   * Blocked identifier columns and full-name columns stay untouched. */
  /* ---------------- Gemini FULL-SHEET answers (per row) ----------------
   * Gemini read every row and says, per source line, which official column
   * each piece of information belongs to. The value it gives back is only
   * accepted when it REBUILDS the exact source text of the line (word by
   * word): it can move, split or reformat a value, never invent or drop one.
   * The local engine keeps reading the row; the AI answer then places what
   * it classified and the rest of the information goes to « Commentaire ». */

  function rowWordSet(cells) {
    var set = {};
    (cells || []).forEach(function (c) {
      String(c == null ? '' : c)
        .replace(/\u00a0/g, ' ')
        .toLowerCase()
        .match(/[0-9a-z\u00e0-\u00f6\u00f8-\u00ff]+/g)
        .forEach(function (w) { set[w] = true; });
    });
    return set;
  }

  /* Where a value COULD go, used as the default suggestion of the
   * « À valider » table. It never places anything: it only proposes. */
  function suggestDestination(val, values) {
    if (N.looksLikeEmail(val)) return 'email';
    if (N.isConfirmedAmount && N.isConfirmedAmount(val)) return 'cotisation';
    if (N.looksLikeInsurer && N.isPureInsurer && N.isPureInsurer(val)) return 'mutuelle';
    if (N.looksLikePhone(val) && !N.isUnusablePhone(val)) return values && values.tel1 ? 'tel2' : 'tel1';
    if (N.looksLikePostal(val)) return 'cp';
    if (N.looksLikeAddress(val)) return N.looksLikeCity && N.looksLikeCity(val) ? 'ville' : 'adresse';
    if (N.looksLikePersonName(val)) return 'nom';
    return 'commentaire';
  }

  function aiValuesForRow(aiRow, srcRow) {
    var words = rowWordSet(srcRow);
    var out = {};
    var placedWords = {};
    var dropped = false;
    SEC.TARGET_KEYS.forEach(function (key) {
      var v = aiRow && aiRow[key];
      if (typeof v !== 'string') return;
      v = N.trimClean(v);
      if (!v) return;
      var tokens = v
        .replace(/\u00a0/g, ' ')
        .toLowerCase()
        .match(/[0-9a-z\u00e0-\u00f6\u00f8-\u00ff]+/g) || [];
      /* Every word must come from the source line. */
      for (var t = 0; t < tokens.length; t++) {
        if (!words[tokens[t]]) { dropped = true; return; }
      }
      out[key] = v;
      tokens.forEach(function (w) { placedWords[w] = true; });
    });
    if (!Object.keys(out).length) return null;
    /* A word of the source line that Gemini placed nowhere is not lost: it
     * goes to « Commentaire » (after everything the AI did classify). */
    var leftovers = [];
    (srcRow || []).forEach(function (cell) {
      var text = N.trimClean(cell);
      if (!text || /^\.+$/.test(text)) return;
      var tokens = text
        .replace(/\u00a0/g, ' ')
        .toLowerCase()
        .match(/[0-9a-z\u00e0-\u00f6\u00f8-\u00ff]+/g) || [];
      var missing = tokens.some(function (w) { return !placedWords[w]; });
      if (!missing) return;
      if (out.commentaire && out.commentaire.indexOf(text) !== -1) return;
      leftovers.push(text);
    });
    if (leftovers.length) {
      out.commentaire = (out.commentaire ? out.commentaire + ' | ' : '') + leftovers.join(' | ');
    }
    out.__dropped = dropped;
    return out;
  }

  function applyGeminiHints(columns, hints) {
    columns.forEach(function (col) {
      if (col.blocked || col.fullName || col.excluded || col.learned) return;
      var hint = hints[col.headerNorm];
      if (!hint || typeof hint !== 'string') return;
      /* "name" = a value holding a surname AND a first name: it is read into
       * « Nom » first, then split (locally or with the assistant's parts). */
      if (hint === 'name') hint = 'nom';
      /* The client only sends genuinely AMBIGUOUS columns (header + samples),
       * so its answer is authoritative for them — even when it confirms the
       * (weak) local guess: the column is then decided, not "to verify". */
      col.aiHint = hint;
      col.headerMatched = true;
      col.profileExcluded = false;
      /* The column now has a single leading destination: the review flag that
       * comes from a CONFLICT between two equal guesses no longer applies. */
      col.margin = 100;
      col.mergeInto = null;
      col.rewrittenAsComment = false;
      col.geminiHint = hint;
      var changed = !col.best || col.best !== hint;
      col.best = hint;
      col.aiModel = (SEC.AI && SEC.AI.model && SEC.AI.model()) || '';
      /* Once the AI answered, the column IS resolved: it must not be listed
       * as "unknown column" any more (only 92 keeps it inside the IA band). */
      col.score = Math.max(col.score, 96);
      var aiLabel = ((SEC.AI && SEC.AI.status().label) || 'IA') + (col.aiModel ? ' ' + col.aiModel : '');
      col.reasons.push(changed
        ? 'destination corrigée par l assistant IA (' + aiLabel + ') : « ' + SEC.TARGET_LABELS[hint] + '»'
        : 'destination confirmée par l assistant IA (' + aiLabel + ') : « ' + SEC.TARGET_LABELS[hint] + '»');
    });
  }

  /* One report line per session describing what the optional assistant did
   * (or why it stopped). Never blocks: purely informational. */
  /* A short, human sentence for an AI failure: Google sends a long JSON blob
   * ("Quota exceeded for metric: <api>/generate_content_free_tier_requests,
   * limit: 15 … Please retry in 4.3s") that only hides the actionable part. */
  function aiReasonText(scan) {
    var code = (scan && scan.code) || '';
    var msg = String((scan && scan.message) || '');
    if (code === 'rate') return 'limite de debit gratuite atteinte (nouvelle tentative automatique) — reessayez dans une minute';
    if (code === 'quota') return 'quota gratuit epuise pour aujourd hui';
    if (code === 'auth') return 'cle OpenRouter refusee (invalide ou sans acces)';
    if (code === 'no-key') return 'aucune cle OpenRouter configuree';
    if (code === 'network') return 'internet injoignable';
    if (code === 'disabled') return 'assistant IA desactive';
    /* Otherwise keep the first sentence, cleaned. */
    var first = msg.split(/\n|\*\s|\.\s{2,}/)[0].replace(/\s+/g, ' ').trim();
    return first ? first.slice(0, 160) : 'aucune reponse de l IA';
  }

  function aiFixHint(code) {
    if (code === 'rate') return 'Attendre une minute puis relancer : la limite gratuite est de 15 requetes/minute';
    if (code === 'quota') return 'Quota epuise : attendre demain ou activer la facturation OpenRouter';
    if (code === 'auth') return 'Coller une cle OpenRouter valide puis cliquer « Activer la cle »';
    if (code === 'no-key') return 'Coller une cle OpenRouter dans le panneau de l assistant';
    if (code === 'network') return 'Verifier la connexion internet puis reessayer';
    return 'Verifier la cle OpenRouter, le quota et la connexion internet';
  }

  function recordAiOutcome(session) {
    var AI = SEC.AI;
    if (!AI || typeof AI.status !== 'function') return;
    var st = AI.status();
    /* A failed call disables the assistant, so `enabled` is false exactly
     * when the report line matters most: the tick is the right test. */
    if (!st.enabled && !st.ticked) return;
    var requests = session.aiRequests || 0;
    var hints = session.aiHints || 0;
    var rows = session.aiRows || 0;
    var missing = session.aiMissingRows || 0;
    if (!requests && !hints && !rows && !missing && !st.stopped && !st.mappings) return;
    /* Written once for a summary, but a missing-lines line is added every time
     * a new batch leaves lines Gemini could not classify. */
    var key = 'ia|' + rows + '|' + missing + '|' + st.stopped;
    if (session.__aiReportKey === key) return;
    session.__aiReportKey = key;
    session.__aiReported = true;
    var label = st.label || 'IA';
    var modelText = st.lastModel ? ' (modèle ' + st.lastModel + ')' : '';
    if (rows || missing) {
      session.addReport({
        fileName: '(assistant IA)',
        reason: label + ' a lu ' + (rows + missing) + ' ligne(s) : ' + rows +
          ' classée(s) par l IA' + modelText + (missing ? ', ' + missing +
          ' non classée(s) (moteur local utilisé pour ces lignes)' : ''),
        suggestion: missing
          ? 'Les lignes non classées restent traitées localement et sont signalées dans le rapport'
          : 'Toutes les lignes ont été classées par l assistant IA',
        severity: missing ? 'warning' : 'info'
      });
      return;
    }
    session.addReport({
      fileName: '(assistant IA)',
      reason: hints
        ? label + ' a confirmé ' + hints + ' destination(s) de colonnes sur ' + requests +
          ' bloc(s) analysé(s)' + modelText
        : label + ' n a pas pu répondre : ' + (st.lastError || 'aucune colonne ambiguë') + ' — moteur local utilisé',
      suggestion: hints ? 'Destinations vérifiées par l assistant IA' : 'Vérifier la clé de l assistant et la connexion internet',
      severity: hints ? 'info' : 'warning'
    });
  }

  function createSession(options) {
    options = options || {};
    var session = {

      files: new Map(),
      records: new Map(),          // key -> record (dedup-safe on retry)
      recordsByFile: new Map(),    // fileId -> Array<record> (per-file index, O(1) lookup)
      report: [],
      reportIndex: new Map(),   // reportKey -> index in session.report (O(1) dedup)
      reviewItems: new Map(),      // uncertain columns, resolved automatically (kept for the report)
      learnedRules: {},            // normalized header -> { target, sourceHeader }
      cityCache: new Map(),        // 'city|cp' -> {value, warning} (INSEE lookup cache)
      valueCache: new Map(),       // 'target|raw' -> {value, warning} (transformValue cache)
      fileOrder: [],
      stats: {
        totalRows: 0,
        rowsWithWarnings: 0,
        processed: 0,
        succeeded: 0,
        warnings: 0,
        failed: 0,
        cancelled: 0
      },
      running: false,
      cancelling: false,
      aiRequests: 0,               // number of AI blocks analysed this session
      aiHints: 0,                  // ambiguous columns confirmed by the assistant
      aiRows: 0,                   // source rows placed by Gemini (full-sheet)
      /* Values the provider could not classify with certainty: they are NEVER
       * guessed into a column. They wait here for the user's decision and the
       * export is blocked until they are resolved. */
      manualItems: [],             // [{ id, key, fileName, row, sourceHeader, value, suggestion, destination, status }]
      manualIndex: new Map(),      // recordKey + '|' + field -> item (dedup)
      aiMissingRows: 0,            // rows Gemini could not classify (local engine used)
      aiTruncated: 0,              // truncated answers retried with smaller batches
      headers: (options.headers || SEC.OFFICIAL_HEADERS.slice()).slice()
    };

    /* ---------------- Report helpers (deduped, sample-capped) ---------------- */

    function reportKey(e) {
      return [e.fileName, e.sheet, e.sourceHeader, e.reason].join('§');
    }
    function addReport(entry) {
      var key = reportKey(entry);
      var existingIdx = session.reportIndex.get(key);
      if (existingIdx !== undefined) {
        var existing = session.report[existingIdx];
        existing.count++;
        if (existing.samples.length < SEC.LIMITS.REPORT_SAMPLE_PER_COLUMN && entry.sample) {
          existing.samples.push(String(entry.sample).slice(0, 120));
        }
        return;
      }
      var idx = session.report.push({
        fileName: entry.fileName,
        sheet: entry.sheet || '',
        row: entry.row || '',
        sourceHeader: entry.sourceHeader || '',
        sample: entry.sample ? String(entry.sample).slice(0, 120) : '',
        samples: entry.sample ? [String(entry.sample).slice(0, 120)] : [],
        reason: entry.reason,
        suggestion: entry.suggestion || '',
        severity: entry.severity || 'warning',
        count: 1
      }) - 1;
      session.reportIndex.set(key, idx);
    }

    /* ---------------- File registration ---------------- */

    function registerFile(meta) {
      var id = meta.id || ('f_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 8));
      var entry = {
        id: id,
        name: meta.name,
        ext: (meta.name.split('.').pop() || '').toLowerCase(),
        size: meta.size || 0,
        lastModified: meta.lastModified || 0,
        status: S.WAITING,
        progress: 0,
        stage: '',
        rowsExtracted: 0,
        warnings: 0,
        error: '',
        parsedSheets: null,
        selectedSheets: null,          // array of sheet names to process
        headerOverrides: {},           // sheetName -> header row index
        columnOverrides: {},           // sheetName -> { colIndex: target|'__exclude__' }
        duplicateOf: meta.duplicateOf || null,
        encoding: '',
        delimiter: ''
      };
      session.files.set(id, entry);
      session.fileOrder.push(id);
      return entry;
    }

    function removeFile(id) {
      var entry = session.files.get(id);
      if (!entry) return;
      dropRecords(id);
      session.files.delete(id);
      session.fileOrder = session.fileOrder.filter(function (fid) { return fid !== id; });
      /* purge review items & report rows belonging to the file */
      session.reviewItems.forEach(function (item, key) {
        if (item.fileId === id) session.reviewItems.delete(key);
      });
      session.report = session.report.filter(function (e) { return e.fileName !== entry.name; });
      session.reportIndex.clear();
      session.report.forEach(function (e, i) { session.reportIndex.set(reportKey(e), i); });
      entry.parsedSheets = null;
      recomputeStats();
    }

    function clearFiles() {
      session.files.clear();
      session.records.clear();
      session.recordsByFile.clear();
      session.cityCache.clear();
      session.valueCache.clear();
      session.report = [];
      session.reportIndex.clear();
      session.reviewItems.clear();
      session.fileOrder = [];
      resetStats();
    }

    function dropRecords(fileId) {
      var prefix = fileId + '|';
      session.records.forEach(function (rec, key) {
        if (key.indexOf(prefix) === 0) session.records.delete(key);
      });
      session.recordsByFile.delete(fileId);
    }

    function resetStats() {
      session.stats = { totalRows: 0, rowsWithWarnings: 0, processed: 0, succeeded: 0, warnings: 0, failed: 0, cancelled: 0 };
    }
    /* Recalculate everything that is derived from the records:
     * totals, per-file line counts AND each file's status. Called after every
     * extraction so the numbers can never drift from the real data. */
    function recomputeStats() {
      resetStats();
      session.fileOrder.forEach(function (fid) {
        var recs = session.recordsByFile.get(fid);
        var total = recs ? recs.length : 0;
        var warn = 0;
        if (recs) {
          recs.forEach(function (r) { if (r.hasWarning) warn++; });
        }
        var f = session.files.get(fid);
        if (!f) return;
        f.rowsExtracted = total;
        f.warnings = warn;
        /* Status reflects the real outcome: a file whose lines landed in the
         * final file is a SUCCESS. Only an unfinished file or an error is not. */
        if (f.status === S.COMPLETED || f.status === S.COMPLETED_WARNINGS) f.status = S.COMPLETED;
        if (f.status === S.COMPLETED) { session.stats.processed++; session.stats.succeeded++; }
        if (f.status === S.FAILED) session.stats.failed++;
        if (f.status === S.CANCELLED) session.stats.cancelled++;
      });
      session.stats.totalRows = session.records.size;
      var warnRows = 0;
      session.records.forEach(function (r) { if (r.hasWarning) warnRows++; });
      session.stats.rowsWithWarnings = warnRows;
    }

    /* ---------------- Parsing ---------------- */

    function parseEntry(entry, buffer, preParsed) {
      entry.status = S.VALIDATING;
      entry.stage = 'Validation du fichier';
      entry.progress = 8;
      var ext = entry.ext;
      if (SEC.SUPPORTED_EXTENSIONS.indexOf(ext) === -1) {
        throw new Error('Format non pris en charge : ' + (ext || 'inconnu'));
      }
      entry.stage = 'Lecture du contenu';
      entry.progress = 30;
      var sheets;
      if (preParsed && preParsed.sheets) {
        /* Already decompressed by the Web Worker: no SheetJS on the UI thread. */
        sheets = preParsed.sheets;
        entry.encoding = preParsed.encoding || '';
        entry.delimiter = preParsed.delimiter || '';
      } else if (ext === 'csv' || ext === 'txt') {
        var parsed = SEC.CSV.parse(buffer);
        entry.encoding = parsed.encoding;
        entry.delimiter = parsed.delimiter;
        sheets = [{ name: '(CSV)', rows: parsed.rows, isEmpty: parsed.rows.length === 0 }];
      } else {
        sheets = SEC.XLSXReader.parse(buffer);
      }
      entry.parsedSheets = sheets;
      var nonEmpty = sheets.filter(function (sh) { return !sh.isEmpty; });
      if (!entry._sheetsChosen) entry.selectedSheets = nonEmpty.map(function (sh) { return sh.name; });
      sheets.forEach(function (sh) {
        if (sh.isEmpty) {
          addReport({
            fileName: entry.name, sheet: sh.name, sourceHeader: '',
            reason: 'feuille vide ignoree', suggestion: 'Aucune action requise', severity: 'info'
          });
        }
      });
      return sheets;
    }

    /* ---------------- Extraction ---------------- */

    var currentPriceContext = false;

    function transformValue(target, raw, entry, sheetName, header, rowNum) {
      var v = N.trimClean(raw);
      if (!v) return { value: '', warning: null };
      /* Cache per (target, value): identical cell values across 100K+ rows
       * are validated only once. First occurrence is reported; duplicates
       * reuse the cached result silently. */
      /* A Cotisation column only carries prices when at least one of its
       * values is a CONFIRMED amount (currency or decimals). A column of bare
       * numbers ("70", "80") is ambiguous: those values go to « Commentaire »
       * instead of being guessed as prices. */
      var priceContext = target === 'cotisation' ? !!currentPriceContext : false;
      var cacheKey = (target === 'cotisation' && priceContext ? 'ctx|' : '') + target + '|' + v;
      var cached = session.valueCache.get(cacheKey);
      if (cached) return cached;
      var res;
      var result;
      switch (target) {
        case 'tel1':
        case 'tel2':
          res = N.analyzePhone(v);
          if (res.valid === false) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'numero de telephone invalide', suggestion: 'Verifier le numero ou exclure la colonne', severity: 'warning'
            });
            /* A value that cannot be a phone ("1", "123", a postal code, a
             * date…) must NEVER be exported in Téléphone 1/2: it is moved to
             * « Commentaire » so nothing is lost and the phone column stays
             * clean. */
            if (N.isUnusablePhone && N.isUnusablePhone(v)) {
              result = { value: '', warning: true, reason: 'valeur non telephonique deplacee vers Commentaire', rawValue: v, moveToComment: true };
            } else {
              result = { value: res.value, warning: true, reason: 'numero de telephone invalide', rawValue: v };
            }
          } else {
            result = { value: res.value, warning: res.valid === 'international' ? 'info' : null };
          }
          break;
        case 'cp':
          res = N.analyzePostal(v);
          if (!res.valid) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'code postal invalide', suggestion: 'Verifier le code (5 chiffres)', severity: 'warning'
            });
            result = { value: res.value, warning: true, reason: 'code postal invalide', rawValue: v };
          } else if (SEC.Communes && SEC.Communes.count() && !SEC.Communes.isPostalValid(res.value)) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'code postal inexistant (' + res.value + ')', suggestion: 'Verifier le code postal', severity: 'warning'
            });
            result = { value: res.value, warning: true, reason: 'code postal inexistant', rawValue: v };
          } else {
            result = { value: res.value, warning: null };
          }
          break;
        case 'email':
          res = N.analyzeEmail(v);
          if (!res.valid) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'format email suspect', suggestion: 'Verifier l adresse email', severity: 'warning'
            });
            result = { value: res.value, warning: true, reason: 'format email suspect', rawValue: v };
          } else {
            result = { value: res.value, warning: null };
          }
          break;
        case 'cotisation':
          /* An insurer name ("aesio", "alliance", "apivia"…) is NOT a
           * cotisation: it belongs to « Mutuelle », even in a Cotisation
           * column. An age ("63 ans", "70+75", "1950-1948") belongs to
           * « Commentaire ». Both are detected from the value itself. */
          if (N.looksLikeInsurer(v) && !N.looksLikeAmount(v)) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'organisme de mutuelle detecte dans la cotisation — place dans « Mutuelle »',
              suggestion: 'Aucune action requise', severity: 'info'
            });
            result = { value: N.matchInsurer(v) || v, warning: null, moveTo: 'mutuelle' };
            break;
          }
          /* A bare number ("70", "80") may be a price OR an age: without a
           * currency or decimals it is not confirmed, so it is never stated
           * as a cotisation — it goes to Commentaire unless the column
           * context proves prices (see requiresPriceContext below). */
          if (N.isAmbiguousNumber && N.isAmbiguousNumber(v) && !priceContext) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'nombre sans devise ni centimes: age ou prix ambigu — place dans « Commentaire »',
              suggestion: 'Si ces valeurs sont bien des cotisations, choisir « Cotisation » pour la colonne',
              severity: 'warning'
            });
            result = { value: v, warning: null, moveTo: 'commentaire' };
            break;
          }
          if (N.looksLikeAge(v) && !N.looksLikeAmount(v) && !/^\d{3,4}[.,]\d/.test(v)) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'age detecte dans la cotisation — place dans « Commentaire »',
              suggestion: 'Aucune action requise', severity: 'info'
            });
            result = { value: v, warning: null, moveTo: 'commentaire' };
            break;
          }
          res = N.analyzeAmount(v);
          /* Recovery: values like "A· 45,60 €" or "55 euros" extracted with
           * their numeric part so a cotisation is never lost to a comment. */
          if (!res.valid) {
            var numPart = N.extractAmount(v);
            if (numPart) {
              addReport({
                fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
                sample: v, reason: 'montant recupere dans une valeur composee : « ' + numPart + ' »',
                suggestion: 'Le montant numerique a ete place dans Cotisation', severity: 'info'
              });
              result = { value: numPart, warning: null };
              break;
            }
          }
          if (!res.valid) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: 'montant de cotisation non numerique', suggestion: 'Verifier le format du montant', severity: 'warning'
            });
            result = { value: res.value, warning: true, reason: 'montant de cotisation non numerique', rawValue: v };
          } else {
            result = { value: res.value, warning: null };
          }
          break;
        case 'ville':
          /* Resolved later, row by row, together with the postal code of the
           * same row (see resolveCity): the commune database is authoritative. */
          result = { value: v, warning: null };
          break;
        case 'nom':
        case 'prenom':
          /* A note / date / age never belongs to a name: it is moved to
           * « Commentaire » right away (never lost, never misplaced). */
          if (N.isNoteLike(v) || N.isInvalidPersonName(v)) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: (target === 'nom' ? 'Nom' : 'Prénom') + ' : note ou valeur sans rapport — placee dans « Commentaire »',
              suggestion: 'Aucune action requise', severity: 'info'
            });
            result = { value: '', warning: null, rawValue: v, moveToComment: true };
            break;
          }
          if (N.isInvalidPersonName(v)) {
            addReport({
              fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
              sample: v, reason: (target === 'nom' ? 'Nom' : 'Prénom') + ' non conforme : contient des chiffres, dates ou notes de rappel (« ' + v.slice(0, 80) + ' »)',
              suggestion: 'Vérifier la colonne et choisir « Commentaire » ou « Adresse »', severity: 'warning'
            });
            result = {
              value: v,
              warning: true,
              isInvalidName: true,
              reason: (target === 'nom' ? 'Nom' : 'Prénom') + ' non conforme (chiffres/rappel)',
              rawValue: v
            };
          } else {
            result = { value: v, warning: null };
          }
          break;
        default:
          /* adresse, mutuelle, commentaire : preserve original text.
           * The Mutuelle column uses the OFFICIAL name from the database
           * ("swisselife" -> "SwissLife", "aesio" -> "Aésio"). */
          if (target === 'mutuelle') {
            var official = N.matchInsurer(v);
            result = { value: official || v, warning: null };
            break;
          }
          /* Keep the ORIGINAL text of a comment (comma kept: "Dupont, Jean")
           * and its line breaks. */
          result = {
            value: String(raw).replace(/\u00A0/g, ' ').replace(/\r\n?/g, ' \u2022 ').trim(),
            warning: null,
          };
          break;
      }
      session.valueCache.set(cacheKey, result);
      return result;
    }

    /* A destination column keeps only its own TYPE — whatever (human or AI)
     * wrote the value. Text, dates, reminders and junk never stay in
     * « Téléphone 1/2 », « Mutuelle », « Cotisation » or « Code postal »: they
     * move to « Commentaire » (nothing is ever lost). Applied to EVERY row,
     * idempotent. This is the rule that keeps the export free of misplaced
     * values whatever the assistant answered. */
    /* Does this value TYPE-FIT the slot? Used to rescue a value the assistant
     * left unclassified: the local engine's correct placement survives. */
    function slotAcceptsValue(key, v) {
      var t = N.trimClean(v);
      if (!t) return false;
      switch (key) {
        case 'tel1':
        case 'tel2':
          return N.looksLikePhone(t) && !(N.isUnusablePhone && N.isUnusablePhone(t));
        case 'cp':
          return N.looksLikePostal(t);
        case 'email':
          return N.looksLikeEmail(t);
        case 'cotisation':
          return N.looksLikeAmount(t);
        case 'mutuelle':
          return !!(N.matchInsurer(t) || (N.looksLikeMutuelle && N.looksLikeMutuelle(t)));
        case 'ville':
          if (SEC.Communes && SEC.Communes.count()) {
            var res = SEC.Communes.validate(t, '');
            if (res.status === 'ok' || res.status === 'corrected') return true;
          }
          return !!(N.looksLikeCity && N.looksLikeCity(t) && !N.looksLikePersonName(t));
        default:
          return false;
      }
    }

    /* Two texts meaning the same value, ignoring accents, case and every
     * separator: "06 11 00 10 05" and "0611001005" are the SAME phone. */
    function sameLoose(a, b) {
      var norm = function (s) {
        var t = N.stripAccents ? N.stripAccents(String(s == null ? '' : s)) : String(s == null ? '' : s);
        return t.toLowerCase().replace(/[^a-z0-9]/g, '');
      };
      var x = norm(a), y = norm(b);
      return !!x && x === y;
    }

    /* Text that is a date in any common French notation. */
    function isDateLikeText(s) {
      var t = N.trimClean(s);
      if (!t || t.length > 40) return false;
      return /\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}/.test(t) ||
        /^\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}$/.test(t);
    }
    /* Typed destinations and the values they refuse outright — used to vet
     * what the assistant wrote before it can overwrite the local reading.
     * These rules are deliberately STRICT about the wrong type and permissive
     * about the right one (an insurer name that is not in the dictionary is
     * still accepted as a Mutuelle) so a correct value is never thrown away. */
    var TYPED_SLOTS = {
      tel1: function (t) { return !(N.looksLikePhone(t) && !(N.isUnusablePhone && N.isUnusablePhone(t))); },
      tel2: function (t) { return !(N.looksLikePhone(t) && !(N.isUnusablePhone && N.isUnusablePhone(t))); },
      cp: function (t) { return !N.looksLikePostal(t); },
      email: function (t) { return !N.looksLikeEmail(t); },
      cotisation: function (t) { return isDateLikeText(t) || !N.looksLikeAmount(t); },
      mutuelle: function (t) {
        if (isDateLikeText(t) || N.looksLikeAmount(t) || N.looksLikePhone(t)) return true;
        if (SEC.Communes && SEC.Communes.count && SEC.Communes.count() && SEC.Communes.isKnownCity(t)) return true;
        return !!(N.looksLikePersonName && N.looksLikePersonName(t) && !(N.isPureInsurer && N.isPureInsurer(t)));
      },
      ville: function (t) {
        if (isDateLikeText(t) || N.looksLikeAmount(t) || N.looksLikePhone(t)) return true;
        return !!(N.looksLikePersonName && N.looksLikePersonName(t) && !(N.looksLikeCity && N.looksLikeCity(t)));
      }
    };

    function enforceDestinationTypes(values) {
      var pushComment = function (t) {
        var x = N.trimClean(t);
        if (!x) return;
        values.commentaire = values.commentaire ? (values.commentaire + ' | ' + x) : x;
      };
      var isDateLike = function (s) {
        return /\d{1,2}[\/\-.]\d{1,2}[\/\-.]\d{2,4}/.test(s) ||
          /^\d{4}[\/\-.]\d{1,2}[\/\-.]\d{1,2}$/.test(s);
      };
      /* 1. Téléphone 1/2 = phone numbers ONLY (never text, dates, reminders). */
      ['tel1', 'tel2'].forEach(function (slot) {
        if (!values[slot]) return;
        var segs = String(values[slot]).split(/\s*\|\s*/).filter(Boolean);
        var keep = [];
        segs.forEach(function (seg) {
          if (N.looksLikePhone(seg) && !(N.isUnusablePhone && N.isUnusablePhone(seg))) {
            var an = N.analyzePhone(seg);
            keep.push(an.value || seg);
          } else {
            pushComment(seg);
          }
        });
        values[slot] = keep.join(' | ');
      });
      /* 2. Mutuelle = insurer names only: official name when recognized, an
       * unknown small mutuelle name is kept, junk (notes, dates, placeholders)
       * moves to Commentaire. */
      if (values.mutuelle) {
        var mSegs = String(values.mutuelle).split(/\s*\|\s*/).filter(Boolean);
        var mKeep = [];
        mSegs.forEach(function (seg) {
          var t = N.trimClean(seg);
          if (!t) return;
          var official = N.matchInsurer(t);
          if (official) { mKeep.push(official); return; }
          if (N.isPlaceholder && N.isPlaceholder(t)) return;
          if (/^(autre|aucun(e)?|sans objet|sans|n[\/.]?a[\/.]?|neant|ras|x|-+|\.+)$/i.test(t)) return;
          if (N.isNoteLike(t) || (N.looksLikeCommentText && N.looksLikeCommentText(t)) ||
              N.looksLikePhone(t) || N.looksLikeAmount(t) || isDateLike(t)) {
            pushComment(t);
            return;
          }
          /* Not an insurer name at all: a city ("PARIS") or a person name
           * must never stay in « Mutuelle » — it moves to Commentaire (an
           * unknown small mutual is kept, real data is never discarded). */
          if (SEC.Communes && SEC.Communes.isKnownCity && SEC.Communes.isKnownCity(t)) {
            pushComment(t);
            return;
          }
          if (N.looksLikePersonName && N.looksLikePersonName(t) &&
              !(N.looksLikeInsurer && N.looksLikeInsurer(t))) {
            pushComment(t);
            return;
          }
          mKeep.push(t);
        });
        values.mutuelle = mKeep.join(' | ');
      }
      /* 3. Cotisation = money only: a date, an age or a reminder never. */
      if (values.cotisation) {
        var c = N.trimClean(values.cotisation);
        var money = N.looksLikeAmount(c);
        if (!money && (isDateLike(c) || (N.isAmbiguousNumber && N.isAmbiguousNumber(c)) ||
            (N.looksLikeAge && N.looksLikeAge(c)) || N.isNoteLike(c) || !/\d/.test(c))) {
          pushComment(c);
          values.cotisation = '';
        } else if (!money) {
          /* A composed value carries an amount ("A· 45,60 €"): keep the
           * number, never the wrapper. */
          var numPart = /[.,]\d{2}|\d\s*(?:€|euros?|eur|fcfa|francs?)\b/i.test(c) ? N.extractAmount(c) : '';
          if (numPart) values.cotisation = numPart;
          else { pushComment(c); values.cotisation = ''; }
        }
      }
      /* 4. Code postal = never a date or a note. */
      if (values.cp) {
        var cp = N.trimClean(values.cp);
        if (!N.looksLikePostal(cp) && (isDateLike(cp) || (N.isNoteLike && N.isNoteLike(cp)))) {
          pushComment(cp);
          values.cp = '';
        }
      }
    }

    /* City validation against the official commune database.
     * - identified with certainty -> official INSEE name (accents, spacing)
     * - not identified -> Ville left EMPTY, the raw text is reported
     * - a postal code alone never creates a city unless it designates one
     *   single commune AND is compatible with the written name.
     * The value is never invented. */
    function resolveCity(rawCity, cpValue, entry, sheetName, header, rowNum) {
      var C = SEC.Communes;
      if (!C || !C.count()) {
        return { value: N.trimClean(rawCity), warning: null };
      }
      var cityKey = N.trimClean(rawCity).toLowerCase();
      var cpKey = N.trimClean(cpValue || '');
      var cacheKey = cityKey + '|' + cpKey;
      var cached = session.cityCache.get(cacheKey);
      if (cached) return cached;
      var res = C.validate(rawCity, cpValue);
      if (res.status === 'ok') {
        cached = { value: res.value, warning: null };
      } else if (res.status === 'corrected') {
        addReport({
          fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
          sample: String(rawCity).slice(0, 120),
          reason: 'ville corrigee selon la base officielle des communes (INSEE) : « ' + res.value + ' »' +
            (res.insee ? ' [' + res.insee + ']' : ''),
          suggestion: 'Orthographe officielle appliquee', severity: 'info'
        });
        cached = { value: res.value, warning: null };
      } else if (res.status === 'empty') {
        cached = { value: '', warning: null };
      } else {
        var why = res.status === 'ambiguous'
          ? 'plusieurs communes portent ce nom et aucun code postal ne les distingue'
          : 'commune introuvable dans la base officielle (INSEE / codes postaux)';
        addReport({
          fileName: entry.name, sheet: sheetName, row: rowNum, sourceHeader: header,
          sample: String(rawCity).slice(0, 120),
          reason: 'ville non identifiee : ' + why + ' — champ Ville laisse vide',
          suggestion: 'Verifier la ville dans le fichier source', severity: 'warning'
        });
        cached = { value: '', warning: true, reason: 'ville non identifiee', rawValue: N.trimClean(rawCity) };
      }
      session.cityCache.set(cacheKey, cached);
      return cached;
    }

    function effectiveTarget(entry, sheetName, col) {
      var ov = entry.columnOverrides[sheetName] && entry.columnOverrides[sheetName][col._index];
      if (ov !== undefined) return ov;
      if (col.learned && col.excluded) return '__exclude__';
      return col.best;
    }


    function extractSheet(entry, sheet, yieldFn, preparedColumns, aiHints, aiRows) {
      var headerIdx = entry.headerOverrides[sheet.name];
      var analysis;
      if (preparedColumns) {
        /* Re-use the analysis the AI step already saw: SAME columns, SAME
         * header row. Running analyzeSheet twice made the hints land on a
         * different object (and doubled the CPU cost of every sheet). */
        analysis = { columns: preparedColumns, headerRowIndex: preparedColumns[0] ? preparedColumns[0].headerRowIndex : (typeof headerIdx === 'number' ? headerIdx : SEC.Agent.detectHeaderRow(sheet.rows).index) };
      } else {
        analysis = SEC.Agent.analyzeSheet(sheet.rows, typeof headerIdx === 'number' ? headerIdx : undefined);
      }
      var columns = analysis.columns;
      /* True when Gemini had this sheet but could not answer (no key, quota,
       * offline): the local engine then keeps its best evidence instead of
       * demoting a value-driven guess that nothing else can decide. */
      var aiUnavailable = !!(aiHints && aiHints.__unavailable);
      if (aiHints) applyGeminiHints(columns, aiHints);
      /* Deliberate exclusions (identifier/date/year, out-of-scope profiles) are
       * applied AFTER Gemini answered: Gemini gets one chance on those columns
       * first (it is exactly the "date ou identifiant ?" judgement), then a
       * column it could not decide is excluded and reported, never dumped. */
      columns.forEach(function (col) {
        if (col.aiHint) return;
        if (col.outOfScopePending) {
          col.excluded = true;
          col.profileExcluded = true;
          return;
        }
        if (col.profileExcludedPending) {
          col.excluded = true;
          col.profileExcluded = true;
        }
      });
      entry.columnOverrides = entry.columnOverrides || {};
      entry.columnOverrides[sheet.name] = entry.columnOverrides[sheet.name] || {};
      var dataRows = sheet.rows.slice(analysis.headerRowIndex + 1);
      var newReview = 0;
      /* Source row number (1-based, as in the original file) of each record:
       * the AI answers are keyed by it, so a line can never be mixed with
       * another one. */
      var dataStartRow = analysis.headerRowIndex + 2;

      /* Period context for Cotisation headers (no silent monthly/annual conversion). */
      columns.forEach(function (col) {
        if (col.best === 'cotisation') {
          var hn = col.headerNorm;
          var annual = SEC.PERIOD_TERMS_ANNUAL.some(function (t) { return hn.indexOf(N.normalizeHeader(t)) !== -1; });
          var monthly = SEC.PERIOD_TERMS_MONTHLY.some(function (t) { return hn.indexOf(N.normalizeHeader(t)) !== -1; });
          if (annual || monthly) {
            addReport({
              fileName: entry.name, sheet: sheet.name, sourceHeader: col.header,
              reason: 'periode de cotisation detectee : ' + (annual ? 'annuelle' : 'mensuelle') + ' (valeur conservee sans conversion)',
              suggestion: 'Aucune conversion automatique n est appliquee', severity: 'info'
            });
          }
        }
      });

      /* Detect the destinations of every column. A column the engine flags
       * uncertain is resolved automatically (no manual validation step). */
      var uncertainItems = [];
      columns.forEach(function (col) {
        var target = effectiveTarget(entry, sheet.name, col);
        var itemId = entry.id + '|' + sheet.name + '|' + col._index;
        var hasManualOverride = entry.columnOverrides[sheet.name] &&
          entry.columnOverrides[sheet.name][col._index] !== undefined;
        /* A destination the user already chose is final: remove the open
         * review item (the override stays in columnOverrides). Excluded
         * columns stay visible as 'excluded' review items. */
        if (hasManualOverride) {
          var userChoice = entry.columnOverrides[sheet.name][col._index];
          if (userChoice === '__exclude__') {
            var excludedItem = {
              id: itemId, fileId: entry.id, fileName: entry.name, sheet: sheet.name,
              colIndex: col._index, headerNorm: col.headerNorm,
              status: 'excluded', decision: '__exclude__',
              suggestion: col.best || null,
              header: col.header,
              samples: col.profile.samples,
              confidence: 100,
              validation: col.validation,
              reasons: ['Colonne exclue par l utilisateur'],
              blocked: col.blocked,
              fullName: col.fullName
            };
            session.reviewItems.set(itemId, excludedItem);
          } else {
            session.reviewItems.delete(itemId);
          }
          return;
        }
        var hasRappel = (col.profile.samples || []).some(function (s) {
          return /rappel|rdv|date\s+et\s+heure/i.test(s);
        });
        var isAmbiguousHeader = !col.header || /date|heure|retour|age|nombre|personne|statut/i.test(col.headerNorm);
        /* A CLEAR destination (strong header match, values accepted, single
         * leading destination) needs no manual verification, even at 88/89:
         * listing it made whole batches look "unfinished" for nothing. A
         * real conflict / weak lead still goes to review. */
        var margin = typeof col.margin === 'number' ? col.margin : 100;
        var profile = col.profile || {};
        var clearDestination = !!(col.best && col.headerMatched && col.score >= 85 &&
          margin >= 20 && !(col.validation && col.validation.status === 'warn') &&
          !col.fullName && !col.blocked && !hasRappel);
        /* A SHORT header that is an exact synonym the dictionary does not know
         * ("Insurer", "Mutual", "Town", "Montant") with coherent values is
         * a clear destination too. */
        if (!clearDestination && col.best && col.shortHeaderMatch && margin >= 20 &&
            !(col.validation && col.validation.status === 'warn') && !col.fullName &&
            !col.blocked && !hasRappel) clearDestination = true;
        /* Confirmed by the optional assistant: decided, no need to re-ask. */
        if (col.aiHint && !hasRappel && !col.fullName && !col.blocked) clearDestination = true;
        /* A high-confidence INHERENT type (a real e-mail / phone / postal code /
         * amount, or a 90 %+ insurer name) is certain by content: only values
         * that fail their own format (validation 'warn') need review. */
        var ratioKey = { email: 'email', tel1: 'phone', tel2: 'phone', cp: 'postal', cotisation: 'amount', mutuelle: 'mutuelle' }[col.best];
        var inherentType = !!(ratioKey && profile.ratios && (profile.ratios[ratioKey] || 0) >= 0.9);
        if ((col.aiHint || inherentType) && !hasRappel && !col.fullName && !col.blocked &&
            !(col.best === 'cp' && col.validation && col.validation.validRatio < 1)) clearDestination = true;
        /* A column that only ADDS information to a destination already fully
         * filled by a stronger column is merged automatically: information is
         * preserved, the row is not held back. */
        var mergeable = !!(col.mergeInto && col.mergeInto !== 'commentaire');
        /* A column decided by the profile heuristics (a strict running id, a
         * pure date/years/ages column) is EXCLUDED on purpose: its values are
         * reported once but they cannot land in any of the 11 official
         * columns. Asking the user to "choose a destination" for it on every
         * run made batches look permanently unfinished. */
        var deliberate = !!col.profileExcluded;
        /* A destination computed ONLY from the shape of the values ("Titre
         * client" -> Ville because Dupont looks like a city) is not proof.
         * Gemini decides it online; only when Gemini cannot answer at all
         * does the value shape stay the last local evidence. */
        var valueOnly = !!col.valueOnlyGuess;
        if (valueOnly && col.aiHint) valueOnly = false;
        if (valueOnly && !aiUnavailable && col.best && col.best !== 'commentaire' &&
            SEC.TARGET_LABELS[col.best]) {
          col.best = 'commentaire';
          col.validation = { status: 'info', validRatio: 1, reason: '' };
          col.reasons = (col.reasons || []).concat(['en-tete non reconnu : valeur conservee dans « Commentaire »']);
          target = 'commentaire';
        }
        var needsReview =
          !deliberate && (
          col.fullName ||
          col.blocked ||
          (!col.best || valueOnly) ||
          (col.score < 90 && !clearDestination) ||
          hasRappel ||
          (isAmbiguousHeader && !clearDestination) ||
          /* Age / age-range columns are shown in « Vérification des
           * correspondances » so the user can choose the destination
           * himself (default suggestion: Commentaire). */
          col.rewrittenAsComment ||
          (col.validation && col.validation.status === 'warn') ||
          /* Détection forcée : une colonne « Code postal » qui contient
           * ne serait-ce qu'un seul code postal invalide doit passer en
           * revue (sinon ces valeurs erronées sont exportées en l'état). */
          (col.best === 'cp' && col.validation && col.validation.validRatio < 1) ||
          (col.reasons || []).some(function (r) { return /conflit|supplementaire|decoupage|non reconnu/.test(r); }));
        if (mergeable) needsReview = false;
        if (needsReview) {
          var existing = session.reviewItems.get(itemId);
          var item = existing || {
            id: itemId, fileId: entry.id, fileName: entry.name, sheet: sheet.name,
            colIndex: col._index, headerNorm: col.headerNorm, status: 'open', decision: null,
            suggestion: null
          };
          item.header = col.header;
          item.samples = col.profile.samples;
          item.confidence = col.score;
          item.validation = col.validation;
          item.reasons = (col.reasons || []).slice();
          item.suggestion = hasRappel ? 'commentaire' : ((target && target !== '__exclude__') ? target : (col.best || 'commentaire'));
          if (hasRappel && !item.reasons.some(function(r) { return /rappel/.test(r); })) {
            item.reasons.push('contient des notes de rappel / dates : suggéré pour « Commentaire »');
          }
          item.blocked = col.blocked;
          item.fullName = col.fullName;
          if (!existing) {
            session.reviewItems.set(itemId, item);
            newReview++;
            if (item.status === 'open') uncertainItems.push(item);
          }
        } else {
          session.reviewItems.delete(itemId);
        }
      });
      /* Automatic resolution: each uncertain column gets its final destination
       * now, exactly as the « Correction » pass would have decided it. The
       * decision is then applied to the file like a user choice. */
      if (uncertainItems.length) {
        resolveOpenItems(uncertainItems);
        uncertainItems.forEach(function (it) {
          if (it.decision) entry.columnOverrides[it.sheet][it.colIndex] = it.decision;
        });
      }

      /* Extract rows. Each value goes to the destination of its column and
       * nowhere else: no automatic split, redirect or rewrite. */
      var mappedCols = columns.map(function (col) {
        return { col: col, target: effectiveTarget(entry, sheet.name, col) };
      }).filter(function (m) {
        if (!m.target || m.target === '__exclude__') return false;
        /* A profile exclusion only applies when nothing stronger was decided:
         * an AI hint on a genuinely useful header wins over the heuristic. */
        if (m.col.profileExcluded && !m.col.headerMatched) return false;
        return true;
      });
      /* Does a Cotisation column of this sheet carry confirmed prices
       * (currency / decimals / explicit price header)? Bare numbers are
       * otherwise treated as ambiguous ages and sent to Commentaire. */
      currentPriceContext = (function () {
        var priceHeader = /cotisation|prime|montant|tarif|prix|acompte|mensualite|annuite|contribution|adhesion/i;
        return mappedCols.some(function (mc) {
          if (mc.target !== 'cotisation') return false;
          if (priceHeader.test(String(mc.col.header || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')))
            return true;
          return ((mc.col.profile && mc.col.profile.samples) || []).some(function (v) {
            return N.isConfirmedAmount && N.isConfirmedAmount(v);
          });
        });
      })();

      /* Report unmapped / excluded columns once. */
      columns.forEach(function (col) {
        var target = effectiveTarget(entry, sheet.name, col);
        /* A profile exclusion only applies when NOTHING overrides it (AI hint,
         * user decision, confident local destination). */
        if (col.profileExcluded && !col.headerMatched && !col.aiHint &&
            (target === col.best || !target || target === '__exclude__')) {
          addReport({
            fileName: entry.name, sheet: sheet.name, sourceHeader: col.header,
            sample: col.profile.samples[0] || '',
            reason: 'colonne identifiant/date/annee/age : aucune des 11 colonnes officielles ne la concerne',
            suggestion: 'Choisir une destination dans « Vérification des correspondances » si ces valeurs doivent être conservées',
            severity: 'info'
          });
          return;
        }
        if ((!target || target === '__exclude__') && col.profile.nonEmpty > 0) {
          var excluded = target === '__exclude__' || (col.learned && col.excluded);
          addReport({
            fileName: entry.name, sheet: sheet.name, sourceHeader: col.header,
            sample: col.profile.samples[0] || '',
            reason: excluded ? 'colonne exclue volontairement' : 'colonne non reconnue (non cartographiee)',
            suggestion: excluded ? 'Aucune action (choix utilisateur)' : 'Choisir une destination ou exclure la colonne',
            severity: 'info'
          });
        }
      });

      var rowCount = 0;
      /* Values that never change from row to row are computed ONCE per sheet
       * (they used to be rebuilt for every single row). */
      var villeCol = null;
      mappedCols.forEach(function (m2) { if (m2.target === 'ville' && !villeCol) villeCol = m2; });
      var nomColNorm = '';
      mappedCols.some(function (m3) {
        if (m3.target === 'nom') { nomColNorm = m3.col.headerNorm; return true; }
        return false;
      });
      var cpColHeaderForSheet = '';
      mappedCols.some(function (m4) {
        if (m4.target === 'cp') { cpColHeaderForSheet = m4.col.header || ''; return true; }
        return false;
      });
      var SUBHEADER_WORD = /^(date|mail|courriel|tel|telephone|cp|code|mutuelle|cotisation|nom|prenom|age|nombre)\b/;
      var nomKeyIndex = SEC.TARGET_KEYS.indexOf('nom');
      var prenomKeyIndex = SEC.TARGET_KEYS.indexOf('prenom');

      /* Adaptive batches: keep each synchronous slice around 8 ms so the
       * progress bar moves, the Cancel button reacts and scrolling stays
       * smooth from 100 rows to 100 000 rows (no fixed 1500-row freeze). */
      var batchSize = SEC.LIMITS.YIELD_EVERY_ROWS;
      /* A slice keeps the UI responsive while using most of every frame:
       * ~14 ms of work per paint instead of 8 ms (≈ +75 % throughput), still
       * yielding between slices so progress, Cancel and scrolling stay live. */
      var BATCH_BUDGET_MS = 14;
      var rowMs = 0.02; // measured per-row cost (seeded, refined while running)
      var slowRowSeen = false;

      /* Process rows in batches to avoid building a giant promise chain
       * upfront (which freezes the main thread on large files). Each batch
       * runs synchronously, then yields to the UI before the next batch. */
      var SUBHEADER_EXACT = /^(date(\s+et\s+heure)?|mail|courriel|tel|telephone|cp|code\s+postal|mutuelle|cotisation|nom(\s+et\s+prenom)?|prenom|age|nombre\s+de\s+personne)$/;
      function isSubHeaderRow(srcRow) {
        var nonBlank = 0, headerTerms = 0;
        for (var s = 0; s < srcRow.length; s++) {
          if (N.trimClean(srcRow[s]) === '') continue;
          nonBlank++;
          if (SUBHEADER_EXACT.test(N.normalizeHeader(srcRow[s]))) headerTerms++;
        }
        return nonBlank >= 3 && headerTerms / nonBlank >= 0.5;
      }

      var processBatch = function (start) {
        var t0 = (typeof Date !== 'undefined' ? Date.now() : 0);
        var end = Math.min(start + batchSize, dataRows.length);
        for (var i = start; i < end; i++) {
          if (!slowRowSeen && t0 && Date.now() - t0 > BATCH_BUDGET_MS * 3) {
            /* The batch overran its budget: shrink the next slices. */
            slowRowSeen = true;
            batchSize = Math.max(25, Math.round(Math.min(batchSize, (i - start) / 2) || 50));
          }
          var srcRow = dataRows[i];
          var hasAnySource = srcRow.some(function (c) { return N.trimClean(c) !== ''; });
          if (!hasAnySource) continue; // genuinely empty row -> ignored

          /* 0. Embedded sub-header row detection (e.g. 'date et heure', 'Mail', 'Tel', 'Mutuelle', 'Nom et Prénom').
           * Only a row whose FIRST cell is a known header word can be one: the
           * full check is skipped for every ordinary data row (hot path). */
          var firstNorm = N.normalizeHeader(srcRow[0]);
          if (SUBHEADER_WORD.test(firstNorm) && isSubHeaderRow(srcRow)) {
            /* This is a duplicate / sub-table header row: skip it completely from customer output! */
            continue;
          }
          /* Gemini was consulted for this line but placed nothing in it: the
           * line is kept and handled by the local engine below (counted once,
           * reported once, never dropped without a mention). */
          var aiSeenEmpty = !!(aiRows && aiRows[dataStartRow + i] &&
            !SEC.TARGET_KEYS.some(function (t) { return (aiRows[dataStartRow + i][t] || '') !== ''; }));
          if (aiSeenEmpty) aiStats.empty++;
          var values = {};
          SEC.TARGET_KEYS.forEach(function (t) { values[t] = ''; });
          /* What each mapped column wrote / may be moved later (see FORCE-CORRECT). */
          var written = {};
          var strongDest = {};
          /* Normalised header of the first column feeding « Nom »: used to read
           * a name split already provided by the assistant. */
          var nameColumnNorm = nomColNorm;
          mappedCols.forEach(function (m2) {
            if (m2.col && m2.col.headerMatched && m2.col.score >= 85) strongDest[m2.target] = true;
          });
          var usedCommentParts = [];
          /* True when Gemini classified this line: the local content checks
           * then only verify / normalise, they never re-classify. */
          var aiPlaced = false;
          /* Set when the assistant left pieces in « Commentaire »: they are
           * proposed for validation once the row has been cleaned. */
          var aiPiecesPending = null;

          var problemList = [];   // per-row human-readable problems

          mappedCols.forEach(function (m) {
            var raw = srcRow[m.col._index];
            if (N.trimClean(raw) === '') return;
            /* Placeholder junk (".", "...", "//", "N/A") means EMPTY: it must
             * never be written to the final file nor flagged as an error.
             * Exception: a value destined for « Commentaire » is preserved
             * exactly — a note like "RAS" is a legitimate comment, not junk. */
            if (N.isPlaceholder && N.isPlaceholder(raw) && m.target !== 'commentaire') return;
            var target = m.target;

            /* The destination chosen by the user (or detected with certainty)
             * is applied EXACTLY: the value is never redirected, merged or
             * copied into another column here. Formatting is only applied to
             * the value itself (phone digits, postal code, amount). */
            var out = transformValue(target, raw, entry, sheet.name, m.col.header, analysis.headerRowIndex + i + 2);
            /* Value whose real nature contradicts the destination
             * (insurer in Cotisation -> Mutuelle, age -> Commentaire). */
            if (out.moveTo) {
              var dest = out.moveTo;
              if (dest === 'commentaire') {
                usedCommentParts.push(out.value);
              } else if (values[dest] === '') {
                values[dest] = out.value;
              } else {
                values[dest] = values[dest] + ' | ' + out.value;
              }
              return;
            }
            if (out.moveToComment) {
              /* Value that cannot be a phone: kept in « Commentaire ». */
              usedCommentParts.push(out.rawValue);
              problemList.push(SEC.TARGET_LABELS[target] + ' : valeur « ' + out.rawValue + ' » non telephonique, placee dans Commentaire');
              return;
            }
            if (out.value === '') return;

            if (target === 'commentaire') {
              usedCommentParts.push(out.value);
              written.commentaire = (written.commentaire ? written.commentaire + '\u0001' : '') + out.value;
            } else if (values[target] !== '') {
              /* Two source columns for one destination: the values are kept in
               * that destination (separator), never moved to another column. */
              values[target] = values[target] + ' | ' + out.value;
            } else {
              values[target] = out.value;
            }
            written[target] = (written[target] ? written[target] + '\u0001' : '') + out.value;
            if (out.warning === true) {
              /* Value-format quirks (bad phone, bad CP…) are recorded in the
               * report and in the line detail, but they do NOT turn the row
               * into a problem: the value is kept exactly where the user put
               * it, so the line counts as added successfully. */
              problemList.push(SEC.TARGET_LABELS[target] + ' : ' + (out.reason || 'valeur non conforme') +
                (out.rawValue ? ' (« ' + out.rawValue + ' »)' : ''));
            }
          });

          /* A mapped value that resolved to nothing (placeholder / junk) is
           * an empty cell, never a warning. */
          Object.keys(values).forEach(function (k) {
            if (values[k] && (values[k] === '.' || /^\.+$/.test(values[k]))) values[k] = '';
          });

          if (usedCommentParts.length) values.commentaire = usedCommentParts.join(' | ');

          /* 0bis. Gemini read this line: it takes over the placement (every
           * source word is placed exactly once, Nom / Prénom are separated by
           * Gemini itself). An answer that does not rebuild the source text is
           * ignored and the local reading stays. */
          if (aiRows) {
            var sourceRowNumAi = dataStartRow + i;
            var aiRow = aiRows[sourceRowNumAi];
            if (aiRow) {
              var aiValues = aiValuesForRow(aiRow, srcRow);
              if (aiValues) {
                if (aiValues.__dropped) {
                  addReport({
                    fileName: entry.name, sheet: sheet.name, row: sourceRowNumAi,
                    sourceHeader: '(assistant IA)',
                    reason: 'reponse IA partiellement rejetee : une valeur absente du fichier source a ete ignoree',
                    suggestion: 'Classement local conserve pour cette valeur', severity: 'info'
                  });
                }
                /* Gemini classifies the whole line: its placement REPLACES
                 * the local reading (the local heuristics are what
                 * mis-placed values in the first place). Nothing is lost: a
                 * source word Gemini placed nowhere was already appended to
                 * « Commentaire » above. */
                var localBefore = {};
                SEC.TARGET_KEYS.forEach(function (key) {
                  localBefore[key] = values[key];
                  values[key] = '';
                });
                SEC.TARGET_KEYS.forEach(function (key) {
                  if (typeof aiValues[key] === 'string') values[key] = aiValues[key];
                });
                /* A typed destination keeps only its own type, even when the
                 * assistant wrote it: a text in « Téléphone 1 », a city in
                 * « Mutuelle » or a date in « Cotisation » is moved back to
                 * « Commentaire » right away, so the local reading of that
                 * same slot (rescued just below) is kept instead of being
                 * dropped. This is what corrupted Téléphone / Mutuelle /
                 * Cotisation on the rows the assistant got wrong. */
                SEC.TARGET_KEYS.forEach(function (key) {
                  if (key === 'commentaire' || !values[key]) return;
                  if (!TYPED_SLOTS[key]) return;
                  var segs = String(values[key]).split(/\s*\|\s*/);
                  var keep = [];
                  segs.forEach(function (seg) {
                    var s = N.trimClean(seg);
                    if (!s) return;
                    if (!TYPED_SLOTS[key](s)) { keep.push(s); return; }
                    values.commentaire = values.commentaire ? (values.commentaire + ' | ' + s) : s;
                  });
                  values[key] = keep.join(' | ');
                });
                /* The assistant answered a LATER sheet-style line and left a
                 * destination empty (e.g. no city given): the local engine's
                 * correct placement for that slot is NOT thrown away into
                 * « Commentaire » — it is kept. A value the assistant
                 * explicitly moved elsewhere (it is covered by its answer)
                 * is not rescued: its decision wins. */
                SEC.TARGET_KEYS.forEach(function (key) {
                  if (values[key]) return;
                  var local = N.trimClean(localBefore[key]);
                  if (!local || local.indexOf('|') !== -1) return;
                  /* Only a value the assistant EXPLICITLY placed in another
                   * real column beats the local reading. A value that merely
                   * fell into « Commentaire » as an unplaced leftover is
                   * rescued back to its correct column. */
                  var covered = SEC.TARGET_KEYS.some(function (k2) {
                    if (k2 === 'commentaire') return false;
                    return typeof values[k2] === 'string' && values[k2] &&
                      N.trimClean(values[k2]).indexOf(local) !== -1;
                  });
                  if (covered) return;
                  if (!slotAcceptsValue(key, local)) return;
                  values[key] = local;
                  /* the same text was appended to Commentaire as a leftover:
                   * remove the duplicate (its column now holds it). */
                  if (values.commentaire) {
                    values.commentaire = values.commentaire
                      .split(/\s*\|\s*/)
                      .filter(function (p) { return !sameLoose(N.trimClean(p), local); })
                      .join(' | ');
                  }
                });
                aiPlaced = true;
                /* The « À valider » items are built AFTER the type checks
                 * (below): the row is first cleaned, so a value the app has
                 * already placed in its correct column is never proposed for
                 * validation as well (that flooded the table and blocked the
                 * export for hundreds of perfectly placed rows). */
                aiPiecesPending = { row: sourceRowNumAi };
              }
            }
          }

          /* Ville: validated against the official commune database, using the
           * postal code of the same row. Not identified -> left EMPTY. */
          if (villeCol && values.ville) {
            var city = resolveCity(values.ville, values.cp, entry, sheet.name, villeCol.col.header, analysis.headerRowIndex + i + 2);
            values.ville = city.value;
            /* An unidentified city is left empty and reported, but the row is
             * still added successfully (nothing is lost or misplaced). */
          }

          /* 0. Destination columns keep only their own type — whatever
           * wrote the value (local engine or the assistant). */
          enforceDestinationTypes(values);

          /* What the assistant could not identify is proposed for validation
           * ONLY now, from the CLEANED row: a piece that already sits in its
           * right column (phone, mutuelle, cotisation…) is not uncertain, it
           * is placed — proposing it again blocked the export for nothing. */
          if (aiPiecesPending && values.commentaire) {
            var aiRowNum = aiPiecesPending.row;
            var aiRecKey = entry.id + '|' + sheet.name + '|' + aiRowNum;
            var placedPieces = [];
            SEC.TARGET_KEYS.forEach(function (k3) {
              if (k3 === 'commentaire' || !values[k3]) return;
              String(values[k3]).split(/\s*\|\s*/).forEach(function (p) {
                var t = N.trimClean(p);
                if (t) placedPieces.push(t);
              });
            });
            values.commentaire.split(/\s*\|\s*/).forEach(function (piece, pi) {
              var val = N.trimClean(piece);
              if (!val || (N.isPlaceholder && N.isPlaceholder(val))) return;
              /* Already in its column (same digits, accents and separators
               * ignored): nothing to decide. */
              if (placedPieces.some(function (pv) { return sameLoose(pv, val); })) return;
              var guess = suggestDestination(val, values);
              if (guess === 'commentaire') return;
              /* A phone never contains a word: a note such as
               * « Rappel 30/04/2024 » holds 10 digits but is not a phone —
               * proposing « Téléphone » for it flooded the review table and
               * blocked the export for nothing. */
              if ((guess === 'tel1' || guess === 'tel2') && /[a-zA-Z\u00c0-\u00ff]/.test(val)) return;
              /* The suggested column is already filled by another value: not
               * a decision the user can make either. */
              if (values[guess]) return;
              addManualItem({
                recordKey: aiRecKey, field: 'ai' + pi,
                fileName: entry.name, sheet: sheet.name, row: aiRowNum,
                sourceHeader: '(classe par l IA dans Commentaire)',
                value: val, suggestion: guess,
                reason: 'l IA n a pas pu identifier cette valeur avec certitude'
              });
            });
          }

          /* 1. Nom / Prénom hygiene — a person name NEVER contains a note.
           * Any value with a date, an age, a time or a workflow note
           * ("Rappel mardi 30/04/2024 à 17h30 | 87 ans") is moved to
           * « Commentaire » and never lost. The check is on the CONTENT, not on
           * the column: a scrambled file cannot push a note into a name. */
          var toComment = [];
          ['nom', 'prenom'].forEach(function (slot) {
            if (!values[slot]) return;
            var parts = values[slot].split(/\s*\|\s*/).filter(Boolean);
            var keep = [];
            parts.forEach(function (p) {
              if (N.isNoteLike(p) || N.isInvalidPersonName(p)) {
                toComment.push(N.trimClean(p));
              } else {
                keep.push(N.trimClean(p));
              }
            });
            values[slot] = keep.join(' ');
          });
          if (toComment.length) {
            values.commentaire = (toComment.join(' | ') +
              (values.commentaire ? ' | ' + values.commentaire : '')).trim();
          }


          /* 1bis. SPLIT MERGED VALUES (two source columns landed in the same
           * destination): "73+69 | swisselife" -> "73+69" stays in
           * Commentaire, "swisselife" moves to Mutuelle. Every segment is
           * classified by its real nature; unrecognized text stays where it
           * is. This is what makes "64 | assurema" and "61+72 | axa"
           * correct without any manual step. */
          (function splitMergedValues() {
            /* Gemini already placed every part of the line: re-splitting its
             * answer with local rules would re-introduce the mis-placements
             * (an age offered as a price, an insurer in Cotisation). */
            if (aiPlaced) return;
            SEC.TARGET_KEYS.forEach(function (slot) {
              var raw = values[slot];
              /* Only short merged values are re-classified: a long text is a
               * note and stays in its column untouched (it may legitimately
               * contain " | "). */
              if (!raw || raw.indexOf('|') === -1 || raw.length > 300) return;
              var parts = raw.split(/\s*\|\s*/).filter(function (p) { return N.trimClean(p) !== ''; });
              if (parts.length < 2) return;
              var keep = [];
              var moved = false;
              parts.forEach(function (p) {
                var t = N.trimClean(p);
                var dest = null;
                if (N.looksLikeInsurer(t) && !N.looksLikeAmount(t)) { dest = 'mutuelle'; t = N.matchInsurer(t) || t; }
                else if (N.looksLikeAge(t) && !N.looksLikeAmount(t) && !/^\d{3,4}[.,]\d/.test(t)) dest = 'commentaire';
                else if (N.looksLikeEmail(t)) dest = 'email';
                else if (N.looksLikeAmount(t)) dest = 'cotisation';
                else if (N.looksLikePhone(t) && !N.isUnusablePhone(t)) dest = values.tel1 ? (values.tel2 ? null : 'tel2') : 'tel1';
                if (dest && dest !== slot) {
                  if (values[dest]) {
                    if (values[dest].indexOf(t) === -1) values[dest] += ' | ' + t;
                  } else {
                    values[dest] = t;
                  }
                  moved = true;
                } else {
                  keep.push(t);
                }
              });
              if (moved) values[slot] = keep.join(' | ');
            });
          })();

          /* 2. Recover client name, amount or e-mail from a PACKED cell
           * ("Jean-Pierre Andrée | 110 €") — but ONLY when the text itself
           * carries that type of value. A column the user (or the local engine)
           * deliberately mapped to « Commentaire » keeps EVERYTHING it holds:
           * moving a name out of it was silent data loss. */
          if (!aiPlaced && values.commentaire) {
            /* A cell that is EXCLUSIVELY an amount ("1 002,84 euro", "361 euro")
             * in a Commentaire column IS a cotisation: these files carry
             * thousands of them, and leaving them in Commentaire lost the price.
             * The whole value is moved (no reference, no note involved). */
            if (!values.cotisation && N.isPureAmount && N.isPureAmount(values.commentaire)) {
              values.cotisation = values.commentaire;
              values.commentaire = '';
            }
          }
          if (!aiPlaced && values.commentaire && values.commentaire.indexOf('|') !== -1) {
            var cParts = values.commentaire.split(/\s*\|\s*/);
            var keptCommentParts = [];
            cParts.forEach(function (p) {
              var trimmed = N.trimClean(p);
              if (!trimmed) return;
              /* Only a CONFIRMED price (currency / decimals) or a pure amount
               * leaves a note; a bare "70" stays in Commentaire (it may be an
               * age). */
              var confirmed = N.isPureAmount && N.isPureAmount(trimmed) &&
                (N.isConfirmedAmount ? N.isConfirmedAmount(trimmed) : true);
              if (confirmed && !values.cotisation) {
                values.cotisation = trimmed;
              } else if (!values.email && N.looksLikeEmail(trimmed)) {
                values.email = trimmed;
              } else {
                keptCommentParts.push(trimmed);
              }
            });
            values.commentaire = keptCommentParts.join(' | ');
          }
          /* 2bis. PACKED record in a free-text cell ("test happ -- date ----
           * n° sécu -- GROUPAMA -- naissance -- 139,6 -- IBAN"): the price and
           * the insurer are extracted from their own field. The rest of the
           * note is kept, so nothing is lost. */
          if (!aiPlaced && values.commentaire && N.parsePackedRecord) {
            var packed = N.parsePackedRecord(values.commentaire);
            if (packed.amount && !values.cotisation) values.cotisation = packed.amount;
            if (packed.insurer && !values.mutuelle) values.mutuelle = packed.insurer;
          }
          /* A Commentaire cell that is EXCLUSIVELY an insurer name is a
           * Mutuelle, not a note ("C.MUT", "MGEN", "APREVA"). */
          if (!aiPlaced && !values.mutuelle && values.commentaire && N.isPureInsurer
              && N.isPureInsurer(values.commentaire)) {
            values.mutuelle = N.matchInsurer(values.commentaire) || values.commentaire;
            values.commentaire = '';
          }
          /* A Mutuelle cell holding an AMOUNT is not an insurer: it moves to
           * Cotisation (some files mix both in the "Mutuelle" column). */
          if (!aiPlaced && values.mutuelle && N.isPureAmount && N.isPureAmount(values.mutuelle)) {
            if (!values.cotisation) values.cotisation = values.mutuelle;
            values.mutuelle = '';
          }

          /* 3. Inverted field cleanup (e.g. Email in Tel, Mutuelle in Cotisation, Phone/CP merged in CP) */
          if (values.tel1 && values.tel1.indexOf('@') !== -1 && !values.email) {
            values.email = values.tel1;
            values.tel1 = '';
          }
          if (values.tel2 && values.tel2.indexOf('@') !== -1 && !values.email) {
            values.email = values.tel2;
            values.tel2 = '';
          }
          /* AI mode: Gemini already decided Mutuelle vs Cotisation from the
           * value's meaning; the local name test cannot contradict it. */
          if (!aiPlaced && values.cotisation && N.looksLikeMutuelle(values.cotisation) && !N.looksLikeAmount(values.cotisation)) {
            if (values.mutuelle && /^\d+$/.test(values.mutuelle.trim())) {
              values.commentaire = (values.commentaire ? values.commentaire + ' | ' : '') + values.mutuelle.trim() + ' ans';
            }
            values.mutuelle = values.cotisation;
            values.cotisation = '';
          }
          if (values.cp && values.cp.indexOf('|') !== -1) {
            var cpParts = values.cp.split(/\s*\|\s*/);
            var pureCp = '', pureTel = '', extraCp = [];
            cpParts.forEach(function (cpp) {
              var ct = N.trimClean(cpp);
              if (/^\d{5}$/.test(ct) && !pureCp) pureCp = ct;
              else if (N.looksLikePhone(ct) && !pureTel) pureTel = ct;
              else extraCp.push(ct);
            });
            if (pureCp) values.cp = pureCp;
            if (pureTel && !values.tel1) values.tel1 = pureTel;
            if (extraCp.length) {
              values.commentaire = (values.commentaire ? values.commentaire + ' | ' : '') + extraCp.join(' | ');
            }
            /* Phone number or e-mail accidentally left in the Code postal
             * column: move it to the right destination so nothing is lost. */
            if (values.cp && !/^\d{5}$/.test(N.trimClean(values.cp))) {
              var cpTrimmed = N.trimClean(values.cp);
              if (N.looksLikePhone(cpTrimmed)) {
                if (!values.tel1) { values.tel1 = cpTrimmed; values.cp = ''; }
                else if (!values.tel2) { values.tel2 = cpTrimmed; values.cp = ''; }
              } else if (N.looksLikeEmail(cpTrimmed)) {
                if (!values.email) { values.email = cpTrimmed; values.cp = ''; }
              }
            }

          }
          /* Détection forcée : une valeur de « Code postal » qui n'est pas un
           * code postal (format invalide) n'y reste pas. Un numéro de téléphone
           * ou un e-mail tombe dans sa vraie colonne ; toute autre valeur
           * invalide est déplacée dans « Commentaire » pour être vérifiée.
           * Elle n'est jamais exportée en l'état sous « Code postal ». */
          if (values.cp && values.cp.indexOf('|') === -1 && !/^\d{5}$/.test(N.trimClean(values.cp))) {
            var cpBad = N.trimClean(values.cp);
            if (N.looksLikePhone(cpBad)) {
              if (!values.tel1) { values.tel1 = cpBad; values.cp = ''; }
              else if (!values.tel2) { values.tel2 = cpBad; values.cp = ''; }
            } else if (N.looksLikeEmail(cpBad)) {
              if (!values.email) { values.email = cpBad; values.cp = ''; }
            } else {
              values.commentaire = (values.commentaire ? values.commentaire + ' | ' : '') + cpBad;
              values.cp = '';
              var cpColHeader = cpColHeaderForSheet;
              addReport({
                fileName: entry.name, sheet: sheet.name, row: analysis.headerRowIndex + i + 2,
                sourceHeader: cpColHeader,
                sample: cpBad,
                reason: 'code postal invalide — déplacé dans « Commentaire »',
                suggestion: 'Verifier et replacer la valeur à la main si nécessaire',
                severity: 'warning'
              });
            }
          }

          /* Téléphone : une cellule contient plusieurs numéros ;
           * le premier va dans Téléphone 1, le second dans Téléphone 2,
           * les troisième et suivants sont déplacés dans « Commentaire ». */
          var telValues = [values.tel1, values.tel2];
          for (var _pi = 0; _pi < telValues.length; _pi++) {
            if (telValues[_pi]) {
              var phones = SEC.Normalizers.splitPhoneValues(telValues[_pi]);
              if (phones.length >= 2) {
                values.tel1 = phones[0];
                values.tel2 = phones[1];
                // Additional numbers (3rd+) go to commentaire
                for (var _ai2 = 2; _ai2 < phones.length; _ai2++) {
                  values.commentaire = (values.commentaire ? values.commentaire + ' | ' : '') + phones[_ai2];
                }
              } else if (phones.length === 1) {
                // Single number: put in tel1 only if empty; never move to tel2
                if (!values.tel1) { values.tel1 = phones[0]; }
              }
            }
          }

          /* FORCE-CORRECT (row level): any value sitting in a destination its
           * real type contradicts is moved to its TRUE column — nothing ever
           * stays misplaced, whatever the source header or the IA said.
           * SAFETY NET (zero data loss): a value is never taken away from the
           * only cell that holds it when its real destination is already
           * filled by an EXPLICIT mapping: it would be silently lost. Such a
           * value stays where it is (review flags it). At row level (no
           * explicit mapping) the recovery below re-places it instead of
           * dropping it.
           *  - an email found in Nom/Ville/Adresse/Mutuelle -> Email
           *  - a phone number found in Nom/Ville/Adresse/Cotisation -> Téléphone
           *  - an amount found in Ville/Email -> Cotisation
           *  - a known insurer name found in Email/Ville/Cotisation -> Mutuelle
           *  - an age / age range ("70+75", "63 ans", "1950-1948") in any
           *    column -> Commentaire (never Cotisation, never Code postal) */
          (function forceCorrectRow() {
            /* Gemini's line-level classification is authoritative in AI mode:
             * these local rules exist for the local engine, not to undo the
             * assistant's answer (they turned "70" into a price again). */
            if (aiPlaced) return;
            var slots = ['nom', 'prenom', 'ville', 'adresse', 'mutuelle', 'commentaire', 'email', 'cotisation', 'tel1', 'tel2', 'cp'];
            function place(dest, t) {
              if (!values[dest]) values[dest] = t;
              else if (values[dest].indexOf(t) === -1) values[dest] = values[dest] + ' | ' + t;
            }
            slots.forEach(function (from) {
              var v = values[from];
              if (!v) return;
              var t = N.trimClean(v);
              /* Only a value that can be written SOMEWHERE without
               * overwriting an explicit value may be moved. */
              var movable = !(strongDest[from] && written[from] && written[from].indexOf(t) !== -1);
              if (from !== 'commentaire' && N.looksLikeAge(t) && !N.looksLikeAmount(t) &&
                  !/^\d{3,4}[.,]\d/.test(t)) {
                if (movable) { place('commentaire', t); values[from] = ''; }
                return;
              }
              if ((from === 'cotisation' || from === 'adresse' || from === 'nom' || from === 'prenom') &&
                  N.looksLikeInsurer(t) && !N.looksLikeAmount(t)) {
                place('mutuelle', t);
                values[from] = ''; return;
              }
              if (from !== 'email' && from !== 'cotisation' && N.looksLikeEmail(t) && !values.email && movable) {
                values.email = t; values[from] = ''; return;
              }
              if (from !== 'cotisation' && (from === 'nom' || from === 'prenom' || from === 'ville') &&
                  N.looksLikePhone(t)) {
                if (!values.tel1) { values.tel1 = t; values[from] = ''; return; }
                if (!values.tel2) { values.tel2 = t; values[from] = ''; return; }
                /* Both phone slots are taken: keep it where it is rather than
                 * losing it (the review flags the column). */
              }
              if (from !== 'cotisation' && (from === 'ville' || from === 'email') &&
                  N.looksLikeAmount(t) && !values.cotisation && movable) {
                values.cotisation = t; values[from] = ''; return;
              }
              if (from !== 'mutuelle' && (from === 'email' || from === 'ville') &&
                  N.looksLikeMutuelle(t) && !N.looksLikeAmount(t) && !values.mutuelle) {
                values.mutuelle = t; values[from] = ''; return;
              }
              /* A DESTINATION column (a phone, a postal code, an amount…) that
               * received free text ("Appelé", a remark) would export a comment
               * under « Téléphone 2 » and LOSE it. The text moves to
               * « Commentaire »; the value is never discarded. */
              if ((from === 'tel1' || from === 'tel2' || from === 'cp') &&
                  !N.looksLikePhone(t) && !N.looksLikePostal(t) && N.looksLikeCommentText(t)) {
                place('commentaire', t);
                values[from] = '';
              }
            });
          })();

          /* 4. Separate Nom and Prénom.
           * A value holding a surname AND a first name is split when the
           * evidence is clear (known first name, ALL-CAPS surname, "Dupont,
           * Jean"). Otherwise the value stays WHOLE in the column it came
           * from: no guess, no text loss. */
          if (values.nom && values.prenom && !N.isNoteLike(values.nom + ' ' + values.prenom)) {
            /* A single-value column was merged into a pair: keep both words. */
            if (values.nom.split(/\s+/).length > 2) {
              var pairBoth = N.splitNamePair(values.nom);
              if (pairBoth) {
                values.prenom = (values.prenom + ' ' + pairBoth.prenom).trim();
                values.nom = pairBoth.nom;
              }
            }
          } else if (values.nom && !values.prenom) {
            /* Gemini (when it answered) knows which part is the surname; its
             * split is used ONLY when it rebuilds the exact source text. */
            var nameSplit = null;
            var colHeaderNorm = nameColumnNorm || '';
            var aiParts = colHeaderNorm && SEC.AI && SEC.AI.nameParts
              ? SEC.AI.nameParts(colHeaderNorm, entry.name)
              : null;
            if (aiParts && aiParts.nom && aiParts.prenom) {
              var rebuilt = (aiParts.nom + ' ' + aiParts.prenom).replace(/\s+/g, ' ').trim().toLowerCase();
              var source = String(values.nom).replace(/\s+/g, ' ').replace(/\s*,\s*/g, ' ').trim().toLowerCase();
              if (rebuilt === source) {
                nameSplit = { nom: aiParts.nom, prenom: aiParts.prenom };
              }
            }
            if (!nameSplit) {
              nameSplit = N.splitNamePair(values.nom);
              if (!nameSplit && !N.hasLeadingParticle(values.nom)) {
                nameSplit = N.splitFullName(values.nom);
              }
            }
            if (nameSplit && nameSplit.nom && nameSplit.prenom) {
              values.nom = nameSplit.nom;
              values.prenom = nameSplit.prenom;
            }
          } else if (!values.nom && values.prenom) {
            var splitPrenom = N.splitNamePair(values.prenom) ||
              (N.hasLeadingParticle(values.prenom) ? null : N.splitFullName(values.prenom));
            if (splitPrenom && splitPrenom.nom && splitPrenom.prenom) {
              values.nom = splitPrenom.nom;
              values.prenom = splitPrenom.prenom;
            }
          }
          /* Final guard: nothing note-like may survive in the name columns. */
          ['nom', 'prenom'].forEach(function (slot) {
            if (values[slot] && (N.isNoteLike(values[slot]) || N.isInvalidPersonName(values[slot]))) {
              values.commentaire = (N.trimClean(values[slot]) +
                (values.commentaire ? ' | ' + values.commentaire : '')).trim();
              values[slot] = '';
            }
          });

          var mappedAny = SEC.TARGET_KEYS.some(function (t) { return values[t] !== ''; });
          var sourceRowNum = analysis.headerRowIndex + i + 2;
          if (!mappedAny) {
            /* Never discard a data row: put every non-empty source cell into
             * « Commentaire » so the final file always contains all information. */
            var leftovers = [];
            columns.forEach(function (col) {
              var rawLeft = srcRow[col._index];
              if (N.trimClean(rawLeft) === '') return;
              leftovers.push((col.header ? col.header + ' : ' : '') + String(rawLeft).replace(/ /g, ' ').trim());
            });
            if (leftovers.length) {
              values.commentaire = leftovers.join(' | ');
            }
            addReport({
              fileName: entry.name, sheet: sheet.name, row: sourceRowNum,
              sourceHeader: '(ligne complete)', sample: (srcRow.filter(Boolean)[0] || '').slice(0, 120),
              reason: 'aucune colonne reconnue : donnees conservees dans la colonne « Commentaire »',
              suggestion: 'Cartographier les colonnes de ce fichier pour un placement precis', severity: 'warning'
            });
          }
          /* Gemini read this line but did not classify a single value: the
           * line is kept (as above) and the failure is named, never hidden. */
          if (aiSeenEmpty) {
            aiStats.missing++;
            if (aiStats.missing <= 50) {
              addReport({
                fileName: entry.name, sheet: sheet.name, row: sourceRowNum,
                sourceHeader: '(assistant IA)',
                sample: String(srcRow.filter(function (c) { return N.trimClean(c) !== ''; })[0] || '').slice(0, 120),
                reason: 'l IA n a pas pu identifier cette ligne : elle a ete traitee par le moteur local',
                suggestion: 'Verifier le classement de cette ligne dans le fichier final',
                severity: 'warning'
              });
            }
          }

          var key = entry.id + '|' + sheet.name + '|' + sourceRowNum;
          var rec = {
            key: key, fileId: entry.id, fileName: entry.name, sheet: sheet.name,
            sourceRow: sourceRowNum, values: values, hasWarning: problemList.length > 0,
            problems: problemList
          };
          session.records.set(key, rec);
          var byFile = session.recordsByFile.get(entry.id);
          if (!byFile) { byFile = []; session.recordsByFile.set(entry.id, byFile); }
          byFile.push(rec);
          rowCount++;
        }

        /* Calibrate the batch size from what actually happened. */
        var processedCount = Math.max(1, end - start);
        var elapsed = t0 ? Date.now() - t0 : 0;
        if (elapsed > 0) {
          var perRow = elapsed / processedCount;
          rowMs = rowMs ? rowMs * 0.6 + perRow * 0.4 : perRow;
          batchSize = Math.max(25, Math.min(4000, Math.round(BATCH_BUDGET_MS / Math.max(perRow, 0.0005))));
        }
        if (entry.parsedSheets) {
          /* Honest per-file progress (45 % -> 95 %) instead of a frozen 45 %. */
          entry.progress = Math.min(95, 45 + Math.round((end / Math.max(1, dataRows.length)) * 50));
        }
        if (end < dataRows.length) {
          /* Yield to the UI before processing the next batch so the
           * interface stays responsive during large extractions. */
          if (yieldFn) return yieldFn().then(function () { return processBatch(end); });
          return processBatch(end);
        }
        return Promise.resolve();
      };

      return processBatch(0).then(function () {
        return { rows: rowCount, warnings: 0, review: newReview, headerRowIndex: analysis.headerRowIndex };
      });
    }

    /* Deliberate (research-backed) EXCLUSIONS: a column that is only a running
     * id, a pure date, a birth/joining year or an age carries no value in any
     * of the 11 official columns. It used to pour dozens of "59 ans" values
     * into « Commentaire » and push every row to "à vérifier". It is now
     * reported once (never silently dropped) without polluting the data. */
    function excludedByProfile(col) {
      var samples = (col.profile && col.profile.samples) || [];
      if (!samples.length) return false;
      if (looksLikeDateColumn(samples) || looksLikeYearColumn(samples)) return true;
      var usable = samples.filter(function (v) { return N.trimClean(v) !== ''; });
      if (!usable.length) return false;
      var ages = usable.filter(function (v) {
        var s = N.trimClean(v);
        return N.looksLikeAge(s) && !N.looksLikeAmount(s) && !/^\d{3,4}[.,]\d/.test(s);
      }).length;
      return ages / usable.length >= 0.7;
    }

    function processBuffer(entry, buffer, yieldFn, preParsed, onAiProgress) {
      dropRecords(entry.id);
      /* A new buffer (re-read file / chosen sheets) rebuilds the analysis: the
       * destinations are attached to column POSITIONS, so decisions from a
       * previous analysis are dropped. Decisions taken by the user for the
       * CURRENT analysis are kept (they live in the review items). */
      entry.columnOverrides = entry.columnOverrides || {};
      entry.headerOverrides = entry.headerOverrides || {};
      /* Same buffer (retry / re-extraction): the user's destinations are
       * intentionally preserved, they survive retries and rules re-apply. */
      var sheets = parseEntry(entry, buffer, preParsed);
      entry.error = '';
      entry.status = S.PROCESSING;
      entry.progress = 45;
      entry.stage = 'Analyse semantique et cartographie';

      /* Every batch starts a fresh AI run: the counters of a previous batch
       * never describe this one. */
      var AIrun = aiAssistant();
      if (AIrun && typeof AIrun.startRun === 'function' &&
          typeof AIrun.enabled === 'function' && AIrun.enabled()) AIrun.startRun();

      var chain = Promise.resolve();
      var totalRows = 0, totalWarnings = 0;
      var aiStats = { rows: 0, missing: 0, truncated: 0, empty: 0 };
      /* A new run rebuilds its own review items: stale ones (from another
       * header row or from a previous state) must not survive. */
      Array.from(session.reviewItems.keys()).forEach(function (id) {
        if (id.indexOf(entry.id + '|') === 0) session.reviewItems.delete(id);
      });
      (entry.selectedSheets || []).forEach(function (sheetName) {
        var sheet = sheets.find(function (s) { return s.name === sheetName; });
        if (!sheet || sheet.isEmpty) return;
        chain = chain.then(function () {
          /* Announce the sheet BEFORE yielding, so the paint the yield
           * triggers already shows it during the CPU-heavy analysis. */
          entry.stage = 'Extraction : ' + sheetName;
          /* Analysis of a big worksheet is itself CPU-heavy: paint first. */
          return (yieldFn || nextTick)().then(function () {
          /* Optional Gemini consultation BEFORE extraction: only the columns
           * the local agent could not resolve confidently are sent (headers +
           * a few samples). Any failure leaves the local mapping untouched. */
          var preAnalysis = SEC.Agent.analyzeSheet(sheet.rows,
            typeof entry.headerOverrides[sheet.name] === 'number' ? entry.headerOverrides[sheet.name] : undefined);
          /* Actually remove the deliberate exclusions from the pipeline. */
          preAnalysis.columns.forEach(function (col) {
            if (col.outOfScope && !col.blocked) {
              col.outOfScopePending = true;
              return;
            }
            if (!col.best && excludedByProfile(col) && !col.blocked) {
              col.profileExcludedPending = true;
              col.reasons.push('colonne hors perimetre (identifiant/dates/annees/ages) : exclue, signalee dans le rapport');
            }
          });
          var ai = aiAssistant();
          /* Any AI call is a promise to the user: a worker thread (and a
           * worker client) must exist before the optional consultation starts.
           * fetch is unavailable inside a Worker, so the assistants must be in
           * place here. */
          if (ai && ai.enabled() && SEC.WorkerClient && typeof SEC.WorkerClient.mode === 'function') {
            try { SEC.WorkerClient.prewarm(); } catch (e) { /* inline fallback */ }
          }
          /* The backend may be down: a failed probe must never block. */
          var ready = ai && ai.enabled() && ai.status().checked === false && ai.refreshStatus
            ? ai.refreshStatus().catch(function () { return null; })
            : Promise.resolve(null);
          /* Gemini is a fast accelerator, never a blocker: if the backend or
           * the network stalls past this deadline, the local engine carries on
           * (the late answer is simply ignored) and the loop measures it. */
          var AI_DEADLINE_MS = 20000;
          var askAi = function (cols) {
            if (!ai || !ai.enabled()) return Promise.resolve(null);
            return new Promise(function (resolve) {
              var timer = setTimeout(function () {
                resolve({ __unavailable: true });
              }, AI_DEADLINE_MS);
              ai.suggestSheet(entry.name, sheetName, cols).then(function (hints) {
                clearTimeout(timer);
                resolve(hints);
              }, function () {
                clearTimeout(timer);
                resolve(null); // AI must never break a batch
              });
            });
          };
          /* FULL-SHEET mode: Gemini reads the REAL text of every row of this
           * sheet (not the header and a few samples). Enabled whenever the
           * assistant is ticked and the client provides mapSheet. A failure is
           * returned as __error and reported — never silently ignored. */
          /* Ticked = the user wants Gemini to read the lines. `enabled()` also
           * requires a working backend; when it is not usable the call below
           * returns its explicit reason and the report says so. */
          var aiState = ai && typeof ai.status === 'function' ? ai.status() : null;
          var aiTicked = !!(ai && (aiState && aiState.ticked !== undefined ? aiState.ticked : ai.enabled()));
          var fullScan = !!(ai && aiTicked && typeof ai.mapSheet === 'function');
          var aiData = null;
          var scanPromise = fullScan
            ? new Promise(function (resolve) {
                entry.stage = "Analyse par l'IA de toutes les lignes : " + sheetName;
                entry.aiTotal = sheet.rows.length - preAnalysis.headerRowIndex - 1;
                entry.aiDone = 0;
                if (yieldFn) yieldFn();
                ai.mapSheet(entry.name, sheetName,
                  (sheet.rows[preAnalysis.headerRowIndex] || []),
                  sheet.rows.slice(preAnalysis.headerRowIndex + 1),
                  preAnalysis.headerRowIndex + 2,
                  /* Live progress: rows the provider has answered so far. */
                  function (doneRows, totalRows) {
                    entry.aiDone = doneRows;
                    if (totalRows) entry.aiTotal = totalRows;
                    entry.stage = 'Analyse IA : ' + doneRows + ' / ' + totalRows + ' lignes (' + sheetName + ')';
                    var pct = totalRows ? Math.round((doneRows / totalRows) * 100) : 0;
                    entry.progress = Math.min(95, 45 + Math.round(pct * 0.5));
                    if (onAiProgress) onAiProgress(entry, doneRows, totalRows);
                  }
                ).then(resolve, function (err) {
                  resolve({ __error: true, code: 'network', message: err && err.message });
                });
              })
            : Promise.resolve(null);
          return ready.then(function () {
            return askAi(preAnalysis.columns);
          }).then(function (hints) {
            /* Some columns were not PROFILED by the local engine (no values,
             * hence no guess at all). They carry real samples: give Gemini a
             * second, explicit chance on them instead of dropping their data. */
            var blind = preAnalysis.columns.filter(function (col) {
              return !col.best && !col.profileExcluded && !col.blocked && !col.fullName &&
                !col.excluded && !col.learned && !col.mergeInto &&
                col.profile && col.profile.samples && col.profile.samples.length;
            });
            if (!blind.length || !ai || !ai.enabled()) return hints;
            return askAi(blind).then(function (extra) {
              if (!extra) return hints;
              if (extra.__unavailable && !hints) return extra;
              var merged = Object.assign({}, hints || {});
              Object.keys(extra).forEach(function (k) { merged[k] = extra[k]; });
              return merged;
            });
          }).then(function (hints) {
            /* Every call (even an empty/cached answer) counts as an AI
             * request so the report always states whether it took part. */
            if (hints) {
              /* Only real destinations count: __unavailable is not a hint. */
              var count = Object.keys(hints).filter(function (k) {
                return typeof hints[k] === 'string' && SEC.TARGET_LABELS[hints[k]];
              }).length;
              if (count) {
                session.aiRequests = (session.aiRequests || 0) + 1;
                session.aiHints = (session.aiHints || 0) + count;
              }
            }
            /* ONE analysis per sheet: the hints are applied to exactly the
             * columns extractSheet() will use (no second pass can diverge).
             * The full-sheet answer (when it came) is applied row by row. */
              return scanPromise.then(function (scan) {
                if (scan && scan.__error && fullScan) {
                  /* Gemini did not answer this sheet at all. The reason is
                   * shown in plain words (no raw Google JSON) and the count of
                   * lines is exact, so nothing is hidden. */
                  addReport({
                    fileName: entry.name, sheet: sheetName, sourceHeader: '(assistant IA)',
                    reason: 'assistant IA indisponible pour cette feuille : ' + aiReasonText(scan) +
                      ' — ' + (scan.failed || sheet.rows.length) +
                      ' ligne(s) traitee(s) par le moteur local',
                    suggestion: aiFixHint(scan.code),
                    severity: 'warning'
                  });
                } else if (scan && scan.rows) {
                  aiData = {};
                  scan.rows.forEach(function (r) { aiData[r.index] = r.values || {}; });
                  aiStats.rows += scan.placed || 0;
                  if (scan.missing) aiStats.missing += scan.missing;
                  if (scan.truncated) aiStats.truncated = (aiStats.truncated || 0) + scan.truncated;
                  if (scan.error) {
                    addReport({
                      fileName: entry.name, sheet: sheetName, sourceHeader: '(assistant IA)',
                      reason: 'assistant IA incomplet : ' + aiReasonText(scan.error) +
                        ' — ' + (scan.missing || 0) + ' ligne(s) non classée(s) par l IA (traitées localement)',
                      suggestion: aiFixHint(scan.error.code),
                      severity: 'warning'
                    });
                  }
                }
                return extractSheet(entry, sheet, yieldFn, preAnalysis.columns, hints || null, aiData).then(function (r) {
                  totalRows += r.rows;
                  totalWarnings += r.warnings;
                });
              });
            });
          });
        });
      });

      return chain.then(function () {
        entry.rowsExtracted = totalRows;
        entry.warnings = totalWarnings;
        entry.progress = 100;
        /* A file whose rows DID land in the final file is a SUCCESS, even when
         * some values were flagged for checking. "Completed with warnings" is
         * no longer a separate end state: warnings live on the lines and in
         * the report, while the file itself counts as succeeded. */
        entry.status = S.COMPLETED;
        entry.stage = 'Termine';
        entry.parsedSheets = null; // release heavy arrays; re-read from File on retry if needed
        /* Keep at most ONE raw buffer alive: several large workbooks would
         * otherwise all sit in memory at once (freeze, corrupted last file).
         * A reprocess re-reads the File object via readBuffer instead. */
        session.files.forEach(function (e2) {
          if (e2 !== entry) e2._lastBuffer = null;
        });
        entry._lastBuffer = buffer; // in-session retry without re-reading
        /* Full-sheet statistics: how many lines Gemini really classified, how
         * many it could not (kept by the local engine) and how many answers
         * were truncated and retried. Shown in the report, never hidden. */
        entry.aiRows = aiStats.rows;
        entry.aiMissing = aiStats.missing;
        entry.aiTruncated = aiStats.truncated;
        session.aiRows = (session.aiRows || 0) + aiStats.rows;
        session.aiMissingRows = (session.aiMissingRows || 0) + aiStats.missing;
        session.aiTruncated = (session.aiTruncated || 0) + aiStats.truncated;
        /* The optional assistant is recorded in the report the moment it took
         * part (a wrong key or an exhausted quota is never silent, even when
         * the rest of the batch runs locally). */
        recordAiOutcome(session);
        /* Sync the per-file counters with the records that were just written. */
        recomputeStats();
        return entry;
      });
    }

    /* ---------------- Batch orchestration ---------------- */

    function startBatch(entries, opts) {
      opts = opts || {};
      if (session.running) return Promise.reject(new Error('Un traitement est deja en cours'));
      session.running = true;
      session.cancelling = false;
      currentEntries = entries;
      var yieldFn = opts.yieldFn || nextTick;
      var onProgress = opts.onProgress || function () {};
      var readBuffer = opts.readBuffer; // fn(entry) -> Promise<ArrayBuffer>
      /* fn(entry, buffer) -> Promise<{sheets,encoding,delimiter}>: parsing off
       * the UI thread (Web Worker). Optional: falls back to in-thread parsing. */
      var parseFn = opts.parseFn;

      var i = 0;
      currentIndex = 0;
      var lastProgressTime = 0;
      var PROGRESS_THROTTLE_MS = 300; // don't refresh UI more than ~3 times/sec during batch

      /* One file ahead, at most: bounded memory, no unbounded queue. */
      function prefetchNext() {
        if (!readBuffer || prefetchFor || !entries[i]) return;
        var next = entries[i];
        if (next.duplicateOf || next._lastBuffer || next._bufferPromise) return;
        prefetchFor = next.id;
        try {
          next._bufferPromise = Promise.resolve(readBuffer(next)).catch(function (err) {
            next._bufferPromise = null;
            throw err;
          });
        } catch (e) {
          prefetchFor = null;
        }
      }

      function maybeProgress(entry) {
        var now = Date.now();
        var isLast = i >= entries.length;
        if (isLast || now - lastProgressTime >= PROGRESS_THROTTLE_MS) {
          lastProgressTime = now;
          onProgress(entry);
          recomputeStats();
        }
      }

      var prefetchFor = null;

      function runNext() {
        if (prefetchFor && (!entries[i] || entries[i].id !== prefetchFor)) prefetchFor = null;
        if (session.cancelling) {
          entries.slice(i).forEach(function (e) {
            if (e.status === S.WAITING) { e.status = S.CANCELLED; e.stage = 'Annule'; }
          });
          session.stats.cancelled++;
          session.running = false;
          return Promise.resolve();
        }
        if (i >= entries.length) { session.running = false; return Promise.resolve(); }
        var entry = entries[i++];
        currentIndex = i;
        if (entry.duplicateOf) {
          entry.status = S.CANCELLED;
          entry.stage = 'Doublon ignore';
          addReport({
            fileName: entry.name, sourceHeader: '',
            reason: 'fichier duplique (meme nom et taille) ignore',
            suggestion: 'Supprimer le doublon si necessaire', severity: 'info'
          });
          maybeProgress(entry);
          return runNext();
        }
        return Promise.resolve()
          .then(function () {
            /* The bytes of the NEXT file are read while this one is being
             * processed: reading (a disk / memory copy) is no longer a
             * serial wait between two files of a big batch. */
            prefetchNext();
            if (entry._bufferPromise) {
              var pending = entry._bufferPromise;
              entry._bufferPromise = null;
              prefetchFor = null; // consumed: the next file may be prefetched
              return pending;
            }
            if (entry._lastBuffer) return entry._lastBuffer;
            if (readBuffer) return readBuffer(entry);
            throw new Error('Aucun lecteur de fichier fourni');
          })
          .then(function (buffer) {
            var parsed = parseFn ? parseFn(entry, buffer).catch(function () { return null; }) : Promise.resolve(null);
            return parsed.then(function (preParsed) {
              return processBuffer(entry, buffer, yieldFn, preParsed || null, opts.onAiProgress);
            });
          })
          .catch(function (err) {
            entry.status = S.FAILED;
            entry.error = err && err.message ? err.message : String(err);
            entry.stage = 'Echec';
            entry.progress = 100;
            addReport({
              fileName: entry.name, sourceHeader: '', reason: 'echoc du fichier : ' + entry.error,
              suggestion: 'Verifier que le fichier n est pas corrompu puis reessayer', severity: 'error'
            });
          })
          .then(function () {
            maybeProgress(entry);
            return runNext();
          });
      }
      return runNext();
    }

    /* The files of the CURRENT batch are remembered so cancel() can stop the
     * file in progress even when its read/parse promise never settles. */
    var currentEntries = [];
    var currentIndex = 0;

    function cancel() {
      if (!session.running || session.cancelling) return;
      session.cancelling = true;
      var entries = currentEntries;
      var i = currentIndex;
      /* The file IN PROGRESS must stop now, not after it finishes: a pending
       * readBuffer / parse promise that never resolves (unreadable file, slow
       * disk, AI deadline) would otherwise keep `running` true forever and the
       * Cancel button dead. Whatever stage the current file is in is marked
       * cancelled here; runNext() honours the flag as soon as a promise
       * settles. */
      entries.forEach(function (e) {
        if (e.status === S.WAITING) { e.status = S.CANCELLED; e.stage = 'Annule'; }
      });
      var current = entries[i] || (entries[i - 1] && entries[i - 1].status === S.PROCESSING ? entries[i - 1] : null);
      if (current && (current.status === S.PROCESSING || current.status === S.VALIDATING)) {
        current.status = S.CANCELLED;
        current.stage = 'Annule';
        current.progress = 100;
        dropRecords(current.id);
        addReport({
          fileName: current.name, sourceHeader: '',
          reason: 'traitement annule par l utilisateur : fichier stoppe en cours de lecture',
          suggestion: 'Relancer « Démarrer le traitement » pour reprendre ce fichier', severity: 'info'
        });
      }
      session.running = false;
      session.stats.cancelled++;
    }

    function retryFile(entry, readBuffer) {
      entry.status = S.WAITING;
      entry.error = '';
      entry.progress = 0;
      return startBatch([entry], { readBuffer: readBuffer, yieldFn: nextTick });
    }

    /* ---------------- Review decisions ---------------- */

    /* The decision applies to THIS column of THIS file only: nothing is
     * reused on other headers (no « entêtes similaires » step), so the
     * chosen destination is respected exactly. */
    function decideReview(itemId, decision) {
      var item = session.reviewItems.get(itemId);
      var parts = itemId.split('|');
      var fileId = item ? item.fileId : parts[0];
      var sheet = item ? item.sheet : parts[1];
      var colIndex = item ? item.colIndex : parseInt(parts[2], 10);
      var entry = session.files.get(fileId);
      if (!entry) return;

      entry.columnOverrides[sheet] = entry.columnOverrides[sheet] || {};
      entry.columnOverrides[sheet][colIndex] = decision;

      if (!item) {
        item = {
          id: itemId, fileId: fileId, fileName: entry.name, sheet: sheet,
          colIndex: colIndex, status: decision === '__exclude__' ? 'excluded' : 'changed',
          decision: decision, suggestion: null, confidence: 100
        };
      } else {
        item.decision = decision; // target key or '__exclude__'
        item.status = decision === '__exclude__' ? 'excluded'
          : (decision === item.suggestion ? 'confirmed' : 'changed');
      }
      session.reviewItems.set(itemId, item);
    }

    /* Automatic mode: no manual validation step. Every column the engine
     * flagged as uncertain is resolved with autoFixReview's decision (clear
     * suggestion, value profile, exclusion, or « Commentaire » so nothing is
     * lost) and the destination is applied to the file like a user choice. */
    function resolveColumn(entry, sheet, col) {
      var itemId = entry.id + '|' + sheet.name + '|' + col._index;
      var item = {
        id: itemId, fileId: entry.id, fileName: entry.name, sheet: sheet.name,
        colIndex: col._index, headerNorm: col.headerNorm, status: 'open', decision: null,
        header: col.header, samples: col.profile.samples,
        confidence: col.score, validation: col.validation,
        reasons: (col.reasons || []).slice(), suggestion: col.best || null,
        blocked: col.blocked, fullName: col.fullName
      };
      session.reviewItems.set(itemId, item);
      return item;
    }

    function openReviewCount() {
      var n = 0;
      session.reviewItems.forEach(function (it) { if (it.status === 'open') n++; });
      return n;
    }

    /* ---------------- Manual validation (« À valider ») ----------------
     * A value the provider did not classify stays OUT of the 11 columns and
     * lands here instead. The user picks its destination; nothing is ever
     * guessed in its place. */

    function addManualItem(item) {
      var key = item.recordKey + '|' + item.field;
      if (session.manualIndex.has(key)) return session.manualIndex.get(key);
      var entry = {
        id: 'm' + (session.manualItems.length + 1),
        recordKey: item.recordKey,
        field: item.field,
        fileName: item.fileName || '',
        sheet: item.sheet || '',
        row: item.row || '',
        sourceHeader: item.sourceHeader || '',
        value: String(item.value == null ? '' : item.value).slice(0, 200),
        suggestion: item.suggestion || 'commentaire',
        destination: item.suggestion || 'commentaire',
        status: 'open',
        reason: item.reason || 'non identifie avec certitude par l IA',
      };
      session.manualItems.push(entry);
      session.manualIndex.set(key, entry);
      return entry;
    }

    /* Number of values still waiting for a decision (blocks the export). */
    function openManualCount() {
      var n = 0;
      session.manualItems.forEach(function (it) { if (it.status === 'open') n++; });
      return n;
    }

    /* Apply the validated destinations: the value leaves « Commentaire » and
     * lands in the chosen column of ITS OWN record. */
    function applyManualDecision(id, destination) {
      var item = null;
      session.manualItems.forEach(function (it) { if (it.id === id) item = it; });
      if (!item) return null;
      if (!destination) return item;
      /* Changing a decision on an applied item: revert first. */
      var rec = session.records.get(item.recordKey);
      if (rec && item.applied) {
        var cur = rec.values[item.appliedTo] || '';
        var parts = cur.split(/\s*\|\s*/).filter(function (p) { return p !== item.value; });
        rec.values[item.appliedTo] = parts.join(' | ');
        if (item.appliedTo !== 'commentaire'
            && (!rec.values.commentaire || rec.values.commentaire.indexOf(item.value) === -1)) {
          rec.values.commentaire = rec.values.commentaire
            ? rec.values.commentaire + ' | ' + item.value
            : item.value;
        }
        item.applied = false;
      }
      item.destination = destination;
      item.status = destination === item.suggestion ? 'confirmed' : 'changed';
      if (rec) {
        var v = rec.values[destination] || '';
        if (destination === 'commentaire') {
          var rest = v.split(/\s*\|\s*/).filter(function (p) { return p !== item.value; });
          rec.values.commentaire = rest.join(' | ');
        } else if (v.indexOf(item.value) === -1) {
          rec.values[destination] = v ? v + ' | ' + item.value : item.value;
          var c = (rec.values.commentaire || '').split(/\s*\|\s*/)
            .filter(function (p) { return p !== item.value; });
          rec.values.commentaire = c.join(' | ');
        }
        item.recordKey = item.recordKey;
        item.applied = true;
        item.appliedTo = destination;
      }
      return item;
    }

    function acceptAllManual() {
      session.manualItems.forEach(function (it) {
        if (it.status === 'open') applyManualDecision(it.id, it.suggestion || 'commentaire');
      });
      return openManualCount();
    }

    function sendAllManualToComment() {
      session.manualItems.forEach(function (it) {
        applyManualDecision(it.id, 'commentaire');
      });
      return openManualCount();
    }

    /* Decisions are taken before this call (decideReview), then the files are
     * re-extracted with their destinations applied. A decision can reveal a
     * new uncertain column of the file (the one it used to mask): those are
     * resolved by autoFixReview and the file is processed once more. */
    function reprocessAffected(affectedIds, readBuffer, yieldFn, autoFixFn) {
      if (session.running) return Promise.reject(new Error('Un traitement est deja en cours'));
      session.running = true;
      var ids = affectedIds || session.fileOrder.slice();
      return processIds(ids, readBuffer, yieldFn).then(function () {
        if (!autoFixFn) return null;
        var open = 0;
        ids.forEach(function (id) {
          session.reviewItems.forEach(function (it) {
            if (it.fileId === id && it.status === 'open') open++;
          });
        });
        if (!open) return null;
        var fixed = autoFixFn();
        if (!fixed || !fixed.fixed) return null;
        return processIds(ids, readBuffer, yieldFn);
      }).then(function () {
        recomputeStats();
        session.running = false;
        return session;
      }, function (err) {
        session.running = false;
        throw err;
      });
    }

    function processIds(ids, readBuffer, yieldFn) {
      var chain = Promise.resolve();
      ids.forEach(function (id) {
        var entry = session.files.get(id);
        if (!entry) return;
        chain = chain.then(function () {
          if (entry._lastBuffer) return processBuffer(entry, entry._lastBuffer, yieldFn || nextTick);
          if (readBuffer) return readBuffer(entry).then(function (buf) { return processBuffer(entry, buf, yieldFn || nextTick); });
          return null;
        });
      });
      return chain;
    }

    /* ---------------- Output helpers ---------------- */

    /* Final review before output/export: every record is re-checked as a
     * whole, and only PROVEN misplacements are corrected — always by MOVING
     * the text to « Commentaire », never deleting it. Idempotent, so the
     * preview and the export see exactly the same values. */
    function reviewRecords() {
      var moved = 0;
      session.records.forEach(function (r) {
        var v = r.values;
        var push = function (text) {
          var t = N.trimClean(text);
          if (!t) return;
          v.commentaire = v.commentaire ? (t + ' | ' + v.commentaire) : t;
        };
        /* 1. A note is never a name. */
        ['nom', 'prenom'].forEach(function (slot) {
          if (!v[slot]) return;
          var parts = v[slot].split(/\s*\|\s*/).filter(Boolean);
          var keep = [];
          parts.forEach(function (p) {
            if (N.isNoteLike(p) || N.isInvalidPersonName(p)) { push(p); moved++; }
            else keep.push(N.trimClean(p));
          });
          v[slot] = keep.join(' ');
        });
        /* 2. Nom and Prénom both filled by ONE full-name value: split it. */
        if (v.nom && !v.prenom) {
          var pair = N.splitNamePair(v.nom);
          if (pair && pair.nom && pair.prenom) { v.nom = pair.nom; v.prenom = pair.prenom; }
        }
        /* 3. A phone / e-mail / postal code inside a name column is moved to
         * the right column (or to Commentaire when that column is taken). */
        ['nom', 'prenom'].forEach(function (slot) {
          if (!v[slot]) return;
          var value = v[slot];
          var dest = null;
          if (N.looksLikeEmail(value)) dest = 'email';
          else if (N.looksLikePhone(value) && !N.isUnusablePhone(value)) dest = v.tel1 ? 'tel2' : 'tel1';
          else if (N.looksLikePostal(value)) dest = 'cp';
          if (!dest) return;
          if (!v[dest]) v[dest] = value;
          else push(value);
          v[slot] = '';
          moved++;
        });
        /* 4. A person name inside « Ville » or « Adresse » is a person name. */
        ['ville', 'adresse'].forEach(function (slot) {
          if (!v[slot] || !N.looksLikePersonName(v[slot]) || N.isNoteLike(v[slot])) return;
          if (!v.nom && !v.prenom) {
            var split = N.splitNamePair(v[slot]);
            v.nom = split && split.nom ? split.nom : v[slot];
            v.prenom = split && split.prenom ? split.prenom : '';
            v[slot] = '';
            moved++;
          } else {
            push(v[slot]);
            v[slot] = '';
            moved++;
          }
        });
      });
      return moved;
    }

    /* Consolidated rows in file order, then source-row order. Each row is the
     * exact content of its destination columns: nothing is cleared, moved or
     * reordered at output time. maxRows caps output for preview rendering. */
    var SAFE_CELL = 32767;
    var rowCountForContinuations = 0;

    function buildConsolidatedRows(maxRows) {
      reviewRecords(); // final check, applied in place (idempotent)
      var rows = [];
      rowCountForContinuations = 0;
      var remaining = maxRows == null ? -1 : maxRows;
      session.fileOrder.forEach(function (fid) {
        var recs = session.recordsByFile.get(fid);
        if (!recs || !recs.length) return;
        recs.sort(function (a, b) {
          if (a.sheet !== b.sheet) return a.sheet < b.sheet ? -1 : 1;
          return a.sourceRow - b.sourceRow;
        });
        recs.forEach(function (r) {
          /* Excel cannot hold more than 32 767 characters in one cell. A
           * longer text is CUT HERE and its continuation is exposed separately
           * (buildContinuations), so the main data rows stay exactly one per
           * record: the preview, the export and the end-to-end verification
           * keep comparing record for record, without shifts. */
          var parts = {};
          var row = SEC.TARGET_KEYS.map(function (t) {
            var v = r.values[t];
            var s = v == null ? '' : String(v);
            if (s.length <= SAFE_CELL) return s;
            parts[t] = s.slice(SAFE_CELL);
            return s.slice(0, SAFE_CELL);
          });
          if (parts && Object.keys(parts).length) {
            r.longParts = parts;
            if (r.values) {
              Object.keys(parts).forEach(function (t) {
                r.values[t] = String(r.values[t]).slice(0, SAFE_CELL);
              });
            }
            if (rowCountForContinuations === 0) {
              addReport({
                fileName: r.fileName, sheet: r.sheet, row: r.sourceRow,
                sourceHeader: Object.keys(parts).join(', '),
                sample: String(Object.values(parts)[0]).slice(0, 80),
                reason: 'texte plus long que la limite Excel de ' + SAFE_CELL +
                  ' caracteres : la suite est ecrite dans la feuille « Suite des textes longs » (aucune perte)',
                suggestion: 'Aucune action requise', severity: 'info'
              });
            }
            rowCountForContinuations++;
          }
          if (remaining === 0) return;
          rows.push(row);
          if (remaining > 0) remaining--;
        });
      });
      return rows;
    }

    /* ---------------- Smart auto-fix ----------------
     * One click resolves every open review item intelligently:
     *  - ID/reference columns  -> excluded (can never be phone/postal/amount)
     *  - "Nom complet" columns -> kept whole in « Nom » (never split)
     *  - clear suggestion      -> accepted as-is (data lands in the right column)
     *  - no suggestion         -> value-profile detection from real samples;
     *                             if the majority of values look like one type,
     *                             map there, otherwise keep everything in
     *                             « Commentaire » so NO information is ever lost. */

    /* A 4-digit year (1900-2035) is NEVER a postal code nor an amount:
     * "annee de naissance" columns must never be mapped to Code postal. */
    function isYearLike(v) {
      var s = N.trimClean(v);
      if (!/^\d{4}$/.test(s)) return false;
      var n = parseInt(s, 10);
      return n >= 1900 && n <= 2035;
    }

    function smartTargetFromSamples(samples) {
      var cnt = {};
      var list = samples || [];
      list.forEach(function (v) {
        if (isYearLike(v)) return;             // year -> no destination
        var target = null;
        if (N.looksLikePhone(v)) target = 'tel1';
        else if (N.looksLikeEmail(v)) target = 'email';
        else if (N.looksLikePostal(v)) target = 'cp';
        else if (N.looksLikeAmount(v)) target = 'cotisation';
        else if (N.looksLikeAddress(v)) target = 'adresse';
        else if (N.looksLikeCity(v)) target = 'ville';
        else if (N.looksLikeMutuelle(v)) target = 'mutuelle';
        else if (N.looksLikePersonName(v)) target = 'nom';
        if (target) cnt[target] = (cnt[target] || 0) + 1;
      });
      var best = null, bestN = 0;
      Object.keys(cnt).forEach(function (k) { if (cnt[k] > bestN) { bestN = cnt[k]; best = k; } });
      /* Strict majority of the real sample values must agree, and a year-only
       * column must never become a postal code. */
      if (best === 'cp' && list.every(isYearLike)) return null;
      return (bestN > 0 && bestN * 2 > list.length) ? best : null;
    }

    /* A pure running sequence (1, 2, 3, … or ID-001, ID-002) carries no data
     * worth keeping in the final file: exclude it instead of dumping it into
     * « Commentaire ». */
    function looksLikeSequenceId(samples) {
      if (!samples || !samples.length) return false;
      var plain = samples.filter(function (v) { return /^\d{1,9}$/.test(N.trimClean(v)); });
      if (plain.length !== samples.length) return false;
      var nums = plain.map(function (v) { return parseInt(N.trimClean(v), 10); });
      for (var i = 1; i < nums.length; i++) {
        if (nums[i] - nums[i - 1] !== 1) return false;
      }
      return true;
    }

    /* A pure date column (date d'adhésion, date de naissance…) is not one of
     * the 11 official columns: exclude it cleanly instead of crowding
     * « Commentaire » with values the final file cannot use. */
    function looksLikeDateColumn(samples) {
      if (!samples || !samples.length) return false;
      var nonEmpty = samples.filter(function (v) { return N.trimClean(v) !== ''; });
      if (!nonEmpty.length) return false;
      var dateCount = 0;
      nonEmpty.forEach(function (v) {
        var s = N.trimClean(v);
        if (/^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$/.test(s) ||
            /^\d{4}-\d{2}-\d{2}/.test(s) ||
            /^\d{1,2}\s+(janv|fevr|mars|avr|mai|juin|juil|aout|sept|oct|nov|dec)/i.test(N.stripAccents(s))) {
          dateCount++;
        }
      });
      return dateCount / nonEmpty.length >= 0.7;
    }

    /* A column of 4-digit birth/joining years is not one of the 11 official
     * columns: exclude it rather than putting years into « Commentaire ». */
    function looksLikeYearColumn(samples) {
      if (!samples || !samples.length) return false;
      var years = 0, usable = 0;
      samples.forEach(function (v) {
        var s = N.trimClean(v);
        if (!s) return;
        usable++;
        if (isYearLike(s)) years++;
      });
      return usable > 0 && years / usable >= 0.6;
    }

    /* A column of short, single-word person names ("dn", "prenom", "given")…
     * Recognised by its VALUES, not only its header. */
    function looksLikePrenomColumn(item) {
      var samples = (item.samples || []).filter(function (v) { return N.trimClean(v); });
      if (!samples.length || samples.length > 6) return false;
      var single = samples.filter(function (v) {
        var s = N.trimClean(v);
        return s.split(/\s+/).length === 1 && s.length >= 2 && s.length <= 20 &&
          /^[A-Za-zÀ-ÿ'’\-]+$/.test(s);
      }).length;
      return single === samples.length;
    }

    function looksLikeNomColumn(item) {
      var samples = (item.samples || []).filter(function (v) { return N.trimClean(v); });
      if (!samples.length) return false;
      return samples.every(function (v) { return N.looksLikePersonName(v); });
    }

    /* Resolve a set of undecided columns in ONE pass: reprocessing a file
     * rebuilds its items, so deciding file by file would re-open columns and
     * leave the session permanently "to verify". */
    function resolveOpenItems(open) {
      var result = { fixed: 0, accepted: 0, preserved: 0, excluded: 0 };
      var hasMappedPrenom = false, hasMappedNom = false;
      /* One column is allowed to land in Commentaire: supplementary comment
       * columns must stay there, never be promoted to Nom carte blanche. */
      var nothingMapped = false;
      session.reviewItems.forEach(function (it) {
        if (it.status === 'open' || !it.decision) return;
        if (it.decision === 'prenom') hasMappedPrenom = true;
        if (it.decision === 'nom') hasMappedNom = true;
        if (it.decision === 'commentaire') nothingMapped = true;
      });
      open.forEach(function (item) {
        var target = item.suggestion;
        /* If suggestion was nom or prenom, but values contain invalid person names (like rappel / numbers),
         * do NOT map to nom/prenom! Preserve in commentaire instead. */
        if ((target === 'nom' || target === 'prenom') && (item.samples || []).some(N.isInvalidPersonName)) {
          target = 'commentaire';
          result.preserved++;
          session.decideReview(item.id, target);
          result.fixed++;
          return;
        }
        /* A full-name column goes whole into « Nom »: the value is never
         * split or copied into another column. */
        if (item.fullName) {
          target = 'nom'; result.accepted++;
        } else if (item.blocked) {
          target = '__exclude__'; result.excluded++;
        } else if (!target) {
          var smart = smartTargetFromSamples(item.samples);
          if (smart && !(item.samples || []).some(function (v) {
            return N.looksLikeAmount(v) && !N.looksLikePhone(v) && /[.,€]/.test(v) && smart === 'cp';
          })) {
            /* « Commentaire » keeps everything: one extra comment column
             * ("Observations", "Return") must not be routed to Nom/Ville. */
            if (smart === 'commentaire') nothingMapped = true;
            target = smart; result.accepted++;
          } else if (looksLikeSequenceId(item.samples) || looksLikeDateColumn(item.samples) ||
            looksLikeYearColumn(item.samples)) { target = '__exclude__'; result.excluded++; }
          else if (looksLikePrenomColumn(item) && !hasMappedPrenom) {
            /* A short column of single first names ("dn") is Prénom. */
            target = 'prenom'; result.accepted++; hasMappedPrenom = true;
          } else if (looksLikeNomColumn(item) && !hasMappedNom && !nothingMapped &&
            !(item.samples || []).some(N.isKnownPrenom)) {
            target = 'nom'; result.accepted++; hasMappedNom = true;
          } else { target = 'commentaire'; result.preserved++; }
        } else {
          result.accepted++;
        }
        session.decideReview(item.id, target);
        result.fixed++;
      });
      return result;
    }

    function autoFixReview() {
      var open = [];
      session.reviewItems.forEach(function (item) {
        if (item.status === 'open') open.push(item);
      });
      return resolveOpenItems(open);
    }

    /* Every text that did not fit in one Excel cell, kept in full so NOTHING
     * is lost: [{ fileName, sheet, sourceRow, column, text }]. One row per
     * continuation (each ≤ 32 767 characters). */
    function buildContinuations() {
      var out = [];
      session.fileOrder.forEach(function (fid) {
        var recs = session.recordsByFile.get(fid) || [];
        recs.forEach(function (r) {
          if (!r.longParts) return;
          Object.keys(r.longParts).forEach(function (t) {
            var rest = r.longParts[t];
            while (rest && rest.length) {
              out.push({
                fileName: r.fileName,
                sheet: r.sheet,
                sourceRow: r.sourceRow,
                column: SEC.TARGET_LABELS[t] || t,
                text: rest.slice(0, SAFE_CELL),
              });
              rest = rest.slice(SAFE_CELL);
            }
          });
        });
      });
      return out;
    }

    /* ---------------- Success / error line statistics ---------------- */

    function buildFileStats() {
      var files = [];
      var successRows = 0, warnRows = 0, failedFiles = 0, skippedFiles = 0;
      session.fileOrder.forEach(function (fid) {
        var f = session.files.get(fid);
        var success = 0, warn = 0;
        var problems = [];
        /* Per-file index (O(1)): scanning every record for every file made the
         * export of 100 files x 10 000 rows quadratic (the "export freeze"). */
        var ownRecords = session.recordsByFile.get(fid) || [];
        ownRecords.forEach(function (r) {
          if (r.hasWarning) {
            warn++;
            /* Cap the text: 10 000 basically identical problems made the
             * workbook (and its memory) explode for no added information. */
            if (problems.length < 500) {
              problems.push('Ligne ' + r.sourceRow + ' : ' + ((r.problems || []).join(' ; ') || 'valeur à vérifier'));
            }
          } else {
            success++;
          }
        });
        if (warn > problems.length) problems.push('… ' + (warn - problems.length) + ' autre(s) ligne(s) avec problème (voir la feuille « Detail des lignes »)');
        successRows += success;
        warnRows += warn;
        if (!f) return;
        if (f.status === S.FAILED) failedFiles++;
        if (f.status === S.CANCELLED) skippedFiles++;
        files.push({
          fileName: f.name,
          statusLabel: SEC.STATUS_LABELS[f.status] || f.status,
          success: success,
          warn: warn,
          rowsExtracted: success + warn,
          failed: f.status === S.FAILED,
          error: f.error || '',
          problems: problems
        });
      });
      return {
        files: files,
        successRows: successRows,
        warnRows: warnRows,
        totalRows: successRows + warnRows,
        failedFiles: failedFiles,
        skippedFiles: skippedFiles
      };
    }

    /* Every consolidated line with its content and its problems, so the UI and
     * the exported Bilan sheet can show exactly WHICH lines are OK and which
     * are flagged, and WHY. */
    function buildLineDetails() {
      var out = [];
      session.fileOrder.forEach(function (fid) {
        var f = session.files.get(fid);
        var recs = session.recordsByFile.get(fid);
        if (!recs || !recs.length) return;
        recs.sort(function (a, b) { return a.sourceRow - b.sourceRow; });
        recs.forEach(function (r) {
          out.push({
            fileName: f.name,
            sourceRow: r.sourceRow,
            sheet: r.sheet,
            ok: !r.hasWarning,
            problems: r.problems || [],
            values: r.values
          });
        });
      });
      return out;
    }



    session.registerFile = registerFile;
    session.removeFile = removeFile;
    session.clearFiles = clearFiles;
    session.parseEntry = parseEntry;
    session.processBuffer = processBuffer;
    session.startBatch = startBatch;
    session.cancel = cancel;
    session.retryFile = retryFile;
    session.decideReview = decideReview;
    session.resolveColumn = resolveColumn;
    session.openReviewCount = openReviewCount;
    /* Manual validation table (« Correction manuelle / À valider »). */
    session.addManualItem = addManualItem;
    session.openManualCount = openManualCount;
    session.applyManualDecision = applyManualDecision;
    session.acceptAllManual = acceptAllManual;
    session.sendAllManualToComment = sendAllManualToComment;
    session.reprocessAffected = reprocessAffected;
    session.buildConsolidatedRows = buildConsolidatedRows;
    session.reviewRecords = reviewRecords;
    session.autoFixReview = autoFixReview;
    session.buildFileStats = buildFileStats;
    session.buildLineDetails = buildLineDetails;
    session.buildContinuations = buildContinuations;
    session.addReport = addReport;
    session.recomputeStats = recomputeStats;
    return session;
  }

  /* Shared prerequisite checks (used by the UI and by automated tests). */
  function validatePrerequisites(opts) {
    var errors = [];
    if (!opts.fileCount) errors.push('Aucun fichier sélectionné. Ajoutez au moins un fichier CSV ou XLSX.');
    if (!opts.headers || opts.headers.length !== 11) errors.push('Les 11 entêtes officiels ne sont pas configurés.');
    if (opts.running) errors.push('Un traitement est déjà en cours.');
    if (!opts.hasProcessable) errors.push('Aucun fichier en attente de traitement.');
    return { ok: errors.length === 0, errors: errors };
  }

  /* Validate that a parsed workbook contains the exact 11 official headers. */
  function validateTemplate(sheets) {
    var targetNorms = SEC.OFFICIAL_HEADERS.map(SEC.Normalizers.normalizeHeader);
    var best = null;
    sheets.forEach(function (sheet) {
      sheet.rows.slice(0, SEC.LIMITS.MAX_SCAN_HEADER_ROWS).forEach(function (row, rowIdx) {
        var cells = row.map(SEC.Normalizers.normalizeHeader);
        var matches = 0;
        targetNorms.forEach(function (t) { if (cells.indexOf(t) !== -1) matches++; });
        if (!best || matches > best.matches) {
          best = { sheet: sheet.name, rowIdx: rowIdx, matches: matches, row: row };
        }
      });
    });
    if (!best || best.matches !== 11) return { ok: false, found: best ? best.matches : 0 };
    return { ok: true, row: best.row.map(String), sheet: best.sheet, rowIdx: best.rowIdx };
  }

  SEC.Processor = {
    createSession: createSession,
    validatePrerequisites: validatePrerequisites,
    validateTemplate: validateTemplate
  };
})(typeof self !== 'undefined' ? self : globalThis);

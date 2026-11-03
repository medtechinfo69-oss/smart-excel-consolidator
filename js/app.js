/* ============================================================
 * Smart Excel - Application controller
 * All state is in-memory only. No localStorage/sessionStorage/
 * IndexedDB, no network calls, no external services.
 * ============================================================ */
(function (root) {
  'use strict';
  var SEC = root.SEC;
  var UI = SEC.UI;
  var $ = UI.$;

  var session = SEC.Processor.createSession();
  var fileMap = new Map();      // entry id -> browser File (never persisted)
  /* Exports are always built with the standard 11-column workbook: no
   * external template file is loaded anymore. */
  var templateBuffer = null;
  var templateSheet = null;
  var modalEntryId = null;
  var rafPending = false;
  var lastRefreshTime = 0;
  var REFRESH_MIN_INTERVAL = 300; // ms between UI refreshes during processing

  /* ---------------- Utilities ---------------- */

  function downloadBlob(buffer, filename, mime) {
    var blob = new Blob([buffer], { type: mime || 'application/octet-stream' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }



  /* COALESCED refresh: processing announces every batch, but the UI is only
   * rebuilt at most ~3 times per second, and always at least once after the
   * last event (trailing call). Without the trailing call a lesson learned
   * at the end of a batch could be dropped; without the throttle 100 files x
   * 40 batches rebuilt every table thousands of times (the freeze). */
  var refreshTimer = 0;
  function scheduleRefresh() {
    if (rafPending) return;
    var wait = Math.max(0, REFRESH_MIN_INTERVAL - (Date.now() - lastRefreshTime));
    if (refreshTimer) return;
    refreshTimer = setTimeout(function () {
      refreshTimer = 0;
      var run = function () {
        rafPending = false;
        lastRefreshTime = Date.now();
        refreshAll();
      };
      rafPending = true;
      /* requestAnimationFrame is paused in background tabs: the timeout
       * fallback keeps the interface (and the export button) in sync. */
      if (typeof requestAnimationFrame === 'function') {
        requestAnimationFrame(run);
        setTimeout(function () { if (rafPending) run(); }, 120);
      } else {
        run();
      }
    }, wait);
  }

  function refreshAll() {
    UI.renderFileTables(session);
    UI.renderSummary(session);
    updateControls();
    updateOverall();
    renderOkWarnCounts();
  }

  function renderOkWarnCounts() {
    UI.syncRowCounters(session);
  }

  function updateOverall() {
    var total = session.fileOrder.length;
    if (!total) { $('#overallProgressWrap').hidden = true; return; }
    var done = 0, progressSum = 0, active = null;
    session.files.forEach(function (f) {
      progressSum += f.progress || 0;
      if (['completed', 'completed_warnings', 'failed', 'cancelled'].indexOf(f.status) !== -1) done++;
      else if (f.status === 'processing' || f.status === 'validating') active = f;
    });
    var pct = Math.round(progressSum / total);
    /* The active file's stage (« Extraction : Feuille1 », « Analyse IA :
     * 45 / 300 lignes ») is part of the line: without it the bar moved but
     * the text never changed and the run looked frozen. */
    var stage = session.running
      ? 'Traitement des fichiers… (' + done + '/' + total + ')' +
        (active && active.stage ? ' — ' + active.stage : '')
      : done === total ? 'Lot terminé' : 'En attente de démarrage';
    UI.setOverallProgress(pct, stage);
  }

  function updateControls() {
    var hasFiles = session.fileOrder.length > 0;
    var processable = session.fileOrder.some(function (id) {
      var s = session.files.get(id).status;
      return s === 'waiting' || s === 'failed' || s === 'cancelled';
    });
    var hasFailed = session.fileOrder.some(function (id) { return session.files.get(id).status === 'failed'; });
    $('#btnStart').disabled = session.running || !processable || session.headers.length !== 11;
    $('#btnCancel').disabled = !session.running;
    UI.updateExportAvailability(session, session.running);
    /* EXPORT GATE: an open « À valider » value blocks the Excel export until
     * the user validates it (or explicitly accepts every suggestion). */
    var open = session.openManualCount ? session.openManualCount() : 0;
    var exp = $('#btnExportExcel');
    if (exp && open > 0) {
      exp.disabled = true;
      exp.title = open + ' valeur(s) à valider dans « Correction manuelle / À valider » avant l\'export';
    }
    var step = $('#stepValidation');
    if (step) step.classList.toggle('has-open', open > 0);
  }

  function readBuffer(entry) {
    var file = fileMap.get(entry.id);
    if (!file) return Promise.reject(new Error('File reference unavailable (reselect the file)'));
    return file.arrayBuffer();
  }

  /* Parsing (SheetJS / CSV) runs in a Web Worker: the page never freezes.
   * Falls back transparently to the main thread when no worker is allowed. */
  var busySeq = 0;
  function parseInWorker(entry, buffer) {
    var seq = ++busySeq;
    UI.busy(true, 'Analyse de « ' + entry.name + ' »…', 25);
    return SEC.WorkerClient.parse(buffer, entry.ext, function (pct, stage) {
      if (seq === busySeq) UI.busy(true, (stage || 'Analyse') + ' — ' + entry.name, pct);
    }).finally(function () {
      if (seq === busySeq) UI.busy(false);
    });
  }

  function yieldToUI() { return new Promise(function (r) { setTimeout(r, 0); }); }

  /* ---------------- Automatic resolution ----------------
   * The engine decides every uncertain column by itself, inside the processing
   * loop (see Processor.resolveColumn): ID/reference columns are excluded,
   * "Nom complet" stays whole in « Nom », a clear suggestion is accepted, and
   * anything else is detected from the real values or kept in « Commentaire »
   * so NO information is ever lost. No manual validation step is needed. */

  /* ---------------- Step 1 : header configuration ---------------- */

  function setConfigStatus(text, kind) {
    $('#configStatusText').textContent = text;
    var dot = $('#configStatus .dot');
    dot.className = 'dot ' + (kind || 'dot-blue');
  }

  /* The manual editor is the only header configuration: fill the 11 inputs
   * with the active headers so they can be edited right away. */
  function renderManualInputs() {
    $('#manualInputs').innerHTML = session.headers.map(function (h, i) {
      return '<label>Header ' + (i + 1) + '<input data-manual="' + i + '" value="' + UI.escapeHtml(h) + '"></label>';
    }).join('');
  }
  renderManualInputs();

  $('#btnManualEdit').addEventListener('click', function () {
    /* The editor is always on screen: the button scrolls back to it. */
    var panel = $('#manualPanel');
    panel.hidden = false;
    renderManualInputs();
    if (typeof panel.scrollIntoView === 'function') panel.scrollIntoView({ behavior: 'smooth', block: 'center' });
    var first = panel.querySelector('[data-manual]');
    if (first && typeof first.focus === 'function') first.focus({ preventScroll: true });
  });
  $('#btnSaveManual').addEventListener('click', function () {
    var inputs = UI.$all('[data-manual]');
    var headers = inputs.map(function (i) { return i.value.trim(); });
    if (headers.some(function (h) { return !h; }) || headers.length !== 11) {
      UI.toast('Les 11 entêtes doivent être renseignés', 'error'); return;
    }
    session.headers = headers;
    UI.renderHeaderList(headers);
    setConfigStatus('Configuration manuelle active (11 entêtes)', 'dot-amber');
    UI.toast('Entêtes enregistrés', 'success');
  });
  $('#btnResetManual').addEventListener('click', function () {
    session.headers = SEC.OFFICIAL_HEADERS.slice();
    UI.renderHeaderList(session.headers);
    $('#manualInputs').querySelectorAll('input').forEach(function (inp, i) { inp.value = session.headers[i]; });
    setConfigStatus('Configuration officielle intégrée active (11 entêtes)', 'dot-blue');
  });



  /* ---------------- Optional AI assistant (OpenRouter) -----------------
   * The provider is called DIRECTLY from this page. The API key the user
   * pastes is stored in this browser only (localStorage) and sent only to
   * the provider itself. Without a key the assistant reports itself
   * unavailable and everything stays local. */

  var aiCheckSeq = 0;
  var AI = SEC.AI;

  var aiBusy = false;
  function setAiStatus(text, kind) {
    var el = $('#aiStatus');
    if (!el) return;
    el.textContent = text;
    el.className = 'ai-status' + (kind ? ' ' + kind : '');
    el.title = text;
    lockAiControls();
  }

  /* Controls are locked while a real generation is being tested: a second
   * click used to launch a parallel request and duplicate the work. */
  function lockAiControls() {
    var testBtn = $('#btnAiTest');
    var modelSel = $('#aiModel');
    if (testBtn) { testBtn.disabled = aiBusy; testBtn.textContent = aiBusy ? 'Test en cours…' : 'Tester la connexion'; }
    if (modelSel) modelSel.disabled = aiBusy;
    var retry = $('#btnAiRetry');
    if (retry) retry.disabled = aiBusy;
  }

  function providerLabel() {
    return AI.PROVIDER_LABELS[AI.status().provider] || 'OpenRouter';
  }

  function updateAiStatus() {
    var st = AI.status();
    var retry = $('#btnAiRetry');
    var keyHint = $('#aiKeyHint');
    /* The tick in the interface is the ONLY source of truth: a status line
     * that says "active" while the box is clear (or the other way round) is
     * exactly what makes a user think Gemini ran when it did not. */
    var box = $('#aiEnabled');
    if (box) box.checked = !!st.ticked;
    if (!st.ticked) {
      /* The box is clear: nothing is sent to Gemini for the next batch. */
      if (retry) retry.hidden = true;
      if (keyHint) keyHint.hidden = true;
      setAiStatus('Désactivé — traitement 100 % local, aucun appel à l IA');
      return;
    }
    if (st.stopped) {
      if (box) box.checked = true; // still ticked, but the AI failed
      setAiStatus(st.label + ' : ' + (st.lastError || 'indisponible') + ' — les lignes non classées sont traitées par le moteur local', 'warn');
      if (retry) retry.hidden = false;
      return;
    }
    if (st.checked && !st.reachable) {
      setAiStatus('Connexion au fournisseur IA impossible' + (st.lastError ? ' : ' + st.lastError : '') + ' — traitement local', 'warn');
      if (retry) retry.hidden = false;
      return;
    }
    if (!st.available[st.provider]) {
      /* Ticked but no key: the fix is in the box right above the status. */
      var keyBoxRow = $('#aiKeyBox');
      if (keyBoxRow) keyBoxRow.hidden = false;
      setAiStatus('Aucune clé ' + providerLabel() + ' active — collez-la ci-dessus puis cliquez « Activer la clé » (traitement local en attendant)', 'warn');
      if (keyHint) keyHint.hidden = true;
      return;
    }

    var shownModel = st.model && st.model !== 'auto'
      ? st.model + (st.usedModel && st.usedModel !== st.model ? ' (dernier appel : ' + st.usedModel + ')' : '')
      : (st.usedModel ? 'auto → ' + st.usedModel : 'auto');
    setAiStatus(st.label + ' actif — modèle : ' + shownModel + ' — analyse de toutes les lignes', 'ok');
    var keyBox = $('#aiKeyBox');
    if (keyBox) keyBox.hidden = false;
    if (retry) retry.hidden = true;
  }

  /* Real connectivity check: proves the key works end to end and shows the
   * exact model used (or the exact error) otherwise. */
  function checkAiConnection(silent) {
    if (!AI.status().enabled) return;
    var seq = ++aiCheckSeq;
    aiBusy = true;
    setAiStatus('Vérification de la connexion à ' + providerLabel() + '…');
    AI.test().then(function (res) {
      if (seq !== aiCheckSeq) return;
      aiBusy = false;
      if (res.ok) {
        setAiStatus(providerLabel() + ' connecté (' + res.model + ') — analyse de toutes les lignes', 'ok');
        fillModelSelect(res.models || [], res.model);
        if (!silent) UI.toast(providerLabel() + ' connecté (' + res.model + ')', 'success');
        /* Connection proven: retry a previously stopped assistant. */
        if (AI.status().stopped) { AI.reset(); updateAiStatus(); }
      } else {
        var msg = res.code === 'quota' ? 'quota gratuit épuisé' : (res.error || 'clé invalide');
        /* A failed test must never block the application: the local engine
         * (and the export) continue exactly the same way. */
        setAiStatus(providerLabel() + ' indisponible : ' + msg + ' — moteur local utilisé', 'err');
        if (!silent) UI.toast(providerLabel() + ' indisponible : ' + msg + ' — le moteur local prend le relais', 'warning');
      }
    });
  }

  /* ---------------- Paste & activate a provider API key ----------------
   * The key is validated directly with the provider, then stored in THIS
   * browser (localStorage) so it survives a reload. It is never echoed back:
   * only a masked hint is displayed. */
  var keyBusy = false;
  function activateKey() {
    var input = $('#aiKeyInput');
    var btn = $('#btnAiKey');
    if (!input || keyBusy) return;
    var value = String(input.value || '').trim();
    if (!value) {
      UI.toast('Collez votre clé OpenRouter avant de l\'activer', 'warning');
      input.focus();
      return;
    }
    keyBusy = true;
    if (btn) { btn.disabled = true; btn.textContent = 'Vérification…'; }
    setAiStatus('Vérification de la clé auprès de OpenRouter…');
    AI.setKey(value).then(function (res) {
      keyBusy = false;
      if (btn) { btn.disabled = false; btn.textContent = 'Activer la clé'; }
      if (!res.ok) {
        var msg = res.code === 'quota'
          ? 'quota gratuit épuisé pour cette clé'
          : (res.error || 'clé refusée');
        /* "network" is NOT a key problem: say what to fix. */
        if (res.code === 'network') {
          setAiStatus('Connexion impossible : ' + msg, 'err');
          UI.toast('Impossible de joindre ' + providerLabel() + ' : ' + msg, 'error');
          return;
        }
        setAiStatus('Clé refusée : ' + msg + ' — traitement local actif', 'err');
        UI.toast('Clé OpenRouter refusée : ' + msg, 'error');
        return;
      }
      /* Activated: turn the assistant on and prove it answers. */
      var cb = $('#aiEnabled');
      if (cb && !cb.checked) { cb.checked = true; }
      AI.configure({ enabled: true, model: selectedModel || 'auto' });
      var keyRow = $('#aiKeyRow');
      if (keyRow) keyRow.hidden = false;
      input.value = '';       // never keep the secret in the page
      input.type = 'password';
      input.placeholder = 'Clé active (' + (res.masked || 'enregistrée') + ') — collez-en une autre pour la remplacer';
      var hint = $('#aiKeyHint');
      if (hint) hint.hidden = true;
      fillModelSelect(res.models || [], res.model);
      setAiStatus(providerLabel() + ' connecté (' + (res.model || 'auto') + ') — analyse de toutes les lignes', 'ok');
      UI.toast('Clé OpenRouter activée' + (res.model ? ' (' + res.model + ')' : ''), 'success');
    });
  }
  /* The key input only appears when the assistant is ticked AND has no usable
   * key right now: with the box clear the panel stays one clear line. */
  function setAiKeyRowVisible(visible) {
    var row = $('#aiKeyRow');
    if (!row) return;
    var st = AI.status();
    var needsKey = st.ticked && (!st.checked || !st.reachable || !st.available[st.provider]);
    row.hidden = !(visible && needsKey);
  }

  if ($('#btnAiKey')) $('#btnAiKey').addEventListener('click', activateKey);
  if ($('#aiKeyInput')) {
    $('#aiKeyInput').addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); activateKey(); }
    });
    $('#aiKeyInput').addEventListener('input', function (e) {
      /* Leaving the field must never submit a half-pasted key. */
      if (e.target.value && /\s/.test(String(e.target.value).trim())) {
        e.target.value = String(e.target.value).replace(/\s+/g, '');
      }
    });
  }
  if ($('#btnAiKeyToggle')) {
    $('#btnAiKeyToggle').addEventListener('click', function () {
      var input = $('#aiKeyInput');
      if (!input) return;
      var show = input.type === 'password';
      input.type = show ? 'text' : 'password';
      this.textContent = show ? 'Masquer' : 'Afficher';
    });
  }

  /* Model picker. The CHOICE of the user is authoritative: a list refresh
   * never changes it, and the selected model (even one not in the current
   * list any more) is kept as an option instead of silently falling back to
   * "auto". */
  var selectedModel = 'auto'; // user's choice, remembered here

  function fillModelSelect(models, current) {
    var sel = $('#aiModel');
    if (!sel || !models || !models.length) return;
    var want = selectedModel || 'auto';
    if (want === 'auto' && current && current !== 'auto') want = current;
    var list = models.slice(0, 40);
    if (want !== 'auto' && list.indexOf(want) === -1) list.unshift(want);
    var wasAuto = sel.value === 'auto' || !sel.value;
    sel.innerHTML =
      '<option value="auto">Automatique — premier modèle disponible le plus rapide</option>' +
      list
        .map(function (m) {
          return '<option value="' + UI.escapeHtml(m) + '">' + UI.escapeHtml(m) + '</option>';
        })
        .join('');
    sel.value = want;
    /* First fill: show the model that will really be used by "auto". */
    if (want === 'auto' && wasAuto && current && current !== 'auto') {
      sel.options[0].textContent =
        'Automatique (utilisé : ' + current + ') — choisissez un modèle fixe ci-dessous';
    }
  }

  function loadAiModels() {
    AI.models().then(function (res) {
      if (res && res.models && res.models.length) fillModelSelect(res.models, res.model);
    });
  }

  /* Free-tier model picker before any network round-trip: the backend already
   * lists the provider's models when it starts. */
  function seedModelSelect() {
    var st = AI.status();
    if (st.models && st.models.length) fillModelSelect(st.models, st.lastModel || 'auto');
  }

  /* Give the picker a full list as soon as the page loads (the backend lists
   * the models in the background), so every choice is available immediately. */
  function loadFullModelList() {
    AI.models().then(function (res) {
      if (res && res.models && res.models.length) fillModelSelect(res.models, res.lastModel || res.model);
    });
  }

  /* Provider selector: the chosen provider is the one that processes the
   * data (OpenRouter). Switching reloads its key state and models. */
  function currentProvider() {
    var sel = $('#aiProvider');
    return (sel && sel.value) || 'openrouter';
  }

  function applyProvider(provider, silent) {
    AI.configure({ provider: provider });
    selectedModel = 'auto';
    var keyLabel = $('#aiKeyLabel');
    var keyInput = $('#aiKeyInput');
    var modelSel = $('#aiModel');
    if (keyLabel) keyLabel.textContent = 'Clé API OpenRouter (sk-or-v1-…)';
    if (keyInput) keyInput.placeholder = 'Collez ici votre clé OpenRouter (sk-or-v1-…) — facultative, une clé est déjà incluse';
    if (keyInput) keyInput.setAttribute('data-provider', provider);
    // Show the OpenRouter hint
    var hOr = $('#aiKeyHintOpenRouter');
    if (hOr) hOr.hidden = false;
    if (modelSel) {
      modelSel.innerHTML = '<option value="auto">Automatique — meilleur modèle disponible</option>';
      modelSel.value = 'auto';
    }
    AI.configure({ model: 'auto' });
    AI.reset();
    setAiStatus('Changement de fournisseur : ' + (AI.PROVIDERS_LABELS ? AI.PROVIDERS_LABELS[provider] : provider) + '…');
    AI.refreshStatus().then(function () {
      updateAiStatus();
      seedModelSelect();
      var st = AI.status();
      if (st.reachable && st.available[st.provider]) {
        loadAiModels();
        checkAiConnection(true);
      }
      if (!silent) {
        UI.toast('Fournisseur IA : ' + (AI.PROVIDERS_LABELS ? AI.PROVIDERS_LABELS[provider] : provider), 'success');
      }
    });
  }

  if ($('#aiProvider')) {
    $('#aiProvider').addEventListener('change', function (e) {
      applyProvider(e.target.value);
    });
  }
  if ($('#aiStrict')) {
    $('#aiStrict').addEventListener('change', function (e) {
      AI.configure({ strict: e.target.checked });
      UI.toast(e.target.checked
        ? 'IA obligatoire : les valeurs incertaines iront dans « À valider »'
        : 'Les valeurs incertaines seront proposées automatiquement', 'info');
    });
  }

  $('#aiEnabled').addEventListener('change', function (e) {
    var hint = $('#aiKeyHint');
    AI.configure({ enabled: e.target.checked, model: $('#aiModel').value, provider: currentProvider() });
    if (e.target.checked) {
      /* A tick always starts from a clean slate: a previous "stopped" flag
       * must not turn off the assistant the user just switched back on. */
      AI.reset();
      setAiKeyRowVisible(true);
      if (hint) hint.hidden = false;
      setAiStatus('Vérification de la connexion à ' + providerLabel() + '…');
      AI.refreshStatus().then(function (st) {
        updateAiStatus();
        seedModelSelect();
        if (st.reachable && st.available[st.provider]) {
          loadAiModels();
          checkAiConnection(true);
        }
      });
      UI.toast('Assistant IA OpenRouter activé : toutes les lignes seront analysées par l IA', 'success');
    } else {
      setAiKeyRowVisible(false);
      if (hint) hint.hidden = true;
      aiCheckSeq++;
      /* Disabled means: no request at all for the next batch. */
      setAiStatus('Désactivé — traitement 100 % local, aucun appel à l IA');
      UI.toast('Assistant IA désactivé : traitement et export 100 % locaux', 'info');
    }
  });
  /* The chosen model is used FROM NOW ON, exactly as selected. */
  $('#aiModel').addEventListener('change', function (e) {
    selectedModel = e.target.value || 'auto';
    AI.configure({ model: selectedModel });
    /* A previous failure/stop with another model must not disable the new
     * choice: the next request really tries the selected one. */
    if (AI.status().stopped) AI.reset();
    updateAiStatus();
    UI.toast(
      selectedModel === 'auto'
        ? 'Modèle : automatique (le plus rapide disponible)'
        : 'Modèle sélectionné : ' + selectedModel + ' (sera utilisé pour les prochaines analyses)',
      'success');
  });
  $('#btnAiTest').addEventListener('click', function () { checkAiConnection(false); });
  $('#btnAiRetry').addEventListener('click', function () {
    AI.reset();
    checkAiConnection(false);
  });

  /* Surface what the assistant did (or why it stopped) in the report, so a
   * wrong key or an exhausted quota is never silently ignored. */
  function recordAiResult(session) {
    var st = AI.status();
    /* A failure disables the assistant, so `enabled` is false exactly when a
     * report line matters most: the tick is the right test here. */
    if (!st.enabled && !st.ticked) return;
    /* Full-sheet mode: how many source lines the AI really classified. */
    if (st.sheetRows || st.sheetMissing) {
      var missing = st.sheetMissing || 0;
      session.addReport({
        fileName: '(assistant IA)',
        reason: st.label + ' a analysé ' + st.sheetRows + ' ligne(s) en ' + st.sheetCalls +
          ' requête(s)' + (st.usedModel ? ' (modèle ' + st.usedModel + ')' : '') +
          ' — ' + st.sheetPlaced + ' classée(s) par l IA' +
          (missing ? ', ' + missing + ' non classée(s) (moteur local utilisé)' : ''),
        suggestion: missing
          ? 'Lignes non classées : ' + (st.lastSheetError || 'voir les lignes concernées') + ' — elles ont été traitées localement'
          : 'Toutes les lignes ont été classées par l assistant IA',
        severity: missing ? 'warning' : 'info'
      });
      if (missing) UI.toast(missing + ' ligne(s) n ont pas pu être classées par l IA — traitées par le moteur local (voir le rapport)', 'warning');
      return;
    }
    if (!session.aiRequests && !st.mappings) return;
    var label = st.label || 'IA';
    session.addReport({
      fileName: '(assistant IA)',
      reason: st.mappings
        ? label + ' a confirmé ' + st.mappings + ' destination(s) de colonnes sur ' + session.aiRequests +
          ' bloc(s) analysé(s)' + (st.lastModel ? ' (modèle ' + st.lastModel + ')' : '')
        : label + ' n a pas pu répondre : ' + (st.lastError || 'erreur inconnue') + ' — moteur local utilisé',
        suggestion: st.mappings ? 'Destinations vérifiées par l assistant IA' : 'Vérifier la clé de l assistant et la connexion internet',
      severity: st.mappings ? 'info' : 'warning'
    });
  }

  /* ---------------- Step 2 : file selection ---------------- */

  var dropzone = $('#dropzone');
  $('#btnBrowse').addEventListener('click', function (e) { e.stopPropagation(); $('#fileInput').click(); });
  dropzone.addEventListener('click', function () { $('#fileInput').click(); });
  dropzone.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') $('#fileInput').click(); });
  ['dragover', 'dragenter'].forEach(function (evt) {
    dropzone.addEventListener(evt, function (e) { e.preventDefault(); dropzone.classList.add('dragover'); });
  });
  ['dragleave', 'drop'].forEach(function (evt) {
    dropzone.addEventListener(evt, function (e) { e.preventDefault(); dropzone.classList.remove('dragover'); });
  });
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) { e.preventDefault(); });
  dropzone.addEventListener('drop', function (e) { addFiles(e.dataTransfer.files); });
  $('#fileInput').addEventListener('change', function (e) {
    addFiles(e.target.files);
    e.target.value = '';
  });

  function addFiles(fileList) {
    var added = 0, dup = 0, unsupported = 0;
    Array.prototype.forEach.call(fileList, function (file) {
      var ext = (file.name.split('.').pop() || '').toLowerCase();
      if (SEC.SUPPORTED_EXTENSIONS.indexOf(ext) === -1) { unsupported++; return; }
      var duplicateOf = null;
      session.files.forEach(function (existing) {
        if (existing.name === file.name && existing.size === file.size) duplicateOf = existing.id;
      });
      var entry = session.registerFile({
        name: file.name, size: file.size, lastModified: file.lastModified, duplicateOf: duplicateOf
      });
      fileMap.set(entry.id, file);
      if (duplicateOf) dup++;
      added++;
    });
    UI.renderFileTables(session);
    updateControls();
    if (unsupported) UI.toast(unsupported + ' fichier(s) non pris en charge ignoré(s) (CSV, TXT, XLSX, XLSM, XLS uniquement)', 'warning');
    if (dup) UI.toast(dup + ' doublon(s) détecté(s) — ils seront ignorés pendant le traitement', 'warning');
    if (added) UI.toast(added + ' fichier(s) ajouté(s) au lot', 'success');
  }

  $('#btnRemoveAll').addEventListener('click', function () {
    if (session.running) { UI.toast('Impossible de vider la liste pendant le traitement', 'warning'); return; }
    session.clearFiles();
    fileMap.clear();
    UI.clearPreview();
    $('#previewSection').hidden = true;
    $('#validationList').innerHTML = '';
    UI.renderFileTables(session);
    UI.renderSummary(session);
    updateControls();
    updateOverall();
  });

  document.addEventListener('click', function (e) {
    var btn = e.target.closest('[data-action]');
    if (!btn) return;
    var id = btn.getAttribute('data-id');
    var action = btn.getAttribute('data-action');
    if (action === 'remove') {
      if (session.running) { UI.toast('Attendez la fin du lot ou annulez-le d\'abord', 'warning'); return; }
      session.removeFile(id);
      fileMap.delete(id);
      refreshAll();
      if (!session.records.size) { $('#previewSection').hidden = true; UI.clearPreview(); $('#validationList').innerHTML = ''; }
    } else if (action === 'details') {
      openDetails(id);
    } else if (action === 'show-lines') {
      var entry = session.files.get(id);
      var lines = session.buildLineDetails().filter(function (l) {
        return l.fileName === (entry && entry.name);
      });
      if (!lines.length) {
        UI.toast('Aucune ligne ajoutée pour ce fichier', 'info');
        return;
      }
      UI.renderLineDetails(lines, 'Lignes ajoutées — ' + (entry ? entry.name : '') + ' (' + lines.length + ')');
      document.querySelector('#lineDetailSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  });

  $('#btnCloseLineDetail').addEventListener('click', function () {
    document.querySelector('#lineDetailSection').hidden = true;
  });

  /* ---------------- Step 3 : processing ---------------- */

  $('#btnStart').addEventListener('click', function () {
    if (session.running) return;
    /* EVERY file in the list is (re)processed: completed ones included, so a
     * restart after a cancel — or after adding more files — never leaves an
     * old file behind. */
    var processableIds = session.fileOrder.slice();
    if (!processableIds.length) { UI.toast('Aucun fichier a traiter', 'error'); return; }
    var prereq = SEC.Processor.validatePrerequisites({
      fileCount: session.fileOrder.length,
      headers: session.headers,
      running: session.running,
      hasProcessable: processableIds.length > 0
    });
    if (!prereq.ok) { UI.toast(prereq.errors[0], 'error'); return; }
    var entries = processableIds.map(function (id) { return session.files.get(id); });
    /* Every previous record is dropped: a restart rebuilds everything from
     * the current file list (no duplicates, no stale rows left over). */
    session.records.clear();
    session.recordsByFile.clear();
    entries.forEach(function (en) {
      en.status = 'waiting'; en.progress = 0; en.error = ''; en.stage = '';
    });
    updateControls();
    UI.setOverallProgress(0, 'Initialisation de la file…');

    var batchPromise = session.startBatch(entries, {
      readBuffer: readBuffer,
      parseFn: parseInWorker,
      /* Every yield also repaints (scheduleRefresh is throttled to ~3/s):
       * without this the tables froze for the whole extraction — the
       * classic « no progression » report. */
      yieldFn: function () {
        updateOverall();
        scheduleRefresh();
        return yieldToUI();
      },
      /* Cheap local repaint only: the full tables are refreshed by the
       * coalesced scheduleRefresh below. */
      onProgress: function () { updateOverall(); scheduleRefresh(); },
      /* Real-time AI progress: the stage text and the percentage follow the
       * rows the provider has actually answered (file reading, AI analysis,
       * corrections and export each report their own stage). */
      onAiProgress: function () {
        /* updateOverall() now includes the file's stage (the AI row counter);
         * scheduleRefresh repaints the file table's progress bar. */
        updateOverall();
        scheduleRefresh();
      }
    });
    /* Lock the controls synchronously: no race window for a second click. */
    updateControls();
    batchPromise.then(function () {
      recordAiResult(session);
      renderManualReview();
      entries.forEach(function (en) { if (en._lastBuffer && fileMap.has(en.id)) en._lastBuffer = null; });
      scheduleRefresh();
      renderCurrentPreview();
      if (session.cancelling) UI.toast('Traitement annulé — interface restaurée, résultats partiels conservés', 'warning');
      else UI.toast('Lot traité avec succès — colonnes détectées et remplies automatiquement', 'success');
      updateControls();
    }).catch(function (err) {
      UI.toast('Erreur de traitement : ' + err.message, 'error');
    renderManualReview();
      updateControls();
    }).finally(function () {
      UI.busy(false);
    });
  });

  $('#btnCancel').addEventListener('click', function () {
    session.cancel();
    /* The state changed NOW: refresh the buttons immediately instead of
     * waiting for the next throttled table repaint (which could leave the
     * Cancel button looking active and Start inactive for seconds). */
    updateControls();
    updateOverall();
    UI.renderFileTables(session);
    UI.toast('Traitement annulé — cliquez « Démarrer le traitement » pour reprendre (les nouveaux fichiers seront inclus)', 'warning');
  });



  /* Preview shows EVERY consolidated line; the table is virtualised (only the
   * visible rows exist in the DOM) so it scrolls without freezing. */
  function renderCurrentPreview() {
    var total = session.records.size;
    UI.renderPreview(session.headers, session.buildConsolidatedRows(), total);
    var p = $('#previewPlaceholder');
    if (p) p.hidden = total > 0;
  }

  /* ---------------- File details modal ---------------- */

  function openDetails(id) {
    modalEntryId = id;
    var entry = session.files.get(id);
    var parseWith = function (buffer) {
      UI.busy(true, 'Lecture de « ' + entry.name + ' »…', 40);
      return SEC.WorkerClient.parse(buffer, entry.ext).then(function (res) {
        renderModal(entry, res.sheets);
      }).finally(function () { UI.busy(false); });
    };
    if (entry._lastBuffer) parseWith(entry._lastBuffer);
    else readBuffer(entry).then(parseWith).catch(function (err) { UI.toast('Impossible d\'ouvrir le fichier : ' + err.message, 'error'); });
  }

  function renderModal(entry, sheets) {
    $('#modalTitle').textContent = 'Fiche du fichier — ' + entry.name;
    var html = '';
    if (entry.encoding) html += '<p class="muted">Encodage : <strong>' + UI.escapeHtml(entry.encoding) + '</strong> · Séparateur : <strong>' + UI.escapeHtml(entry.delimiter === '\t' ? 'TAB' : entry.delimiter) + '</strong></p>';
    sheets.forEach(function (sheet) {
      var detected = SEC.Agent.detectHeaderRow(sheet.rows).index;
      var current = entry.headerOverrides[sheet.name];
      var checked = !entry.selectedSheets || entry.selectedSheets.indexOf(sheet.name) !== -1;
      html += '<div class="sheet-block">' +
        '<label class="sheet-check"><input type="checkbox" data-sheet-name="' + UI.escapeHtml(sheet.name) + '" ' + (checked ? 'checked' : '') + (sheet.isEmpty ? ' disabled' : '') + '> ' +
        '<strong>' + UI.escapeHtml(sheet.name) + '</strong>' + (sheet.isEmpty ? ' <span class="muted">(vide)</span>' : '') + '</label>' +
        '<label class="field-label">Ligne d\'entête (1 = première ligne ; détectée automatiquement : ' + (detected + 1) + ')</label>' +
        '<input type="number" min="1" data-header-row="' + UI.escapeHtml(sheet.name) + '" value="' + ((typeof current === 'number') ? current + 1 : detected + 1) + '">' +
        '</div>';
    });
    if (entry.error) html += '<p style="color:var(--red)">Erreur : ' + UI.escapeHtml(entry.error) + '</p>';
    $('#modalBody').innerHTML = html;
    $('#modalBackdrop').hidden = false;
  }

  $('#btnCloseModal').addEventListener('click', closeModal);
  $('#modalBackdrop').addEventListener('click', function (e) { if (e.target === $('#modalBackdrop')) closeModal(); });
  function closeModal() { $('#modalBackdrop').hidden = true; modalEntryId = null; }

  $('#btnSaveModal').addEventListener('click', function () {
    var entry = session.files.get(modalEntryId);
    if (!entry) return closeModal();
    var selected = UI.$all('[data-sheet-name]').filter(function (c) { return c.checked; }).map(function (c) { return c.getAttribute('data-sheet-name'); });
    if (!selected.length) { UI.toast('Sélectionnez au moins une feuille', 'warning'); return; }
    entry.selectedSheets = selected;
    entry._sheetsChosen = true;
    UI.$all('[data-header-row]').forEach(function (inp) {
      var name = inp.getAttribute('data-header-row');
      var v = parseInt(inp.value, 10);
      if (v >= 1) entry.headerOverrides[name] = v - 1;
    });
    closeModal();
    if (session.running) { UI.toast('Les modifications s\'appliqueront après la fin du lot en cours', 'warning'); return; }
    entry.status = 'waiting';
    session.reprocessAffected([entry.id], readBuffer, yieldToUI).then(function () {
      scheduleRefresh(); renderCurrentPreview();
      updateControls();
      UI.toast('Fichier retraité avec vos réglages de feuilles/entêtes', 'success');
    }).catch(function (err) {
      updateControls();
      UI.toast('Échec du retraitement : ' + err.message, 'error');
    });
  });

  /* ---------------- Correction manuelle / À valider ----------------
   * Every value the assistant could not classify with certainty is listed
   * here. The user picks its destination; the export stays locked until none
   * is left open. Nothing is ever guessed into a column silently. */

  var REVIEW_LIMIT = 500; // rows rendered at once (the rest is summarised)

  function renderManualReview() {
    var items = session.manualItems || [];
    var open = session.openManualCount ? session.openManualCount() : 0;
    var banner = $('#reviewBanner');
    var countEl = $('#reviewCount');
    var tableWrap = $('#reviewTableWrap');
    var empty = $('#reviewEmpty');
    var more = $('#reviewMore');
    var actions = $('#reviewActions');
    if (!items.length) {
      if (banner) banner.hidden = true;
      if (tableWrap) tableWrap.hidden = true;
      if (actions) actions.hidden = true;
      if (empty) empty.hidden = false;
      if (more) more.hidden = true;
      updateControls();
      return;
    }
    if (banner) banner.hidden = false;
    if (countEl) countEl.textContent = String(open);
    if (banner) banner.classList.toggle('all-done', open === 0);
    if (empty) empty.hidden = true;
    if (actions) actions.hidden = false;
    if (tableWrap) tableWrap.hidden = false;

    var shown = items.slice(0, REVIEW_LIMIT);
    var options = SEC.TARGET_KEYS.map(function (k) {
      return '<option value="' + k + '">' + UI.escapeHtml(SEC.TARGET_LABELS[k]) + '</option>';
    }).join('');
    $('#reviewTableBody').innerHTML = shown.map(function (it) {
      var cls = it.status === 'open' ? '' : ' class="review-done"';
      return '<tr' + cls + '>' +
        '<td>' + UI.escapeHtml(it.fileName) + '</td>' +
        '<td>' + UI.escapeHtml(String(it.row)) + '</td>' +
        '<td>' + UI.escapeHtml(it.sourceHeader) + '</td>' +
        '<td class="val" title="' + UI.escapeHtml(it.value) + '">' + UI.escapeHtml(it.value) + '</td>' +
        '<td>' + UI.escapeHtml(SEC.TARGET_LABELS[it.suggestion] || 'Commentaire') + '</td>' +
        '<td><select data-manual-dest="' + it.id + '">' +
          options.replace('value="' + it.destination + '"', 'value="' + it.destination + '" selected') +
        '</select></td>' +
        '<td><span class="review-state ' + (it.status === 'open' ? 'open' : 'done') + '">' +
          (it.status === 'open' ? 'À valider' : 'Validé') + '</span></td>' +
      '</tr>';
    }).join('');
    if (more) {
      if (items.length > REVIEW_LIMIT) {
        more.hidden = false;
        more.textContent = 'Affichage limité aux ' + REVIEW_LIMIT + ' premières valeurs sur ' +
          items.length + ' — utilisez « Accepter les suggestions » pour traiter le reste en une fois.';
      } else more.hidden = true;
    }
    updateControls();
  }

  /* A destination chosen in the table applies immediately to the record. */
  document.addEventListener('change', function (e) {
    var sel = e.target.closest('[data-manual-dest]');
    if (!sel) return;
    session.applyManualDecision(sel.getAttribute('data-manual-dest'), sel.value);
    renderManualReview();
    renderCurrentPreview();
  });

  if ($('#btnReviewAcceptAll')) {
    $('#btnReviewAcceptAll').addEventListener('click', function () {
      var left = session.acceptAllManual();
      renderManualReview();
      renderCurrentPreview();
      UI.toast(left ? left + ' valeur(s) restante(s)' : 'Suggestions acceptées pour toutes les valeurs', 'success');
    });
  }
  if ($('#btnReviewSendComment')) {
    $('#btnReviewSendComment').addEventListener('click', function () {
      session.sendAllManualToComment();
      renderManualReview();
      renderCurrentPreview();
      UI.toast('Toutes les valeurs restent dans « Commentaire »', 'info');
    });
  }
  if ($('#btnReviewApply')) {
    $('#btnReviewApply').addEventListener('click', function () {
      renderManualReview();
      renderCurrentPreview();
      updateControls();
      UI.toast('Corrections appliquées', 'success');
    });
  }

  /* ---------------- Step 4 : export ---------------- */

  /* Export: the consolidated rows are written exactly as they are displayed.
   * Open review items are NOT auto-decided here (that would silently choose
   * a destination for the user): the user is told how many remain. */
  var downloading = false;

  function exportWorkbook() {
    if (session.running) return;
    if (downloading) return;
    /* Never export with unvalidated values. */
    var openManual = session.openManualCount ? session.openManualCount() : 0;
    if (openManual > 0) {
      UI.toast(openManual + ' valeur(s) à valider dans « Correction manuelle / À valider » — l\'export est bloqué tant qu\'elles ne sont pas confirmées', 'warning');
      var v = $('#stepValidation');
      if (v && typeof v.scrollIntoView === 'function') v.scrollIntoView({ behavior: 'smooth', block: 'start' });
      return;
    }
    var rows = session.buildConsolidatedRows();
    if (!rows.length) { UI.toast('Aucun enregistrement à exporter', 'warning'); return; }
    /* Final verification BEFORE writing the workbook: a value whose real
     * nature contradicts its column (an insurer in Cotisation, a bare number
     * that may be an age) is moved to Mutuelle / Commentaire instead of being
     * exported misplaced. The user is told how many were corrected. */
    var repaired = SEC.Exporter.repairPlacement
      ? SEC.Exporter.repairPlacement(session.headers, rows)
      : { rows: rows, moved: { toMutuelle: 0, toCommentaire: 0 } };
    rows = repaired.rows;
    if (repaired.moved.toMutuelle || repaired.moved.toCommentaire) {
      UI.toast('Vérification avant export : ' + repaired.moved.toMutuelle +
        ' valeur(s) déplacée(s) vers Mutuelle, ' + repaired.moved.toCommentaire +
        ' vers Commentaire (mauvaise colonne détectée)', 'warning');
    }

    downloading = true;
    var exportBtn = $('#btnExportExcel');
    if (exportBtn) exportBtn.disabled = true; // no double click on a long export
    UI.busy(true, 'Génération et vérification du classeur…', 20);
    return yieldToUI().then(function () {
      /* Every data sheet (tableau final + Rapport/Bilan/Détail) is rebuilt with
       * the SAME cell writer as the preview, then re-opened and verified.
       * The build yields between steps so the page never stops responding. */
      var stats = session.buildFileStats();
      UI.busy(true, 'Préparation des feuilles de rapport…', 35);
      return yieldToUI().then(function () {
        /* The per-line sheet is only built when a line actually has a
         * problem: on a clean batch it duplicated the whole dataset. */
        var details = stats.warnRows > 0 ? session.buildLineDetails() : null;
        /* Texts longer than one Excel cell: exported in full, separately. */
        var continuations = session.buildContinuations ? session.buildContinuations() : [];
        UI.busy(true, 'Écriture du classeur Excel…', 55);
        return yieldToUI().then(function () {
          /* Built in the Web Worker: the page keeps scrolling and responding
           * during a big export instead of freezing behind the spinner. */
          return SEC.WorkerClient.exportXlsx({
            headers: session.headers,
            rows: rows,
            report: session.report,
            expectedCount: rows.length,
            fileStats: stats,
            lineDetails: details,
            continuations: continuations
          }, function (pct, stage) {
            UI.busy(true, stage || 'Écriture du classeur Excel…', Math.min(79, 55 + Math.round(pct / 4)));
          });
        }).then(function (out) {
          UI.busy(true, 'Vérification du fichier exporté…', 80);
          /* The workbook builder already re-opens its own bytes and runs every
           * integrity check (headers, row count, alignment, dates, empty
           * cells) in the worker; verifying a SECOND time here doubled the
           * export time for the same result. */
          return { buffer: out.buffer, validation: out.validation || { ok: true, checks: [] } };
        });
      });
    }).then(function (result) {
      UI.busy(true, 'Vérification du fichier exporté…', 80);
      UI.renderValidation(result.validation);
      if (!result.validation.ok) {
        /* Keep the download available: the user must be able to get their
         * data even when one control is negative. The failing checks are
         * named, and the file is offered in CSV as a guaranteed fallback. */
        var failed = result.validation.checks.filter(function (c) { return !c.pass; })
          .map(function (c) { return c.name + (c.detail ? ' (' + c.detail + ')' : ''); });
        UI.toast('Vérification incomplète : ' + failed.join(' ; ') +
          ' — le fichier est tout de même téléchargé, plus un CSV de secours.', 'warning');
        downloadBlob(result.buffer, 'fichier_final.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        downloadBlob(rowsToCsv(session.headers, rows), 'fichier_final_secours.csv', 'text/csv;charset=utf-8');
        return null;
      }
      /* Template: reuse the loaded official « En tete finale.xlsx » so its title
       * rows, header styling, column widths and every other sheet survive. */
      if (!templateBuffer) return { buffer: result.buffer, validation: result.validation };
      /* Template-preserving local rebuild, then plain validated workbook.
       * Each step is verified by re-opening the bytes. */
      UI.busy(true, 'Application de la mise en forme du modèle…', 70);
      return buildTemplateWorkbook(rows, result.validation, continuations).then(function (tpl) {
        if (tpl && tpl.buffer) {
          UI.renderValidation(tpl.validation);
          var loader = $('#templateLoader');
          if (loader) loader.setAttribute('data-template-applied', '1');
          return tpl;
        }
        UI.toast('La mise en forme du modèle n\'a pas pu être appliquée ; classeur standard validé utilisé', 'warning');
        return { buffer: result.buffer, validation: result.validation };
      });
    }).then(function (final) {
      if (!final) return;
      downloadBlob(final.buffer, 'fichier_final.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      UI.toast('Classeur Excel consolidé exporté (' + rows.length + ' lignes) — toutes les colonnes détectées automatiquement', 'success');
    }).catch(function (err) {
      /* Never leave the user without their data: a CSV is always produced. */
      try {
        downloadBlob(rowsToCsv(session.headers, rows), 'fichier_final_secours.csv', 'text/csv;charset=utf-8');
        UI.toast('Export Excel impossible (' + err.message + ') — un CSV de secours a été téléchargé', 'warning');
      } catch (e2) {
        UI.toast('Export impossible : ' + err.message, 'error');
      }
    }).finally(function () {
      downloading = false;
      UI.busy(false);
      updateControls();
    });
  }

  /* CSV fallback: UTF-8 BOM + ';' (Excel FR), quotes escaped, no data lost. */
  function rowsToCsv(headers, rows) {
    var esc = function (v) {
      var s = v == null ? '' : String(v);
      return /[";\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
    };
    var lines = [headers.map(esc).join(';')];
    rows.forEach(function (r) { lines.push(r.map(esc).join(';')); });
    return '\uFEFF' + lines.join('\r\n');
  }

  $('#btnExportExcel').addEventListener('click', function (e) {
    e.preventDefault();
    e.stopPropagation();
    exportWorkbook();
  });

  /* Template-preserving export (Web Worker).
   * The template is snapshotted cell by cell (raw value + style index), the
   * data rows are rewritten with the same writer as the standard export and
   * the result is re-opened and fully validated before the download. */
  function buildTemplateWorkbook(rows, baseValidation, continuations) {
    var total = rows.length;
    UI.busy(true, 'Lecture du modèle officiel…', 40);
    return SEC.WorkerClient.readModel(templateBuffer).then(function (model) {
      UI.busy(true, 'Application du modèle officiel (mise en forme conservée)…', 60);
      return SEC.WorkerClient.writeModel(templateBuffer, model, rows, function (pct) {
        UI.busy(true, 'Application du modèle officiel…', pct);
      });
    }).then(function (out) {
      var sheetName = templateSheetName();
      var validation = SEC.Exporter.validateBuffer(out.buffer, session.headers, total, sheetName, false, rows);
      if (!validation.ok) return { buffer: null, validation: validation };
      /* Template exports keep every original element: prove nothing was lost. */
      validation.checks = validation.checks.concat(baseValidation.checks.filter(function (c) {
        return /Report worksheet included|Textes longs/.test(c.name);
      }), [{ name: 'Modèle officiel conservé (mise en forme, lignes de titre, feuilles)', pass: true, detail: sheetName }]);
      return { buffer: out.buffer, validation: validation };
    }).catch(function (err) {
      return { buffer: null, validation: { ok: false, checks: [{ name: 'Modèle : ' + err.message, pass: false }] } };
    });
  }

  /* The sheet the data rows are written into: the model already located the
   * official 11-header row, so use THAT sheet (not blindly the first one).
   * Reading only the sheet names also avoids re-parsing the workbook. */
  function templateSheetName() {
    if (templateSheet && templateSheet.name) return templateSheet.name;
    try {
      var X = SEC.XLSXReader.getXLSX();
      var wb = X.read(templateBuffer, { type: 'array', bookSheets: true });
      return wb.SheetNames[0];
    } catch (e) { return 'tableau final'; }
  }

  /* ---------------- Init ---------------- */

  UI.renderHeaderList(session.headers);
  refreshAll();
  $('#overallProgressWrap').hidden = true;
  /* The assistant is on by default: it only runs when a key is stored in
   * this browser, and the local engine takes over the instant it cannot
   * answer. */
  AI.configure({
    enabled: $('#aiEnabled').checked,
    model: selectedModel || 'auto',
    provider: currentProvider(),
    strict: $('#aiStrict') ? $('#aiStrict').checked : true
  });
  updateAiStatus();

  /* Pre-warm the parsing worker so the first file is never slowed down. */
  try { SEC.WorkerClient.prewarm(); } catch (e) { /* inline fallback */ }

  /* Discover the key stored in this browser and the provider's model list. */
  if (typeof fetch === 'function') {
    AI.refreshStatus().then(function (st) {
      updateAiStatus();
      if (!st.reachable) return;
      setAiKeyRowVisible($('#aiEnabled').checked);
      seedModelSelect();
      if (st.available[st.provider]) { loadFullModelList(); loadAiModels(); }
    });
  }
})(typeof self !== 'undefined' ? self : globalThis);

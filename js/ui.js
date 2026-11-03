/* ============================================================
 * Smart Excel - DOM rendering layer
 * Pure rendering helpers; event wiring lives in app.js.
 * ------------------------------------------------------------
 * Performance rules (large batches / large previews):
 *  - file tables are rendered in CHUNKS inside requestAnimationFrame
 *    so 100+ rows never block the interface;
 *  - the consolidated preview is VIRTUALISED: only the rows in the
 *    viewport exist in the DOM (10 000+ rows scroll smoothly);
 *  - every other table uses incremental row updates instead of a full
 *    innerHTML rebuild on each refresh.
 * ============================================================ */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});

  function $(sel, ctx) {
    return (ctx || document).querySelector(sel);
  }
  function $all(sel, ctx) {
    return Array.prototype.slice.call((ctx || document).querySelectorAll(sel));
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function formatSize(bytes) {
    if (!bytes) return "0 KB";
    var units = ["B", "KB", "MB", "GB"];
    var i = Math.floor(Math.log(bytes) / Math.log(1024));
    return (bytes / Math.pow(1024, i)).toFixed(i ? 1 : 0) + " " + units[i];
  }

  /* Remove every visible toast: a correction invalidates the previous
   * messages ("rien à corriger", "aucune ligne…"), and stacking stale
   * notices over the refreshed UI is confusing. */
  function clearToasts() {
    var box = $("#toastContainer");
    if (box) box.innerHTML = "";
  }

  function toast(message, type) {
    var el = document.createElement("div");
    el.className = "toast " + (type || "info");
    el.textContent = message;
    $("#toastContainer").appendChild(el);
    setTimeout(function () {
      el.style.opacity = "0";
      el.style.transition = "opacity .3s ease";
      setTimeout(function () {
        el.remove();
      }, 320);
    }, 4200);
  }

  function renderHeaderList(headers) {
    $("#headerList").innerHTML = headers
      .map(function (h, i) {
        return (
          '<li><span class="pos">' +
          (i + 1) +
          "</span><span>" +
          escapeHtml(h) +
          "</span></li>"
        );
      })
      .join("");
  }

  function statusPill(status) {
    var label = SEC.STATUS_LABELS[status] || status;
    var check = status === "completed" ? " &#10003;" : "";
    return (
      '<span class="pill ' +
      escapeHtml(status) +
      '">' +
      escapeHtml(label) +
      check +
      "</span>"
    );
  }

  function fileRowHtml(entry) {
    var progress =
      entry.status === "completed" ||
      entry.status === "completed_warnings" ||
      entry.status === "needs_review"
        ? 100
        : entry.progress || 0;
    var dup = entry.duplicateOf
      ? '<span class="dup-tag">Doublon — ignoré au traitement</span>'
      : "";
    var err = entry.error
      ? '<div class="small" style="color:var(--red);max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="' +
        escapeHtml(entry.error) +
        '">' +
        escapeHtml(entry.error) +
        "</div>"
      : "";
    var rows = entry.rowsExtracted || 0;
    return (
      '<tr data-file-id="' +
      entry.id +
      '">' +
      '<td><div class="file-name-cell" title="' +
      escapeHtml(entry.name) +
      '">' +
      escapeHtml(entry.name) +
      "</div>" +
      dup +
      err +
      "</td>" +
      "<td>" +
      escapeHtml(entry.ext.toUpperCase()) +
      "</td>" +
      "<td>" +
      formatSize(entry.size) +
      "</td>" +
      "<td>" +
      statusPill(entry.status) +
      "</td>" +
      '<td><div class="mini-progress"><div class="mini-track"><div class="mini-fill" style="width:' +
      progress +
      '%"></div></div><div class="mini-pct">' +
      progress +
      "%</div></div></td>" +
      '<td class="cell-ok"><button class="link-btn num" data-action="show-lines" data-id="' +
      entry.id +
      '" data-only="ok" title="Voir les ' +
      rows +
      ' ligne(s) ajoutée(s)">' +
      rows +
      "</button></td>" +
      '<td class="actions-cell">' +
      '<button class="link-btn danger" data-action="remove" data-id="' +
      entry.id +
      '"' +
      (entry.status === "processing" || entry.status === "validating"
        ? " disabled"
        : "") +
      ">Retirer</button>" +
      "</td></tr>"
    );
  }

  function totalsRowHtml(totalRows) {
    return (
      '<tr class="totals-row">' +
      "<td><strong>Total</strong></td>" +
      "<td></td><td></td><td></td><td></td>" +
      '<td><strong class="cell-ok">' +
      totalRows +
      "</strong></td>" +
      "<td></td></tr>"
    );
  }

  /* Render a list of rows into a tbody in chunks, yielding to the browser
   * between chunks (requestAnimationFrame): the interface stays interactive
   * even with several hundred files. */
  function renderChunked(tbodyId, rowsHtml, totalsHtml) {
    var tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    var token = (tbody.__renderToken = (tbody.__renderToken || 0) + 1);
    if (!rowsHtml.length) {
      tbody.innerHTML = totalsHtml || "";
      return;
    }
    var CHUNK = 60;
    var start = 0;
    tbody.innerHTML = "";
    var frag = document.createDocumentFragment();
    var totals = document.createElement("template");
    if (totalsHtml) totals.innerHTML = totalsHtml;
    var step = function () {
      if (tbody.__renderToken !== token) return; // a newer render superseded this one
      var html = rowsHtml.slice(start, start + CHUNK).join("");
      if (html) {
        var tmp = document.createElement("tbody");
        tmp.innerHTML = html;
        while (tmp.firstChild) frag.appendChild(tmp.firstChild);
      }
      start += CHUNK;
      if (start < rowsHtml.length) {
        tbody.appendChild(frag);
        if (typeof requestAnimationFrame === "function")
          requestAnimationFrame(step);
        else setTimeout(step, 0);
        return;
      }
      if (totalsHtml) {
        while (totals.content.firstChild)
          frag.appendChild(totals.content.firstChild);
      }
      tbody.appendChild(frag);
    };
    step();
  }

  /* Detail of every consolidated line of one file (or of the whole batch),
   * so "Lignes OK" and "Lignes avec problème" can be inspected line by line. */
  function renderLineDetails(lines, title) {
    var section = $("#lineDetailSection");
    if (!lines.length) {
      section.hidden = true;
      return;
    }
    section.hidden = false;
    $("#lineDetailTitle").textContent = title;
    var rows = lines.map(function (l) {
      var v = l.values || {};
      var remark = (l.problems || []).join(" ; ");
      return (
        "<tr>" +
        '<td class="file-name-cell" title="' +
        escapeHtml(l.fileName) +
        '">' +
        escapeHtml(l.fileName) +
        "</td>" +
        "<td>" +
        l.sourceRow +
        "</td>" +
        '<td class="reason-cell">' +
        (remark ? escapeHtml(remark) : '<span class="muted">—</span>') +
        "</td>" +
        "<td>" +
        escapeHtml(v.nom || "") +
        "</td>" +
        "<td>" +
        escapeHtml(v.prenom || "") +
        "</td>" +
        "<td>" +
        escapeHtml(v.tel1 || "") +
        "</td>" +
        "<td>" +
        escapeHtml(v.tel2 || "") +
        "</td>" +
        "<td>" +
        escapeHtml(v.cotisation || "") +
        "</td>" +
        "</tr>"
      );
    });
    renderChunked("lineDetailBody", rows, "");
  }

  /* Incremental rewrite: only the rows whose HTML actually changed are
   * touched. Rebuilding every row of 100 files on every progress tick was a
   * main cause of the freeze (and of the scroll position being lost). */
  function setTbodyHtmlIncremental(tbodyId, rowsHtml, totalsHtml) {
    var tbody = document.getElementById(tbodyId);
    if (!tbody) return;
    /* A previous chunked render still in progress: let it finish first. */
    tbody.__renderToken = (tbody.__renderToken || 0) + 1;
    var existing = {};
    for (var i = 0; i < tbody.children.length; i++) {
      var tr = tbody.children[i];
      var id = tr.getAttribute("data-file-id");
      if (id) existing[id] = tr;
    }
    var wanted = {};
    var seen = {};
    var changed = false;
    rowsHtml.forEach(function (html) {
      var m = html.match(/^<tr data-file-id="([^"]+)"/);
      var id = m ? m[1] : null;
      if (!id) return;
      seen[id] = true;
      var tr = existing[id];
      if (tr && tr.__html === html) {
        wanted[id] = tr;
        return;
      }
      changed = true;
      if (!tr) tr = document.createElement("tr");
      tr.innerHTML = html.replace(/^<tr[^>]*>/, "").replace(/<\/tr>$/, "");
      tr.setAttribute("data-file-id", id);
      tr.__html = html;
      wanted[id] = tr;
    });
    Object.keys(existing).forEach(function (id) {
      if (!seen[id]) { existing[id].remove(); changed = true; }
    });
    if (!changed && tbody.children.length === rowsHtml.length + (totalsHtml ? 1 : 0)) return;
    tbody.innerHTML = "";
    Object.keys(wanted).forEach(function (id) { tbody.appendChild(wanted[id]); });
    var totals = tbody.querySelector("tr.totals-row");
    if (totalsHtml) {
      if (!totals) {
        var t = document.createElement("tbody");
        t.innerHTML = totalsHtml;
        tbody.appendChild(t.firstChild);
      } else {
        totals.outerHTML = totalsHtml;
      }
    } else if (totals) totals.remove();
  }

  function renderFileTables(session) {
    var rows = [];
    var totalRows = 0;
    var bytes = 0;
    session.fileOrder.forEach(function (id) {
      var f = session.files.get(id);
      if (!f) return;
      rows.push(fileRowHtml(f));
      totalRows += f.rowsExtracted || 0;
      bytes += f.size;
    });
    var totals = totalsRowHtml(totalRows);
    setTbodyHtmlIncremental("fileTableBody", rows, totals);
    setTbodyHtmlIncremental("processTableBody", rows, totals);
    $("#fileCounter").textContent = session.fileOrder.length;
    $("#totalSize").textContent = formatSize(bytes);

    var warn = $("#memoryWarning");
    if (
      bytes > SEC.LIMITS.MEMORY_WARN_BYTES ||
      session.fileOrder.length > SEC.LIMITS.MEMORY_HARD_HINT_FILES
    ) {
      warn.hidden = false;
      warn.textContent =
        "Large batch detected (" +
        session.fileOrder.length +
        " files, " +
        formatSize(bytes) +
        "). Processing is sequential and memory-conscious, but very large workbooks may approach browser memory limits. Close other tabs and process in smaller batches if the browser becomes slow.";
    } else {
      warn.hidden = true;
    }
  }

  /* Headline cards, kept simple:
   *  - « Fichiers réussis »  : files whose rows all landed in the final file
   *  - « Fichiers en échec » : unreadable / unsupported files only
   *  - « Correspondances à vérifier » : mapping decisions still open
   *  - « Lignes ajoutées avec succès » : every extracted line (they all land
   *    in the final file, in the chosen columns) */
  function syncRowCounters(session) {
    if ($("#sumOkRows")) $("#sumOkRows").textContent = session.records.size;
    var okCard = $("#sumOkRows") && $("#sumOkRows").closest(".summary-card");
    if (okCard) okCard.className = "summary-card sc-green";
    return { ok: session.records.size, warn: 0 };
  }

  function renderSummary(session) {
    var success = 0,
      failed = 0;
    session.files.forEach(function (f) {
      if (
        f.status === "completed" ||
        f.status === "completed_warnings" ||
        f.status === "needs_review"
      )
        success++;
      if (f.status === "failed") failed++;
    });
    $("#sumTotal").textContent = session.fileOrder.length;
    $("#sumSuccess").textContent = success;
    $("#sumFailed").textContent = failed;
    /* Lines added = every record (nothing is discarded). */
    if ($("#sumOkRows")) $("#sumOkRows").textContent = session.records.size;
  }



  /* ---------------- Virtual scrolling preview ----------------
   * Only the rows visible in the scroll container exist in the DOM, with a
   * top and bottom spacer of the exact height of the hidden rows: 100 000
   * consolidated lines scroll perfectly smoothly and never freeze the page. */
  var vp = {
    headers: [],
    rows: [],
    rowHeight: 33,
    buffer: 8,
    bound: false,
    frame: 0,
  };

  function placeholder() {
    var section = $("#previewSection");
    if (!section) return null;
    var p = $("#previewPlaceholder");
    if (!p) {
      p = document.createElement("div");
      p.id = "previewPlaceholder";
      p.className = "preview-placeholder";
      p.textContent = "Aucune donnée consolidée pour le moment.";
      section.appendChild(p);
    }
    return p;
  }

  function measureRowHeight() {
    var table = $("#previewTable");
    var tr = table && table.querySelector("tbody tr.pv-row");
    if (tr && tr.getBoundingClientRect().height > 0)
      vp.rowHeight = Math.ceil(tr.getBoundingClientRect().height);
    return vp.rowHeight;
  }

  /* Window to render: ROW BUFFERING keeps about 24 rows above and below the
   * viewport so fast scrolling never shows a white gap. */
  function renderWindow() {
    vp.frame = 0;
    var wrap = $("#previewWrap");
    var table = $("#previewTable");
    if (!wrap || !table) return;
    var H = vp.rowHeight;
    /* The TOTAL comes from the dataset (a streamed preview may render a
     * capped slice): using the slice length made the scrollbar and the row
     * counter wrong while more rows were still arriving. */
    var total = Math.max(vp.total || 0, vp.rows.length);
    var viewTop = wrap.scrollTop;
    var visible = Math.ceil(wrap.clientHeight / H) || 20;
    var start = Math.max(0, Math.floor(viewTop / H) - vp.buffer);
    var end = Math.min(total, start + visible + vp.buffer * 2);
    var tbody = $("#previewBody");
    if (!tbody) {
      table.innerHTML = "";
      var thead = document.createElement("thead");
      thead.innerHTML =
        "<tr>" +
        vp.headers
          .map(function (head) {
            return "<th>" + escapeHtml(head) + "</th>";
          })
          .join("") +
        "</tr>";
      table.appendChild(thead);
      tbody = document.createElement("tbody");
      tbody.id = "previewBody";
      table.appendChild(tbody);
    }
    var html = "";
    /* Spacer = a few real <tr> of at most ~50 000 px in total: a single very
     * tall row is capped by browsers and then the scrollbar thumbs/jumps
     * (the "scrolling jumps" symptom), while 100 000 spacer rows would cost
     * more DOM than the data. `table-layout: fixed` + `!height` keeps each
     * spacer's height EXACT, so the offset stays correct. */
    var spacerTop = Math.floor(start);
    var spacerBottom = Math.floor(total - end);
    var addSpacers = function (count, extraPx) {
      if (count <= 0 && !extraPx) return "";
      var totalPx = count * H + (extraPx || 0);
      var html = "";
      while (totalPx > 0) {
        var px = Math.min(totalPx, 50000);
        html +=
          '<tr class="pv-spacer" aria-hidden="true" style="height:' + px +
          'px"><td colspan="' + vp.headers.length +
          '" style="padding:0;border:0;height:' + px + 'px"></td></tr>';
        totalPx -= px;
      }
      return html;
    };
    if (spacerTop)
      html += addSpacers(spacerTop, (start - spacerTop) * H);
    /* Never build a row that does not exist yet (streamed preview). */
    end = Math.min(end, vp.rows.length);
    for (var i = start; i < end; i++) {
      var r = vp.rows[i];
      if (!r) continue;
      html +=
        '<tr class="pv-row">' +
        r
          .map(function (c) {
            var s = c == null ? "" : String(c);
            return (
              '<td title="' +
              escapeHtml(s) +
              '">' +
              escapeHtml(s === "" ? " " : s) +
              "</td>"
            );
          })
          .join("") +
        "</tr>";
    }
    if (spacerBottom)
      html += addSpacers(spacerBottom, (total - end - spacerBottom) * H);
    else if (total > vp.rows.length)
      html += addSpacers(total - vp.rows.length, 0);
    tbody.innerHTML = html;
    measureRowHeight();
    var info = $("#previewInfo");
    if (info)
      info.textContent = total
        ? "lignes " +
          (start + 1) +
          "–" +
          end +
          " sur " +
          total +
          " (défilement virtuel)"
        : "";
  }

  function requestWindow() {
    if (vp.frame) return;
    if (typeof requestAnimationFrame === "function")
      vp.frame = requestAnimationFrame(renderWindow);
    else vp.frame = setTimeout(renderWindow, 16);
  }

  function bindPreviewScroll() {
    if (vp.bound) return;
    var wrap = $("#previewWrap");
    if (!wrap) return;
    vp.bound = true;
    wrap.addEventListener("scroll", requestWindow, { passive: true });
    if (typeof ResizeObserver === "function") {
      new ResizeObserver(function () {
        requestWindow();
      }).observe(wrap);
    }
  }

  /* Render the consolidated rows (ALL of them: export and preview share the
   * same data; only the visible slice reaches the DOM). */
  function renderPreview(headers, rows, total) {
    var section = $("#previewSection");
    if (!rows || !rows.length) {
      section.hidden = true;
      vp.rows = [];
      return;
    }
    section.hidden = false;
    var p = placeholder();
    if (p) p.hidden = true;
    vp.headers = headers.slice();
    vp.rows = rows;
    vp.total = total == null ? rows.length : total;
    bindPreviewScroll();
    /* A refresh must never throw the user back to the top of a long table. */
    renderWindow();
  }

  function clearPreview() {
    vp.rows = [];
    vp.headers = [];
    var table = $("#previewTable");
    if (table) table.innerHTML = "";
    var p = placeholder();
    if (p) p.hidden = false;
    var info = $("#previewInfo");
    if (info) info.textContent = "";
  }

  /* The export button must be enabled as soon as there is something to
   * export (rows landed in the final file), and only then. */
  function updateExportAvailability(session, running) {
    var btn = $("#btnExportExcel");
    if (!btn) return;
    btn.disabled = !!running || session.records.size === 0;
    btn.title = btn.disabled
      ? "Traitez au moins un fichier pour pouvoir exporter"
      : "";
  }



  function renderValidation(validation) {
    $("#validationList").innerHTML = validation.checks
      .map(function (c) {
        return (
          '<li class="' +
          (c.pass ? "pass" : "fail") +
          '"><span class="v-icon">' +
          (c.pass ? "✓" : "✕") +
          "</span>" +
          escapeHtml(c.name) +
          (c.detail
            ? ' <span class="muted">(' + escapeHtml(c.detail) + ")</span>"
            : "") +
          "</li>"
        );
      })
      .join("");
  }

  function setOverallProgress(pct, stage) {
    var wrap = $("#overallProgressWrap");
    wrap.hidden = false;
    $("#overallBar").style.width = pct + "%";
    $("#overallPct").textContent = pct + "%";
    if (stage) $("#overallStage").textContent = stage;
  }

  /* ---------------- Busy overlay (long parse / export) ----------------
   * Shown only for operations the user cannot see progress of otherwise
   * (worker parsing a big workbook, export build). The interface stays
   * interactive: the overlay is pointer-transparent and can be dismissed. */
  function busy(show, message, pct) {
    var el = $("#busyOverlay");
    if (!el) return;
    if (!show) {
      el.hidden = true;
      return;
    }
    el.hidden = false;
    var text = $("#busyText");
    if (text) text.textContent = message || "Traitement en cours…";
    var bar = $("#busyBar");
    if (bar)
      bar.style.width =
        (pct == null ? 35 : Math.max(0, Math.min(100, pct))) + "%";
  }

  var UI = {
    $: $,
    $all: $all,
    escapeHtml: escapeHtml,
    formatSize: formatSize,
    toast: toast,
    clearToasts: clearToasts,
    renderHeaderList: renderHeaderList,
    renderFileTables: renderFileTables,
    renderSummary: renderSummary,
    syncRowCounters: syncRowCounters,
    renderPreview: renderPreview,
    clearPreview: clearPreview,
    renderValidation: renderValidation,
    renderLineDetails: renderLineDetails,
    updateExportAvailability: updateExportAvailability,
    setOverallProgress: setOverallProgress,
    busy: busy,
    statusPill: statusPill,
  };
  SEC.UI = UI;
})(typeof self !== "undefined" ? self : globalThis);

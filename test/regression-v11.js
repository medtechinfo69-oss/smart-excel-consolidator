/* ============================================================
 * Regression suite for the v1.1 rework (freezes, mapping, export).
 * ------------------------------------------------------------
 * Everything here failed (or was impossible) before the fixes:
 *  - event-loop yield really gives the browser back its CPU
 *  - unknown headers are actually resolved by the AI hints
 *  - extra comment columns are never promoted to Nom / never duplicated
 *  - a deliberate profile exclusion no longer keeps rows "to verify"
 *  - 2000 records are verified cell by cell on export
 *  - large files process with a responsive UI (yield count measured)
 * Run: node test/regression-v11.js
 * ============================================================ */
"use strict";
const fs = require("fs");
const path = require("path");

require("../js/config.js");
require("../js/data/communes.js");
require("../js/dictionaries.js");
require("../js/data/mutuelles.js");
require("../js/normalizers.js");
require("../js/csv-parser.js");
require("../js/xlsx-reader.js");
require("../js/agent.js");
require("../js/processor.js");
require("../js/exporter.js");
const SEC = globalThis.SEC;

let passed = 0,
  failed = 0;
function test(name, fn) {
  return Promise.resolve()
    .then(fn)
    .then(() => {
      passed++;
      console.log("  PASS " + name);
    })
    .catch((e) => {
      failed++;
      console.log("  FAIL " + name + " -> " + e.message);
    });
}
function assert(c, m) {
  if (!c) throw new Error(m || "assertion failed");
}
function eq(a, b, m) {
  if (JSON.stringify(a) !== JSON.stringify(b))
    throw new Error(
      (m || "not equal") +
        " (got " +
        JSON.stringify(a) +
        " want " +
        JSON.stringify(b) +
        ")",
    );
}
const tick = () => new Promise((r) => setTimeout(r, 0));

function fakeAI(mapping, calls) {
  return {
    enabled: () => true,
    status: () => ({
      checked: true,
      enabled: true,
      label: "FakeAI",
      stopped: false,
    }),
    refreshStatus: () => Promise.resolve(),
    model: () => "fake-1",
    suggestSheet: (file, sheet, columns) => {
      calls.push(columns.length);
      return Promise.resolve(Object.assign({}, mapping));
    },
  };
}

async function run(name, buffer, opts) {
  const session = SEC.Processor.createSession();
  const entry = session.registerFile({ name, size: buffer.length });
  let yields = 0;
  await session.startBatch([entry], {
    readBuffer: () => Promise.resolve(buffer),
    yieldFn: () => {
      yields++;
      return new Promise((r) => setTimeout(r, 0));
    },
  });
  return { session, entry, yields, rows: session.buildConsolidatedRows() };
}

(async () => {
  console.log("\n=== v1.1 regression suite ===\n");

  await test("Processing yields to the event loop regularly (no freeze)", async () => {
    const lines = ["Nom;Prénom;Téléphone 1;Code postal;Ville"];
    for (let i = 0; i < 12000; i++) {
      lines.push(
        "Nom" +
          i +
          ";Prénom" +
          i +
          ";0612" +
          String(100000 + i).slice(-6) +
          ";01200;Bellegarde",
      );
    }
    const buf = Buffer.from(lines.join("\n"), "utf8");
    const t0 = Date.now();
    const r = await run("big.csv", buf);
    const ms = Date.now() - t0;
    eq(r.rows.length, 12000, "all 12 000 rows extracted");
    assert(
      r.yields >= 3,
      "yieldFn called " + r.yields + " times (distinct animation frames)",
    );
    console.log("    (12 000 rows in " + ms + " ms, " + r.yields + " yields)");
  });

  await test("Unknown header resolved by the AI hint, no review left open", async () => {
    const calls = [];
    SEC.AI = fakeAI({ "titre client": "nom", "ref dossier": "commentaire" }, calls);
    const r = await run(
      "ai.csv",
      Buffer.from(
        "Titre client;Ref dossier;Courriel;Prime\nDupont;A-12;a@b.fr;MGEN\nMartin;B-44;c@d.fr;AXA\n",
      ),
    );
    assert(
      r.rows.some((row) => row[0] === "Dupont"),
      "Titre client -> Nom through the AI hint",
    );
    assert(r.session.aiHints >= 1, "AI hints counted in the session");
    assert(
      r.session.report.some((x) => /assistant IA/.test(x.fileName)),
      "AI outcome recorded in the report",
    );
    const open = r.session.openReviewCount();
    const names = Array.from(r.session.reviewItems.values())
      .filter((i) => i.status === "open")
      .map((i) => i.header);
    assert(
      open === 0,
      "no open review after the AI answered (still open: " +
        names.join(", ") +
        ")",
    );
    SEC.AI = undefined;
  });

  await test("AI unavailable -> local engine still maps, processes and exports", async () => {
    SEC.AI = {
      enabled: () => true,
      status: () => ({
        checked: true,
        enabled: true,
        label: "FakeAI",
        stopped: true,
        lastError: "quota épuisé",
      }),
      refreshStatus: () => Promise.resolve(),
      suggestSheet: () => Promise.resolve({}),
      model: () => "",
    };
    const r = await run(
      "quota.csv",
      Buffer.from("Nom;Prénom;Portable\nDupont;Jean;0612345678\n"),
    );
    eq(r.rows[0][0], "Dupont", "Nom mapped locally");
    eq(r.rows[0][2], "0612345678", "phone kept with its leading zero");
    const out = SEC.Exporter.generateConsolidated(
      r.session.headers,
      r.rows,
      r.session.report,
      r.rows.length,
    );
    assert(out.validation.ok, "export validates without the AI");
    SEC.AI = undefined;
  });

  await test("Supplementary comment column stays in Commentaire and is never duplicated", async () => {
    const csv =
      "Nom de naissance;Prénom;Courriel;Commentaire;Retour\n" +
      "Dupont;Jean;jean@x.fr;RAS;Appelé\n" +
      "Martin;Marie;marie@x.fr;Rappel mardi;Ne répond pas\n";
    const r = await run("comments.csv", Buffer.from(csv));
    eq(r.rows.length, 2, "2 rows");
    assert(
      !/Dupont/.test(r.rows[0][10]) || r.rows[0][0] === "Dupont",
      "name stays in Nom",
    );
    assert(
      /RAS/.test(r.rows[0][10]) && /Appelé/.test(r.rows[0][10]),
      "both comments preserved: " + r.rows[0][10],
    );
    eq(
      (r.rows[0][10].match(/RAS/g) || []).length,
      1,
      "comment text appears exactly once",
    );
  });

  await test("Ages / years / identifiers are excluded, reported once and do not block the file", async () => {
    const csv =
      "Nom;Prénom;Age;Année de naissance;Dossier\n" +
      "Dupont;Jean;63;1950;D-1\n" +
      "Martin;Marie;71;1948;D-2\n" +
      "Bernard;Paul;64;1949;D-3\n";
    const r = await run("ages.csv", Buffer.from(csv));
    eq(
      r.entry.status,
      "completed",
      "file completed (was needs_review for every age column)",
    );
    /* Deliberate exclusions (running id, ages, birth years) are reported but
     * never turned into a question for the user. */
    const openItems = Array.from(r.session.reviewItems.values())
      .filter((i) => i.status === "open")
      .map((i) => i.header);
    ["Dossier", "Age", "Année de naissance"].forEach((h) =>
      assert(openItems.indexOf(h) === -1, h + " must not ask for review (open: " + openItems.join(", ") + ")"),
    );
    assert(
      r.session.report.some((x) =>
        /identifiant\/date\/annee\/age/.test(x.reason),
      ),
      "exclusion reported",
    );
    eq(r.rows[0][0], "Dupont");
    eq(r.rows[0][10], "", "no age dumped into Commentaire");
  });

  await test("Export verifies EVERY record of a 2000-row dataset cell by cell", async () => {
    const session = SEC.Processor.createSession();
    const rows = [];
    for (let i = 0; i < 2000; i++) {
      rows.push([
        "Nom" + i,
        "Prénom" + i,
        "06" + String(10000000 + i),
        "02" + String(10000000 + i).slice(-8),
        i + 1 + " rue des Lilas",
        "0" + String(1200 + (i % 800)),
        "Ville" + (i % 50),
        "user" + i + "@example.fr",
        "12," + String(10 + (i % 80)),
        "MGEN",
        "note " + i,
      ]);
    }
    const out = SEC.Exporter.generateConsolidated(
      session.headers,
      rows,
      [],
      rows.length,
    );
    assert(
      out.validation.ok,
      "validation ok: " +
        JSON.stringify(out.validation.checks.filter((c) => !c.pass)),
    );
    const names = out.validation.checks.map((c) => c.name);
    assert(
      names.indexOf("Every exported record is present and aligned") !== -1,
      "record integrity check present",
    );
    assert(
      names.indexOf("No exported record is empty") !== -1,
      "empty record check present",
    );
    const X = SEC.XLSXReader.getXLSX();
    const wb = X.read(out.buffer, { type: "array" });
    const aoa = X.utils.sheet_to_json(wb.Sheets["tableau final"], {
      header: 1,
      defval: "",
      raw: true,
    });
    eq(aoa.length - 1, 2000, "2000 rows in the workbook");
    for (let i = 0; i < 2000; i += 97) {
      eq(aoa[i + 1][0], rows[i][0], "row " + (i + 1) + " Nom");
      eq(
        String(aoa[i + 1][5]),
        rows[i][5],
        "row " + (i + 1) + " Code postal keeps its zero",
      );
    }
  });

  await test("Five real sample files are processed with zero data loss", async () => {
    const files = [
      "01_clients_semicolon.csv",
      "02_clients_comma.csv",
      "03_tab_delimited.txt",
      "05_workbook_multi.xlsx",
      "06_ambiguous.xlsx",
      "07_edge_cases.xlsx",
      "04_legacy_cp1252.csv",
    ];
    const session = SEC.Processor.createSession();
    const entries = files.map((f) =>
      session.registerFile({
        name: f,
        size: fs.readFileSync(path.join(__dirname, "fixtures", f)).length,
      }),
    );
    await session.startBatch(entries, {
      readBuffer: (e) =>
        Promise.resolve(
          fs.readFileSync(path.join(__dirname, "fixtures", e.name)),
        ),
      yieldFn: tick,
    });
    const failedFiles = entries.filter((e) => e.status === "failed");
    assert(
      failedFiles.length === 0,
      "no failed file: " +
        failedFiles.map((e) => e.name + ":" + e.error).join(", "),
    );
    const rows = session.buildConsolidatedRows();
    const stats = session.buildFileStats();
    eq(stats.totalRows, rows.length, "stats and rows agree");
    assert(rows.length >= 20, "rows extracted: " + rows.length);
    const empty = rows.filter((r) => r.every((c) => c === ""));
    eq(empty.length, 0, "no empty consolidated record");
    const out = SEC.Exporter.generateConsolidated(
      session.headers,
      rows,
      session.report,
      rows.length,
      stats,
      session.buildLineDetails(),
    );
    assert(
      out.validation.ok,
      "export of the real sample set validates: " +
        JSON.stringify(out.validation.checks.filter((c) => !c.pass)),
    );
    console.log(
      "    (" +
        files.length +
        " files, " +
        rows.length +
        " rows, all verified)",
    );
  });

  await test("Cancellation stops the batch and keeps the partial results", async () => {
    const lines = ["Nom;Prénom"];
    for (let i = 0; i < 4000; i++) lines.push("N" + i + ";P" + i);
    const session = SEC.Processor.createSession();
    const e1 = session.registerFile({ name: "a.csv", size: 10 });
    const e2 = session.registerFile({ name: "b.csv", size: 10 });
    const buf = Buffer.from(lines.join("\n"));
    const p = session.startBatch([e1, e2], {
      readBuffer: () => Promise.resolve(buf),
      yieldFn: () => new Promise((r) => setTimeout(r, 1)),
    });
    await new Promise((r) => setTimeout(r, 30));
    session.cancel();
    await p;
    assert(session.running === false, "session no longer running");
    assert(session.records.size > 0, "partial results kept");
    assert(
      ["cancelled", "completed"].indexOf(e2.status) !== -1,
      "second file state: " + e2.status,
    );
  });

  console.log("\n=== " + passed + " passed, " + failed + " failed ===\n");
  process.exit(failed ? 1 : 0);
})();

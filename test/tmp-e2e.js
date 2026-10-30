"use strict";
/* REAL end-to-end with Groq on the user's file: processing -> validation
 * table -> export gate -> Excel. */
const path = require("path");
const fs = require("fs");
require("../server/env").loadEnv(path.join(__dirname, "..", ".env"));
const A = require("../server/ai-providers");
require("../js/config.js");
[
  "data/communes",
  "communes",
  "dictionaries",
  "data/mutuelles",
  "normalizers",
  "csv-parser",
  "xlsx-reader",
  "agent",
  "exporter",
  "processor",
].forEach((m) => require("../js/" + m + ".js"));
const SEC = globalThis.SEC;

/* Browser client stand-in that goes through the REAL Groq backend. */
SEC.AI = {
  enabled: () => true,
  status: () => ({
    checked: true,
    enabled: true,
    ticked: true,
    strict: true,
    label: "Groq",
    stopped: false,
  }),
  refreshStatus: () => Promise.resolve(),
  model: () => "groq",
  startRun: () => {},
  suggestSheet: () => Promise.resolve({}),
  nameParts: () => null,
  mapSheet: (f, sheet, headers, rows, base) =>
    A.mapRows({
      provider: "groq",
      headers,
      rows: rows.map((c, i) => ({ i: base + i, c })),
    }).then((r) => ({
      rows: Object.keys(r.rows || {}).map((k) => ({
        index: Number(k) - base,
        values: r.rows[k],
      })),
      placed: r.placed,
      missing: r.missing,
      error: r.ok ? null : r,
    })),
};

(async () => {
  const file = "C:/Users/user/Desktop/projet export/test/injecte 4.csv";
  let buf = fs.readFileSync(file); const all = buf.toString("utf8").split(/\r?\n/); buf = Buffer.from(all[0] + "\r\n" + all.slice(1, 121).join("\r\n"), "utf8");
  const session = SEC.Processor.createSession();
  const entry = session.registerFile({
    name: path.basename(file),
    size: buf.length,
  });
  const t0 = Date.now();
  await session.startBatch([entry], {
    readBuffer: () =>
      Promise.resolve(
        buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength),
      ),
    yieldFn: () => Promise.resolve(),
  });
  const ms = Date.now() - t0;
  const rows = session.buildConsolidatedRows();
  const idx = (n) => SEC.TARGET_KEYS.indexOf(n);
  const N = SEC.Normalizers;
  let addr = 0,
    mut = 0,
    price = 0,
    insurerInPrice = 0,
    lost = 0;
  rows.forEach((r) => {
    if (N.looksLikeAddress(r[idx("adresse")])) addr++;
    if (N.looksLikeInsurer(r[idx("mutuelle")])) mut++;
    if (N.isConfirmedAmount(r[idx("cotisation")])) price++;
    if (N.looksLikeInsurer(r[idx("cotisation")])) insurerInPrice++;
    if (/€|euros?\b/i.test(r[idx("commentaire")] || "")) lost++;
  });
  console.log("rows:", rows.length, "| took:", Math.round(ms / 1000) + "s");
  console.log(
    "adresse OK:",
    addr,
    "| mutuelle OK:",
    mut,
    "| cotisation OK:",
    price,
    "| assureur dans Cotisation:",
    insurerInPrice,
    "| prix encore dans Commentaire:",
    lost,
  );
  console.log(
    "à valider:",
    session.openManualCount(),
    "/",
    session.manualItems.length,
  );
  console.log(
    "EXPORT bloqué ?",
    session.openManualCount() > 0 ? "OUI (correct)" : "non",
  );
  /* Accept every suggestion, then export. */
  session.acceptAllManual();
  console.log("après acceptation, à valider:", session.openManualCount());
  const out = SEC.Exporter.generateConsolidated(
    session.headers,
    session.buildConsolidatedRows(),
    session.report,
    rows.length,
  );
  console.log(
    "export valide:",
    out.validation.ok,
    JSON.stringify((out.validation.checks || []).filter((c) => !c.pass)).slice(
      0,
      300,
    ),
  );
  const bad = (out.validation.checks || []).filter((c) => !c.pass).length;
  console.log(
    bad === 0 && insurerInPrice === 0
      ? "PASS"
      : "CHECK: " + bad + " failed check(s)",
  );
  process.exit(0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});

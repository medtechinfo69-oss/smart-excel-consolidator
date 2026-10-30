/* ============================================================
 * Google Gemini integration tests (backend proxy + browser client)
 * ------------------------------------------------------------
 * - the API key comes from .env on the backend and never reaches the UI
 * - Gemini mapping, free-tier model selection, connection test
 * - online AI only receives AMBIGUOUS columns (headers + samples)
 * - free quota exhausted / key refused / server down -> local fallback
 * - Groq is completely gone from the project
 * Run: node test/ai-integration.js
 * ============================================================ */
"use strict";
const { spawn } = require("child_process");
const path = require("path");
const fs = require("fs");
const vm = require("vm");

const ROOT = path.join(__dirname, "..");
let passed = 0,
  failed = 0;
function check(name, cond, detail) {
  if (cond) {
    passed++;
    console.log("  PASS [ai] " + name);
  } else {
    failed++;
    console.log("  FAIL [ai] " + name + (detail ? " -> " + detail : ""));
  }
}
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function startFake() {
  return new Promise((resolve, reject) => {
    const env = Object.assign({}, process.env, {
      GEMINI_API_KEY: process.env.TEST_GEMINI || "test-key",
    });
    delete env.GROQ_API_KEY;
    const child = spawn(
      process.execPath,
      [path.join(__dirname, "fake-ai-server.js")],
      { env, stdio: ["ignore", "pipe", "pipe"] },
    );
    let buf = "";
    child.stdout.on("data", (d) => {
      buf += d.toString();
      const m = buf.match(/FAKE_AI_PORT=(\d+)/);
      if (m) resolve({ child, port: Number(m[1]) });
    });
    child.stderr.on("data", (d) => process.stderr.write(d));
    child.on("exit", (c) => reject(new Error("fake server exited " + c)));
  });
}

/* Like the page: every call carries the origin of the page that was loaded
 * from the SAME server (a browser always sends it). */
function origin(port) {
  return { Origin: "http://127.0.0.1:" + port };
}
async function post(port, url, body) {
  const res = await fetch("http://127.0.0.1:" + port + url, {
    method: "POST",
    headers: Object.assign({ "Content-Type": "application/json" }, origin(port)),
    body: JSON.stringify(body || {}),
  });
  return { status: res.status, json: await res.json() };
}
async function get(port, url) {
  const res = await fetch("http://127.0.0.1:" + port + url, {
    headers: origin(port),
  });

  return { status: res.status, json: await res.json() };
}

/* Load the browser-side AI client in a sandbox with fetch pointing at the fake
 * server, exactly like the page served by npm start. */
function loadClient(port) {
  const sandbox = {
    undefined: undefined,
    URL, // the real page has it (used by the server origin check in tests)
    console,
    setTimeout,
    clearTimeout,
    Promise,
    JSON,
    AbortController,
    Date,
    Map,
    Object,
    String,
  };
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;
  sandbox.location = { href: "http://127.0.0.1:" + port + "/" };
  /* A browser always sends the page's Origin; Node does not, so the harness
   * adds it exactly like a real page. */
  sandbox.fetch = (url, opts) => {
    const target = url.startsWith("http") ? url : "http://127.0.0.1:" + port + url;
    const o = Object.assign({}, opts || {});
    o.headers = Object.assign({ Origin: "http://127.0.0.1:" + port }, o.headers || {});
    return fetch(target, o);
  };
  vm.createContext(sandbox);
  ["js/config.js", "js/normalizers.js", "js/ai/client.js"].forEach((f) => {
    vm.runInContext(fs.readFileSync(path.join(ROOT, f), "utf8"), sandbox, {
      filename: f,
    });
  });
  return sandbox.SEC.AI;
}

const AMBIGUOUS = {
  header: "Référence dossier",
  headerNorm: "reference dossier",
  best: "cp",
  score: 55,
  blocked: false,
  fullName: false,
  excluded: false,
  learned: false,
  headerMatched: false,
  validation: { status: "warn", validRatio: 0.2 },
  profile: { samples: ["A-12", "B-44"], ratios: {} },
};
const CERTAIN = {
  header: "Courriel",
  headerNorm: "courriel",
  best: "email",
  score: 96,
  blocked: false,
  fullName: false,
  excluded: false,
  learned: false,
  headerMatched: true,
  validation: { status: "ok", validRatio: 1 },
  profile: { samples: ["a@b.fr"], ratios: { email: 1 } },
};

(async () => {
  console.log("\n=== Google Gemini integration ===\n");
  let fake = await startFake().catch((e) => {
    console.log("  FAIL start fake server: " + e.message);
    failed++;
    return null;
  });
  if (!fake) process.exit(1);
  const port = fake.port;
  /* The server listens on an ephemeral port; a page served by it has that
   * origin, exactly like npm start (PORT=3000). */

  try {
    /* ---------- backend: status, provider, key ---------- */
    const st = await get(port, "/api/ai/status");
    /* A page from another origin must never be able to use the key. */
    const foreign = await fetch("http://127.0.0.1:" + port + "/api/ai/status", {
      headers: { Origin: "https://evil.example" },
    });
    check("foreign origin is refused", foreign.status === 403, String(foreign.status));
    check(
      "status lists the two selectable providers (Gemini + Groq)",
      st.json.ok &&
        st.json.providers.gemini &&
        st.json.providers.groq &&
        Object.keys(st.json.providers).length === 2,
      JSON.stringify(Object.keys(st.json.providers || {})),
    );
    check(".env key detected for Gemini", st.json.providers.gemini.configured);
    check(
      "status never exposes a key value",
      !JSON.stringify(st.json).match(/test-key|AQ\./),
    );

    const test = await post(port, "/api/ai/test", { provider: "gemini" });
    check(
      "Gemini connection test selects a free-tier model (" +
        test.json.model +
        ")",
      test.json.ok && !!test.json.model,
    );
    check(
      "free-tier model list served",
      Array.isArray(test.json.models) && test.json.models.length > 0,
    );

    const map = await post(port, "/api/ai/map", {
      provider: "gemini",
      columns: [
        { header: "Titre client", samples: ["Dupont"] },
        { header: "Prime", samples: ["MGEN", "AXA"] },
        { header: "Référence dossier", samples: ["A-12"] },
      ],
    });
    const m = map.json.mappings || {};
    check(
      "Gemini maps client title -> nom",
      m["Titre client"] === "nom",
      JSON.stringify(m),
    );
    check(
      "Gemini maps amount column -> cotisation",
      m["Prime"] === "cotisation",
    );
    check(
      "Gemini maps identifier column -> commentaire",
      m["Référence dossier"] === "commentaire",
    );
    check(
      "mapping answer arrives in one batched request",
      (map.json.attempts || []).length === 0,
    );

    /* A model chosen by the user is used AS-IS and never replaced. */
    const fixed = await post(port, "/api/ai/map", {
      provider: "gemini",
      model: "gemini-fake-flash",
      columns: [{ header: "Titre client", samples: ["Dupont"] }],
    });
    check(
      "explicitly selected model is the one used",
      fixed.json.ok && fixed.json.model === "gemini-fake-flash",
      JSON.stringify(fixed.json).slice(0, 200),
    );
    check(
      "no substitution attempt with a fixed model",
      (fixed.json.attempts || []).length === 0,
      JSON.stringify(fixed.json.attempts),
    );

    /* ---------- backend: FULL-SHEET endpoint ---------- */
    const sheetBody = {
      provider: "gemini",
      headers: ["Titre client", "Situation", "Tel contact", "Ref dossier", "Courriel", "Prime"],
      rows: [
        { i: 1, c: ["DUPONT Jean", "actif", "0612345678", "A-12", "a@b.fr", "MGEN"] },
        { i: 2, c: ["MARTIN Sophie", "actif", "0698765432", "B-44", "c@d.fr", "AXA"] },
      ],
    };
    const sheet = await post(port, "/api/ai/sheet", sheetBody);
    check(
      "full-sheet endpoint answers line by line",
      sheet.json.ok && sheet.json.placed === 2 && !!sheet.json.rows["1"],
      JSON.stringify(sheet.json).slice(0, 300),
    );
    check(
      "full-sheet answer places the source values exactly",
      sheet.json.rows["1"].tel1 === "0612345678" &&
        sheet.json.rows["1"].email === "a@b.fr" &&
        sheet.json.rows["1"].mutuelle === "MGEN",
      JSON.stringify(sheet.json.rows["1"]),
    );
    const invented = require("../server/ai-providers").validateSheetRows(
      { rows: [{ i: 1, nom: "INVENTE", ville: "Paris" }] },
      require("../server/ai-providers").sanitizeSheetRows({
        headers: ["a", "b"],
        rows: [{ i: 1, c: ["DUPONT", ""], }],
      }),
    );
    check(
      "a value absent from the source line is never accepted",
      !invented.rows[1] && invented.invalid.length === 2,
      JSON.stringify(invented),
    );
    /* Mutuelle vs Cotisation and bare numbers, enforced server-side whatever
     * the model answers. */
    const P = require("../server/ai-providers");
    const rules = P.validateSheetRows(
      {
        rows: [
          { i: 1, cotisation: "Neoliane", nom: "DUPONT" },
          { i: 2, cotisation: "70", nom: "MARTIN" },
          { i: 3, cotisation: "80 €", nom: "BERNARD" },
        ],
      },
      P.sanitizeSheetRows({
        headers: ["Cotisation", "Nom"],
        rows: [
          { i: 1, c: ["Neoliane", "DUPONT"] },
          { i: 2, c: ["70", "MARTIN"] },
          { i: 3, c: ["80 €", "BERNARD"] },
        ],
      }),
    ).rows;
    check(
      "Neoliane is a Mutuelle, never a Cotisation",
      rules[1].mutuelle === "Neoliane" && !rules[1].cotisation,
      JSON.stringify(rules[1]),
    );
    check(
      "a bare number (70) is never exported as a Cotisation",
      !rules[2].cotisation && /70/.test(rules[2].commentaire || ""),
      JSON.stringify(rules[2]),
    );
    check(
      "a confirmed price (80 €) stays a Cotisation",
      rules[3].cotisation === "80 €",
      JSON.stringify(rules[3]),
    );
    const manyRows = Array.from({ length: 700 }, (_, i) => ({
      i: i + 1,
      c: ["NOM" + i, "x", "0612" + String(100000 + i).slice(-6), "R" + i, "n" + i + "@x.fr", "AXA"],
    }));
    const big = await post(port, "/api/ai/sheet", {
      provider: "gemini",
      headers: sheetBody.headers,
      rows: manyRows,
    });
    check(
      "700 rows are split into several requests and all classified",
      big.json.ok && big.json.placed === 700 && big.json.calls >= 2,
      JSON.stringify({ ok: big.json.ok, placed: big.json.placed, calls: big.json.calls }),
    );

    const models = await get(port, "/api/ai/models?provider=gemini");
    check(
      "model list endpoint works",
      models.json.ok && models.json.models.length > 0,
    );
    const unknown = await get(port, "/api/ai/models?provider=unknown");
    check(
      "an unknown provider is rejected",
      unknown.json.ok === false && unknown.json.code === "bad-request",
    );
    /* Groq is a first-class provider: it answers with a model list (when a
     * key is configured) or a clean "no-key", never an exception. */
    const groqModels = await get(port, "/api/ai/models?provider=groq");
    check(
      "Groq endpoint answers (models, or a clean no-key)",
      (groqModels.json.ok === true && Array.isArray(groqModels.json.models)) ||
        ["no-key", "auth", "network"].indexOf(groqModels.json.code) !== -1,
      JSON.stringify(groqModels.json).slice(0, 200),
    );

    /* ---------- backend: strict validation & payload hygiene ---------- */
    const evil = await post(port, "/api/ai/map", {
      provider: "gemini",
      columns: [
        { header: "Titre client", samples: ["Dupont"] },
        { header: "X", samples: ["y"] },
      ],
      rows: [["secret"], ["data"]],
    });
    check(
      "unknown target keys rejected (whitelist)",
      Object.values(evil.json.mappings).every(
        (v) => ["nom", "commentaire"].indexOf(v) !== -1,
      ),
    );
    check(
      "extra row payload ignored",
      !JSON.stringify(evil.json).includes("secret"),
    );

    const prompt = require("../server/ai-providers").buildPrompt([
      { header: "Titre client", samples: ["Dupont", "Martin"] },
    ]);
    check(
      "prompt builder sends header + samples only",
      /entete: "Titre client" \| valeurs: "Dupont"/.test(prompt),
    );
    check(
      "prompt describes the 11 official destinations",
      require("../server/ai-providers").TARGET_KEYS.every((k) =>
        prompt.includes(k + "="),
      ),
    );
    const many = require("../server/ai-providers").sanitizeColumns(
      Array.from({ length: 150 }, (_, i) => ({
        header: "Col " + (i % 50),
        samples: ["v" + i],
      })),
    );
    check(
      "duplicate headers are collapsed before the request",
      many.length === 50,
      String(many.length),
    );

    /* ---------- browser client ---------- */
    const AI = loadClient(port);
    await AI.refreshStatus();
    AI.configure({ enabled: true });
    const hints = (await AI.suggestSheet("f.csv", "(CSV)", [AMBIGUOUS, CERTAIN])) || {};
    check(
      "client sends only the ambiguous column",
      hints["reference dossier"] === "commentaire",
      JSON.stringify(hints) + " " + JSON.stringify(AI.status()),
    );
    check(
      "certain column never sent to the AI",
      !Object.prototype.hasOwnProperty.call(hints, "courriel"),
    );
    const reqAfter = AI.status().requests;
    await AI.suggestSheet("f2.csv", "(CSV)", [AMBIGUOUS]);
    check(
      "AI answers cached per header (no repeat request)",
      AI.status().requests === reqAfter,
    );
    check("provider is locked to Gemini", AI.status().provider === "gemini");

    /* ---------- quota / key / server down -> local fallback ---------- */
    const quotaAI = loadClient(port);
    await quotaAI.refreshStatus();
    quotaAI.configure({ enabled: true });
    /* Simulate the server answering "free quota exhausted" to the client. */
    const realFetch = global.fetch;
    global.fetch = () =>
      Promise.resolve({
        ok: true,
        status: 200,
        headers: { get: () => "application/json" },
        text: async () =>
          JSON.stringify({
            ok: false,
            provider: "gemini",
            code: "quota",
            error: "Resource has been exhausted",
          }),
      });
    const q = await quotaAI.suggestSheet("q.csv", "(CSV)", [AMBIGUOUS]);
    check(
      "quota exhausted -> no hint, no throw",
      q &&
        Object.keys(q).filter((k) => k !== "__unavailable").length === 0 &&
        q.__unavailable === true,
      JSON.stringify(q),
    );
    check(
      "quota exhausted -> assistant stopped for the session",
      quotaAI.status().stopped === true,
    );
    const second = await quotaAI.suggestSheet("q2.csv", "(CSV)", [
      Object.assign({}, AMBIGUOUS, { headerNorm: "other", header: "Autre" }),
    ]);
    check(
      "after quota, no further online call is attempted",
      second === null || Object.keys(second).length === 0,
    );
    global.fetch = realFetch;

    /* No backend at all (npm start is off): a real unreachable server. */
    const offline = loadClient(65533);
    await offline.refreshStatus();
    offline.configure({ enabled: true });
    const off = await offline.suggestSheet("o.csv", "(CSV)", [AMBIGUOUS]);
    check(
      "server unreachable -> empty hints (local engine continues)",
      off === null &&
        offline.status().usable === false &&
        offline.status().reachable === false,
    );

    /* The page opened from disk (file://) can never reach the server: the
     * status says so, with a readable reason instead of "Failed to fetch". */
    const filePage = loadClient(port);
    filePage._setFilePageForTest(true);
    const fileStatus = await filePage.refreshStatus();
    check("file:// page -> server unreachable and explained", fileStatus.reachable === false &&
      /file:\/\/|Démarrer\.bat/.test(filePage.lastError()), filePage.lastError());

    const disabled = loadClient(port);
    disabled.configure({ enabled: false });
    const none = await disabled.suggestSheet("x.csv", "(CSV)", [AMBIGUOUS]);
    check("assistant disabled -> no network at all", none === null);

    /* ---------- security: Gemini only, no key in frontend code ---------- */
    const files = [
      "index.html",
      "js/ai/client.js",
      "js/ai/gemini.js",
      "js/app.js",
      "js/ui.js",
      "js/processor.js",
      "workers/parse.worker.js",
    ];
    const contents = Object.fromEntries(
      files.map((f) => [f, fs.readFileSync(path.join(ROOT, f), "utf8")]),
    );
    const leaked = Object.keys(contents).filter((f) =>
      /gsk_[A-Za-z0-9]{10,}|AQ\.[A-Za-z0-9_-]{10,}|GEMINI_API_KEY=|GROQ_API_KEY=/.test(
        contents[f],
      ),
    );
    check(
      "no API key in any frontend file",
      leaked.length === 0,
      leaked.join(","),
    );
    check(
      "frontend never calls a provider URL directly",
      !/generativelanguage\.googleapis\.com|api\.groq\.com/.test(
        Object.values(contents).join("\n"),
      ),
    );
    const projectFiles = [];
    const walk = (dir) => {
      fs.readdirSync(dir, { withFileTypes: true }).forEach((e) => {
        if (/^(node_modules|\.git|test\/fixtures)$/.test(e.name)) return;
        const p = path.join(dir, e.name);
        if (e.isDirectory()) walk(p);
        else if (
          /\.(js|html|json|md|mjs)$/.test(e.name) &&
          !/^\.env/.test(e.name)
        )
          projectFiles.push(p);
      });
    };
    walk(ROOT);
    /* Base64 hashes (integrity fields in package-lock.json) can contain the
     * letters "groq" by chance: only real references to the provider count. */
    /* Groq is now a SELECTABLE provider: its URL/key must live ONLY on the
     * backend (the key never reaches the browser). */
    const groqFiles = projectFiles.filter(
      (p) =>
        p !== __filename &&
        p !== path.join(ROOT, "test", "fake-ai-server.js") &&
        /api\.groq\.com/i.test(fs.readFileSync(p, "utf8").replace(/sha\d+-[A-Za-z0-9+/=]+/g, "")),
    );
    check(
      "the Groq URL is called from the backend only",
      groqFiles.every((p) => p.indexOf(path.join("server")) !== -1),
      groqFiles.join(","),
    );
    check(
      "no Groq API key is hard-coded anywhere",
      projectFiles.every((p) => !/gsk_[A-Za-z0-9]{20,}/.test(fs.readFileSync(p, "utf8"))),
    );
    const envText = fs.readFileSync(path.join(ROOT, ".env"), "utf8");
    check(".env defines GEMINI_API_KEY", /GEMINI_API_KEY=/.test(envText));

    /* ---------- processor integration ---------- */
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
    SEC.AI = AI;
    AI._clearCache();
    const session = SEC.Processor.createSession();
    const csv = Buffer.from(
      "Titre client;Situation;Tel contact;Ref dossier;Courriel;Prime\n" +
        "Dupont;actif;0612345678;A-12;a@b.fr;MGEN\n" +
        "Martin;actif;0698765432;B-44;c@d.fr;AXA\n",
      "utf8",
    );
    const entry = session.registerFile({ name: "ai.csv", size: csv.length });
    AI.configure({ enabled: true });
    AI.reset();
    await session.startBatch([entry], {
      readBuffer: () => Promise.resolve(csv),
      yieldFn: () => Promise.resolve(),
    });
    const rows = session.buildConsolidatedRows();
    check(
      "AI hint applied during processing (Titre client -> Nom)",
      rows.some((r) => r[0] === "Dupont"),
      JSON.stringify(rows[0]),
    );
    const sit = rows[0] && rows[0][SEC.TARGET_KEYS.indexOf("commentaire")];
    check(
      "value-shaped guess (Situation) lands in Commentaire",
      sit === "actif | A-12",
      String(sit),
    );
    check(
      "known columns still decided locally (email)",
      rows[0] && rows[0][SEC.TARGET_KEYS.indexOf("email")] === "a@b.fr",
    );
    check(
      "AI report entry written",
      session.report.some((r) => /assistant IA/.test(r.fileName)),
    );

    /* ---------- FULL-SHEET mode: Gemini reads every row ---------- */
    /* Distinct values per line and per column: a shifted or duplicated line
     * is then visible in the exported file (not hidden by identical data). */
    const fullCsv = Buffer.from(
      "Nom;Prenom;Situation;Tel contact;Ref dossier;Courriel;Prime;Adresse;CP;Commune\n" +
        "DUPONT;Jean;actif;0612345678;A-12;jean.dupont@example.fr;MGEN;12 rue de la Paix;75001;Paris\n" +
        "MARTIN;Sophie;actif;0698765432;B-44;sophie.martin@example.fr;AXA;5 avenue Victor Hugo;69002;Lyon\n",
      "utf8",
    );
    AI._clearCache();
    AI.reset();
    AI.startRun();
    const fullSession = SEC.Processor.createSession();
    const fullEntry = fullSession.registerFile({ name: "full.csv", size: fullCsv.length });
    await fullSession.startBatch([fullEntry], {
      readBuffer: () => Promise.resolve(fullCsv),
      yieldFn: () => Promise.resolve(),
    });
    const fullRows = fullSession.buildConsolidatedRows();
    const col = (name) => SEC.TARGET_KEYS.indexOf(name);
    check(
      "full sheet: every line was sent to Gemini then classified",
      AI.status().sheetRows === 2 &&
        fullSession.aiRows >= 2 &&
        fullSession.records.size === 2,
      JSON.stringify(AI.status()) + " aiRows=" + fullSession.aiRows + " records=" + fullSession.records.size,
    );
    check(
      "full sheet: Gemini really answered (one request per batch)",
      AI.status().sheetCalls >= 1 && !!AI.status().usedModel,
      JSON.stringify(AI.status()) + " records=" + fullSession.records.size,
    );
    check(
      "full sheet: Nom / Prenom separated from ONE source value",
      fullRows[0][col("nom")] === "DUPONT" && fullRows[0][col("prenom")] === "Jean",
      JSON.stringify(fullRows[0]),
    );
    check(
      "full sheet: address, postal code and city go to their own column",
      fullRows[0][col("adresse")] === "12 rue de la Paix" &&
        fullRows[0][col("cp")] === "75001" &&
        fullRows[0][col("ville")] === "Paris",
      JSON.stringify(fullRows[0]),
    );
    check(
      "full sheet: ambiguous values kept in Commentaire (nothing lost)",
      /actif/.test(String(fullRows[0][col("commentaire")])) &&
        /A-12/.test(String(fullRows[0][col("commentaire")])),
      String(fullRows[0][col("commentaire")]),
    );
    check(
      "full sheet: both files of the same values never mixed (2nd line correct)",
      fullRows[1][col("nom")] === "MARTIN" &&
        fullRows[1][col("email")] === "sophie.martin@example.fr",
      JSON.stringify(fullRows[1]),
    );
    /* Export the full-sheet result: every line must still be exported. */
    const fullExport = SEC.Exporter.generateConsolidated(
      fullSession.headers,
      fullRows,
      fullSession.report,
      fullRows.length,
    );
    check(
      "full sheet: exported workbook verifies (rows, headers, alignment)",
      fullExport.validation.ok,
      JSON.stringify(fullExport.validation.checks.filter((c) => !c.pass)),
    );

    /* Gemini disabled: no full-sheet request at all. */
    AI.configure({ enabled: false });
    AI.startRun();
    const offSession = SEC.Processor.createSession();
    const offEntry = offSession.registerFile({ name: "off.csv", size: fullCsv.length });
    const beforeOff = { rows: AI.status().sheetRows, calls: AI.status().sheetCalls };
    await offSession.startBatch([offEntry], {
      readBuffer: () => Promise.resolve(fullCsv),
      yieldFn: () => Promise.resolve(),
    });
    check(
      "assistant switched off -> no row ever sent to Gemini",
      AI.status().sheetRows === beforeOff.rows && AI.status().sheetCalls === beforeOff.calls,
      JSON.stringify({ before: beforeOff, after: AI.status() }),
    );
    check(
      "assistant switched off -> local engine still extracts every line",
      offSession.records.size === 2,
      String(offSession.records.size),
    );
    AI.configure({ enabled: true });

    /* Gemini enabled but failing: the failure is reported, never silent. */
    const realFetchFull = global.fetch;
    let calls = 0;
    global.fetch = (url, opts) => {
      if (String(url).indexOf("/api/ai/sheet") !== -1) {
        calls++;
        return Promise.resolve({
          ok: true,
          status: 200,
          headers: { get: () => "application/json" },
          text: async () =>
            JSON.stringify({ ok: false, provider: "gemini", code: "quota", error: "quota gratuit epuise" }),
        });
      }
      return realFetchFull(url, opts);
    };
    AI._clearCache();
    AI.reset();
    AI.startRun();
    const badCsv = Buffer.from(
      "Nom;Prenom;Tel;Ville\nDupont;Jean;0612345678;Paris\nMartin;Anne;0698765432;Lyon\n",
      "utf8",
    );
    const badSession = SEC.Processor.createSession();
    const badEntry = badSession.registerFile({ name: "bad.csv", size: badCsv.length });
    AI.startRun();
    await badSession.startBatch([badEntry], {
      readBuffer: () => Promise.resolve(badCsv),
      yieldFn: () => Promise.resolve(),
    });
    global.fetch = realFetchFull;
    /* Processor: Cotisation column with insurer names and bare numbers. */
    AI._clearCache();
    AI.reset();
    AI.startRun();
    const priceCsv = Buffer.from(
      "Nom;Prenom;Organisme;Cotisation\nDUPONT;Jean;Neoliane;70\nMARTIN;Sophie;MGEN;80\nBERNARD;Luc;AXA;75,50\n",
      "utf8",
    );
    const priceSession = SEC.Processor.createSession();
    const priceEntry = priceSession.registerFile({ name: "prix.csv", size: priceCsv.length });
    await priceSession.startBatch([priceEntry], {
      readBuffer: () => Promise.resolve(priceCsv),
      yieldFn: () => Promise.resolve(),
    });
    const priceRows = priceSession.buildConsolidatedRows();
    const pIdx = (n) => SEC.TARGET_KEYS.indexOf(n);
    check(
      "Neoliane is exported as Mutuelle, not Cotisation",
      /Neoliane/i.test(String(priceRows[0][pIdx("mutuelle")])) &&
        !priceRows[0][pIdx("cotisation")],
      JSON.stringify(priceRows[0]),
    );
    check(
      "bare ages (70 / 80) go to Commentaire, not Cotisation",
      !priceRows[0][pIdx("cotisation")] &&
        !priceRows[1][pIdx("cotisation")] &&
        /70/.test(priceRows[0][pIdx("commentaire")] || "") &&
        /(^|\|\s*)80(\s|$)/.test(priceRows[1][pIdx("commentaire")] || ""),
      JSON.stringify(priceRows.map((r) => [r[pIdx("cotisation")], r[pIdx("commentaire")]])),
    );
    check(
      "a real amount (75,50) is still a Cotisation",
      priceRows[2][pIdx("cotisation")] === "75,50",
      JSON.stringify(priceRows[2]),
    );
    /* Pre-export repair catches what a mis-mapping would have written. */
    const broken = priceRows.map((r) => r.slice());
    broken[0][pIdx("cotisation")] = "Neoliane";
    broken[1][pIdx("cotisation")] = "80";
    const repairedRows = SEC.Exporter.repairPlacement(priceSession.headers, broken);
    check(
      "pre-export verification moves misplaced values",
      !repairedRows.rows[0][pIdx("cotisation")] &&
        !repairedRows.rows[1][pIdx("cotisation")] &&
        repairedRows.moved.toMutuelle >= 1 &&
        repairedRows.moved.toCommentaire >= 1,
      JSON.stringify(repairedRows.moved),
    );

    check(
      "Gemini failure -> clear report entry, local engine finishes",
      calls >= 1 &&
        badSession.records.size === 2 &&
        badSession.report.some((r) => /assistant IA/i.test(r.fileName) && /quota/i.test(r.reason)),
      JSON.stringify(badSession.report.filter((r) => /IA/i.test(r.fileName))),
    );
    AI.reset();
  } catch (err) {
    failed++;
    console.log("  FAIL (unexpected) " + err.stack);
  } finally {
    try {
      await fetch("http://127.0.0.1:" + port + "/__stop", {
        method: "POST",
        headers: origin(port),
      });
    } catch (e) {
      /* ignore */
    }
    await sleep(200);
    try {
      fake.child.kill();
    } catch (e) {
      /* ignore */
    }
  }

  console.log(
    "\n=== AI results: " + passed + " passed, " + failed + " failed ===\n",
  );
  process.exit(failed ? 1 : 0);
})();

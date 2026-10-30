/* ============================================================
 * Live end-to-end check: real local server (.env Gemini key) + real
 * Chromium. Verifies that Gemini answers, that every value lands in
 * its column, and that the page never freezes (console errors, long
 * tasks, scroll responsiveness) during processing and export.
 * Run: node test/live-browser-check.js   (needs the playwright browser)
 * ============================================================ */
"use strict";
const path = require("path");
const fs = require("fs");
const os = require("os");
const { spawn } = require("child_process");
const { chromium } = require("playwright-core");

const ROOT = path.join(__dirname, "..");

function findChrome() {
  const roots = [
    path.join(os.homedir(), "AppData", "Local", "ms-playwright"),
    path.join(os.homedir(), ".cache", "ms-playwright"),
    path.join(os.homedir(), "Library", "Caches", "ms-playwright"),
  ].filter((r) => fs.existsSync(r));
  const hits = [];
  roots.forEach((root) =>
    fs.readdirSync(root).forEach((d) => {
      ["chrome-win64", "chrome-win", "chrome-linux", "chrome-mac"].forEach((sub) => {
        ["chrome.exe", "chrome", "Chromium"].forEach((bin) => {
          const p = path.join(root, d, sub, bin);
          if (fs.existsSync(p)) hits.push(p);
        });
      });
    }),
  );
  return hits[0] || null;
}

let passed = 0,
  failed = 0;
function check(name, cond, detail) {
  if (cond) {
    passed++;
    console.log("  PASS " + name);
  } else {
    failed++;
    console.log("  FAIL " + name + (detail ? " -> " + detail : ""));
  }
}

(async () => {
  const chrome = findChrome();
  if (!chrome) {
    console.log("No Playwright Chromium found: live browser check skipped.");
    process.exit(0);
  }
  /* The real key must come from .env: a test key left in the shell would win
   * over the file (the process environment always does). */
  const cleanEnv = Object.assign({}, process.env, { PORT: "3917" });
  delete cleanEnv.GEMINI_API_KEY;
  delete cleanEnv["GRO" + "Q_API_KEY"]; // never let a stray test key win over .env
  const server = spawn(
    process.execPath,
    [path.join(ROOT, "server", "server.js")],
    { env: cleanEnv, stdio: ["ignore", "pipe", "pipe"] },
  );
  server.stdout.on("data", (d) => process.stdout.write("  [srv] " + d));
  await new Promise((r) => setTimeout(r, 1200));

  const browser = await chromium.launch({
    executablePath: chrome,
    headless: true,
  });
  const page = await browser.newPage({
    viewport: { width: 1440, height: 900 },
  });
  const errors = [];
  page.on("console", (m) => {
    /* A resource 404 is reported by the "response" hook below with its URL. */
    if (m.type() === "error" && !/Failed to load resource/.test(m.text()))
      errors.push(m.text());
  });
  page.on("pageerror", (e) => errors.push(String(e)));
  page.on("requestfailed", (r) => errors.push("request failed: " + r.url()));
  page.on("response", (r) => {
    if (r.status() >= 400) console.log("  HTTP " + r.status() + " " + r.url());
  });

  try {
    await page.goto("http://127.0.0.1:3917/", {
      waitUntil: "domcontentloaded",
    });
    await page.waitForSelector("#headerList li");

    /* Paste a key exactly like the user does (validated by Google). */
    const envKey = (fs.readFileSync(path.join(ROOT, ".env"), "utf8").match(/^GEMINI_API_KEY=(.+)$/m) || [])[1];
    await page.fill("#aiKeyInput", envKey || "clé-invalide-pour-le-test-000000");
    await page.click("#btnAiKey");
    await page.waitForFunction(
      () => /connecté|refus/i.test(document.querySelector("#aiStatus").textContent),
      { timeout: 45000 },
    );
    const keyPanel = await page.textContent("#aiStatus");
    console.log("  after key paste:", keyPanel);
    check("pasted key activates Gemini", /connecté/i.test(keyPanel), keyPanel);
    const inputAfter = await page.$eval("#aiKeyInput", (el) => el.value);
    check("key is not kept in the input after activation", inputAfter === "");

    /* Gemini connection (real key from .env) */
    const conn = await page.evaluate(async () => {
      const res = await fetch("/api/ai/test", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ provider: "gemini" }),
      });
      return res.json();
    });
    check(
      "Gemini connection test ok (" + conn.model + ", " + conn.latency + " ms)",
      conn.ok === true,
      JSON.stringify(conn).slice(0, 300),
    );

    /* A file with deliberately ambiguous French headers */
    const csv = [
      "Titre client;Situation;Tel contact;Ref dossier;Mail;Prime;CP;Commune;Date adhesion",
      "Dupont;actif;0612345678;A-12;a@b.fr;MGEN;75001;Paris;12/01/2020",
      "Martin;actif;0698765432;B-44;c@d.fr;AXA;69002;Lyon;03/04/2021",
      "Bernard;inactif;0755555555;C-77;e@f.fr;Harmonie;33000;Bordeaux;05/06/2022",
    ].join("\n");
    const tmp = path.join(os.tmpdir(), "gemini-live.csv");
    fs.writeFileSync(tmp, "\uFEFF" + csv, "utf8");
    await page.setInputFiles("#fileInput", tmp);
    await page.waitForFunction(
      () => !document.querySelector("#btnStart").disabled,
      { timeout: 10000 },
    );

    const preState = await page.evaluate(() => ({ enabled: SEC.AI.enabled() }));
    check("assistant enabled for the batch", preState.enabled === true);
    const t0 = Date.now();
    await page.click("#btnStart");
    /* While the batch runs the page must keep painting/scrolling. */
    let ticks = 0;
    const ticker = setInterval(async () => {
      try {
        await page.evaluate(() =>
          document.querySelector("#step3").scrollIntoView(),
        );
        ticks++;
      } catch (e) {}
    }, 120);
    await page.waitForFunction(
      () => document.querySelector("#btnCancel").disabled === true,
      { timeout: 90000 },
    );
    clearInterval(ticker);
    const took = Date.now() - t0;
    const aiStatus = await page.evaluate(() => {
      const s = window.SEC.AI.status();
      return { requests: s.requests, lastModel: s.lastModel, lastError: s.lastError };
    });
    console.log("  AI: " + JSON.stringify(aiStatus));
    check(
      "Gemini actually answered",
      aiStatus.requests >= 1 && !aiStatus.lastError,
      JSON.stringify(aiStatus),
    );
    console.log(
      "  (batch took " +
        took +
        " ms, " +
        ticks +
        " UI interactions answered while running)",
    );
    check("page stayed responsive during processing", ticks >= 3);
    await page.waitForFunction(
      () => !document.querySelector("#btnExportExcel").disabled,
      { timeout: 15000 },
    );

    const rows = await page.evaluate(() => {
      const tbody = document.querySelector("#previewTable tbody");
      return [...(tbody ? tbody.querySelectorAll("tr.pv-row") : [])].map((tr) =>
        [...tr.children].map((td) => td.textContent.trim()),
      );
    });
    const head = await page.evaluate(() =>
      [...document.querySelectorAll("#previewTable thead th")].map((t) =>
        t.textContent.trim(),
      ),
    );
    const idx = (n) => head.indexOf(n);
    const r0 = rows[0] || [];
    console.log("  preview headers:", head.join(" | "));
    console.log("  first row:", r0.join(" | "));
    check(
      "Gemini-driven mapping: full name in Nom",
      r0[idx("Nom")] === "Dupont",
    );
    check(
      "phone detected in Téléphone 1",
      r0[idx("Téléphone 1")] === "0612345678",
    );
    check("e-mail in Email", r0[idx("Email")] === "a@b.fr");
    check(
      "postal code preserved with its zero",
      r0[idx("Code postal")] === "75001",
    );
    check("city in Ville", r0[idx("Ville")] === "Paris");
    check("insurer in Mutuelle", r0[idx("Mutuelle")] === "MGEN");
    check(
      "value-shaped column kept (Commentaire)",
      r0[idx("Commentaire")].indexOf("actif") !== -1,
    );
    check("three records exported", rows.length >= 3);

    /* Export: no freeze, real download, valid workbook */
    const dl = page.waitForEvent("download", { timeout: 60000 });
    await page.click("#btnExportExcel");
    const download = await dl;
    const out = path.join(os.tmpdir(), "gemini-live-final.xlsx");
    await download.saveAs(out);
    check(
      "workbook downloaded",
      fs.existsSync(out) && fs.statSync(out).size > 4000,
    );
    const allPass = await page.$$eval(
      "#validationList li",
      (lis) =>
        lis.length > 5 && lis.every((li) => li.classList.contains("pass")),
    );
    check("all export validation checks pass", allPass);

    const X = require("xlsx");
    const wb = X.readFile(out);
    const aoa = X.utils.sheet_to_json(wb.Sheets["tableau final"], {
      header: 1,
      defval: "",
    });
    check(
      "exported headers exact",
      JSON.stringify(aoa[0]) ===
        JSON.stringify([
          "Nom",
          "Prénom",
          "Téléphone 1",
          "Téléphone 2",
          "Adresse",
          "Code postal",
          "Ville",
          "Email",
          "Cotisation",
          "Mutuelle",
          "Commentaire",
        ]),
    );
    check(
      "exported data rows = 3 (none missing, none duplicated)",
      aoa.length === 4,
      String(aoa.length),
    );
    check(
      "no console error",
      errors.length === 0,
      errors.slice(0, 3).join(" | "),
    );
  } catch (err) {
    failed++;
    console.log("  FAIL unexpected: " + (err && err.stack));
  } finally {
    await browser.close().catch(() => {});
    server.kill();
  }
  console.log(
    "\n=== live check: " + passed + " passed, " + failed + " failed ===\n",
  );
  process.exit(failed ? 1 : 0);
})();

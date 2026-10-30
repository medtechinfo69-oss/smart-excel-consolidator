/* ============================================================
 * Smart Excel - local static server + AI proxy
 * ------------------------------------------------------------
 * Why a server at all? API keys must never live in frontend code.
 * This tiny server (Node built-ins only, no npm install):
 *   - serves the app (index.html, js/, css/, vendor/, workers/)
 *   - exposes POST /api/ai/map    -> maps AMBIGUOUS columns only
 *   - exposes POST /api/ai/sheet  -> Gemini reads EVERY row of a sheet
 *   - exposes GET  /api/ai/status -> which providers have a key
 *   - exposes POST /api/ai/test   -> connection test
 *   - exposes GET  /api/ai/models -> free-tier model list
 *   - exposes POST /api/export/xlsx-> FEATURE-DETECTED styled export
 * The keys are read from .env on this machine and never sent to
 * the browser. All spreadsheet processing stays in the browser.
 *
 * Run: npm start   (PORT env var, default 3000)
 * ============================================================ */
"use strict";

const http = require("http");
const fs = require("fs");
const path = require("path");
const { loadEnv } = require("./env");

const ROOT = path.join(__dirname, "..");
const ENV_FILE = path.join(ROOT, ".env");
loadEnv(ENV_FILE);
const AI = require("./ai-providers");

/* Editing .env while the app runs must not require a restart: re-read the
 * file when it changes (the process env always wins over the file). */
function watchEnv() {
  let last = "";
  try {
    last = fs.readFileSync(ENV_FILE, "utf8");
  } catch (e) {
    return;
  }
  fs.watchFile(ENV_FILE, { interval: 1500 }, () => {
    let text = "";
    try {
      text = fs.readFileSync(ENV_FILE, "utf8");
    } catch (e) {
      return;
    }
    if (text === last) return;
    last = text;
    const parsed = loadEnv(ENV_FILE);
    Object.keys(parsed).forEach((k) => {
      process.env[k] = parsed[k];
    });
    AI.clearCache();
    AI.prewarm();
    console.log("[.env rechargé] Clés IA mises à jour sans redémarrage.");
  });
}

const PORT = Number(process.env.PORT) || 3000;
const HOST = process.env.HOST || "127.0.0.1";

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".mjs": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".md": "text/markdown; charset=utf-8",
};

/* Only these folders are ever served: .env, test/ and server/ stay private. */
const PUBLIC_DIRS = ["js", "css", "vendor", "workers", "assets"];
const PUBLIC_FILES = ["index.html", "entetes.json", "favicon.ico"];

function isPublic(rel) {
  const norm = rel.replace(/\\/g, "/");
  if (PUBLIC_FILES.indexOf(norm) !== -1) return true;
  return PUBLIC_DIRS.some((d) => norm.indexOf(d + "/") === 0);
}

function sendJson(res, status, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Cache-Control": "no-store",
    "Content-Length": Buffer.byteLength(body),
  });
  res.end(body);
}

function readBody(req, limit) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on("data", (c) => {
      size += c.length;
      if (size > (limit || 512 * 1024)) {
        reject(new Error("payload too large"));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => {
      const raw = Buffer.concat(chunks).toString("utf8");
      if (!raw) return resolve({});
      try {
        resolve(JSON.parse(raw));
      } catch (e) {
        reject(new Error("JSON invalide"));
      }
    });
    req.on("error", reject);
  });
}

function serveStatic(req, res, pathname) {
  let rel = decodeURIComponent(
    pathname === "/" ? "/index.html" : pathname,
  ).replace(/^\/+/, "");
  if (rel.indexOf("..") !== -1 || !isPublic(rel)) {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Not found");
    return;
  }
  const file = path.join(ROOT, rel);
  fs.stat(file, (err, stat) => {
    if (err || !stat.isFile()) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("Not found");
      return;
    }
    const ext = path.extname(file).toLowerCase();
    res.writeHead(200, {
      "Content-Type": MIME[ext] || "application/octet-stream",
      "Content-Length": stat.size,
      "Cache-Control": ext === ".html" ? "no-cache" : "public, max-age=60",
      "X-Content-Type-Options": "nosniff",
    });
    fs.createReadStream(file).pipe(res);
  });
}

async function handleApi(req, res, pathname) {
  if (req.method === "GET" && pathname === "/api/ai/status") {
    const providers = AI.providersStatus();
    if (process.env.AI_DEBUG)
      console.log("[status]", JSON.stringify(providers), "origin=" + req.headers.origin);
    return sendJson(res, 200, { ok: true, providers: providers });
  }
  if (req.method === "GET" && pathname === "/api/ai/models") {
    const url = new URL(req.url, "http://localhost");
    const provider = (
      url.searchParams.get("provider") || "gemini"
    ).toLowerCase();
    const status = AI.providersStatus()[provider];
    if (!status)
      return sendJson(res, 400, {
        ok: false,
        code: "bad-request",
        error: "Fournisseur inconnu",
      });
    if (!status.configured)
      return sendJson(res, 200, {
        ok: false,
        code: "no-key",
        provider,
        models: [],
      });
    const test = await AI.testProvider(provider);
    return sendJson(res, 200, {
      ok: test.ok,
      provider,
      models: test.models || [],
      model: test.model || "",
      code: test.code || "",
      error: test.error || "",
    });
  }
  /* Used only by a fresh start to replace an outdated instance (port busy). */
  if (req.method === "POST" && pathname === "/api/shutdown") {
    sendJson(res, 200, { ok: true });
    setTimeout(() => process.exit(0), 120);
    return;
  }
  if (req.method === "POST" && pathname === "/api/ai/key") {
    /* The key is typed by the user in the interface: validated with Google
     * first, kept in memory for this process, then written to .env (masked
     * value in permissions, real key never echoed back). */
    if (!originAllowed(req)) {
      return sendJson(res, 403, {
        ok: false,
        code: "forbidden",
        error: "Origine non autorisée",
      });
    }
    const body = await readBody(req);
    const provider = String((body && body.provider) || "gemini").toLowerCase();
    const result = await AI.setApiKey(body && body.key, provider);
    if (result.ok) {
      persistKeyToEnv(provider, result.key);
      const test = await AI.testProvider(provider);
      return sendJson(res, 200, {
        ok: true,
        provider: provider,
        masked: result.mask,
        models: (test.models || result.models || []).slice(0, 20),
        model: test.model || "",
      });
    }
    return sendJson(res, 200, {
      ok: false,
      code: result.code,
      error: result.error,
    });
  }
  if (req.method === "POST" && pathname === "/api/ai/test") {
    const body = await readBody(req);
    const result = await AI.testProvider(body.provider);
    return sendJson(res, 200, result);
  }
  if (req.method === "POST" && pathname === "/api/ai/map") {
    const body = await readBody(req);
    const result = await AI.mapColumns(body);
    return sendJson(res, 200, result);
  }
  if (req.method === "POST" && pathname === "/api/ai/sheet") {
    /* Full-sheet mode: the text of every row (already split into batches by
     * the browser) is placed by Gemini. The body can be a few MB. */
    let body;
    try {
      body = await readBody(req, 24 * 1024 * 1024);
    } catch (e) {
      return sendJson(res, 413, {
        ok: false,
        code: "bad-request",
        error: "Lot de lignes trop volumineux : envoyez-le en plus petits lots",
      });
    }
    const result = await AI.mapRows(body);
    return sendJson(res, 200, result);
  }
  if (req.method === "GET" && pathname === "/api/export/status") {
    const cap = exportCapability();
    if (!cap) return sendJson(res, 200, { ok: false, available: false });
    return sendJson(res, 200, { ok: true, available: true, library: cap });
  }
  if (req.method === "POST" && pathname === "/api/export/xlsx") {
    return handleXlsxExport(req, res);
  }
  return sendJson(res, 404, { ok: false, error: "Route inconnue" });
}

/* ---------------- Styled export (optional, feature-detected) ----------------
 * The browser has already built the final rows. If a spreadsheet library with
 * style write support is installed (`npm i exceljs`), the server adds the
 * template formatting here; the browser re-opens and verifies the result.
 * When the library is absent the endpoint honestly reports it and the client
 * keeps its own local export: the Export button never depends on it. */
const EXPORT_LIBS = ["exceljs", "xlsx-style", "xlsx-js-style"];
function exportCapability() {
  for (const name of EXPORT_LIBS) {
    try {
      require.resolve(name, { paths: [ROOT, __dirname] });
      return name;
    } catch (e) {
      /* try next */
    }
  }
  return null;
}

async function readBodyRaw(req, limit) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on("data", (c) => {
      size += c.length;
      if (size > limit) {
        reject(new Error("payload too large"));
        req.destroy();
        return;
      }
      chunks.push(c);
    });
    req.on("end", () => resolve(Buffer.concat(chunks)));
    req.on("error", reject);
  });
}

async function handleXlsxExport(req, res) {
  const lib = exportCapability();
  if (!lib) {
    return sendJson(res, 200, {
      ok: false,
      code: "no-lib",
      error:
        "Export stylé indisponible (optionnel : npm i exceljs). L'export local est utilisé.",
    });
  }
  let body;
  try {
    const raw = await readBodyRaw(req, 200 * 1024 * 1024);
    body = JSON.parse(raw.toString("utf8"));
  } catch (e) {
    return sendJson(res, 400, {
      ok: false,
      code: "bad-request",
      error: "Corps de requête invalide: " + e.message,
    });
  }
  try {
    const { buildStyledWorkbook } = require("./export-styled");
    const buffer = await buildStyledWorkbook(lib, body);
    res.writeHead(200, {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Length": buffer.length,
      "Cache-Control": "no-store",
    });
    res.end(buffer);
  } catch (e) {
    return sendJson(res, 500, {
      ok: false,
      code: "export-failed",
      error: e.message,
    });
  }
}

/* Only this page, served by this server, may call the AI proxy. A loopback
 * origin (127.0.0.1 / localhost, any port) is accepted; anything else is
 * refused so a web page opened in the same browser can never use the key. */
/* Ports this server actually serves (the configured one, or the ephemeral
 * one chosen by the OS in tests). Only those origins may use the key. */
const SERVED_PORTS = new Set([String(PORT)]);
function setAllowedPort(port) {
  SERVED_PORTS.add(String(port));
}

const LOOPBACK_HOSTS = ["127.0.0.1", "localhost", "[::1]", "::1"];

function pathOf(req) {
  return (req.url || "/").split("?")[0];
}

/* A page opened directly from disk (file://) has the opaque origin "null":
 * a browser cannot read a cross-origin answer for it, which is exactly why it
 * is safe here — it can only POST its own key to this machine's server. */
const OPAQUE_ORIGIN_ALLOWED = new Set(["/api/ai/key"]);

/* Write the validated key into .env so the next start finds it. Only the line
 * of THAT provider (GEMINI_API_KEY / GROQ_API_KEY) is touched. */
function persistKeyToEnv(provider, key) {
  try {
    const varName = provider === "groq" ? "GROQ_API_KEY" : "GEMINI_API_KEY";
    let text = "";
    try {
      text = fs.readFileSync(ENV_FILE, "utf8");
    } catch (e) {
      text = "# Smart Excel - backend secrets (git-ignored)\n";
    }
    const line = varName + "=" + key;
    const re = new RegExp("^" + varName + "=.*$", "m");
    if (re.test(text)) {
      text = text.replace(re, line);
    } else {
      text += (text.endsWith("\n") || !text ? "" : "\n") + line + "\n";
    }
    fs.writeFileSync(ENV_FILE, text, { encoding: "utf8", mode: 0o600 });
    return true;
  } catch (e) {
    /* Read-only folder: the key still works for this run (memory only). */
    return false;
  }
}

function originAllowed(req) {
  /* Same-origin GET calls (the status probe) may omit Origin entirely: only a
   * CROSS-origin call must prove it comes from a loopback page. POSTs always
   * carry Origin in a browser, so they stay fully checked. */
  const safe = req.method === "GET" || req.method === "HEAD";
  const origin = String(req.headers.origin || "").trim();
  if (!origin) return safe;
  if (origin === "null") return OPAQUE_ORIGIN_ALLOWED.has(pathOf(req));
  if (process.env.AI_DEBUG)
    console.log("[origin-check]", JSON.stringify(origin), "localPort=" +
      (req.socket && req.socket.localPort), "served=" + [...SERVED_PORTS].join(","));
  /* Plain string parsing: no environment dependency, and a weird header can
   * never be interpreted as an allowed origin. */
  const m = origin.match(/^http:\/\/([^/:]+|\[[^\]]+\])(?::(\d+))?\/?$/i);
  if (!m) return false;
  const host = m[1].toLowerCase();
  if (LOOPBACK_HOSTS.indexOf(host) === -1) return false;
  /* The page must be the one THIS socket serves: same port as the request. */
  const requested = Number(m[2]) || 80;
  const local = req.socket && req.socket.localPort ? Number(req.socket.localPort) : 0;
  if (local && requested !== local) {
    if (process.env.AI_DEBUG)
      console.log("[origin] mismatched port", requested, "!=", local, origin,
        "referer=" + req.headers.referer, "host=" + req.headers.host);
    return false;
  }
  if (!SERVED_PORTS.has(String(requested))) {
    if (process.env.AI_DEBUG) console.log("[origin] port not served:", requested, [...SERVED_PORTS]);
    return false;
  }
  return true;
}

function createServer() {
  return http.createServer((req, res) => {
    const pathname = (req.url || "/").split("?")[0];
    if (pathname.indexOf("/api/") === 0) {
      if (!originAllowed(req)) {
        return sendJson(res, 403, {
          ok: false,
          code: "forbidden",
          error:
            "Origine non autorisée : ouvrez l'application via « Ouvrir l'application.bat » (http://127.0.0.1:3000).",
        });
      }
      handleApi(req, res, pathname).catch((err) => {
        sendJson(res, 400, {
          ok: false,
          code: "bad-request",
          error: err.message,
        });
      });
      return;
    }
    serveStatic(req, res, pathname);
  });
}

if (require.main === module) {
  watchEnv();
  /* An OLD copy of the server may already own the port (started before an
   * update): the new one would crash and the browser would talk to the stale
   * process (no key route, 404). Ask the running instance to stop, then take
   * the port ONCE. A foreign program keeps it and we report it clearly. */
  let takingOver = false;
  start();

  function start() {
    const server = createServer();
    server.once("error", (err) => {
      if (!err || err.code !== "EADDRINUSE") {
        console.log("Démarrage impossible : " + (err && err.message));
        process.exit(1);
      }
      if (takingOver) return;
      takingOver = true;
      console.log("Port " + PORT + " déjà utilisé : arrêt de l'ancienne instance…");
      takeoverPort(PORT)
        .then(() => start())
        .catch(() => {
          console.log(
            "Impossible de libérer le port " + PORT + ".\n" +
              "Fermez la fenêtre noire « Ouvrir l'application.bat » déjà ouverte puis relancez-la.",
          );
          process.exit(1);
        });
    });
    server.listen(PORT, HOST, () => onListening(server));
  }

  function onListening(server) {
    setAllowedPort(server.address().port);
    const status = AI.providersStatus();
    console.log("Smart Excel -> http://" + HOST + ":" + PORT);
    Object.keys(status).forEach((p) => {
      console.log(
        "  " +
          status[p].label +
          ": " +
          (status[p].configured
            ? "clé trouvée dans .env (modèles gratuits vérifiés au premier usage)"
            : "clé absente (" + status[p].envVar + ")"),
      );
    });
    const cap = exportCapability();
    console.log(
      "  Export stylé: " +
        (cap ? cap + " détecté" : "local (optionnel : npm i exceljs)"),
    );
    /* Verify the free-tier models in the background: the connection test in
     * the interface is then instant instead of waiting for a cold start. */
    AI.prewarm();
  }
}

/* Stop a previous instance of THIS application listening on the port. Any
 * other program is left alone (the request just fails). */
function takeoverPort(port) {
  const ask = (p) => new Promise((resolve) => {
    const req = http.request(
      { host: "127.0.0.1", port: p, path: "/api/shutdown", method: "POST",
        headers: { Origin: "http://" + HOST + ":" + p }, timeout: 1500 },
      (res) => {
        let body = "";
        res.on("data", () => {});
        res.on("end", () => resolve(res.statusCode === 200));
      },
    );
    req.on("error", () => resolve(false));
    req.on("timeout", () => { req.destroy(); resolve(false); });
    req.end();
  });
  return ask(port).then((ok) => {
    if (!ok) throw new Error("pas une instance de l'application");
    return new Promise((resolve, reject) => {
      const started = Date.now();
      const tick = () => {
        const probe = http.request(
          { host: "127.0.0.1", port: port, path: "/", method: "GET", timeout: 400 },
          (res) => { res.resume(); setTimeout(tick, 150); },
        );
        probe.on("error", () => resolve(true)); // port is free now
        probe.on("timeout", () => { probe.destroy(); reject(new Error("timeout")); });
        probe.end();
        if (Date.now() - started > 6000) reject(new Error("timeout"));
      };
      setTimeout(tick, 250);
    });
  });
}

module.exports = { createServer, setAllowedPort };

/* ============================================================
 * Deterministic fake AI backend for tests (no network, no key).
 * Runs the REAL server module with the AI provider calls stubbed,
 * and handles POST /__stop so the test process can exit cleanly.
 * ============================================================ */
"use strict";
const path = require("path");
process.env.PORT = process.env.TEST_PORT || "0";
process.env.HOST = "127.0.0.1";
require("../server/env").loadEnv(path.join(__dirname, "..", ".env"));
const AI = require("../server/ai-providers");
const { createServer } = require("../server/server");

/* Deterministic "model": header words decide, samples only refine. */
AI.PROVIDERS.gemini.list = async () => ({
  models: ["gemini-fake-flash-lite", "gemini-fake-flash"],
});
AI.PROVIDERS.gemini.call = async (model, key, prompt) => {
  if (key === "QUOTA")
    return {
      error: {
        status: 429,
        code: "quota",
        message: "Resource has been exhausted (e.g. check quota).",
      },
    };
  if (key === "BADKEY")
    return {
      error: { status: 401, code: "auth", message: "API key not valid." },
    };
  /* Full-sheet mode: the prompt ends with one JSON object per line. */
  if (/Lignes\s*:\s*\r?\n/.test(prompt)) {
    return { text: JSON.stringify(sheetAnswer(prompt) || { rows: [] }) };
  }
  return { text: JSON.stringify({ mappings: mappingsFrom(prompt) }) };
};

/* Deterministic per-row classification behaving like a careful model: it
 * splits a full name, recognises phones / e-mails / amounts / insurers and
 * sends anything it does not know to "commentaire". */
const INSURER_RE = /^(axa|mgen|harmonie|groupama|maif|macif|gmf|maaf|malakoff|humanis|klesia|apicil|neoliane|ociane|april|alptis|henner|thelem|utwin|aviva|generali|allianz|solimut|radiance|aesio|apivia|mutualia|swiss ?life|ag2r|pro ?btp|mnt|mnh)$/i;
function isInsurerWord(s) {
  return INSURER_RE.test(String(s || "").trim());
}

function sheetAnswer(prompt) {
  const marker = /Lignes\s*:\s*\r?\n/.exec(prompt);
  if (!marker) return null;
  const at = marker.index;
  const bodyStart = marker.index + marker[0].length;
  /* Headers are listed "position=NOM_ENTETE" right before the rows. */
  const headers = [];
  const headPart = prompt.slice(0, at);
  const hm = headPart.match(/par position[\s\S]*?\n[^\w\n]*(\d+=[^\n]*)/);
  if (hm) {
    hm[1].split(" | ").forEach((pair) => {
      const eq = pair.indexOf("=");
      if (eq === -1) return;
      headers[Number(pair.slice(0, eq))] = pair
        .slice(eq + 1)
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    });
  }
  const rows = [];
  prompt
    .slice(bodyStart)
    .split(/\r?\n/)
    .forEach((line) => {
      let row;
      try {
        row = JSON.parse(line);
      } catch (e) {
        return;
      }
      if (!row || !Array.isArray(row.c)) return;
      const out = { i: row.i };
      let seen = false;
      row.c.forEach((v, pos) => {
        const s = String(v == null ? "" : v).trim();
        if (!s) return;
        const head = headers[pos] || "";
        /* A header the model recognises is stronger than a value shape. */
        if (/^(nom|nom prenom|nom et prenom|titre client|adherent|assure|client)$/.test(head) && !out.nom && /\s/.test(s) && /^[A-Z\u00c0-\u00de]+\s+[A-Z\u00c0-\u00de][a-z\u00e0-\u00ff]+$/.test(s)) {
          const parts = s.split(/\s+/);
          out.nom = parts.pop();
          out.prenom = parts.join(" ");
          seen = true;
          return;
        }
        if (/@/.test(s) && !out.email) out.email = s;
        else if (/^(?:\+33|0033|0)\s?[1-9](?:[\s.-]?\d{2}){4}$/.test(s) && !out.tel1) out.tel1 = s;
        else if (/^\d{5}$/.test(s) && !out.cp) out.cp = s;
        else if (/^\d+[.,]\d{2}\s*(?:€|EUR)?$/i.test(s) && !out.cotisation) out.cotisation = s;
        else if (/^(mgen|axa|harmonie|swisslife|malakoff humanis|ag2r|pro btp)$/i.test(s) && !out.mutuelle)
          out.mutuelle = s;
        else if ((/^(organisme|mutuelle|caisse|assurance)$/.test(head) || isInsurerWord(s)) && !out.mutuelle) {
          out.mutuelle = s;
        }
        else if (/^(actif|inactif|oui|non|true|false)$/i.test(s)) {
          out.commentaire = out.commentaire ? out.commentaire + " | " + s : s;
          return;
        } else if (/^(nom|nom prenom|nom et prenom|titre client|adherent|assure|client)$/.test(headers[pos] || "") &&
          !out.nom) {
          if (!/\s/.test(s)) {
            out.nom = s;
            return;
          }
          /* A name header: split the full name of this very cell. */
          const parts = s.split(/\s+/);
          out.nom = parts.pop();
          if (parts.length) out.prenom = parts.join(" ");
          return;
        } else if (/^(commune|ville|city)$/.test(headers[pos] || "") && !out.ville) {
          out.ville = s;
        } else if (/^(adresse|address|rue)$/.test(headers[pos] || "") && !out.adresse) {
          out.adresse = s;
        } else if (/^[A-Z]{1,3}[-\s]?\d+$|^\d+[-\s]?[A-Z]{1,3}$/i.test(s)) {
          /* A reference like "A-12": never a name. */
          out.commentaire = out.commentaire ? out.commentaire + " | " + s : s;
          return;
        } else if (/\b(rue|avenue|boulevard|chemin|impasse|place|route|all[ée]e)\b/i.test(s) && !out.adresse) {
          out.adresse = s;
        } else if (/^(commune|ville|city)$/.test(headers[pos] || "") === false &&
          !isInsurerWord(s) &&
          /^[A-Z\u00c0-\u00de][a-z\u00e0-\u00ff-]+$/.test(s) && !out.ville && out.nom && out.prenom) {
          /* "Paris", "Lyon": a single capitalised word after the name. */
          out.ville = s;
        } else if (/\s/.test(s) && /^[A-Z\u00c0-\u00de' -]{2,}$/.test(s) && !out.nom) {
          /* "DUPONT Jean": surname + first name on one line. */
          const parts = s.split(/\s+/);
          out.nom = parts.pop();
          if (parts.length && !out.prenom) out.prenom = parts.join(" ");
        } else if (/^(nom|nom prenom|nom et prenom|titre client|adherent|assure|client|prenom|prenom nom)$/.test(headers[pos] || "") && !out.nom) {
          /* A name header: this cell is the name even if it is not ALL CAPS. */
          out.nom = s;
        } else if (!out.nom) {
          out.nom = s;
        } else if (/^(prenom|prenom nom)$/.test(headers[pos] || "") && !out.prenom) {
          out.prenom = s;
        } else if (!out.prenom && !/\d/.test(s)) {
          out.prenom = s;
        } else {
          out.commentaire = out.commentaire ? out.commentaire + " | " + s : s;
        }
        seen = true;
      });
      rows.push(out);
    });
  return rows.length ? { rows: rows } : null;
}
function mappingsFrom(prompt) {
  const out = {};
  const re = /- entete: "((?:[^"\\]|\\.)*)" \| valeurs: (.*)/g;
  let m;
  while ((m = re.exec(prompt))) {
    const header = JSON.parse('"' + m[1] + '"');
    const samples = m[2];
    const h = header.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (/e-?mail|courriel|mail/.test(h) || /@/.test(samples))
      out[header] = "email";
    else if (header === "Titre client") out[header] = "nom";
    else if (/prenom/.test(h)) out[header] = "nom";
    else if (/prime|cotisation|montant/.test(h)) out[header] = "cotisation";
    else if (/^\s*(ref|num|no)\b/.test(h)) out[header] = "commentaire";
    else out[header] = "commentaire";
  }
  return out;
}

const { setAllowedPort } = require("../server/server");
const server = createServer();
server.listen(0, "127.0.0.1", () => {
  const { port } = server.address();
  setAllowedPort(port); // its own origin, exactly like PORT=3000 in real use
  process.send && process.send({ port });
  process.stdout.write("FAKE_AI_PORT=" + port + "\n");
});

/* POST /__stop is handled before the static layer. */
const origEmit = server.emit.bind(server);
server.emit = function (event, req, res) {
  if (event === "request" && req.method === "POST" && req.url === "/__stop") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end('{"ok":true}');
    server.close(() => process.exit(0));
    return true;
  }
  return origEmit(event, req, res);
};

/* ============================================================
 * Smart Excel - AI providers (browser, direct calls)
 * ------------------------------------------------------------
 * Google Gemini and Groq are called DIRECTLY from this page:
 * no local server, no proxy. The API key the user pastes is kept
 * in THIS browser only (localStorage, one entry per provider)
 * and is sent only to the provider itself.
 *
 * Free-tier model selection is dynamic: the model list is fetched
 * (cached 10 min), non-text models are dropped, remaining models
 * are ranked and tried in order until one answers.
 *
 * Every function resolves — it never throws — with
 *   { ok: true, ... }  or  { ok: false, code, error }
 * so the application can always fall back to 100% local processing.
 *
 * Two modes:
 *   - mapColumns (headers + samples of ambiguous columns) — cheap
 *     column-level check kept for the connection test;
 *   - mapRows (FULL sheet: every row, every cell) — the model reads
 *     each piece of information, places it in one of the 11 official
 *     columns, separates Nom / Prénom and sends whatever is not
 *     identifiable to Commentaire.
 * ============================================================ */
(function (root) {
"use strict";

/* "name" is a virtual key: a column holding a surname AND a first name. It is
 * converted to nom/prenom before anything reaches the browser. */
const VIRTUAL_KEYS = ["name"];
const TARGET_KEYS = [
  "nom",
  "prenom",
  "tel1",
  "tel2",
  "adresse",
  "cp",
  "ville",
  "email",
  "cotisation",
  "mutuelle",
  "commentaire",
];
const TARGET_LABELS = {
  nom: "Nom",
  prenom: "Prénom",
  tel1: "Téléphone 1",
  tel2: "Téléphone 2",
  adresse: "Adresse",
  cp: "Code postal",
  ville: "Ville",
  email: "Email",
  cotisation: "Cotisation",
  mutuelle: "Mutuelle",
  commentaire: "Commentaire",
};

/* Ranking hints (a hint, never a requirement: the live list is authoritative).
 * A preferred model that is absent or retired is simply skipped. */
const PROVIDERS = {
  gemini: {
    label: "Google Gemini",
    envKey: "GEMINI_API_KEY",
    preferredModels: [
      "gemini-flash-lite-latest",
      "gemini-3.1-flash-lite",
      "gemini-3-flash-preview",
      "gemini-3.1-flash-lite-preview",
      "gemini-2.5-flash-lite",
      "gemini-2.0-flash-lite",
      "gemini-flash-latest",
      "gemini-2.5-flash",
      "gemini-2.0-flash",
    ],
    listUrl: "https://generativelanguage.googleapis.com/v1beta/models",
    list: listGeminiModels,
    call: callGemini,
  },
  groq: {
    label: "Groq (Llama)",
    envKey: "GROQ_API_KEY",
    /* Groq serves OpenAI-compatible chat completions. The 8B models are the
     * fastest with the highest free-tier throughput; 70B is a fallback when
     * the small one is unavailable. The LIVE list stays authoritative. */
    preferredModels: [
      "llama-3.1-8b-instant",
      "llama-3.3-70b-versatile",
      "openai/gpt-oss-20b",
      "gemma2-9b-it",
      "llama3-8b-8192",
    ],
    listUrl: "https://api.groq.com/openai/v1/models",
    list: listGroqModels,
    call: callGroq,
  },
};

/* Per-call budget. A free-tier "lite" model sometimes queues for a while, so
 * the budget is bounded: a dead endpoint can never freeze processing or the
 * export (the local engine takes over). */
const TIMEOUT_MS = 10000;
const LIST_TIMEOUT_MS = 10000;
const LIST_TTL_MS = 10 * 60 * 1000; // the model list is cached 10 minutes
const MAX_ANSWER_TOKENS = 4096;
/* One request carries at most this many columns: fewer, bigger requests are
 * faster (one round-trip for a whole sheet) but a huge prompt is slower and
 * more likely to be truncated. */
const MAX_COLUMNS_PER_CALL = 24;
/* Full-sheet mode: rows per request and per-request budgets. A request must
 * stay short enough for the free tier (15 RPM / 250k TPM) AND produce an
 * answer small enough to never be truncated. */
const SHEET_MODES = ["column", "full"];
/* Rows per AI request. Measured: 60 rows answer in ~7 s and are never
 * truncated; 300 rows take >2 min and risk a cut JSON, which costs MORE
 * (retry + split) than sending two smaller requests. */
const ROWS_PER_CALL = 60;
const MAX_ROWS_PER_REQUEST = 4000; // safety cap on a single /api/ai/sheet call
const MAX_CELL_CHARS = 300; // a cell is sent in full up to this length
/* A 180-row prompt is ~6 000 tokens in and up to 8 000 out: on a slow link
 * (or a cold free-tier model) 45 s was not enough and the request was ABORTED,
 * which is exactly the "internet injoignable" the user saw. */
const SHEET_CALL_TIMEOUT_MS = 120000;
const SHEET_CALL_TOKENS = 16384;
/* Free tier = 15 requests/min. Two parallel batches are safe; the per-minute
 * limiter below handles the bursts. */
/* Requests in flight. Gemini free tier allows 15 RPM, Groq is far more
 * generous: two parallel batches roughly halve the wall-clock time of a big
 * file without tripping the limiter (the retry handles the occasional 429). */
const SHEET_CONCURRENCY = 2;
/* A per-MINUTE rate limit is TRANSIENT: wait for the delay Gemini asks for
 * and retry, instead of abandoning the whole sheet. */
const SHEET_RATE_RETRIES = 5;
const SHEET_RATE_MAX_WAIT_MS = 20000;
/* A batch that fails on a transient error is retried and then SPLIT in half:
 * a lost batch is never reported as "thousands of unclassified lines". */
const SHEET_JOB_RETRIES = 2;

const listCache = new Map(); // provider -> { at, models }
const REMOVED_MODELS = new Set(); // provider|model answers that the model is retired
const COOLDOWN = new Map(); // provider -> timestamp until which models are known busy
const LAST_OK = new Map(); // provider -> { model, at, latency }
let MAINTENANCE_KEY = ""; // key refused (401/403) during this session
/* In-flight identical requests (two tickets asking the same columns because a
 * refresh re-sent them) share one network call. */
const INFLIGHT = new Map();
/* Keys pasted in the interface (validated first) are stored PER PROVIDER in
 * this browser's localStorage so they survive a reload. They are sent only to
 * the provider itself; only a masked hint is ever displayed. */
const KEY_STORE = "sec.ai.keys";
const KEY_FORMAT = /^[A-Za-z0-9._\-]{20,200}$/;
const RUNTIME_KEYS = Object.create(null);
let KEY_STORE_LOADED = false;
function loadKeyStore() {
  if (KEY_STORE_LOADED) return;
  KEY_STORE_LOADED = true;
  try {
    const raw = root.localStorage && root.localStorage.getItem(KEY_STORE);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === "object") {
      ["gemini", "groq"].forEach((p) => {
        if (typeof parsed[p] === "string" && parsed[p]) RUNTIME_KEYS[p] = parsed[p];
      });
    }
  } catch (e) { /* storage disabled: keys live in memory for this session */ }
}
function persistKeyStore() {
  try {
    if (root.localStorage) root.localStorage.setItem(KEY_STORE, JSON.stringify(RUNTIME_KEYS));
  } catch (e) { /* storage disabled: the key still works for this session */ }
}

/* ---------------- Provider configuration ---------------- */

function apiKey(provider) {
  const def = PROVIDERS[provider];
  if (!def) return "";
  loadKeyStore();
  return RUNTIME_KEYS[provider] || "";
}

/* Only a prefix/suffix is ever shown ("AIza…4f2a"): enough to recognise the
 * key, useless to anyone who reads it. */
function maskKey(key) {
  const s = String(key || "");
  if (s.length <= 10) return s ? "****" : "";
  return s.slice(0, 4) + "…" + s.slice(-4);
}

/* Accept a key pasted by the user: trim, strip quotes/spaces left by a copy,
 * refuse obvious mistakes before touching Google. */
function normalizeKey(raw) {
  let key = String(raw == null ? "" : raw).trim();
  key = key.replace(/^["']+|["']+$/g, "").trim();
  key = key.replace(/^key\s*=\s*/i, "").replace(/^bearer\s+/i, "").trim();
  if (!key) return { error: "Clé vide" };
  if (/\s/.test(key)) return { error: "La clé contient des espaces (copie incomplète ?)" };
  if (!KEY_FORMAT.test(key)) return { error: "Format de clé invalide" };
  return { key: key };
}

/* Validate the key with Google (model list + one tiny generation), then make
 * it active for this process. Never stores an invalid key. */
async function setApiKey(rawKey, providerName) {
  const provider = String(providerName || "gemini").toLowerCase();
  const def = PROVIDERS[provider];
  if (!def)
    return { ok: false, code: "bad-request", error: "Fournisseur IA inconnu: " + provider };
  const norm = normalizeKey(rawKey);
  if (norm.error) return { ok: false, code: "bad-request", error: norm.error };
  const key = norm.key;
  const previous = RUNTIME_KEYS[provider];
  RUNTIME_KEYS[provider] = key;
  MAINTENANCE_KEY = "";
  clearCache(true);
  const listed = await getModels(provider, key, true);
  if (listed.error) {
    if (previous) RUNTIME_KEYS[provider] = previous;
    else delete RUNTIME_KEYS[provider];
    clearCache(true);
    const code = listed.error.code === "auth" ? "auth" : listed.error.code;
    return {
      ok: false,
      code: code,
      error:
        code === "auth"
          ? "Clé refusée par " + def.label + " (invalide ou sans accès à l'API)"
          : listed.error.message,
    };
  }
  persistKeyStore();
  return { ok: true, provider: provider, key: key, models: listed.models, mask: maskKey(key) };
}

function providersStatus() {
  const out = {};
  Object.keys(PROVIDERS).forEach((p) => {
    const cached = listCache.get(p);
    const last = LAST_OK.get(p);
    const key = apiKey(p);
    out[p] = {
      label: PROVIDERS[p].label,
      configured: !!key && key !== MAINTENANCE_KEY,
      masked: key ? maskKey(key) : "",
      source: key ? "navigateur" : "",
      model: last ? last.model : "",
      models: cached ? cached.models : [],
      cachedAt: cached ? cached.at : 0,
      cooldown: (COOLDOWN.get(p) || 0) > Date.now(),
    };
    if (MAINTENANCE_KEY && key === MAINTENANCE_KEY) {
      /* A key refused for authentication never becomes valid by retrying. */
      out[p].configured = false;
      out[p].authError = true;
    }
  });
  return out;
}

function markKeyRejected(provider, key) {
  if (key) MAINTENANCE_KEY = key;
  COOLDOWN.delete(provider);
}

/* ---------------- Payload / prompt ---------------- */

/* Only what is strictly needed to decide: the source header and a few
 * short, distinct, REAL sample values (never the whole file). */
function sanitizeColumns(columns) {
  if (!Array.isArray(columns)) return [];
  const seenHeader = Object.create(null); // one entry per header text
  const seenSig = Object.create(null); // identical (header, samples) pairs
  const out = [];
  columns.slice(0, 200).forEach((c) => {
    const header = String((c && c.header) || "").slice(0, 120);
    const samples = [];
    (Array.isArray(c && c.samples) ? c.samples : []).forEach((s) => {
      const v = String(s == null ? "" : s).slice(0, 60);
      if (!v || samples.indexOf(v) !== -1) return; // distinct values only
      if (samples.length < 6) samples.push(v);
    });
    if (!header && !samples.length) return;
    const key = header || "(sans entete)";
    const sig = key + "\u0000" + samples.join("\u0001");
    if (seenSig[sig]) return; // the exact same column twice in one request
    if (seenHeader[key]) {
      /* Same header, other samples (two sheets with the same column): the
       * answer is per header, so one representative entry is enough. */
      if (seenHeader[key].samples.length >= samples.length) return;
      const at = out.indexOf(seenHeader[key]);
      if (at !== -1) out.splice(at, 1); // keep the richest sample set
    }
    seenSig[sig] = true;
    const entry = { header: header, samples: samples };
    seenHeader[key] = entry;
    out.push(entry);
  });
  return out;
}

function buildPrompt(columns) {
  const targets = TARGET_KEYS.map((k) => k + "=" + TARGET_LABELS[k]).join("; ");
  const lines = columns
    .map(
      (c) =>
        "- entete: " +
        JSON.stringify(c.header) +
        " | valeurs: " +
        (c.samples.length
          ? c.samples.map((s) => JSON.stringify(s)).join(", ")
          : "(aucune)"),
    )
    .join("\n");
  return (
    "Tu es un expert en fichiers CSV/Excel français (clients, assurances santé, mutuelles). " +
    "Pour CHAQUE colonne, donne la SEULE destination correcte parmi ces clés exactes: " +
    targets +
    ".\n" +
    "Méthode: décide d'abord d'après les VALEURS RÉELLES (format, contenu), puis d'après l'entête. " +
    "Une entête déjà explicite (« Nom », « Prénom », « Nom de famille », « First name ») est prioritaire. " +
    "Des valeurs longues, des phrases, des dates, des âges ou des notes de suivi sont des notes, " +
    "pas des noms. " +
    "Les entêtes peuvent être abrégés, mal orthographiés, en majuscules ou en anglais (nom, NAME, last name, " +
    "prenom, first name, tel, GSM, mobile, portable, fixe, mail, courriel, CP, zip, commune, city, " +
    "adherent, assure, client, montant, prime, cotisation, caisse, organisme, observation, remarque).\n" +
    "Règles:\n" +
    "1. Email = une adresse avec @. Téléphone = 10 chiffres (0X XX XX XX XX), +33, 33X, espaces/tirets/points. " +
    "Le 1er numéro va dans tel1, un 2e numéro DIFFÉRENT dans tel2.\n" +
    "2. Code postal = exactement 5 chiffres (01000-99999), jamais une année ni un âge. " +
    "Une colonne « CP » + « Ville » séparées -> cp et ville.\n" +
    "3. Montant (75,50 / 1 250,00 € / 12.5EUR) -> cotisation. Nom d'organisme (MGEN, AXA, Harmonie, " +
    "Aesio, Apivia, Malakoff Humanis, Groupama, MAIF, AG2R, Klesia, Pro BTP, Almerys, Itelis…) -> mutuelle.\n" +
    "4. Adresse = numéro + type de voie (rue, avenue, boulevard, impasse, chemin, place) ou mention BP/CS. " +
    "Ville = nom de commune seul (Paris, Lyon, Marseille, Nantes, Toulouse…).\n" +
    "5. Nom de famille seul -> nom. Prénom usuel seul -> prenom.\n" +
    "   Une colonne dont les valeurs sont des NOMS DE PERSONNES («DUPONT Jean», «Jean MARTIN», " +
    "«Dupont, Jean») va dans name, MÊME si son entête est vague (Titre client, Adherent, Assure, " +
    "Client, Personne, Titulaire, Beneficiaire, Nom, NOM PRENOM) et pas dans commentaire.\n" +
    "   Exemples: valeurs «DUPONT Jean» -> name ; «Dupont» seul -> nom ; «Jean» seul -> prenom.\n" +
    "6. Identifiants/références (n° client, contrat, adhérent, dossier, IBAN, matricule), dates, " +
    'années, âges et tranches d\'âge ("64", "70+75", "63 ans", "1950-1948"), statuts, scores, ' +
    "indicateurs Oui/Non -> commentaire. Ce ne sont JAMAIS des montants ni des codes postaux.\n" +
    "7. Toute information qui n'est pas clairement un nom de personne, une adresse, un téléphone, " +
    "un e-mail, un code postal, une ville, un montant CONFIRMÉ ou une mutuelle -> commentaire " +
    "(ex. \"Rappel mardi 30/04/2024 à 17h30 | 87 ans\"). Une note va TOUJOURS dans commentaire, " +
    "jamais dans nom ni prenom.\n" +
    "8. Ne mets JAMAIS une valeur contenant des chiffres dans nom ou prenom (téléphone, " +
    "code postal, montant, date, âge, référence) : ces valeurs vont dans leur colonne ou commentaire.\n" +
    "9. Donne aussi fullName: {\"nom\":\"...\",\"prenom\":\"...\"} quand la colonne contient un " +
    "nom ET un prénom (\"DUPONT Jean\" -> nom DUPONT, prenom Jean ; \"Jean Dupont\" -> nom Dupont, " +
    "prenom Jean ; \"Dupont, Jean\" -> nom Dupont, prenom Jean). N'invente rien et ne découpe jamais " +
    "une note, un identifiant, une ville ou une société (fullName.prenom vide si inconnu).\n" +
    "10. Un NOM D'ORGANISME DE SANTÉ est TOUJOURS mutuelle, JAMAIS cotisation, même placé dans une " +
    "colonne « Cotisation », « Prime » ou « Montant » : MGEN, AXA, Harmonie, Malakoff Humanis, Klesia, " +
    "SwissLife, Apicil, Groupama, MAAF, MAIF, MACIF, AG2R, Pro BTP, MNT, MNH, Radiance, Solimut, " +
    "Neoliane, Ociane, April, Alptis, Henner, Verspieren, Gras Savoye, Filhet Allard, Thelem, Utwin… " +
    "Cotisation = un MONTANT D'ARGENT uniquement.\n" +
    "11. Un NOMBRE NU sans symbole monétaire et sans centimes (70, 80, 75, 1950) est AMBIGU : " +
    "c'est probablement un âge, une année ou un nombre de personnes, PAS une cotisation. Il va dans " +
    "commentaire, sauf si l'entête et le contexte du fichier prouvent qu'il s'agit d'un prix (" +
    "« Cotisation annuelle », « Montant prime », montants voisins avec €). En cas de doute : commentaire.\n" +
    "12. Ne laisse jamais une destination vide et n'invente pas de clé: si rien ne convient, commentaire.\n" +
    'Réponds UNIQUEMENT par un objet JSON {"mappings":{"<entete exact>":"<cle>"},' +
    '"fullName":{"<entete exact>":{"nom":"...","prenom":"..."}}}, ' +
    "une entrée par entête fourni, sans texte autour.\n\n" +
    "Colonnes:\n" +
    lines
  );
}

function parseAnswer(text) {
  let s = String(text || "").trim();
  if (!s) return null;
  s = s
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/```\s*$/i, "")
    .trim();
  try {
    return JSON.parse(s);
  } catch (e) {
    /* try harder */
  }
  const start = s.indexOf("{");
  if (start !== -1) {
    const body = s.slice(start);
    const end = body.lastIndexOf("}");
    if (end !== -1) {
      try {
        return JSON.parse(body.slice(0, end + 1));
      } catch (e) {
        /* keep trying below */
      }
    }
    /* A cut-off answer (token limit reached) ends in the middle of a row.
     * Repair it to the last COMPLETE element and flag it: the caller then
     * retries the missing lines instead of losing them silently. */
    const repaired = repairTruncatedJson(body);
    if (repaired) return repaired;
  }
  return null;
}

function repairTruncatedJson(body) {
  const text = String(body || "");
  if (!/"rows"\s*:/.test(text)) return null;
  /* Keep every complete object of the "rows" array. */
  const key = text.indexOf('"rows"');
  const arr = text.indexOf("[", key);
  if (arr === -1) return null;
  const items = [];
  let depth = 0;
  let begin = -1;
  let inString = false;
  let escaped = false;
  for (let i = arr + 1; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === "\\") escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') {
      inString = true;
      continue;
    }
    if (ch === "{") {
      if (depth === 0) begin = i;
      depth++;
    } else if (ch === "}") {
      depth--;
      if (depth === 0 && begin !== -1) {
        items.push(text.slice(begin, i + 1));
        begin = -1;
      }
    } else if (ch === "]" && depth === 0) {
      break;
    }
  }
  if (!items.length) return null;
  const rows = [];
  items.forEach((it) => {
    try {
      const parsed = JSON.parse(it);
      if (parsed && typeof parsed === "object") rows.push(parsed);
    } catch (e) {
      /* an incomplete object: dropped, its lines are retried */
    }
  });
  if (!rows.length) return null;
  return { rows: rows, __truncated: true };
}

/* Strict whitelist: only the 11 official target keys are accepted. */
function validateMappings(parsed, columns) {
  const maps = parsed && (parsed.mappings || parsed);
  if (!maps || typeof maps !== "object") return {};
  const known = Object.create(null);
  columns.forEach((c) => {
    known[c.header] = true;
  });
  /* Models sometimes answer with a normalised header ("titre client") or drop
   * an accent: map the answer back to the EXACT source header. */
  const byNorm = Object.create(null);
  columns.forEach((c) => {
    const k = String(c.header || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
    if (k) byNorm[k] = c.header;
  });
  const out = {};
  Object.keys(maps).forEach((h) => {
    const t = String(maps[h] == null ? "" : maps[h])
      .trim()
      .toLowerCase();
    if (TARGET_KEYS.indexOf(t) === -1 && VIRTUAL_KEYS.indexOf(t) === -1) return;
    if (known[h]) {
      out[h] = t;
      return;
    }
    const k = String(h)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .trim();
    if (byNorm[k] && !out[byNorm[k]]) out[byNorm[k]] = t;
  });
  return { mappings: out, fullName: validateFullName(parsed, columns, out) };
}

/* Keep only real nom/prenom splits for columns the model labelled "name"
 * (or a plain name column): a split that is empty, note-like or longer than
 * the source value is discarded, never applied. */
function validateFullName(parsed, columns, mappings) {
  const given = (parsed && parsed.fullName) || {};
  const byHeader = {};
  columns.forEach((c) => {
    byHeader[c.header] = true;
  });
  const out = {};
  Object.keys(given || {}).forEach((h) => {
    if (!byHeader[h]) return;
    const v = given[h];
    if (!v || typeof v !== "object") return;
    const nom = String(v.nom == null ? "" : v.nom).trim().slice(0, 80);
    const prenom = String(v.prenom == null ? "" : v.prenom).trim().slice(0, 80);
    if (!nom && !prenom) return;
    /* A note is never a name part. */
    if (/\d/.test(nom + prenom)) return;
    out[h] = { nom: nom, prenom: prenom };
  });
  return out;
}

/* ---------------- Google Gemini ---------------- */

function isTextGeminiModel(name) {
  return !/embedding|aqa|imagen|veo|tts|live|vision|banana|image|robotics|transcribe|computer-use|deep-research|lyria|antigravity|omni|native-audio|learnlm/i.test(
    name,
  );
}

async function listGeminiModels(key) {
  const res = await rawFetch(
    PROVIDERS.gemini.listUrl +
      "?key=" +
      encodeURIComponent(key) +
      "&pageSize=200",
    {},
    LIST_TIMEOUT_MS,
  );
  if (!res.ok) return { error: await httpError(res) };
  const json = await res.json();
  const models = (json.models || [])
    .filter(
      (m) =>
        Array.isArray(m.supportedGenerationMethods) &&
        m.supportedGenerationMethods.indexOf("generateContent") !== -1,
    )
    .map((m) => String(m.name || "").replace(/^models\//, ""))
    .filter(Boolean)
    .filter(isTextGeminiModel);
  return { models };
}

/* ---------------- Groq (OpenAI-compatible chat completions) ---------------- */

function isTextGroqModel(name) {
  /* Groq lists audio (whisper), speech (orpheus/playai/tts), moderation and
   * vision-only models: they cannot classify text rows, and some need a
   * terms acceptance that makes every request fail. Only real CHAT text
   * models are kept. */
  return !/whisper|tts|guard|distil-whisper|playai|orpheus|prompt-guard|moderation|safety|embed|vision|audio|speech|synthesize|allam/i.test(
    name,
  );
}

async function listGroqModels(key) {
  const res = await rawFetch(PROVIDERS.groq.listUrl, {
    headers: { Authorization: "Bearer " + key },
  }, LIST_TIMEOUT_MS);
  if (!res.ok) return { error: await httpError(res) };
  const json = await res.json();
  const models = (json.data || [])
    .map((m) => String((m && m.id) || ""))
    .filter(Boolean)
    .filter(isTextGroqModel);
  return { models };
}

async function callGroq(model, key, prompt, settings) {
  settings = settings || {};
  const isSmall = /70b|versatile|120b|large/i.test(model);
  const res = await rawFetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + key,
    },
    body: JSON.stringify({
      model: model,
      temperature: 0,
      top_p: 0.1,
      max_tokens: Math.min(
        settings.maxOutputTokens || MAX_ANSWER_TOKENS,
        /* the small models cap lower; asking for more is rejected */
        isSmall ? 8000 : 8192,
      ),
      /* NO response_format: Groq rejects it (400) unless the word "json" is
       * in the messages, which cost a failed round-trip and a wrong-model
       * retry on every batch. The prompt already demands JSON only. */
      messages: [
        {
          role: "system",
          content:
            "Tu classes des lignes de fichiers clients français. Tu réponds UNIQUEMENT en JSON valide, sans texte autour.",
        },
        { role: "user", content: prompt },
      ],
    }),
  }, settings.timeoutMs || TIMEOUT_MS);
  if (!res.ok) return { error: await httpError(res) };
  const json = await res.json();
  const choice = (json.choices || [])[0] || {};
  const answer = (choice.message && choice.message.content) || "";
  if (!answer)
    return {
      error: {
        status: 502,
        code: choice.finish_reason === "length" ? "truncated" : "empty",
        message: "réponse vide (« " + (choice.finish_reason || "aucun contenu") + " »)",
      },
    };
  /* Truncated by the token limit: the caller retries with smaller batches. */
  if (choice.finish_reason === "length") {
    return {
      error: { status: 502, code: "truncated", message: "réponse tronquée (limite de tokens)" },
    };
  }
  return { text: answer };
}

async function callGemini(model, key, prompt, settings) {
  settings = settings || {};
  const res = await rawFetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" +
      encodeURIComponent(model) +
      ":generateContent?key=" +
      encodeURIComponent(key),
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ role: "user", parts: [{ text: prompt }] }],
        generationConfig: {
          temperature: 0,
          topP: 0.1,
          candidateCount: 1,
          responseMimeType: "application/json",
          maxOutputTokens: settings.maxOutputTokens || MAX_ANSWER_TOKENS,
        },
      }),
    },
    settings.timeoutMs || TIMEOUT_MS,
  );
  if (!res.ok) return { error: await httpError(res) };
  const json = await res.json();
  const candidate = (json.candidates || [])[0] || {};
  const parts = (candidate.content || {}).parts;
  const answer = Array.isArray(parts)
    ? parts
        .filter((p) => !p.thought) // thinking parts are not the answer
        .map((p) => p.text || "")
        .join("")
    : "";
  if (!answer)
    return {
      error: {
        status: 502,
        code: candidate.finishReason === "MAX_TOKENS" ? "truncated" : "empty",
        message:
          "réponse vide (« " +
          (candidate.finishReason ||
            (json.promptFeedback || {}).blockReason ||
            "aucun contenu") +
          " »)",
      },
    };
  return { text: answer };
}

/* ============================================================
 * FULL-SHEET MODE — Gemini reads every row, every cell
 * ============================================================ */

function buildSheetPrompt(payload) {
  const headers = payload.headers;
  const rows = payload.rows;
  const targets = TARGET_KEYS.map((k) => k + "=" + TARGET_LABELS[k]).join("; ");
  const headerLine = headers
    .map(
      (h, i) =>
        i + "=" + (String(h == null ? "" : h).trim() || "(sans entete)"),
    )
    .join(" | ");
  const rowLines = rows
    .map((r) => JSON.stringify({ i: r.i, c: r.c }))
    .join("\n");
  return (
    "Tu es un expert en consolidation de fichiers clients français (assurances santé, mutuelles). " +
    "Voici des LIGNES RÉELLES extraites d'un fichier source. Chaque ligne a un indice « i » et un tableau « c » " +
    "(une valeur par position de colonne, dans l'ordre). Voici les entêtes, par position (position=NOM_ENTETE) : " +
    headerLine +
    "\n" +
    "Objectif : classer chaque information présente dans CHAQUE ligne dans la bonne colonne de sortie, " +
    "parmi ces clés exactes : " +
    targets +
    ".\n" +
    "Règles :\n" +
    "1. Lis la valeur elle-même (son format, son contenu) : une adresse e-mail (@) va dans email, un numéro de " +
    "téléphone (10 chiffres, 0X, +33, avec espaces/tirets/points) dans tel1 (le 1er) ou tel2 (un 2e numéro DIFFÉRENT), " +
    "un code postal (exactement 5 chiffres) dans cp, un MONTANT CONFIRMÉ (75,50 / 1 250,00 € / 80 €) dans cotisation, " +
    "un nom d'organisme de santé (MGEN, AXA, Harmonie, SwissLife, Malakoff Humanis, AG2R, Pro BTP, Klesia, " +
    "Apicil, Groupama, MAAF, MAIF, MACIF, MNT, MNH, Radiance, Solimut, Neoliane, Ociane, April, Alptis, Henner, " +
    "Verspieren, Thelem, Utwin, Almerys…) dans mutuelle, " +
    "un numéro + type de voie (rue, avenue, boulevard, chemin, impasse, place) dans adresse, un nom de commune seul dans ville.\n" +
    "2. Sépare TOUJOURS le nom et le prénom : « DUPONT Jean » ou « Jean DUPONT » ou « Dupont, Jean » -> nom=DUPONT, " +
    "prenom=Jean. Une valeur qui contient un nom ET un prénom ne doit jamais rester entière dans nom.\n" +
    "3. Corrige les informations mal placées ; ne perds aucune information. Toute information non identifiable, " +
    "ambiguë ou hors périmètre (référence, identifiant, date, âge, statut, note de suivi, texte libre) va dans commentaire.\n" +
    "3a. MUTUELLE vs COTISATION : un nom d'organisme de santé est TOUJOURS mutuelle, jamais cotisation, même s'il " +
    "arrive dans une colonne « Cotisation », « Prime » ou « Montant ». cotisation ne reçoit QUE de l'argent.\n" +
    "3b. Un NOMBRE NU sans symbole monétaire ni centimes (70, 80, 75, 65, 1950) est AMBIGU : c'est très " +
    "probablement un âge, une année ou un nombre de personnes. Mets-le dans commentaire, SAUF si l'entête et les " +
    "autres valeurs de la colonne prouvent un prix (ex. « Cotisation annuelle » avec 320,50 et 415,00). " +
    "Un même nombre (70) ne peut pas être un âge dans une ligne et un prix dans une autre : le contexte de la " +
    "COLONNE décide, pas la valeur isolée.\n" +
    "3c. Comprends le SENS de chaque valeur avant de la placer (contexte de la ligne + de la colonne). " +
    "En cas de doute, commentaire : ne devine jamais.\n" +
    "4. Une cellule ne peut aller que dans une seule destination : les autres champs de la ligne restent vides " +
    "si la ligne ne les contient pas. N'utilise JAMAIS une valeur qui n'est pas dans la ligne (pas d'invention).\n" +
    "5. Ne recopie pas deux fois la même information.\n" +
    "6. Si une valeur est déjà correctement placée, ne la recopie pas inutilement (elle est conservée par défaut).\n" +
    'Réponds UNIQUEMENT par un objet JSON de la forme {"rows":[{"i":<indice>,"nom":"...","prenom":"...",' +
    '"tel1":"...","tel2":"...","adresse":"...","cp":"...","ville":"...","email":"...","cotisation":"...",' +
    '"mutuelle":"...","commentaire":"..."}]} avec UNE entrée par ligne fournie, sans texte autour.\n\n' +
    "Lignes:\n" +
    rowLines
  );
}

/* The row payload a client may send: indices, positions, raw texts. Anything
 * else (a nested object, a huge text) is dropped instead of forwarded. */
function sanitizeSheetRows(payload) {
  if (!payload || !Array.isArray(payload.rows)) return null;
  const headers = Array.isArray(payload.headers)
    ? payload.headers.slice(0, 120).map((h) => String(h == null ? "" : h).slice(0, 120))
    : [];
  if (!headers.length) return null;
  const rows = [];
  payload.rows.slice(0, MAX_ROWS_PER_REQUEST).forEach((r) => {
    if (!r || !Array.isArray(r.c)) return;
    const i = Number(r.i);
    if (!Number.isFinite(i) || i < 0) return;
    const c = r.c.slice(0, headers.length).map((v) =>
      v == null ? "" : String(v).slice(0, MAX_CELL_CHARS),
    );
    rows.push({ i: Math.floor(i), c: c });
  });
  if (!rows.length) return null;
  return { headers: headers, rows: rows };
}

/* Words of a value / of one source row — the SAME tokenizer the processor
 * uses (rowWordSet), so both layers accept and reject identical answers. */
function textWords(s) {
  return (
    String(s == null ? "" : s)
      .replace(/\u00a0/g, " ")
      .toLowerCase()
      .match(/[0-9a-z\u00e0-\u00f6\u00f8-\u00ff]+/g) || []
  );
}

function normText(s) {
  return String(s == null ? "" : s)
    .toLowerCase()
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[.;:,]+$/g, "")
    .trim();
}

/* Per-ROW source index: the exact cells AND the words of THAT row only. */
function sourceTextsForRow(row) {
  const cells = Object.create(null);
  const words = Object.create(null);
  (row.c || []).forEach((v) => {
    const n = normText(v);
    if (n) cells[n] = true;
    textWords(v).forEach((w) => {
      words[w] = true;
    });
  });
  return { cells: cells, words: words };
}

/* Keep only what the row itself contains. A value is accepted when it is an
 * exact cell OR when every one of its words exists on THAT row: the model is
 * TOLD to split "DUPONT Jean" into nom="DUPONT" + prenom="Jean", and the old
 * whole-cell check threw those correct splits away (names ended up in
 * Commentaire, fake review items blocked the export). A value borrowed from
 * ANOTHER row is still rejected — provenance is per row, never per batch. */
function splitFromSource(value, src) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return { kept: "", invented: "" };
  if (src.cells[normText(raw)]) return { kept: raw, invented: "" };
  const tokens = textWords(raw);
  if (tokens.length && tokens.every((w) => src.words[w]))
    return { kept: raw, invented: "" };
  /* A value assembled with a separator is accepted segment by segment. */
  const parts = raw.split(/\s*(?:\||;|\/|\n)\s*/).filter(Boolean);
  if (parts.length > 1) {
    const kept = [];
    const invented = [];
    parts.forEach((p) => {
      const pt = textWords(p);
      if (src.cells[normText(p)] || (pt.length && pt.every((w) => src.words[w])))
        kept.push(p);
      else invented.push(p);
    });
    return { kept: kept.join(" | "), invented: invented.join(" | ") };
  }
  return { kept: "", invented: raw };
}

/* "70", "80", "1950": a bare integer with no currency and no decimals is
 * never stated as a price by the assistant (age / year / count). */
function isBareAmbiguousNumber(v) {
  const s = String(v == null ? "" : v).trim();
  if (!s) return false;
  if (/[€$£%]|\b(euros?|euro|eur|fcfa|francs?)\b/i.test(s)) return false;
  if (!/^\d{1,3}$/.test(s)) return false;
  const n = parseInt(s, 10);
  return n >= 1 && n <= 110; // a plausible human age; "999" is not one
}

/* Health insurer names the assistant must never turn into a price. Small
 * local list (the browser has the full database): the assistant is also told
 * the rule, this is the safety net. */
const INSURER_WORDS = /(mgen|axa|harmonie|swisslife|swiss life|malakoff|humanis|klesia|apicil|groupama|maaf|maif|macif|matmut|gmf|ag2r|pro btp|mnt|mnh|radiance|solimut|neoliane|ociane|april|alptis|henner|verspieren|thelem|utwin|aesio|apivia|mutualia|allianz|generali|aviva|almerys|itelis|santevet|myriade|filhet|gras savoye|miltis|b2v|carcept|audien|fidelia|partea|unilia|vauban|exance|ipsec|smacl|sma assurances)/i;

function isInsurerName(v) {
  const s = String(v == null ? "" : v).trim();
  if (!s || s.length > 60) return false;
  if (/[\d€$£%]/.test(s)) return false; // a price/reference is not a name
  if (s.split(/\s+/).length > 4) return false;
  return INSURER_WORDS.test(s);
}

function validateSheetRows(parsed, payload) {
  if (!parsed || typeof parsed !== "object") return { rows: [], invalid: [] };
  const list = Array.isArray(parsed.rows) ? parsed.rows : [];
  /* Keyed by the source index, so an entry answered out of order still lands
   * on its own line (two lines can hold the SAME values). */
  const byIndex = Object.create(null);
  payload.rows.forEach((r) => {
    byIndex[r.i] = r;
  });
  const srcByRow = Object.create(null);
  payload.rows.forEach((r) => {
    srcByRow[r.i] = sourceTextsForRow(r);
  });
  const rows = {};
  const invalid = [];
  list.forEach((entry) => {
    if (!entry || typeof entry !== "object") return;
    const i = Number(entry.i);
    if (!Number.isFinite(i) || !byIndex[Math.floor(i)]) return;
    const values = {};
    TARGET_KEYS.forEach((key) => {
      /* Only the 11 official keys are read, and an array/object is never a
       * cell value (a malformed answer must not break the line). */
      if (typeof entry[key] !== "string") return;
      const row = byIndex[Math.floor(i)];
      const part = splitFromSource(entry[key], srcByRow[row.i]);
      if (part.kept) values[key] = part.kept;
      if (part.invented) {
        invalid.push({ i: Math.floor(i), key: key, value: part.invented });
      }
    });
    /* HARD RULES enforced whatever the model answers (the prompt is not a
     * guarantee): a cotisation is money, a bare number is ambiguous, and a
     * health insurer is never a price. */
    if (values.cotisation) {
      if (isInsurerName(values.cotisation)) {
        if (!values.mutuelle) values.mutuelle = values.cotisation;
        else if (values.mutuelle.indexOf(values.cotisation) === -1)
          values.mutuelle += " | " + values.cotisation;
        delete values.cotisation;
      } else if (isBareAmbiguousNumber(values.cotisation)) {
        values.commentaire = values.commentaire
          ? values.commentaire + " | " + values.cotisation
          : values.cotisation;
        delete values.cotisation;
      }
    }
    /* A name part is never a note: such a value goes to Commentaire. */
    ["nom", "prenom"].forEach((key) => {
      if (values[key] && /\d/.test(values[key])) {
        values.commentaire = values.commentaire
          ? values.commentaire + " | " + values[key]
          : values[key];
        delete values[key];
      }
    });
    if (Object.keys(values).length) rows[Math.floor(i)] = values;
  });
  return { rows: rows, invalid: invalid };
}

/* One batched request. Resolves { text } or { error } like callGemini. */
async function runSheetCall(provider, def, key, payload, candidates, attempts, stats) {
  const prompt = buildSheetPrompt(payload);
  let last = null;
  let answered = null;
  for (let n = 0; n < candidates.length; n++) {
    const candidate = candidates[n];
    const started = Date.now();
    let answer = await def.call(candidate, key, prompt, {
      timeoutMs: SHEET_CALL_TIMEOUT_MS,
      maxOutputTokens: SHEET_CALL_TOKENS,
    });
    if (answer.error && answer.error.code === "busy") {
      await sleep(500);
      answer = await def.call(candidate, key, prompt, {
        timeoutMs: SHEET_CALL_TIMEOUT_MS,
        maxOutputTokens: SHEET_CALL_TOKENS,
      });
    }
    /* Per-MINUTE rate limit (the free tier allows 15 requests/min): Gemini
     * says how long to wait. Waiting is what makes the assistant finish the
     * batch — stopping here is why whole sheets fell back to the local engine
     * after a "retry in 4.3s". */
    let rateTries = 0;
    while (
      answer.error &&
      answer.error.code === "rate" &&
      rateTries < SHEET_RATE_RETRIES
    ) {
      rateTries++;
      const waitMs = Math.max(500, Math.min(SHEET_RATE_MAX_WAIT_MS, answer.error.retryAfterMs || 5000));
      if (rateTries > 1) { /* a second hit: back off longer */
        await sleep(Math.min(SHEET_RATE_MAX_WAIT_MS, waitMs * rateTries));
      } else {
        await sleep(waitMs);
      }
      answer = await def.call(candidate, key, prompt, {
        timeoutMs: SHEET_CALL_TIMEOUT_MS,
        maxOutputTokens: SHEET_CALL_TOKENS,
      });
    }
    if (answer.error && RETIRED_RE.test(answer.error.message || "")) {
      answer.error.code = "retired";
      REMOVED_MODELS.add(provider + "|" + candidate);
    }
    if (answer.error) {
      attempts.push({
        model: candidate,
        code: answer.error.code,
        message: answer.error.message,
        ms: Date.now() - started,
      });
      last = answer.error;
      if (answer.error.code === "auth") {
        markKeyRejected(provider, key);
        break;
      }
      if (
        answer.error.code === "quota" &&
        /per day|daily|quota exceeded|exceeded your current quota/i.test(
          String(answer.error.message || ""),
        )
      )
        break;
      continue;
    }
    const parsed = parseAnswer(answer.text);
    const invalidJson = !parsed || !Array.isArray(parsed.rows);
    const validated = invalidJson
      ? { rows: {}, invalid: [] }
      : validateSheetRows(parsed, payload);
    /* The answer IS cut: a truncated JSON was repaired by parseAnswer and
     * may contain a partially placed row. Only the complete rows are used. */
    if (invalidJson || parsed.__truncated) {
      attempts.push({
        model: candidate,
        code: invalidJson ? "empty" : "truncated",
        message: invalidJson ? "réponse non JSON" : "réponse tronquée",
        ms: Date.now() - started,
      });
      last = {
        code: invalidJson ? "empty" : "truncated",
        message: invalidJson
          ? "réponse non JSON"
          : "réponse tronquée : lignes relues avec un lot plus petit",
      };
      if (invalidJson) continue;
      /* Truncated: some lines ARE placed, keep them and retry what is left. */
      answered = {
        rows: validated.rows,
        invalid: validated.invalid,
        model: candidate,
        ms: Date.now() - started,
        truncated: true,
      };
      break;
    }
    answered = {
      rows: validated.rows,
      invalid: validated.invalid,
      model: candidate,
      ms: Date.now() - started,
      truncated: false,
    };
    break;
  }
  if (!answered) return { error: last || { code: "network", message: "aucun modèle disponible" } };
  return answered;
}

/* Map EVERY row of one worksheet.
 * body: { provider?, model?, headers:[...], rows:[{i, c:[...]}] }
 * resolves { ok, model, rows:{ i: {target: value} }, placed, missing, invalid, attempts } */
async function mapRows(body) {
  const provider = String((body && body.provider) || "gemini").toLowerCase();
  const def = PROVIDERS[provider];
  if (!def)
    return { ok: false, provider, code: "bad-request", error: "Fournisseur IA inconnu: " + provider };
  const key = apiKey(provider);
  if (!key)
    return {
      ok: false,
      provider,
      code: "no-key",
      error: "Aucune clé " + def.label + " configurée — collez-la dans le panneau IA",
    };
  if (MAINTENANCE_KEY === key)
    return {
      ok: false,
      provider,
      code: "auth",
      error: "Clé " + def.label + " refusée (401/403) — collez une nouvelle clé dans le panneau IA",
    };
  const payload = sanitizeSheetRows(body);
  if (!payload)
    return { ok: false, provider, code: "bad-request", error: "aucune ligne à analyser" };
  const requested = body && body.model ? String(body.model) : "";
  let candidates;
  if (requested && requested !== "auto") {
    candidates = [requested];
  } else {
    const listed = await getModels(provider, key, false);
    if (listed.error)
      return { ok: false, provider, code: listed.error.code, error: listed.error.message };
    candidates = candidateModels(listed.models, provider);
  }

  /* A row split evenly over the requests keeps each answer small enough to
   * be complete (a truncated answer costs a retry, not a lost line). */
  const total = payload.rows.length;
  const calls = Math.max(1, Math.ceil(total / ROWS_PER_CALL));
  const perCall = Math.ceil(total / calls);
  const jobs = [];
  for (let at = 0; at < payload.rows.length; at += perCall) {
    jobs.push(payload.rows.slice(at, at + perCall));
  }

  const attempts = [];
  const stats = { calls: 0, failures: 0, placed: 0, missing: 0, invalid: 0, truncated: 0 };
  const answeredBy = new Set();
  const allRows = {};
  let model = "";
  let lastError = null;
  let stopped = false;

  /* A batch that fails is retried a few times with a growing delay, then
   * SPLIT: a transient timeout ("internet injoignable") or a
   * payload the model chokes on used to mark the WHOLE batch as unclassified,
   * which is how a single hiccup turned into "9 440 lignes non classées".
   * Only a really failing batch — after retries AND split — is reported. */
  const runJob = async (rows, depth) => {
    depth = depth || 0;
    if (stopped) {
      stats.missing += rows.length;
      return;
    }
    stats.calls++;
    let out = await runSheetCall(
      provider,
      def,
      key,
      { headers: payload.headers, rows: rows },
      candidates,
      attempts,
      stats,
    );
    /* Retry a NETWORK / EMPTY / TRUNCATED answer: those are transient. */
    const retriable = out.error && ["network", "empty", "truncated", "busy"].indexOf(out.error.code) !== -1;
    if (retriable && depth < SHEET_JOB_RETRIES) {
      await sleep(Math.min(8000, 700 * Math.pow(2, depth)));
      return runJob(rows, depth + 1);
    }
    if (out.error) {
      /* Still failing: SPLIT the batch and retry the halves. A smaller request
       * succeeds for every size / timeout / JSON-length problem, and a
       * "bad-request" (payload refused, model quirk) is exactly that. The old
       * code treated bad-request as fatal and dropped every remaining row of
       * the sheet — 882 of 1 000 lines lost in one batch. */
      if (rows.length > 1 && depth < SHEET_JOB_RETRIES + 3) {
        const mid = Math.ceil(rows.length / 2);
        await runJob(rows.slice(0, mid), depth + 1);
        await runJob(rows.slice(mid), depth + 1);
        return;
      }
      stats.failures++;
      lastError = out.error;
      stats.missing += rows.length;
      /* Only these stop EVERYTHING: a wrong key, an exhausted day, no key.
       * "bad-request" is not one of them any more (it is per-request). */
      if (["auth", "quota", "no-key"].indexOf(out.error.code) !== -1) stopped = true;
      return;
    }
    model = out.model || model;
    answeredBy.add(out.model);
    if (out.truncated) stats.truncated++;
    /* A truncated answer leaves lines unplaced: retry only those, smaller. */
    const left = [];
    rows.forEach((r) => {
      if (out.rows[r.i]) {
        allRows[r.i] = out.rows[r.i];
        stats.placed++;
      } else left.push(r);
    });
    if (left.length && depth < SHEET_JOB_RETRIES + 3) {
      /* Every line the model did not place is re-sent, in a batch small enough
       * to be answered completely (a truncated answer is usually the cause). */
      const chunkSize = Math.max(1, Math.min(40, Math.ceil(left.length / 2)));
      for (let at = 0; at < left.length; at += chunkSize) {
        await runJob(left.slice(at, at + chunkSize), depth + 1);
      }
      return;
    }
    if (left.length) stats.missing += left.length;
    stats.invalid += (out.invalid || []).length;
  };

  /* Bounded concurrency: the free tier allows a handful of requests per
   * minute, so a whole 25 000-row batch never overruns the quota in seconds. */
  let next = 0;
  const worker = async () => {
    while (next < jobs.length) {
      const job = jobs[next++];
      await runJob(job, 0);
    }
  };
  const workers = [];
  for (let w = 0; w < Math.min(SHEET_CONCURRENCY, jobs.length); w++) workers.push(worker());
  await Promise.all(workers);

  if (!Object.keys(allRows).length && lastError) {
    const code = lastError.code === "retired" ? "network" : lastError.code || "network";
    return {
      ok: false,
      provider,
      code: code,
      error: lastError.message || "aucun modèle disponible",
      attempts: attempts,
      stats: stats,
    };
  }
  if (!(requested && requested !== "auto") && model) {
    LAST_OK.set(provider, { model: model, at: Date.now(), latency: 0 });
  }
  COOLDOWN.delete(provider);
  return {
    ok: true,
    provider: provider,
    model: model,
    requestedModel: requested || "auto",
    rows: allRows,
    placed: stats.placed,
    missing: stats.missing,
    invalid: stats.invalid,
    truncated: stats.truncated,
    calls: stats.calls,
    failures: stats.failures,
    /* The last error is returned even when SOME rows were placed: without it
     * the browser could only say "incomplete" with no reason. */
    error: lastError
      ? {
          code: lastError.code === "retired" ? "network" : lastError.code || "network",
          message: lastError.message || "erreur inconnue",
        }
      : null,
    attempts: attempts,
  };
}

/* ---------------- HTTP helpers ---------------- */

function rawFetch(url, options, timeoutMs) {
  const controller =
    typeof AbortController === "function" ? new AbortController() : null;
  const timer = controller
    ? setTimeout(() => controller.abort(), timeoutMs || TIMEOUT_MS)
    : null;
  const opts = Object.assign({}, options || {});
  if (controller) opts.signal = controller.signal;
  return fetch(url, opts)
    .catch((err) => ({
      ok: false,
      status: err && err.name === "AbortError" ? 504 : 502,
      aborted: !!(err && err.name === "AbortError"),
      text: async () => (err && err.message) || "réseau indisponible",
    }))
    .then((res) => {
      if (timer) clearTimeout(timer);
      return res;
    });
}

async function httpError(res) {
  let raw = "";
  try {
    raw = await res.text();
  } catch (e) {
    raw = "";
  }
  let message = String(raw || "").slice(0, 300);
  try {
    const j = JSON.parse(raw);
    message = (j.error && (j.error.message || j.error)) || j.message || message;
    if (typeof message === "object") message = JSON.stringify(message);
  } catch (e) {
    /* keep raw */
  }
  let code = "bad-request";
  if (res.status === 401 || res.status === 403) code = "auth";
  else if (res.status === 429) code = "rate";
  else if (res.status === 402) code = "quota";
  else if (res.status === 503) code = "busy";
  else if (res.status === 404) code = "retired";
  else if (res.status === 504 || res.status === 502) code = "network";
  else if (
    /quota|rate.?limit|resource.?exhausted|exceeded your current quota/i.test(
      String(message),
    )
  )
    code = "rate";

  /* A PER-DAY / per-plan exhaustion is final for today: nothing retried saves
   * it. A PER-MINUTE rate limit ("limit: 15 … per minute … retry in 4.3s")
   * clears by itself, so it must NOT disable the assistant: it is reported as
   * "rate" with the delay Gemini itself asks for. */
  const msg = String(message || "");
  const daily = /per day|daily|per[_ ]day|quota exceeded for metric/i.test(msg) &&
    !/per minute|per_minute|requests per minute/i.test(msg);
  if (code === "rate" && daily) code = "quota";
  /* "Please retry in 4.312644263s" / "retryDelay":"4s" */
  let retryAfterMs = 0;
  const delay = msg.match(/retry in ([\d.]+)\s*s/i) || msg.match(/"retryDelay"\s*:\s*"([\d.]+)s"/i);
  if (delay) retryAfterMs = Math.min(60000, Math.ceil(parseFloat(delay[1]) * 1000));
  else if (res.status === 429) retryAfterMs = 6000;
  const headerDelay = res.headers && res.headers.get ? res.headers.get("retry-after") : null;
  if (headerDelay && /^\d+$/.test(String(headerDelay).trim()))
    retryAfterMs = Math.min(60000, Number(String(headerDelay).trim()) * 1000);

  return {
    status: res.status,
    code,
    /* A human sentence for the interface, without the blob of Google's JSON. */
    message:
      code === "rate"
        ? "limite de débit gratuite atteinte (" + Math.round(retryAfterMs / 1000) +
          "s) : nouvelle tentative automatique"
        : String(msg || "erreur inconnue"),
    retryAfterMs,
  };
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

/* ---------------- Model ranking ---------------- */

const RETIRED_RE =
  /no longer available|not found|does not exist|do not have access|is not supported|unsupported model|decommissioned|deprecated|not available to new users|invalid model/i;

/* Rank the live list. Rank, never filter to a fixed list: a new model released
 * tomorrow still works, a retired one is skipped. */
function rankModels(provider, available) {
  const pref = PROVIDERS[provider].preferredModels;
  const score = (m) => {
    const p = pref.indexOf(m);
    if (p !== -1) return -100 + p; // preferred first, in order
    const low = m.toLowerCase();
    let s = 0;
    if (/lite|instant|small|mini/.test(low)) s -= 30; // cheap/fast first
    if (/flash/.test(low)) s -= 12;
    if (/pro|max|large|ultra|thinking/.test(low)) s += 25; // slow premium
    if (/preview|exp|alpha|beta/.test(low)) s += 40;
    /* Groq: the "instant" models are the fastest and the most reliable for
     * strict JSON output; gpt-oss is a solid second. Anything else is a
     * fallback only. */
    if (provider === "groq") {
      /* Measured: gpt-oss-120b takes ~50 s for a 120-row sheet; a "large"
       * suffix also matched the /large/ rule (+25). The Llama models are an
       * order of magnitude faster and produce strict JSON. */
      if (/llama-3\.1-8b-instant/.test(low)) s -= 100;
      else if (/llama-3\.3-70b-versatile/.test(low)) s -= 80;
      else if (/gpt-oss-20b/.test(low)) s -= 20;
      else s += 60;
    }
    /* Gemma writes its analysis instead of JSON and ignores JSON mode. */
    if (/^(gemma|learnlm|gemini-omni)/.test(low)) s += 500;
    return s;
  };
  return available
    .filter((m) => !REMOVED_MODELS.has(provider + "|" + m))
    .slice()
    .sort((a, b) => score(a) - score(b) || (a < b ? -1 : 1));
}

async function getModels(provider, key, forceRefresh) {
  const cached = listCache.get(provider);
  if (!forceRefresh && cached && Date.now() - cached.at < LIST_TTL_MS)
    return { models: cached.models };
  const listed = await PROVIDERS[provider].list(key);
  if (listed.error) {
    /* A transient failure must not wipe a usable cached list. */
    if (cached && cached.models.length && listed.error.code !== "auth") {
      return { models: cached.models, stale: true };
    }
    return { error: listed.error };
  }
  const models = rankModels(provider, listed.models || []);
  if (!models.length)
    return {
      error: {
        status: 503,
        code: "empty",
        message: REMOVED_MODELS.size
          ? "modèles gratuits déjà refusés ou quota épuisé"
          : "aucun modèle texte disponible",
      },
    };
  listCache.set(provider, { at: Date.now(), models });
  return { models };
}

/* ---------------- Public API ---------------- */

/* Warm up the model list in the background when the server starts. */
function prewarm() {
  Object.keys(PROVIDERS).forEach((p) => {
    const key = apiKey(p);
    if (!key || MAINTENANCE_KEY === key) return;
    getModels(p, key, true).catch(() => {});
  });
}

/* Order of the models to try: the model that answered last time first (it
 * works and it is warm), then the ranked free-tier list. */
function candidateModels(ranked, provider) {
  /* Groq: SPEED matters most here (the 120b model takes ~50 s per batch).
   * Always start from the ranked fast model, not from whatever answered last. */
  if (provider === "groq") return ranked.slice(0, 2);
  const last = LAST_OK.get(provider);
  if (last && last.model && ranked.indexOf(last.model) !== -1) {
    return [last.model].concat(
      ranked.filter((m) => m !== last.model).slice(0, 2),
    );
  }
  return ranked.slice(0, 3);
}

/* Map the ambiguous columns of ONE worksheet.
 * body: { provider?, columns: [{header, samples[]}], model? }
 * resolves { ok, provider, model, mappings, models, attempts } (never throws). */
async function mapColumns(body) {
  const provider = String((body && body.provider) || "gemini").toLowerCase();
  const def = PROVIDERS[provider];
  if (!def)
    return {
      ok: false,
      code: "bad-request",
      error: "Fournisseur IA inconnu: " + provider,
      provider,
    };
  const key = apiKey(provider);
  if (!key) {
    return {
      ok: false,
      provider,
      code: "no-key",
      error: "Aucune clé " + def.label + " configurée — collez-la dans le panneau IA",
    };
  }
  if (MAINTENANCE_KEY === key) {
    return {
      ok: false,
      provider,
      code: "auth",
      error: "Clé " + def.label + " refusée (401/403) — collez une nouvelle clé dans le panneau IA",
    };
  }
  const columns = sanitizeColumns(body && body.columns);
  const warmup = !!(body && body.warmup); // connection test: no retry needed
  if (!columns.length)
    return {
      ok: false,
      provider,
      code: "empty",
      error: "aucune colonne à analyser",
    };

  /* Identical request already running (two sheets with the same columns):
   * share its answer instead of paying for a second round-trip. */
  const signature = JSON.stringify({
    p: provider,
    m: body && body.model,
    c: columns,
  });
  if (!warmup && INFLIGHT.has(signature)) return INFLIGHT.get(signature);

  const job = runMapping(provider, def, key, columns, body, warmup);
  if (!warmup) {
    INFLIGHT.set(signature, job);
    job.then(
      () => INFLIGHT.delete(signature),
      () => INFLIGHT.delete(signature),
    );
  }
  return job;
}

async function runMapping(provider, def, key, columns, body, warmup) {
  const requested = body && body.model ? String(body.model) : "";
  let candidates;
  if (requested && requested !== "auto") {
    candidates = [requested];
  } else {
    const listed = await getModels(provider, key, false);
    if (listed.error)
      return {
        ok: false,
        provider,
        code: listed.error.code,
        error: listed.error.message,
      };
    candidates = candidateModels(listed.models, provider);
  }

  /* One request per chunk of columns: a whole sheet usually fits in ONE
   * request (fast), and a very wide sheet is still fully covered. */
  const chunks = [];
  for (let i = 0; i < columns.length; i += MAX_COLUMNS_PER_CALL)
    chunks.push(columns.slice(i, i + MAX_COLUMNS_PER_CALL));

  const attempts = [];
  let last = null;
  let model = "";
  let mappings = {};
  const fullNames = {};
  let totalLatency = 0;

  for (const chunk of chunks) {
    const prompt = buildPrompt(chunk);
    let answered = null;
    for (let i = 0; i < candidates.length; i++) {
      const candidate = candidates[i];
      const started = Date.now();
      let answer = await def.call(candidate, key, prompt);
      if (
        answer.error &&
        answer.error.code === "busy" &&
        attempts.length < 2 &&
        !warmup
      ) {
        /* "High demand": one short retry, then move on to the next model. */
        await sleep(400);
        answer = await def.call(candidate, key, prompt);
      }
      if (answer.error && RETIRED_RE.test(answer.error.message || "")) {
        answer.error.code = "retired";
        REMOVED_MODELS.add(provider + "|" + candidate);
      }
      if (answer.error) {
        attempts.push({
          model: candidate,
          code: answer.error.code,
          message: answer.error.message,
          ms: Date.now() - started,
        });
        last = answer.error;
        if (answer.error.code === "auth") {
          markKeyRejected(provider, key);
          break;
        }
        /* Web quota for the whole day: trying more models cannot help. */
        if (
          answer.error.code === "quota" &&
          /per day|daily|quota exceeded|exceeded your current quota/i.test(
            answer.error.message || "",
          )
        )
          break;
        continue;
      }
      const parsed = parseAnswer(answer.text);
      if (!parsed) {
        attempts.push({
          model: candidate,
          code: "empty",
          message: "réponse non JSON",
          ms: Date.now() - started,
        });
        last = { code: "empty", message: "réponse non JSON" };
        continue;
      }
      answered = { parsed: parsed, model: candidate, ms: Date.now() - started };
      break;
    }
    if (!answered) break; // no model answered this chunk: report the failure
    model = answered.model;
    totalLatency += answered.ms;
    const validated = validateMappings(answered.parsed, chunk);
    Object.assign(mappings, validated.mappings);
    Object.assign(fullNames, validated.fullName);
    /* A model EXPLICITLY chosen by the user is never replaced by another one:
     * the remaining chunks keep using it. Only "auto" re-orders candidates. */
    if (!(requested && requested !== "auto")) {
      candidates = candidateModels(
        [answered.model].concat(candidates.filter((m) => m !== answered.model)),
        provider,
      );
    }
    if (last && last.code === "auth") break;
  }

  if (!model) {
    if (last && last.code === "busy")
      COOLDOWN.set(provider, Date.now() + 60000);
    const code =
      last && last.code === "retired"
        ? "network"
        : last && last.code === "truncated"
          ? "empty"
          : (last && last.code) || "network";
    return {
      ok: false,
      provider,
      code,
      error: (last && last.message) || "aucun modèle disponible",
      attempts,
    };
  }

  /* The "last good" model is only remembered for the automatic mode: in
   * fixed mode it must not influence the next choice, nor the dropdown. */
  if (!(requested && requested !== "auto")) {
    LAST_OK.set(provider, { model, at: Date.now(), latency: totalLatency });
  }
  COOLDOWN.delete(provider);
  return {
    ok: true,
    provider,
    model,
    requestedModel: requested || "auto",
    latency: totalLatency,
    mappings,
    fullName: fullNames,
    models: candidates,
    attempts,
  };
}

/* Connectivity / key test without touching customer data. */
async function testProvider(provider) {
  const p = String(provider || "gemini").toLowerCase();
  const def = PROVIDERS[p];
  if (!def)
    return {
      ok: false,
      provider: p,
      code: "bad-request",
      error: "Fournisseur IA inconnu",
    };
  const key = apiKey(p);
  if (!key)
    return {
      ok: false,
      provider: p,
      code: "no-key",
      error: "Aucune clé " + def.label + " configurée — collez-la dans le panneau IA",
    };
  if (MAINTENANCE_KEY === key)
    return {
      ok: false,
      provider: p,
      code: "auth",
      error: "Clé " + def.label + " refusée (401/403)",
    };
  const listed = await getModels(p, key, true);
  if (listed.error) {
    return {
      ok: false,
      provider: p,
      code: listed.error.code,
      error: listed.error.message,
    };
  }
  /* A real (tiny) generation proves the model is usable, not just listed. */
  const result = await mapColumns({
    provider: p,
    warmup: true,
    columns: [{ header: "Nom client", samples: ["Dupont"] }],
  });
  if (!result.ok) {
    return {
      ok: false,
      provider: p,
      code: result.code,
      error: result.error,
      models: listed.models,
      attempts: result.attempts || [],
    };
  }
  return {
    ok: true,
    provider: p,
    model: result.model,
    models: listed.models,
    latency: result.latency,
  };
}

/* clearKey: a new key means the previous "retired model" verdicts must be
 * forgotten too (they were made with the old key). */
function clearCache(clearKey) {
  listCache.clear();
  COOLDOWN.clear();
  LAST_OK.clear();
  REMOVED_MODELS.clear();
  INFLIGHT.clear();
  MAINTENANCE_KEY = "";
}

root.SEC = root.SEC || {};
root.SEC.AIProviders = {
  PROVIDERS,
  TARGET_KEYS,
  TARGET_LABELS,
  providersStatus,
  setApiKey,
  normalizeKey,
  maskKey,
  apiKey,
  mapColumns,
  testProvider,
  buildPrompt,
  parseAnswer,
  validateMappings,
  sanitizeColumns,
  mapRows,
  buildSheetPrompt,
  sanitizeSheetRows,
  validateSheetRows,
  isBareAmbiguousNumber,
  isInsurerName,
  getModels,
  rankModels,
  prewarm,
  clearCache,
};
})(typeof self !== "undefined" ? self : globalThis);

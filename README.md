# Smart Excel Consolidator

A **fully local, standalone web application** that consolidates 100+ differently structured
French CSV and XLSX customer/health-insurance files into **one standardized Excel workbook**
built around the 11 official French headers.

> No database, no login, no upload, **no server to install**. Every byte is processed in your
> browser. An optional **Google Gemini / Groq** assistant (your own API key, kept in your
> browser) helps classify ambiguous columns and read every row; without it everything keeps
> working 100% locally.

---

## 1. Requirements (end users)

- A modern Windows browser: **Microsoft Edge, Google Chrome or Firefox** (any version released in the last few years).
- **Nothing to install.** No Node.js, Python, PHP, server, or runtime is required.
- The complete project folder (the application needs its `css/`, `js/` and `vendor/` sub-folders).

## 2. Windows quick start

1. Unzip `smart-excel-consolidator.zip` anywhere on your computer (e.g. `Documents\`).
2. Open the folder and **double-click `index.html`** — it opens in your default browser.
3. Follow the numbered steps shown on screen.
4. To move the application to another computer or a USB key, copy the **whole folder**,
   not only `index.html` (the local libraries live in `vendor/` and `js/`).

The page can also be opened with *Right click → Open with → your browser*.
It works offline: airplane mode is fine (the optional AI assistant is simply skipped).

**Optional AI assistant (Google Gemini or Groq).** Paste your free API key
(Gemini: https://aistudio.google.com/apikey — Groq: https://console.groq.com)
into the assistant panel and click **Activer la clé**. The key is validated directly
with the provider, then stored **in your browser only** (`localStorage`) so it is
reused on the next opening; it is sent only to the provider itself, never anywhere
else. The free-tier model is picked automatically. If the key is missing, the quota
is exhausted or the network is down, the local engine finishes the job.

---

## 3. The five-step workflow

### Step 1 — Configure Headers

- The **built-in official 11-header configuration is active immediately**, so you can start
  without any template.
- **Manual configuration** — the 11 header labels are always editable in step 1 (order is locked);
  save them, or reset to the official values at any time. No template file and no JSON
  import/export is needed.

The exact mandatory output columns, always in this order:

| # | Official header | # | Official header |
|---|---|---|---|
| 1 | Nom | 7 | Ville |
| 2 | Prénom | 8 | Email |
| 3 | Téléphone 1 | 9 | Cotisation |
| 4 | Téléphone 2 | 10 | Mutuelle |
| 5 | Adresse | 11 | Commentaire |
| 6 | Code postal | | |

### Step 2 — Upload Files

- **Browse Files** or **drag & drop** multiple `.csv`, `.txt`, `.xlsx`, `.xlsm`, `.xls` files.
- File name, type, size and status are listed; a counter and total size are shown.
- **Duplicate files** (same name and size) are flagged with a warning and skipped during processing.
- Remove individual files or use **Remove all** before processing.
- Unsupported file types are rejected up front.
- A memory warning appears for very large batches (over ~300 MB or 150 files). Processing is
  sequential and memory-conscious, but browser limits still apply — split very large batches if needed.

### Step 3 — Process (fully automatic)

- Click **Start Processing**. The button locks immediately (a second concurrent run is impossible),
  and files are processed one at a time so the interface stays responsive.
- Real, honest progress is shown: per-file stage and percentage, plus live summary cards
  (total, successful, failed, rows added).
- CSV handling covers semicolon, comma, tab and pipe delimiters, quoted fields (including embedded
  newlines), UTF-8/BOM, UTF-16 and legacy **Windows-1252** encoded French files.
- XLSX handling covers multiple worksheets, empty sheets, and header rows that are not on row 1.
  Use a file's **Details** button to choose worksheets and correct the detected header row manually.
- **No manual validation step.** Each column is mapped automatically inside the processing loop:
  a known header or recognisable values go straight to their official column; a full-name column
  goes whole into « Nom »; ID/reference/date/year columns are excluded; an unrecognised header whose
  values are only a guess goes to « Commentaire » (or is detected from the real values on the export
  pass) so no information is ever lost. Every decision is listed in the report.
- **Cancel** stops between files, restores the interface and keeps partial results.
- **Retry failed files** reprocesses them without duplicating rows (records are keyed by
  file + worksheet + source row).

### Step 4 — Correction manuelle / À valider

- Any value the assistant could not place **with certainty** is listed here — nothing is ever
  guessed silently into a column.
- Pick the destination for each value (or accept the AI suggestion / send everything to
  « Commentaire »), then **Appliquer et recalculer**.
- The export stays **blocked** while unvalidated values remain; the banner above the table
  shows how many values are still waiting.

### Step 5 — Export Results

- **Export Final Excel** generates `consolidation.xlsx` and **re-opens the generated bytes to
  validate** the exact 11 headers/order, row counts, text preservation of phone numbers, postal
  codes and amounts, and the presence of the report sheet. The download is blocked if any check fails.
- The main worksheet `tableau final` contains exactly the 11 official columns.
- A separate `Rapport` worksheet lists warnings, invalid values, unmapped/excluded columns,
  empty worksheets and file-level results (samples are limited to what is needed for resolution).
- **Export Processing Report** downloads a dedicated `processing-report.xlsx` (warnings + per-file results).
- The workbook is generated with clean, professional formatting (no external template needed).

---

## 4. How the local "smart agent" works

The mapping starts with a **deterministic, rule-based local engine**, and the optional AI
assistant only re-checks the columns this engine could not decide:

1. **Header normalization** — lower-casing, accent stripping, punctuation/underscore removal,
   abbreviation expansion (`cp` → `code postal`, `tel` → `telephone`…), tolerant singular/plural
   matching and word-order independence.
2. **Semantic dictionaries** — hundreds of French and English synonyms per destination
   (e.g. `Nom de famille`, `Surname`, `Last Name` → `Nom`; `Courriel`, `E-mail` → `Email`;
   `Organisme complémentaire santé`, `Assureur` → `Mutuelle`).
3. **Actual cell-value analysis** — every column is profiled for French phone numbers, emails,
   five-digit postal codes, French amount formats (`75,50`, `1 250,00 €`), street vocabulary,
   city/name shapes, known mutual insurers and long-text comments.
4. **Confidence scoring 0–100** combining header evidence and value evidence, with explicit
   human-readable reasons. Contradictory evidence (e.g. invalid values) caps the score.
5. **Blocker rules** — identifier/reference columns (`numéro client`, `matricule`, `IBAN`, `SIRET`…)
   can never be promoted to phone/postal/amount fields just because they contain digits.
6. **Conflict resolution** — primary/secondary phone pairing with header hints and review flags;
   duplicate destinations resolved by strongest evidence; comments merged with clear separators.
7. **Optional AI pass** — for the remaining ambiguous columns only (header + up to 6 short
   sample values, one batched request per worksheet, answers cached per header for the session).
   When the assistant is ticked it also reads **every row of every sheet** and places each piece
   of information in the correct official column. A short deadline guarantees the batch never
   waits on it. An unrecognised header whose values are only a *shape* guess is kept in
   « Commentaire » unless data proves otherwise, so no value is ever silently dropped or sent
   to the wrong column.

### French data rules implemented

- **Phones**: national formats, `+33`/`0033` conversion to the national `0` prefix only when valid,
  landlines `01`–`05`/`09`, mobiles `06`/`07`, international numbers preserved, stored as text with
  leading zeros, never invented; single number always goes to `Téléphone 1`.
- **Postal codes**: five digits including leading zeros and overseas `97/98` codes, kept as text.
- **Emails**: trimmed, preserved otherwise, suspicious formats flagged.
- **Cotisation**: French decimal/thousand formats and currency symbols recognized; the original
  amount text is preserved (no silent monthly/annual conversion; the detected period is reported).
- **Names/addresses/cities/comments**: accents, hyphens, apostrophes and meaningful line breaks
  preserved; unsplittable full-name fields are kept whole instead of being guessed.
- **Formula injection**: exported cells starting with `= + - @ Tab CR` are prefixed with an
  apostrophe so spreadsheets treat them as text.

## 5. Privacy & security

- All processing is in-memory, in the browser tab. The only network calls are the optional
  AI requests sent **directly from your browser to Google Gemini or Groq** when the assistant
  is ticked; no CDN, no analytics, and no customer data is ever sent anywhere else — only
  headers, a handful of sample values per ambiguous column, and (in full-sheet mode) the row
  texts being classified.
- The only thing stored in the browser (`localStorage`) is **your own provider API key**, so
  you do not have to paste it again on the next opening. No customer data is ever persisted:
  clear your browser storage to remove the key.
- Downloaded files are the only data written to disk, and only when you click an export button.
- Workbook formulas/macros are never executed; imported cells are read as formatted text and
  treated as untrusted data.
- The bundled SheetJS library (`vendor/xlsx.full.min.js`, v0.20.3) is a local file, loaded from disk.

## 6. Known limitations (stated honestly)

- Legacy binary `.xls` (OLE) files open, but very old exotic encodings rely on SheetJS codepages;
  CSV encoding auto-detection covers UTF-8, UTF-16 and Windows-1252 (the common French cases).
- Scanned/PDF files and password-protected workbooks are not supported.
- Extremely large workbooks are bounded by browser memory; the app warns early and processes
  sequentially rather than pretending unlimited capacity.
- Parsing and export run on the main thread: the UI shows a progress overlay during heavy
  operations, keeping the architecture simple and fully offline.

## 7. Project structure

```
smart-excel-consolidator/
├── index.html                 # Application shell (double-click to run)
├── css/
│   └── styles.css             # Responsive professional design
├── js/
│   ├── config.js              # Official headers, statuses, limits
│   ├── dictionaries.js        # Synonyms, blockers, French value vocabulary
│   ├── normalizers.js         # French phone/CP/email/amount rules & sanitization
│   ├── csv-parser.js          # RFC 4180 parser, encoding & delimiter detection
│   ├── xlsx-reader.js         # Local SheetJS wrapper (text-only, no formulas)
│   ├── agent.js               # Smart mapping engine (scoring, conflicts, decisions)
│   ├── processor.js           # Batch pipeline, session, automatic resolution, reporting
│   ├── exporter.js            # Workbook generation + re-open validation
│   ├── ui.js                  # DOM rendering
│   ├── worker-client.js       # Parsing/export engine (main thread)
│   ├── ai/
│   │   ├── providers.js       # Direct Gemini/Groq calls (model ranking, prompts, validation)
│   │   └── client.js          # Assistant state machine (full-sheet mode, retries, cache)
│   └── app.js                 # Controller and event wiring
└── vendor/
    └── xlsx.full.min.js       # SheetJS 0.20.3 (local, offline, includes codepages)
```

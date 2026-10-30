/* ============================================================
 * Smart Excel - minimal .env loader (no dependency)
 * Reads KEY=VALUE lines, supports comments, quotes and "export ".
 * Real environment variables always win over the file.
 * ============================================================ */
"use strict";
const fs = require("fs");

function parseEnv(text) {
  const out = {};
  String(text || "")
    .split(/\r?\n/)
    .forEach((line) => {
      const s = line.trim();
      if (!s || s.startsWith("#")) return;
      const m = s.match(/^(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/);
      if (!m) return;
      let value = m[2].trim();
      if (
        (value.startsWith('"') && value.endsWith('"')) ||
        (value.startsWith("'") && value.endsWith("'"))
      ) {
        value = value.slice(1, -1);
      } else {
        value = value.replace(/\s+#.*$/, "").trim();
      }
      out[m[1]] = value;
    });
  return out;
}

function loadEnv(file) {
  let text = "";
  try {
    text = fs.readFileSync(file, "utf8");
  } catch (e) {
    text = "";
  }
  const parsed = parseEnv(text);
  Object.keys(parsed).forEach((k) => {
    if (process.env[k] === undefined || process.env[k] === "")
      process.env[k] = parsed[k];
  });
  return parsed;
}

module.exports = { loadEnv, parseEnv };

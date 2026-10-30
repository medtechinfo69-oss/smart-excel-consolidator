/* ============================================================
 * Smart Excel - RFC 4180 CSV parser
 * - BOM detection (UTF-8 / UTF-16 LE/BE)
 * - Encoding fallback to Windows-1252 for legacy French files
 * - Delimiter sniffing (; , tab |)
 * - Quoted fields, escaped quotes "", embedded newlines
 * Pure: ArrayBuffer -> { rows, encoding, delimiter, hadBom }
 * ============================================================ */
(function (root) {
  'use strict';
  var SEC = (root.SEC = root.SEC || {});

  function newDecoder(label, opts) {
    if (SEC.Encoding && SEC.Encoding.decoder) return SEC.Encoding.decoder(label, opts);
    return new TextDecoder(label, opts);
  }

  function decodeBuffer(bytes) {
    /* BOM detection */
    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) {
      return { text: newDecoder('utf-8').decode(bytes.subarray(3)), encoding: 'utf-8', hadBom: true };
    }
    if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) {
      return { text: newDecoder('utf-16le').decode(bytes.subarray(2)), encoding: 'utf-16le', hadBom: true };
    }
    if (bytes.length >= 2 && bytes[0] === 0xFE && bytes[1] === 0xFF) {
      return { text: newDecoder('utf-16be').decode(bytes.subarray(2)), encoding: 'utf-16be', hadBom: true };
    }
    /* Try strict UTF-8 first; French legacy files often come from Excel in Windows-1252. */
    try {
      var text = newDecoder('utf-8', { fatal: true }).decode(bytes);
      return { text: text, encoding: 'utf-8', hadBom: false };
    } catch (e) {
      return { text: newDecoder('windows-1252').decode(bytes), encoding: 'windows-1252', hadBom: false };
    }
  }

  function detectDelimiter(text) {
    var sample = text.split(/\r?\n/).slice(0, 10).join('\n');
    var candidates = [';', ',', '\t', '|'];
    var best = ';';
    var bestScore = -1;
    candidates.forEach(function (d) {
      var lines = sample.split(/\r?\n/).filter(function (l) { return l.length > 0; });
      var counts = lines.map(function (l) {
        var inQuotes = false, count = 0;
        for (var i = 0; i < l.length; i++) {
          if (l[i] === '"') inQuotes = !inQuotes;
          else if (l[i] === d && !inQuotes) count++;
        }
        return count;
      });
      if (!counts.length) return;
      var avg = counts.reduce(function (a, b) { return a + b; }, 0) / counts.length;
      var consistent = counts.filter(function (c) { return c === counts[0] && c > 0; }).length;
      var score = avg * 10 + consistent;
      if (score > bestScore) { bestScore = score; best = d; }
    });
    return best;
  }

  /* RFC 4180 state machine. */
  function parseCSVText(text, delimiter) {
    var rows = [];
    var row = [];
    var field = '';
    var inQuotes = false;
    var i = 0, c;
    var pushField = function () { row.push(field); field = ''; };
    var pushRow = function () { rows.push(row); row = []; };

    while (i < text.length) {
      c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { inQuotes = true; i++; continue; }
      if (c === delimiter) { pushField(); i++; continue; }
      if (c === '\r') {
        if (text[i + 1] === '\n') i++;
        pushField(); pushRow(); i++; continue;
      }
      if (c === '\n') { pushField(); pushRow(); i++; continue; }
      field += c; i++;
    }
    /* Last field/row (file without trailing newline). */
    if (field.length > 0 || row.length > 0) { pushField(); pushRow(); }

    /* Drop trailing fully-empty rows produced by blank end lines. */
    while (rows.length && rows[rows.length - 1].every(function (f) { return f.trim() === ''; })) {
      rows.pop();
    }
    return rows;
  }

  function parseCSV(arrayBuffer) {
    var bytes = new Uint8Array(arrayBuffer);
    var decoded = decodeBuffer(bytes);
    var delimiter = detectDelimiter(decoded.text);
    var rows = parseCSVText(decoded.text, delimiter);
    return {
      rows: rows,
      encoding: decoded.encoding,
      delimiter: delimiter,
      hadBom: decoded.hadBom
    };
  }

  SEC.CSV = { parse: parseCSV, parseText: parseCSVText, detectDelimiter: detectDelimiter, decodeBuffer: decodeBuffer };
})(typeof self !== 'undefined' ? self : globalThis);

/* ============================================================
 * Smart Excel - XLSX reader (local SheetJS wrapper)
 * Formulas/macros are never evaluated: cells are read as formatted text.
 * ============================================================ */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});

  function getXLSX() {
    if (typeof XLSX !== "undefined" && XLSX && XLSX.read) return XLSX;
    if (root.SEC && root.SEC.XLSXLib && root.SEC.XLSXLib.read)
      return root.SEC.XLSXLib;
    if (typeof require === "function") {
      try {
        return require("../vendor/xlsx.full.min.js");
      } catch (e) {
        /* node tests */
      }
    }
    throw new Error("SheetJS library (xlsx.full.min.js) is missing");
  }

  function checkSignature(arrayBuffer, ext) {
    var bytes = new Uint8Array(arrayBuffer);
    /* Reject files whose extension claims Excel but whose bytes are not a
     * valid workbook: XLSX/XLSM = ZIP (PK..), legacy XLS = OLE compound. */
    var isZip = bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4b;
    var isOle =
      bytes.length >= 4 &&
      bytes[0] === 0xd0 &&
      bytes[1] === 0xcf &&
      bytes[2] === 0x11 &&
      bytes[3] === 0xe0;
    if (!isZip && !isOle) {
      throw new Error(
        "Fichier illisible ou corrompu (signature de classeur Excel introuvable)",
      );
    }
  }

  /* Read every worksheet into plain string arrays.
   * Returns [{ name, rows, isEmpty }] and releases the workbook object. */
  function parseFromLib(X, arrayBuffer, ext) {
    /* Normalise exotic buffers (cross-realm ArrayBuffers, typed-array views):
     * SheetJS only accepts a plain ArrayBuffer for type:'array', and a view
     * would otherwise be read from its pooled root — corrupted sheets. */
    if (!(arrayBuffer instanceof ArrayBuffer)) {
      var tmp = new Uint8Array(arrayBuffer);
      var plain = new ArrayBuffer(tmp.byteLength);
      new Uint8Array(plain).set(tmp);
      arrayBuffer = plain;
    }
    checkSignature(arrayBuffer, ext);
    var wb = X.read(arrayBuffer, {
      type: "array",
      cellDates: false,
      cellFormula: false, // do not retain formulas
      cellHTML: false,
      cellNF: false,
      /* Style/theme parsing is the slowest part of SheetJS and this reader
       * only needs the formatted TEXT: skipping it speeds up big styled
       * workbooks with byte-identical output (verified). */
      cellStyles: false,
      codepage: 1252, // legacy French workbooks
    });

    var sheets = wb.SheetNames.map(function (name) {
      var ws = wb.Sheets[name];
      var rows = X.utils.sheet_to_json(ws, {
        header: 1,
        raw: false, // formatted text -> preserves displayed values
        defval: "",
        blankrows: false,
      });
      /* Normalize ragged rows and coerce everything to string. */
      rows = rows.map(function (r) {
        return (r || []).map(function (cell) {
          if (cell === null || cell === undefined) return "";
          return String(cell);
        });
      });
      var nonEmpty = rows.some(function (r) {
        return r.some(function (c) {
          return String(c).trim() !== "";
        });
      });
      return { name: name, rows: rows, isEmpty: !nonEmpty };
    });

    /* Help garbage collection of the heavy workbook structures. */
    wb.SheetNames.forEach(function (n) {
      wb.Sheets[n] = null;
    });
    wb.Sheets = {};
    return sheets;
  }

  function parseWorkbook(arrayBuffer) {
    return parseFromLib(getXLSX(), arrayBuffer);
  }

  /* Allow the worker (or a test) to inject its own SheetJS instance. */
  function setLib(lib) {
    if (lib && lib.read) SEC.XLSXLib = lib;
  }

  SEC.XLSXReader = {
    parse: parseWorkbook,
    parseFromLib: parseFromLib,
    checkSignature: checkSignature,
    getXLSX: getXLSX,
    setLib: setLib,
  };
})(typeof self !== "undefined" ? self : globalThis);

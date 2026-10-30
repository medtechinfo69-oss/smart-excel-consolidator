/* ============================================================
 * Smart Excel - text decoding shim
 * TextDecoder is a browser/Node global but NOT a property of a
 * jsdom window: this shim keeps CSV parsing working in both.
 * ============================================================ */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});

  SEC.Encoding = {
    has: function () {
      return typeof TextDecoder === "function";
    },
    decoder: function (label, opts) {
      if (typeof TextDecoder === "function")
        return new TextDecoder(label, opts);
      if (typeof require === "function") {
        return new (require("util").TextDecoder)(label, opts);
      }
      throw new Error("TextDecoder indisponible dans cet environnement");
    },
  };
})(typeof self !== "undefined" ? self : globalThis);

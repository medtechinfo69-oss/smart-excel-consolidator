/* Compatibility shim.
 * The AI assistant is Google Gemini and routes every call through the local
 * server so the API key stays in .env.
 * See js/ai/client.js (browser client) and server/ai-providers.js (backend).
 * SEC.Gemini is kept as an alias of SEC.AI for older integrations/tests. */
(function (root) {
  "use strict";
  var SEC = (root.SEC = root.SEC || {});
  SEC.Gemini = SEC.AI || null;
})(typeof self !== "undefined" ? self : globalThis);

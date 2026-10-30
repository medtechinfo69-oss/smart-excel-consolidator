/* Generates edge-adresses.xlsx: address cells carrying trailing/leading
 * whitespace that normalization cleans but the export pipeline rewrites
 * verbatim, reproducing the export-time row/column mismatch. Run once:
 *   node test/fixtures/edge-address.xlsx.js
 */
const XLSX = require("../../vendor/xlsx.full.min.js");
const wb = XLSX.utils.book_new();
/* eslint-disable */
const aoa = [
  ["Nom", "Prenom", "Adresse", "Code postal", "Ville"],
  ["DUPONT", "Jean", "  12 rue des Lilas  ", "01200", "Beligneux"],
  ["MARTIN", "Marie", "3 av. de la Gare", "75011", "  Paris  "],
  ["BERTRAND", "Leo", "  lot. Les Chenes", "30100", "Ales"],
];
const ws = XLSX.utils.aoa_to_sheet(aoa);
XLSX.utils.book_append_sheet(wb, ws, "Clients");
require('fs').writeFileSync(require('path').resolve('test/fixtures/edge-adresses.xlsx'), Buffer.from(XLSX.write(wb, { type: 'array', bookType: 'xlsx' })));
console.log("written");

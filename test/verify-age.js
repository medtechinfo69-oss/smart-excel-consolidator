/* Insurer inside a Cotisation column + merged "age | insurer" values. */
require('../js/config.js');
require('../js/data/communes.js');
require('../js/communes.js');
require('../js/dictionaries.js');
require('../js/data/mutuelles.js');
require('../js/normalizers.js');
require('../js/csv-parser.js');
require('../js/xlsx-reader.js');
require('../js/agent.js');
require('../js/processor.js');
const SEC = globalThis.SEC;

const csv = Buffer.from(
  'Nom;Prenom;COTISATION;CP;Ville\n' +
  'Dupont;Jean;AESIO;75001;Paris\n' +
  'Martin;Luc;alliance sante;31000;Toulouse\n' +
  'Roux;Paul;Swiss Life;69002;Lyon\n' +
  'Petit;Anne;75,50;13001;Marseille\n' +
  'Perrin;Marc;harmonie mutuelle;44000;Nantes\n' +
  'Blanc;Lea;70+75;59000;Lille\n', 'utf8');

(async () => {
  const s = SEC.Processor.createSession();
  const e = s.registerFile({ name: 'mut.csv', size: csv.length });
  await s.startBatch([e], { readBuffer: () => Promise.resolve(csv), yieldFn: () => Promise.resolve() });
  const idx = k => SEC.TARGET_KEYS.indexOf(k);
  console.log('NOM    | COTISATION | MUTUELLE        | COMMENTAIRE | CP    ');
  s.buildConsolidatedRows().forEach(r =>
    console.log([r[idx('nom')], r[idx('cotisation')], r[idx('mutuelle')],
      r[idx('commentaire')], r[idx('cp')]].map(v => String(v || '-').padEnd(12)).join('|')));
})();

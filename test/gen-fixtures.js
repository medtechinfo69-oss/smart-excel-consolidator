/* ============================================================
 * Test fixture generator - realistic French insurance datasets.
 * Run: node test/gen-fixtures.js
 * ============================================================ */
'use strict';
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');

const ROOT = path.join(__dirname, 'fixtures');
fs.rmSync(ROOT, { recursive: true, force: true });
fs.mkdirSync(path.join(ROOT, 'batch'), { recursive: true });

function writeCSV(file, rows, delimiter, encoding) {
  let text = rows.map(r => r.map(csvCell).join(delimiter)).join('\r\n');
  let bytes;
  if (encoding === 'utf8-bom') bytes = Buffer.concat([Buffer.from([0xEF, 0xBB, 0xBF]), Buffer.from(text, 'utf8')]);
  else if (encoding === 'cp1252') bytes = Buffer.from(encodeCP1252(text), 'binary');
  else bytes = Buffer.from(text, 'utf8');
  fs.writeFileSync(path.join(ROOT, file), bytes);
}
function csvCell(v) {
  v = String(v == null ? '' : v);
  return /[;,"\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
}
function encodeCP1252(s) {
  const map = { 'à': 0xE0, 'á': 0xE1, 'â': 0xE2, 'ä': 0xE4, 'å': 0xE5, 'ç': 0xE7, 'è': 0xE8, 'é': 0xE9, 'ê': 0xEA, 'ë': 0xEB, 'î': 0xEE, 'ï': 0xEF, 'ô': 0xF4, 'ö': 0xF6, 'ù': 0xF9, 'û': 0xFB, 'ü': 0xFC, 'É': 0xC9, 'È': 0xC8, 'À': 0xC0, 'Ç': 0xC7, 'Œ': 0x9C, 'œ': 0x9D, '€': 0x80, '«': 0xAB, '»': 0xBB, '’': 0x92, '–': 0x96 };
  let out = '';
  for (const ch of s) {
    if (ch.charCodeAt(0) < 128) out += ch;
    else if (map[ch] !== undefined) out += String.fromCharCode(map[ch]);
    else out += '?';
  }
  return out;
}
function writeXLSX(file, sheets) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(sh => {
    const ws = XLSX.utils.aoa_to_sheet(sh.rows);
    if (sh.cols) ws['!cols'] = sh.cols;
    XLSX.utils.book_append_sheet(wb, ws, sh.name);
  });
  fs.writeFileSync(path.join(ROOT, file), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}

/* ---------- 01 : semicolon CSV, BOM, French headers, all phone formats ---------- */
writeCSV('01_clients_semicolon.csv', [
  ['Nom de famille', 'Prénom', 'Portable', 'Téléphone fixe', 'Adresse', 'CP', 'Commune', 'Courriel', 'Cotisation mensuelle', 'Organisme complémentaire santé', 'Notes client'],
  ['Dupont', 'Jean', '06 12 34 56 78', '01 45 67 89 10', '12 rue de la Paix', '75002', 'Paris', 'jean.dupont@example.fr', '75,50', 'Harmonie Mutuelle', 'Client fidèle'],
  ['Martin', 'Sophie', '+33 6 12 34 56 78', '', '8 avenue des Champs-Élysées', '01200', 'Bellegarde-sur-Valserine', 'sophie.martin@example.com', '1 250,00 €', 'MGEN', ''],
  ['Bernard', 'Élodie', '0033712345678', '02 34 56 78 90', '"10 rue; grande entrée"', '97400', 'Saint-Denis', 'elodie.bernard@example.fr', '120 €', 'AXA Santé', 'Prévenir avant livraison'],
  ['Petit', 'François', '0712345678', '', '1 bis place du Marché', '44000', 'Nantes', 'francois.petit@example.fr', '99,99', 'Malakoff Humanis', ''],
  ['Robert', "Aïcha", '06 98 76 54 32', '04 50 60 70 80', "Bâtiment B, 3 allée des Tilleuls", '69003', 'Lyon', 'aicha.robert@example.fr', '45', 'Groupama', 'Adresse en cours de mise à jour'],
  ['Richard', "Pierre-Yves", '', '05 56 00 11 22', '22 boulevard du Général de Gaulle', '33000', 'Bordeaux', 'py.richard@example.fr', '210,75', 'MAIF', ''],
  ['Durand', 'Marie-Claire', '06.11.22.33.44', '', "Résidence Les Pins, appt 14", '06000', 'Nice', 'marieclaire.durand@example.fr', '88,20', 'Mutuelle Générale', 'Conjoint également assuré'],
  ['Moreau', 'Kévin', '+33712345678', '03 20 12 34 56', '5 impasse des Lilas', '59000', 'Lille', 'kevin.moreau@example.fr', '1 050,00', 'AG2R La Mondiale', '']
], ';', 'utf8-bom');

/* ---------- 02 : comma CSV, English headers, quoted comma + embedded newline ---------- */
writeCSV('02_clients_comma.csv', [
  ['Last Name', 'First Name', 'Mobile Number', 'Phone 2', 'Address', 'Postal Code', 'City', 'Email', 'Monthly Contribution', 'Insurer', 'Comments'],
  ['Smith', 'John', '0612345678', '0145678910', '10 Downing Street, Apt 3', '75001', 'Paris', 'john.smith@example.com', '75.50', 'April', 'Quoted, with comma'],
  ['Garcia', 'Maria', '0712345678', '', 'Calle Mayor 12', '31000', 'Toulouse', 'maria.garcia@example.com', '1200.00', 'SwissLife', 'Line one\nLine two in same cell'],
  ['Nguyen', 'Thi', '0687654321', '0987654321', '3 rue Victor Hugo', '69001', 'Lyon', 'thi.nguyen@example.fr', '59,90', 'MNH', '']
], ',', 'utf8');

/* ---------- 03 : tab-delimited TXT ---------- */
writeCSV('03_tab_delimited.txt', [
  ['Surname', 'Given Name', 'Mobile', 'Fixed line', 'Street', 'Zip', 'Town', 'E-mail', 'Amount', 'Mutual', 'Notes'],
  ['Lefevre', 'Camille', '0612121212', '0123232323', '4 rue de la Gare', '76000', 'Rouen', 'camille.lefevre@example.fr', '82,10', 'Harmonie', 'Tab separated'],
  ['Guerin', 'Lucas', '0734343434', '', '9 chemin Vert', '13001', 'Marseille', 'lucas.guerin@example.fr', '67', 'MGEN', '']
], '\t', 'utf8');

/* ---------- 04 : legacy Windows-1252 encoding ---------- */
writeCSV('04_legacy_cp1252.csv', [
  ['Nom', 'Prénom', 'Portable', 'Adresse', 'Code postal', 'Ville', 'Email', 'Cotisation', 'Mutuelle', 'Commentaire'],
  ['Béranger', 'Élodie', '0612345600', "10 rue de l'église", '44100', 'Nantes', 'elodie@example.fr', '75,50', 'Mutuelle Générale', 'Clientèle très âgée'],
  ['François', 'Jérôme', '0712345600', '2 allée des Châtaigniers', '75018', 'Paris', 'jerome@example.fr', '90,00', 'AXA', 'Envoyer à M. François']
], ';', 'cp1252');

/* ---------- 05 : multi-sheet XLSX, late header row, empty sheet, extra cols ---------- */
writeXLSX('05_workbook_multi.xlsx', [
  {
    name: 'Base clients',
    rows: [
      ['Export du service commercial - octobre 2026'],
      [],
      ['Attention : fichier interne'],
      ['NOM', 'PRENOM', 'TEL MOBILE', 'ADRESSE COMPLETE', 'CODE POSTAL', 'COMMUNE', 'MAIL', 'MONTANT COTISATION', 'CAISSE', 'REMARQUES'],
      ['Leroy', 'Nathalie', '0612000001', '1 rue de Rivoli', '75004', 'Paris', 'nathalie.leroy@example.fr', '75,50', 'MGEN', ''],
      ['Moreau', 'Christophe', '0612000002', '2 rue de Rivoli', '75004', 'Paris', 'christophe@example.fr', '80,00', 'AXA', 'Téléphone sur répondeur'],
      ['Girard', 'Isabelle', '0612000003', '3 rue de Rivoli', '75004', 'Paris', 'isa.girard@example.fr', '110,25', 'Harmonie', '']
    ]
  },
  { name: 'Feuille vide', rows: [['', ''], ['', '']] },
  {
    name: 'Autre format',
    rows: [
      ['Numéro client', 'Nom de famille', 'First name', 'Portable', 'Voie', 'CP', 'Ville', "Date d'adhésion", 'Cotisation annuelle', 'Assureur'],
      ['C-100234', 'Faure', 'Antoine', '0612000004', '5 place Bellecour', '69002', 'Lyon', '2024-03-12', '960,00', 'Groupama'],
      ['C-100235', 'Andre', 'Manon', '0612000005', '6 place Bellecour', '69002', 'Lyon', '2025-01-08', '1 020,00', 'MAIF'],
      ['C-100236', 'Mercier', 'Théo', '', '7 place Bellecour', '69002', 'Lyon', '2025-06-20', '840,00', 'MMA']
    ]
  }
]);

/* ---------- 06 : ambiguous mappings, ID blocker, full name, two phones ---------- */
writeXLSX('06_ambiguous.xlsx', [
  {
    name: 'Doutes',
    rows: [
      ['Nom', 'Prénom', 'Numéro', 'Numéro client', 'Nom complet', 'Téléphone A', 'Téléphone B', 'Email', 'Ville', 'Code Postal'],
      ['Blanc', 'Olivier', '0612000010', '100234', 'Olivier Blanc', '0612000011', '0145000011', 'olivier@example.fr', 'Paris', '75009'],
      ['Noir', 'Caroline', '0612000020', '100235', 'Caroline Noir', '0612000021', '0145000021', 'caroline@example.fr', 'Lyon', '69002'],
      ['Vert', 'Julien', '0612000030', '100236', 'Julien Vert', '0612000031', '0145000031', 'julien@example.fr', 'Lille', '59000']
    ]
  }
]);

/* ---------- 07 : edge cases (blank rows, invalid data, formula injection) ---------- */
writeXLSX('07_edge_cases.xlsx', [
  {
    name: 'Cas limites',
    rows: [
      ['Nom', 'Prénom', 'Téléphone 1', 'Adresse', 'Code postal', 'Ville', 'Email', 'Cotisation', 'Mutuelle', 'Commentaire'],
      ['', '', '', '', '', '', '', '', '', ''],
      ['Carre', 'Lucie', '0612000040', '8 rue des Cerisiers', '75011', 'Paris', 'lucie@example.fr', '75,50', 'MGEN', ''],
      ['Rond', 'Hugo', '123', '9 rue des Cerisiers', '99999', 'Paris', 'pas-un-email', 'abc', 'AXA', '=1+1'],
      ['Triangle', 'Emma', '0612000042', '10 rue des Cerisiers', '01300', 'Béligneux', 'emma@example.fr', '-45,00', 'MAIF', '+CMD injecté'],
      [],
      ['Losange', 'Nathan', '0612000043', '11 rue des Cerisiers', '97400', 'Saint-Denis', 'nathan@example.fr', '120', 'Harmonie', '@injecté']
    ]
  }
]);

/* ---------- 08 : corrupted XLSX ---------- */
fs.writeFileSync(path.join(ROOT, '08_corrupted.xlsx'), Buffer.from('This is definitely not a real zip/xlsx workbook payload.', 'utf8'));

/* ---------- 09 : unsupported extension ---------- */
fs.writeFileSync(path.join(ROOT, '09_unsupported.dat'), Buffer.from('binary-ish data', 'utf8'));

/* ---------- 10 : template with wrong headers (validation failure case) ---------- */
writeXLSX('10_headers_wrong.xlsx', [
  { name: 'Feuille1', rows: [['Nom', 'Prénom', 'Téléphone', 'Adresse', 'Ville']] }
]);

/* ---------- Official template (styled) ---------- */
(function () {
  const wb = XLSX.utils.book_new();
  const headers = ['Nom', 'Prénom', 'Téléphone 1', 'Téléphone 2', 'Adresse', 'Code postal', 'Ville', 'Email', 'Cotisation', 'Mutuelle', 'Commentaire'];
  const ws = XLSX.utils.aoa_to_sheet([
    headers,
    ['Template', 'Exemple', '0600000000', '', '1 rue exemple', '75001', 'Paris', 'exemple@test.fr', '10,00', 'MGEN', 'Ligne de démonstration']
  ]);
  ws['!cols'] = headers.map(() => ({ wch: 18 }));
  XLSX.utils.book_append_sheet(wb, ws, 'En-tête');
  fs.writeFileSync(path.join(ROOT, 'En tete finale.xlsx'), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
})();

/* ---------- 110 randomized batch files (100+ requirement) ---------- */
function mulberry32(a) {
  return function () {
    a |= 0; a = (a + 0x6D2B79F5) | 0;
    let t = Math.imul(a ^ (a >>> 15), 1 | a);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
const rnd = mulberry32(20261027);
const pick = arr => arr[Math.floor(rnd() * arr.length)];
const shuffle = arr => arr.map(v => [rnd(), v]).sort((a, b) => a[0] - b[0]).map(x => x[1]);

const variants = {
  nom: ['Nom', 'Nom de famille', 'Surname', 'Last Name', 'Patronyme'],
  prenom: ['Prénom', 'First Name', 'Given Name', 'Prénoms'],
  tel1: ['Téléphone', 'Portable', 'Mobile', 'Tel 1', 'Phone'],
  tel2: ['Téléphone 2', 'Téléphone fixe', 'Fixe', 'Tel secondaire'],
  adresse: ['Adresse', 'Adresse complète', 'Address', 'Rue', 'Voie'],
  cp: ['Code postal', 'CP', 'Postal Code', 'Zip'],
  ville: ['Ville', 'Commune', 'City', 'Localité'],
  email: ['Email', 'Courriel', 'E-mail', 'Mail', 'Adresse électronique'],
  cotisation: ['Cotisation', 'Cotisation mensuelle', 'Montant cotisation', 'Prime', 'Tarif'],
  mutuelle: ['Mutuelle', 'Organisme complémentaire santé', 'Complémentaire', 'Assureur', 'Caisse'],
  commentaire: ['Commentaire', 'Notes', 'Remarques', 'Observation', 'Comments']
};
const people = [
  ['Dupont', 'Jean', 'Paris', '75001'], ['Martin', 'Sophie', 'Lyon', '69002'], ['Bernard', 'Élodie', 'Nice', '06000'],
  ['Petit', 'François', 'Nantes', '44000'], ['Robert', 'Aïcha', 'Lille', '59000'], ['Richard', 'Pierre-Yves', 'Bordeaux', '33000'],
  ['Durand', 'Marie-Claire', 'Toulouse', '31000'], ['Moreau', 'Kévin', 'Rouen', '76000'], ['Simon', 'Manon', 'Reims', '51100'],
  ['Michel', 'Lucas', 'Saint-Denis', '97400'], ['Leroy', 'Camille', 'Strasbourg', '67000'], ['Roux', 'Théo', 'Rennes', '35000']
];
const mutuelles = ['MGEN', 'AXA Santé', 'Harmonie Mutuelle', 'Malakoff Humanis', 'Groupama', 'MAIF', 'AG2R La Mondiale'];
const phones = ['06 12 34 56 78', '0612345678', '+33 6 12 34 56 78', '0033712345678', '0712345678', '01 45 67 89 10'];
const amounts = ['75,50', '1 250,00 €', '120 €', '99,99', '45', '210,75', '1 050,00', '88.20'];

let csvCount = 0, xlsxCount = 0;
for (let i = 0; i < 110; i++) {
  const keys = shuffle(Object.keys(variants)).slice(0, 6 + Math.floor(rnd() * 6)); // 6..11 columns
  const headerRow = keys.map(k => pick(variants[k]));
  const nRows = 5 + Math.floor(rnd() * 10);
  const rows = [headerRow];
  for (let r = 0; r < nRows; r++) {
    const p = pick(people);
    const row = keys.map(k => {
      switch (k) {
        case 'nom': return p[0];
        case 'prenom': return p[1];
        case 'tel1': return pick(phones);
        case 'tel2': return rnd() > 0.4 ? pick(phones) : '';
        case 'adresse': return (1 + Math.floor(rnd() * 50)) + ' rue ' + pick(['de la Paix', 'Victor Hugo', 'des Lilas', 'du Général de Gaulle']);
        case 'cp': return p[3];
        case 'ville': return p[2];
        case 'email': return p[1].toLowerCase().replace(/[^a-z]/g, '.') + Math.floor(rnd() * 99) + '@example.fr';
        case 'cotisation': return pick(amounts);
        case 'mutuelle': return pick(mutuelles);
        case 'commentaire': return rnd() > 0.5 ? 'Client noté ' + Math.floor(rnd() * 5) + '/5' : '';
        default: return '';
      }
    });
    if (rnd() < 0.08) row.push('VALEUR NON CARTOGRAPHIÉE ' + i + '-' + r); // extra unmapped column sometimes
    rows.push(row);
  }
  if (rnd() < 0.08) rows[0].push('Numéro client'); // ID blocker column sometimes
  const useXlsx = rnd() < 0.35;
  if (useXlsx) {
    writeXLSX(path.join('batch', 'batch_' + String(i).padStart(3, '0') + '.xlsx'), [{ name: 'Data', rows: rows }]);
    xlsxCount++;
  } else {
    writeCSV(path.join('batch', 'batch_' + String(i).padStart(3, '0') + '.csv'), rows, rnd() < 0.7 ? ';' : ',', rnd() < 0.15 ? 'cp1252' : 'utf8');
    csvCount++;
  }
}

console.log('Fixtures generated. CSV:', csvCount, 'XLSX:', xlsxCount, 'in', ROOT);

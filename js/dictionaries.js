/* ============================================================
 * Smart Excel - Local semantic knowledge base
 * Synonyms, industry terms, blockers, French value vocabulary.
 * All keys are normalized headers (lowercase, accents stripped,
 * punctuation converted to spaces). See normalizers.js.
 * ============================================================ */
(function (root) {
  'use strict';
  var SEC = (root.SEC = root.SEC || {});

  /* Exact / near-exact phrase dictionary.
   * Weight: 95 = canonical term, 88 = strong synonym, 80 = industry variant. */
  SEC.PHRASE_DICT = {
    nom: {
      95: ['nom', 'nom de famille', 'nom patronymique', 'nom de naissance', 'surname', 'family name', 'last name', 'lastname'],
      88: ['patronyme', 'nom du client', 'nom client', 'nom usuel', 'nom naissance', 'client nom', 'nom beneficiaire', 'nom assure', 'nom adherent', 'adherent'],
      80: ['nom de famille client', 'familyname', 'customer last name', 'customer surname', 'client', 'customer', 'adherent nom']
    },
    prenom: {
      95: ['prenom', 'first name', 'firstname', 'given name', 'givenname'],
      88: ['prenoms', 'petit nom', 'prenom du client', 'prenom client', 'client prenom', 'prenom beneficiaire'],
      80: ['customer first name', 'forename', 'christian name']
    },
    tel1: {
      95: ['telephone', 'telephone 1', 'tel 1', 'tel1', 'telephone principal', 'tel principal', 'mobile', 'portable', 'numero de telephone', 'phone number', 'mobile number', 'phone', 'telephone mobile', 'telephone portable', 'tel mobile', 'tel portable'],
      88: ['numero portable', 'numero mobile', 'tel', 'contact', 'telephone personnel', 'mobile 1', 'portable 1', 'phone 1', 'cellphone', 'cell', 'numero telephone', 'numero de mobile', 'numero de portable', 'numero de tel'],
      80: ['tel perso', 'telephone perso', 'mobile phone', 'contact telephone', 'contact phone', 'numero contact', 'phone1']
    },
    tel2: {
      95: ['telephone 2', 'tel 2', 'tel2', 'telephone secondaire', 'tel secondaire', 'telephone fixe', 'telephone domicile', 'fixe', 'fix', 'gsm', 'gsm 2', 'tél fixe', 'tel fixe', 'landline', 'land line', 'home phone', 'telephone maison'],
      88: ['mobile 2', 'portable 2', 'phone 2', 'autre telephone', 'autre tel', 'telephone bureau', 'tel bureau', 'tel domicile', 'numero fixe', 'telephone pro', 'tel pro', 'fixe 2', 'telephone residence', 'tel residence', 'ligne fixe', 'ligne 2', 'tel perso 2'],
      80: ['second phone', 'secondary phone', 'phone2', 'tel2 ', 'fixphone', 'fixed', 'fixed line']
    },
    adresse: {
      95: ['adresse', 'adresse 1', 'adresse1', 'adresse postale', 'address', 'street address', 'rue', 'voie'],
      88: ['adresse complete', 'adresse du client', 'adresse client', 'adresse ligne 1', 'adresse ligne1', 'addr', 'street', 'adresse de residence', 'lieu dit', 'lieudit', 'residence'],
      80: ['complement adresse', 'adresse complement', 'adresse 2', 'batiment', 'appartement', 'immeuble', 'customer address', 'postal address', 'adresse geographique']
    },
    cp: {
      95: ['code postal', 'cp', 'postal code', 'zip code', 'zip', 'postcode'],
      88: ['code postale', 'code postal client', 'cp client'],
      80: ['postal zip', 'zipcode', 'customer postal code']
    },
    ville: {
      95: ['ville', 'city', 'commune', 'localite'],
      88: ['town', 'locality', 'ville de residence', 'commune de residence', 'ville client'],
      80: ['municipalite', 'customer city']
    },
    email: {
      95: ['email', 'adresse email', 'adresse mail', 'courriel', 'e mail', 'mail', 'email address', 'e mail address', 'adresse electronique'],
      88: ['mel', 'courriel electronique', 'electronic mail', 'e-mail', 'mail address', 'email client'],
      80: ['courriel client', 'contact mail', 'customer email']
    },
    cotisation: {
      95: ['cotisation', 'cotisation mensuelle', 'cotisation annuelle', 'montant cotisation', 'montant de la cotisation'],
      88: ['contribution', 'prime', 'prime assurance', 'montant', 'tarif', 'mensualite', 'montant mensuel', 'montant annuel', 'prix', 'montant de la prime', 'montant prime'],
      80: ['somme', 'montant cotisation annuelle', 'montant cotisation mensuelle', 'tarif mensuel', 'tarif annuel', 'cout', 'redevance']
    },
    mutuelle: {
      95: ['mutuelle', 'organisme complementaire', 'organisme complementaire sante', 'complementaire sante', 'complementaire', 'mutuelle sante'],
      88: ['organisme', 'assureur', 'assurance', 'caisse', 'caisse maladie', 'nom de la mutuelle', 'organisme sante', 'assureur complementaire', 'compagnie assurance', 'compagnie d assurance', 'contrat mutuelle'],
      80: ['produit sante', 'organisme de sante', 'caisse complementaire', 'institution de prevoyance', 'nom organisme', 'assurance sante']
    },
    commentaire: {
      95: ['commentaire', 'commentaires', 'remarque', 'remarques', 'note', 'notes', 'observation', 'observations'],
      88: ['memo', 'notes client', 'commentaire client', 'remarque client', 'information', 'informations', 'renseignement', 'renseignements', 'details'],
      80: ['detail', 'note client', 'note interne', 'remarques client', 'comment', 'comments', 'customer notes']
    }
  };

  /* Single-token weights used as a bag-of-words fallback (0..1). */
  SEC.TOKEN_WEIGHTS = {
    nom: { nom: 0.95, surname: 0.95, family: 0.7, patronyme: 0.9, famille: 0.6 },
    prenom: { prenom: 0.95, forename: 0.9 },
    tel1: { telephone: 0.9, tel: 0.9, phone: 0.9, mobile: 0.9, portable: 0.9, contact: 0.55, numero: 0.35 },
    tel2: { fixe: 0.85, fix: 0.85, gsm: 0.8, domicile: 0.8, bureau: 0.75, secondaire: 0.7, landline: 0.9 },
    adresse: { adresse: 0.92, address: 0.92, rue: 0.8, voie: 0.75, residence: 0.55, batiment: 0.45, appartement: 0.45 },
    cp: { postal: 0.75, postale: 0.7, code: 0.45, zip: 0.85, postcode: 0.85 },
    ville: { ville: 0.95, city: 0.9, commune: 0.85, localite: 0.8, town: 0.8, municipalite: 0.7 },
    email: { email: 0.95, mail: 0.92, courriel: 0.95, mel: 0.85, electronique: 0.5 },
    cotisation: { cotisation: 0.95, contribution: 0.85, prime: 0.8, montant: 0.75, tarif: 0.75, mensualite: 0.8, prix: 0.6, somme: 0.55, cout: 0.6, redevance: 0.6 },
    mutuelle: { mutuelle: 0.95, complementaire: 0.85, sante: 0.45, assureur: 0.85, assurance: 0.8, organisme: 0.8, caisse: 0.7, prevoyance: 0.6 },
    commentaire: { commentaire: 0.95, remarque: 0.92, note: 0.85, observation: 0.9, memo: 0.85, information: 0.6, renseignement: 0.6, detail: 0.6, comment: 0.9, comments: 0.9 }
  };

  /* Headers that identify records / contracts: their numeric-looking values
   * must never be promoted to phone, postal code or amount based on digits. */
  SEC.ID_BLOCKER_TERMS = [
    'id', 'identifiant', 'identifier', 'reference', 'ref', 'matricule',
    'numero client', 'numero adherent', 'numero contrat', 'numero dossier',
    'numero de dossier', 'numero de contrat', 'numero de client', 'numero adherent client',
    'contrat', 'siret', 'siren', 'code client', 'identifiant client',
    'numero securite sociale', 'numero ss', 'nss', 'iban', 'numero de compte',
    'compte', 'reference dossier', 'numero commande', 'numero de piece',
    'piece identite', 'numero interne', 'cle', 'no client', 'n client',
    'dossier'
  ];

  /* Headers carrying an unsplittable full name: require human review. */
  SEC.FULL_NAME_HEADERS = [
    'nom complet', 'nom prenom', 'nom et prenom', 'nom prenom client',
    'full name', 'fullname', 'name', 'civilite nom prenom',
    'nom et prenom client', 'nom raison sociale', 'raison sociale',
    'client', 'adherent', 'adherents', 'assure', 'assuree', 'personne', 'contact', 'nom du client',
    'nom du adherent', 'nom de l adherent', 'nom prenom adhérent', 'titulaire',
    'assure principal', 'adherent principal', 'beneficiaire', 'patient'
  ];

  /* Abbreviation expansion applied during header normalization. */
  SEC.ABBREVIATIONS = {
    cp: 'code postal',
    tel: 'telephone',
    mut: 'mutuelle',
    addr: 'adresse',
    adr: 'adresse',
    mel: 'email',
    pts: 'points'
  };

  /* French street vocabulary used by the value analyzer. */
  SEC.STREET_TERMS = [
    'rue', 'avenue', 'av', 'boulevard', 'bd', 'place', 'pl', 'chemin', 'chem',
    'impasse', 'allee', 'route', 'rte', 'residence', 'res', 'batiment', 'bat',
    'appartement', 'appart', 'apt', 'immeuble', 'lotissement', 'zac', 'zi',
    'zone industrielle', 'zone artisanale', 'cite', 'clos', 'passage', 'cours',
    'quai', 'square', 'hameau', 'lieu dit', 'voie', 'bvd', 'cote', 'zac '
  ];

  /* Well-known French complementary health insurers / products (value hints). */
  SEC.KNOWN_MUTUELLES = [
    'harmonie', 'mgen', 'axa', 'malakoff', 'humanis', 'groupama', 'maif', 'macif',
    'ag2r', 'april', 'swisslife', 'swiss life', 'gan', 'mma', 'mnh', 'mutuelle generale',
    'mutualia', 'adrea', 'aadp', 'alptis', 'apicil', 'audien', 'b2v', 'carcept',
    'ccmo', 'cgrm', 'ciceo', 'coralis', 'eovi', 'exance', 'fidelia', 'filhet',
    'gras savoye', 'haxo', 'henner', 'historia', 'humanis prevoyance', 'interiale',
    'ipsec', 'itelis', 'kalixia', 'lamie', 'liberte', 'macif mutualite', 'mgefi',
    'micils', 'miltis', 'mutest', 'mutami', 'mutelle', 'myriade', 'neoliane',
    'nov', 'ociane', 'officeo', 'opteo', 'ossate', 'partea', 'prevadies',
    'radiance', 'reunica', 'santevet', 'sereina', 'smam', 'smip', 'smiso',
    'solimut', 'thelem', 'unilia', 'utwin', 'vauban', 'verspieren', 'vitane',
    'mutuelle', 'complementaire', 'sante', 'prevoyance', 'assurance', 'assurances'
  ];

  /* Common French given names used as a value-level signal (small, local list). */
  SEC.COMMON_PRENOMS = [
    'jean', 'pierre', 'michel', 'philippe', 'nicolas', 'david', 'stephane', 'daniel',
    'christian', 'laurent', 'julien', 'francois', 'patrick', 'sebastien', 'olivier',
    'marie', 'nathalie', 'isabelle', 'catherine', 'sandrine', 'stephanie', 'valerie',
    'christine', 'veronique', 'sophie', 'severine', 'julie', 'camille', 'emilie',
    'marine', 'lea', 'manon', 'chloe', 'clara', 'emma', 'lucas', 'hugo', 'nathan',
    'leo', 'gabriel', 'louis', 'raphael', 'arthur', 'jules', 'adam', 'ethan',
    'maelys', 'anna', 'anne', 'jade', 'louise', 'alice', 'inees', 'ines', 'sarah',
    'mohamed', 'mhamed', 'youssef', 'amine', 'fatima', 'aicha', 'khadija', 'omar', 'ali',
    'antoine', 'benoit', 'cedric', 'fabien', 'guillaume', 'kevin', 'thomas', 'remi',
    'brigitte', 'monique', 'gisele', 'jacqueline', 'madeleine', 'simone', 'denise',
    'bernard', 'roger', 'gerard', 'claude', 'henri', 'alain', 'jacques', 'marc',
    'noelle', 'fermande', 'fernande', 'huguette', 'guy', 'serge', 'claudine', 'francoise',
    'didier', 'gilbert', 'yves', 'roland', 'joelle', 'ghislaine', 'christiane', 'claudette',
    'maryse', 'ernest', 'dominique', 'chantal', 'edith', 'odette', 'andre', 'andree',
    'christelle', 'frederique', 'florence', 'micheline', 'suzanne', 'elise', 'marius',
    'danielle', 'annie', 'samuel', 'pascale', 'odile', 'ginette', 'berthe', 'josef',
    'joseph', 'helene', 'francis', 'nadine', 'genevieve', 'rose', 'roselyne', 'rene',
    'renee', 'paulette', 'colette', 'lucien', 'mireille', 'yvette', 'joachim', 'patrice',
    'bruno', 'delphine', 'yvon', 'raymonde', 'eric', 'pascal', 'jeanne', 'albert',
    'beatrice', 'marianne', 'florent', 'thierry', 'corinne', 'robert', 'ulysse', 'paul',
    'georges', 'martine', 'maurice', 'mauricette', 'lise', 'remy', 'jackie', 'jacque'
  ];

  /* French currency / amount vocabulary. */
  SEC.CURRENCY_TOKENS = ['€', 'euro', 'euros', 'eur', '$', '£', 'fcfa', 'franc', 'francs'];
  SEC.PERIOD_TERMS_ANNUAL = ['annuel', 'annuelle', 'annuels', 'annuelles', 'annuellement', 'par an'];
  SEC.PERIOD_TERMS_MONTHLY = ['mensuel', 'mensuelle', 'mensuels', 'mensuelles', 'mensualite', 'par mois', 'mois'];
})(typeof self !== 'undefined' ? self : globalThis);

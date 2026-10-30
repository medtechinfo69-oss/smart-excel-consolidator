/* ============================================================
 * Smart Excel - Local intelligent mapping agent
 * Rule-based "virtual Excel analyst":
 *   1. semantic header recognition (synonyms, abbreviations, accents, order)
 *   2. actual cell-value profiling (French formats, types, consistency)
 *   3. confidence scoring 0-100 with explicit reasons
 *   4. blocker rules (IDs must never become phones/postal codes/amounts)
 *   5. sheet-level conflict resolution (shared destinations, phone pairing)
 * No external AI service: deterministic, auditable local algorithms.
 * ============================================================ */
(function (root) {
  'use strict';
  var SEC = (root.SEC = root.SEC || {});
  var N = SEC.Normalizers;
  var TARGETS = SEC.TARGET_KEYS;

  /* ---------------- Header-row detection ---------------- */

  function headerMatchStrength(text) {
    var norm = N.normalizeHeader(text);
    if (!norm) return 0;
    var phrase = lookupPhrase(norm);
    if (phrase) return phrase.weight;
    var token = scoreTokens(norm);
    var best = 0;
    TARGETS.forEach(function (t) { if (token[t] > best) best = token[t]; });
    return best;
  }

  function detectHeaderRow(rows) {
    var maxScan = Math.min(rows.length, SEC.LIMITS.MAX_SCAN_HEADER_ROWS);
    var bestIdx = 0, bestScore = -1;
    for (var i = 0; i < maxScan; i++) {
      var row = rows[i] || [];
      var nonEmpty = row.filter(function (c) { return N.trimClean(c) !== ''; });
      if (!nonEmpty.length) continue;
      var score = 0;
      nonEmpty.forEach(function (cell) {
        var s = headerMatchStrength(cell);
        if (s > 0) score += s;
        if (/^\d+([.,]\d+)?$/.test(N.trimClean(cell))) score -= 15; // data-like row penalty
      });
      if (score > bestScore) { bestScore = score; bestIdx = i; }
    }
    return { index: bestScore > 0 ? bestIdx : 0, score: Math.max(bestScore, 0) };
  }

  /* ---------------- Header dictionaries ---------------- */

  function lookupPhrase(normHeader) {
    var variants = [normHeader, N.singularize(normHeader)];
    var weights = [95, 88, 80]; // explicit descending order (numeric keys enumerate ascending!)
    for (var t = 0; t < TARGETS.length; t++) {
      var target = TARGETS[t];
      var tiers = SEC.PHRASE_DICT[target];
      /* Prefer the exact header at its real tier before any singular variant. */
      for (var vi = 0; vi < variants.length; vi++) {
        for (var wi = 0; wi < weights.length; wi++) {
          var weight = weights[wi];
          if (tiers[weight].indexOf(variants[vi]) !== -1) {
            return { target: target, weight: weight };
          }
        }
      }
    }
    return null;
  }

  function scoreTokens(normHeader) {
    var tokens = normHeader.split(' ').filter(Boolean);
    var sums = {};
    var matched = 0;
    tokens.forEach(function (tok) {
      var found = false;
      TARGETS.forEach(function (target) {
        var w = (SEC.TOKEN_WEIGHTS[target] || {})[tok];
        if (w) { sums[target] = (sums[target] || 0) + w; found = true; }
      });
      if (found) matched++;
    });
    var out = {};
    if (!matched) return out;
    TARGETS.forEach(function (t) {
      if (sums[t]) out[t] = Math.min(82, Math.round((sums[t] / matched) * 100));
    });
    return out;
  }

  /* Small, explicit extension vocabulary for column names the main
   * synonyms miss (English / insurance legacy files). It only COMPLETES the
   * dictionary: an unknown header still falls back to values + review. */
  var EXTENDED_HEADERS = {
    insurer: 'mutuelle', insure: 'mutuelle', insurance: 'mutuelle',
    'health insurance': 'mutuelle', 'complementary health': 'mutuelle',
    'complementary insurance': 'mutuelle', 'health fund': 'mutuelle',
    'mutual fund': 'mutuelle', 'health insurer': 'mutuelle',
    town: 'ville', city: 'ville', locality: 'ville', county: 'ville',
    surname: 'nom', 'family name': 'nom',
    forename: 'prenom', 'first names': 'prenom', 'firstname': 'prenom',
    mobile: 'tel1', cell: 'tel1', 'cellphone': 'tel1', 'mobile phone': 'tel1',
    phone: 'tel1', telephone: 'tel1', tel: 'tel1', 'phone number': 'tel1',
    'home phone': 'tel2', 'office phone': 'tel2', 'work phone': 'tel2',
    landline: 'tel2', 'fixed line': 'tel2', secondphone: 'tel2',
    'zip': 'cp', 'zipcode': 'cp', 'zip code': 'cp', 'postal': 'cp',
    'post code': 'cp', postcode: 'cp', 'postal code': 'cp',
    amount: 'cotisation', contribution: 'cotisation', premium: 'cotisation',
    fee: 'cotisation', price: 'cotisation', 'subscription': 'cotisation',
    'monthly fee': 'cotisation', 'annual fee': 'cotisation',
    street: 'adresse', address: 'adresse', 'street address': 'adresse',
    mailing: 'adresse', 'postal address': 'adresse',
    comments: 'commentaire', comment: 'commentaire', note: 'commentaire',
    notes: 'commentaire', remark: 'commentaire', remarks: 'commentaire',
    observation: 'commentaire', observations: 'commentaire',
    details: 'commentaire', info: 'commentaire',
    'numero siret': null, siret: null, 'numero client': null,
    'client number': null, 'customer number': null, 'dossier': null,
    'reference': null, 'ref': null, 'contrat': null, 'contract': null
  };

  /* Guard for columns whose header matched a destination only through a WEAK
   * token, but whose values are clearly out of scope (birth years, ages, a
   * running identifier): "Année de naissance" used to become « Code postal »
   * (2 tokens scored 74) and exported 1950, 1948… as postal codes. A strong
   * phrase or a confidently-valued column is never touched. */
  function outOfScopeByValues(profile, headerEvidence, headerPhrase) {
    if (headerPhrase && headerPhrase.weight >= 88) return false;
    if (!profile || !profile.samples || !profile.samples.length) return false;
    if (headerEvidence >= 80) return false;
    var samples = profile.samples.filter(function (v) { return N.trimClean(v) !== ''; });
    if (!samples.length || samples.length > 6) return false;
    var years = samples.filter(function (v) {
      var s = N.trimClean(v);
      if (!/^\d{4}$/.test(s)) return false;
      var n = parseInt(s, 10);
      return n >= 1900 && n <= 2035;
    }).length;
    if (years / samples.length >= 0.6) return 'annees';
    var ages = samples.filter(function (v) { return N.looksLikeAge(v) && !N.looksLikeAmount(v); }).length;
    if (ages / samples.length >= 0.6) return 'ages';
    var sequence = N.looksLikeSequenceSamples ? N.looksLikeSequenceSamples(samples) : false;
    if (sequence && profile.ratios && profile.ratios.numeric >= 0.9) return 'identifiant';
    return false;
  }

  /* Short texts that read like a call log / a free remark ("Appelé",
   * "Ne répond pas", "Rappel mardi", "OK") rather than a town or a name. */
  var COMMENT_WORDS = /\b(rappel|rappeler|repondu|repond|appel[eé]?|appele|injoignable|absent|relance|contact[eé]?|ok|ras|refus[eé]?|annul[eé]?|sans suite|a rappeler|ne repond pas|messagerie|rdv)\b/i;
  function looksLikeCommentValues(values) {
    var list = (values || []).map(N.trimClean).filter(Boolean);
    if (!list.length) return false;
    var hits = list.filter(function (v) {
      return COMMENT_WORDS.test(N.stripAccents(v)) || v.split(/\s+/).length >= 3;
    }).length;
    return hits / list.length >= 0.6;
  }

  function hasBlockerTerm(normHeader) {
    return SEC.ID_BLOCKER_TERMS.some(function (term) {
      var nt = N.normalizeHeader(term);
      return normHeader === nt || normHeader.indexOf(' ' + nt) !== -1 ||
        normHeader.indexOf(nt + ' ') !== -1 || normHeader.indexOf(nt) !== -1 && nt.length >= 4;
    });
  }

  function isFullNameHeader(normHeader) {
    return SEC.FULL_NAME_HEADERS.some(function (h) {
      return normHeader === h || N.singularize(normHeader) === h;
    });
  }

  /* ---------------- Value profiling ---------------- */

  function profileColumn(values) {
    var nonEmpty = [];
    var seen = {};
    var totalLen = 0;
    var phone = 0, email = 0, postal = 0, amount = 0, address = 0,
      city = 0, person = 0, mutuelle = 0, longText = 0, numeric = 0;

    var nonEmptyCount = 0;
    for (var i = 0; i < values.length; i++) {
      var v = N.trimClean(values[i]);
      if (!v) continue;
      nonEmptyCount++;
      if (nonEmpty.length < SEC.LIMITS.MAX_PROFILE_SAMPLE) nonEmpty.push(v);
      totalLen += v.length;
      if (N.looksLikePhone(v)) phone++;
      if (N.looksLikeEmail(v)) email++;
      if (N.looksLikePostal(v)) postal++;
      if (N.looksLikeAmount(v)) amount++;
      if (N.looksLikeAddress(v)) address++;
      if (N.looksLikeCity(v)) city++;
      if (N.looksLikePersonName(v)) person++;
      if (N.looksLikeMutuelle(v)) mutuelle++;
      if (v.length >= 35 || /[.;]/.test(v)) longText++;
      if (v.length <= 40 && /^-?\d+([.,]\d+)?$/.test(v.replace(/\s/g, ''))) numeric++;
    }
    /* The non-empty count comes from the single pass above: a second full
     * pass re-normalised every value (twice the work on long notes). */
    var denom = Math.max(nonEmptyCount, 1);
    var samples = [];
    nonEmpty.forEach(function (v) {
      var key = v.slice(0, 80);
      if (!seen[key] && samples.length < 5) { seen[key] = true; samples.push(v.slice(0, 120)); }
    });

    return {
      total: values.length,
      nonEmpty: nonEmptyCount,
      samples: samples,
      avgLength: nonEmptyCount ? Math.round(totalLen / denom) : 0,
      ratios: {
        phone: phone / denom,
        email: email / denom,
        postal: postal / denom,
        amount: amount / denom,
        address: address / denom,
        city: city / denom,
        person: person / denom,
        mutuelle: mutuelle / denom,
        longText: longText / denom,
        numeric: numeric / denom
      }
    };
  }

  /* Value-based score per destination (0-100), capped so value-only
   * guesses can never reach the auto-accept band on their own. */
  function valueScores(profile, headerNorm) {
    var r = profile.ratios;
    var out = {};
    if (profile.nonEmpty === 0) return out;
    if (r.phone > 0) { out.tel1 = Math.round(45 + 50 * r.phone); out.tel2 = Math.round(45 + 50 * r.phone); }
    if (r.email > 0) out.email = Math.round(45 + 50 * r.email);
    if (r.postal > 0) out.cp = Math.round(45 + 50 * r.postal);
    if (r.amount > 0) out.cotisation = Math.round(42 + 48 * r.amount);
    /* An address is proven by its STRUCTURE (number/type of street, hamlet,
     * lotissement): a column made of them is an Adresse column even when no
     * header is recognised. It used to score 68 and fell into Commentaire. */
    if (r.address > 0) out.adresse = Math.round(45 + 50 * r.address);
    if (r.city > 0) out.ville = Math.round(40 + 48 * r.city);
    if (r.person > 0) {
      out.nom = Math.round(38 + 42 * r.person);
      out.prenom = Math.round(35 + 40 * r.person);
    }
    if (r.mutuelle > 0) out.mutuelle = Math.round(45 + 50 * r.mutuelle);
    if (r.longText > 0.4) out.commentaire = Math.round(40 + 45 * r.longText);
    /* Common-French-first-name signal strengthens Prénom. */
    var prenomHits = N.countCommonPrenom(profile.samples);
    if (prenomHits) out.prenom = Math.max(out.prenom || 0, Math.min(80, 55 + prenomHits * 8));
    return out;
  }

  /* ---------------- Per-target validation ---------------- */

  function validateForTarget(target, values) {
    var nonEmpty = values.map(N.trimClean).filter(Boolean);
    if (!nonEmpty.length) return { status: 'empty', validRatio: 1, reason: 'aucune valeur' };
    var valid = 0, international = 0, reasons = [];
    nonEmpty.forEach(function (v) {
      var res;
      switch (target) {
        case 'tel1':
        case 'tel2':
          res = N.analyzePhone(v);
          if (res.valid === true) valid++;
          else if (res.valid === 'international') { valid++; international++; }
          break;
        case 'cp': if (N.analyzePostal(v).valid) valid++; break;
        case 'email': if (N.analyzeEmail(v).valid) valid++; break;
        case 'cotisation': if (N.analyzeAmount(v).valid) valid++; break;
        case 'nom':
        case 'prenom':
          if (!N.isInvalidPersonName(v)) valid++;
          break;
        default: valid++;
      }
    });
    var ratio = valid / nonEmpty.length;
    var minRatio = (target === 'email' || target === 'tel1' || target === 'tel2') ? 0.7 : 0.8;
    if (ratio >= minRatio) {
      return { status: international ? 'info' : 'ok', validRatio: ratio, reason: international ? 'contient des numeros internationaux' : '' };
    }
    return {
      status: 'warn',
      validRatio: ratio,
      reason: Math.round((1 - ratio) * 100) + '% de valeurs non conformes au format ' + SEC.TARGET_LABELS[target]
    };
  }

  /* ---------------- Single-column analysis ---------------- */

  function analyzeColumn(header, values) {
    var headerNorm = N.normalizeHeader(header);
    var profile = profileColumn(values);
    var reasons = [];
    var blocked = hasBlockerTerm(headerNorm);
    var fullName = isFullNameHeader(headerNorm);
    /* True when the header text itself was recognised (dictionary or tokens):
     * the AI assistant only needs to look at the others. */
    var headerMatched = false;

        /* Header-based scores */
    var headerScores = {};
    var phrase = lookupPhrase(headerNorm);
    if (phrase) {
      headerScores[phrase.target] = phrase.weight;
      reasons.push('en-tete reconnu : synonyme de « ' + SEC.TARGET_LABELS[phrase.target] + '»');
    }
    var token = scoreTokens(headerNorm);
    TARGETS.forEach(function (t) {
      if (token[t]) { headerScores[t] = Math.max(headerScores[t] || 0, token[t]); headerMatched = true; }
    });
    /* Extension vocabulary: an exact (or short, contained) match is decisive. */
    var shortHeaderMatch = false;
    var outOfScope = null;
    var headerTokens = headerNorm.split(' ').filter(Boolean);
    var ext = EXTENDED_HEADERS[headerNorm] !== undefined
      ? EXTENDED_HEADERS[headerNorm]
      : (function () {
          var words = headerTokens;
          if (!words.length || words.length > 3) return null;
          for (var wi = 0; wi < words.length; wi++) {
            if (Object.prototype.hasOwnProperty.call(EXTENDED_HEADERS, words[wi]) &&
                EXTENDED_HEADERS[words[wi]])
              return EXTENDED_HEADERS[words[wi]];
          }
          return null;
        })();
    if (ext) {
      headerScores[ext] = Math.max(headerScores[ext] || 0, 88);
      headerMatched = true;
      shortHeaderMatch = true;
      reasons.push('en-tete reconnu (vocabulaire etendu) : « ' + SEC.TARGET_LABELS[ext] + '»');
    }
    if (phrase || blocked || fullName) headerMatched = true;
    var headerEvidenceForScope = 0;
    Object.keys(headerScores).forEach(function (k) { if (headerScores[k] > headerEvidenceForScope) headerEvidenceForScope = headerScores[k]; });
    /* A sequence of ids is NEVER a value for one of the 11 columns: detect it
     * even when the header ("Dossier", "Ref") is not a known blocker term. */
    var oos = outOfScopeByValues(profile, headerEvidenceForScope, phrase) ||
      (N.looksLikeSequenceSamples && !(phrase && phrase.weight >= 88) &&
        headerEvidenceForScope < 70 &&
        N.looksLikeSequenceSamples((profile.samples || []).filter(function (v) { return N.trimClean(v) !== ''; }))
        ? 'identifiant' : null);
    if (oos) {
      reasons.push(oos === 'annees' ? 'valeurs d annees (naissance/adhésion) : aucune colonne officielle ne la concerne'
        : oos === 'ages' ? 'valeurs d ages : aucune colonne officielle ne la concerne'
        : 'valeurs d identifiant séquentiel : aucune colonne officielle ne la concerne');
    }

    var vScore = valueScores(profile, headerNorm);
    /* Formats that ARE the destination, certain from the values alone: an
     * e-mail, a phone, a postal code, a confirmed amount, a real address, a
     * column made of insurer names. Used below to let such a column be
     * decided without a header (it used to be capped at 72, under review). */
    var VALUES_PROVE = { email: 1, tel1: 1, tel2: 1, cp: 1, cotisation: 1, adresse: 1 };
    if (profile.ratios.mutuelle >= 0.9) VALUES_PROVE.mutuelle = 1;
    if (blocked) {
      /* Identifier-like columns: neither header words nor digit patterns
       * may promote them to phone, postal code or amount. */
      ['tel1', 'tel2', 'cp', 'cotisation'].forEach(function (t) {
        delete vScore[t];
        delete headerScores[t];
      });
      reasons.push('en-tete de type identifiant/reference : la correspondance automatique est bloquee');
    }

    var proposals = TARGETS.map(function (t) {
      var h = headerScores[t] || 0;
      var v = vScore[t] || 0;
      /* A canonical header term ("Courriel", "Email", "Ville"…) is certain
       * by itself: a few odd values must not lower it into the review band. */
      var strongPhrase = phrase && phrase.target === t && phrase.weight >= 95;
      var score;
      if (h && v) {
        score = Math.round(0.55 * h + 0.45 * v);
          /* A canonical/strong dictionary match stays in the high band when
         * value evidence exists (contradictions are capped afterwards). */
        if (strongPhrase) score = Math.max(score, 90);
      } else if (h) {
        /* Header-only: the header IS the strongest signal ("Courriel" with 3
         * odd values is still an e-mail column). Dictionary phrases keep a
         * full-weight score, token guesses stay under review. */
        score = strongPhrase ? Math.max(h, 90) : (phrase && phrase.weight === 80 && !v ? Math.max(h, 94) : h);
      } else if (v) {
        /* VALUE-PROVEN destinations (an e-mail, a phone, a postal code, a real
         * address) reach a decided score from their values alone. A SHAPED
         * guess (a capitalised word that might be a city or a name) stays at
         * 72 so it is always confirmed by the AI / kept in review. */
        var proven = (VALUES_PROVE[t] || (t === 'adresse' && profile.ratios.address >= 0.6));
        score = proven ? Math.min(92, Math.round(v * 0.95 + 12)) : Math.min(72, Math.round(v * 0.8));
      } else {
        score = 0;
      }
      return { target: t, score: score };
    }).filter(function (p) { return p.score > 0; })
      .sort(function (a, b) { return b.score - a.score; });

    if (profile.nonEmpty) {
      if (vScore.tel1 || vScore.email || vScore.cp || vScore.amount) {
        reasons.push('analyse des valeurs reelles : formats de donnees detectes (telephone/email/code postal/montant)');
      }
      if (vScore.adresse) reasons.push('valeurs typiques d adresse postale (rue, numero, complement)');
      if (vScore.mutuelle) reasons.push('valeurs correspondant a des organismes de complementaire sante');
    }

    var best = proposals.length ? proposals[0].target : null;
    var bestScore = proposals.length ? proposals[0].score : 0;
    /* Lead over the runner-up: a column claimed by several destinations is not
     * "certain", even if the winner's score is high. */
    var margin = proposals.length > 1 ? proposals[0].score - proposals[1].score : 100;
    var validation = best ? validateForTarget(best, values) : null;

    /* Think about the WHOLE column, not one value. A value-only guess (no
     * header evidence, e.g. "Titre client" (Dupont, Martin) read as Ville) is
     * kept ONLY when the destination accepts at least 75 % of the real values;
     * otherwise there is no automatic destination and the column is offered to
     * the optional AI and listed in the review table.
     * A header-based match is never dropped here. */
    var headerEvidence = (best && headerScores[best]) || 0;
    /* Out-of-scope column detected from its values: it must never inherit a
     * phone/postal/amount destination from a weak header token. */
    if (oos && ['cp', 'tel1', 'tel2', 'cotisation', 'ville', 'adresse', 'nom', 'prenom'].indexOf(best) !== -1) {
      if (!(phrase && phrase.weight >= 88)) {
        best = null;
        bestScore = 0;
        validation = null;
        headerMatched = false;
        shortHeaderMatch = false;
        outOfScope = oos;
        reasons.push('destination annulee : les valeurs ne sont ni un code postal ni un telephone ni un montant');
      }
    }
    /* Only a strong value guess keeps a value-only destination: "Titre
     * client" (Dupont, Martin) read as Ville gets NO destination here and is
     * handed to the optional AI / the review table.
     * A header-based match is never dropped by this rule. */
    /* A column of real postal addresses is proven by its values: the street
     * structure ("10 AVENUE FELIX FAURE", "1010 ROUTURIER", "HAMEAU X") is as
     * reliable as an e-mail format. Without this rule the column scored 68 and
     * every address ended up in « Commentaire ». */
    if (profile.ratios.address >= 0.6) VALUES_PROVE.adresse = 1;
    /* An insurer name IS the canonical format of « Mutuelle » (MGEN, AXA…):
     * a 100 %-insurer column is proven by its values, header or not. */
    /* A pure guess from the SHAPE of the values must never be presented as a
     * confident destination: "Dupont" in a "Titre client" column was guessed
     * as Ville (all-capital letters ARE a valid city shape). Those proposals
     * stay, but capped into the review band (<= 84) so the column is ALWAYS
     * listed for confirmation and offered to the optional AI. */
    var GUESS_FROM_VALUES = { ville: 1, nom: 1, prenom: 1, adresse: 1, mutuelle: 1, commentaire: 1 };
    /* A PROVEN format is not a guess. "adresse" used to stay inside this map,
     * so a column of real street addresses without a recognised header was
     * later demoted to « Commentaire » — that is exactly the "adresse in
     * commentaire" symptom. E-mail / phone / postal code / insurance name /
     * street address are proved by their own format. */
    var valueOnlyGuess = !!(best && GUESS_FROM_VALUES[best] && !VALUES_PROVE[best] && headerEvidence < 60);
    if (best && !headerEvidence && !VALUES_PROVE[best] && bestScore < 65) {
      var corroborated = validation && validation.status !== 'warn' &&
        validation.status !== 'empty' && validation.validRatio >= 0.75;
      if (!corroborated) {
        reasons.push('en-tete non reconnu et valeurs non conformes : destination non choisie automatiquement');
        best = null;
        bestScore = 0;
        validation = null;
        valueOnlyGuess = false;
      }
    }
    /* "Retour", "Observations", "Remarques": a linear-looking column of calls,
     * answers or short sentences is a COMMENT column, not a Ville. Those short
     * texts used to score as a city (every capitalised word does) and were
     * exported under « Ville », which also cost the real comment. */
    if (!headerEvidence && valueOnlyGuess && (best === 'ville' || best === 'nom' || best === 'prenom') &&
        looksLikeCommentValues(values)) {
      best = 'commentaire';
      bestScore = Math.min(bestScore, 80);
      validation = { status: 'ok', validRatio: 1, reason: '' };
      reasons.push('valeurs de type commentaire/appel : placees dans « Commentaire »');
    }
    /* A value-only guess is not certain, but a LARGE part of the values in
     * the guessed format is real evidence (Gemini only sees 6 samples): it is
     * capped clearly below an automatic decision and still lands somewhere
     * sensible instead of losing the data. */
    if (valueOnlyGuess && best !== 'commentaire') {
      var shapeRatio = (validation && validation.validRatio) || 0;
      bestScore = Math.min(bestScore, shapeRatio >= 0.75 ? 80 : 74);
    }
    /* A column proven by its format gets a decided score so it is never
     * pushed into « Commentaire » by the "not certain enough" rule. */
    if (best && !headerEvidence && VALUES_PROVE[best] && bestScore < 78) {
      var proofRatio = (validation && validation.validRatio) || (profile.ratios[best === 'adresse' ? 'address' : best] || 0);
      if (proofRatio >= 0.6) bestScore = Math.max(bestScore, Math.min(88, Math.round(60 + proofRatio * 28)));
    }
    if (validation && validation.status === 'warn') {
      bestScore = Math.min(bestScore, 85); // contradictory evidence caps confidence
      reasons.push('attention : ' + validation.reason);
      /* If a column mapped to « Code postal » fails postal validation but its
       * values are clearly phone numbers or e-mails, remap to the right
       * destination instead of silently leaving bad data under « Code postal ». */
      /* Values are an insurer name (MGEN, AXA…) while the header says
       * Cotisation/Prime: the value decides, it lands in « Mutuelle ». */
      if (best === 'cotisation') {
        var insurerHits = 0, usable = 0;
        values.forEach(function (v) {
          var s = N.trimClean(v);
          if (!s) return;
          usable++;
          if (N.looksLikeInsurer(s) && !N.looksLikeAmount(s)) insurerHits++;
        });
        if (usable && insurerHits / usable >= 0.7) {
          best = 'mutuelle';
          bestScore = Math.min(bestScore, 85);
          headerMatched = false;   // inferred from values: AI may confirm it
          reasons.push('les valeurs sont des organismes de mutuelle : placé dans « Mutuelle »');
          validation = validateForTarget(best, values);
        }
      }
      if (best === 'cp') {
        var altTarget = null;
        /* Phone first, then e-mail — only remap when values validate cleanly. */
        ['tel1', 'tel2', 'email'].some(function (t) {
          var vt = validateForTarget(t, values);
          if (vt && vt.status === 'ok' && vt.validRatio >= 0.7) {
            altTarget = t;
            return true;
          }
          return false;
        });
        if (altTarget) {
          var originalTarget = best;
          var altV = vScore[altTarget] || 0;
          var altH = headerScores[altTarget] || 0;
          var combined = altH && altV
            ? Math.round(0.55 * altH + 0.45 * altV)
            : Math.min(72, Math.round((altV || 0) * 0.8));
          best = altTarget;
          bestScore = Math.max(combined, 70);
          reasons.push('correction : l\'en-tete suggerait « ' +
            SEC.TARGET_LABELS[originalTarget] + ' » mais les valeurs reelles sont des « ' +
            SEC.TARGET_LABELS[altTarget] + ' » — deplacement automatique');
        }
      }
    }
    /* A value-only guess (no header evidence) must never look certain: the
     * optional AI may still confirm or correct it. */

    if (fullName) {
      /* A full-name column is kept WHOLE in « Nom » (never split into
       * Prénom). A very common name header (Client, Adhérent, Nom complet…)
       * is applied directly; unusual wordings still surface a review item so
       * the destination can be confirmed (never silent). */
      var common = ['client', 'adherent', 'assure', 'nom complet', 'nom prenom', 'nom et prenom', 'full name', 'fullname', 'personne', 'contact'];
      var isCommon = common.some(function (h) {
        return headerNorm === h || headerNorm.indexOf(h + ' ') === 0 || headerNorm.indexOf(' ' + h) !== -1;
      });
      best = 'nom';
      bestScore = isCommon ? 91 : 82;
      reasons.push('nom complet : valeur conservee entiere dans « Nom »');
    }



    return {
      header: header,
      headerNorm: headerNorm,
      profile: profile,
      blocked: blocked,
      fullName: fullName,
      proposals: proposals,
      best: best,
      score: bestScore,
      reasons: reasons,
      learned: false,
      excluded: false,
      outOfScope: outOfScope || null,
      headerMatched: headerMatched,
      shortHeaderMatch: shortHeaderMatch,
      margin: margin,
      validation: validation,
      /* No header evidence: the destination is only a guess from the shape of
       * the values. Callers keep such a column in « Commentaire ». */
      valueOnlyGuess: valueOnlyGuess
    };
  }



  /* ---------------- Sheet-level analysis + conflicts ---------------- */

  function phoneHint(headerNorm) {
    if (/(^|\s)(2|secondaire|fixe|domicile|bureau|pro)($|\s)/.test(headerNorm)) return 'secondary';
    if (/(^|\s)(1|principal|principale|primaire|personnel|perso|mobile|portable)($|\s)/.test(headerNorm)) return 'primary';
    return 'neutral';
  }

  /* ---------- French commune recognition (official database) ---------- */

  var CITY_DB = null;   // normalized official name -> true
  function cityIndex() {
    if (CITY_DB || !SEC.Communes) return CITY_DB;
    CITY_DB = Object.create(null);
    var raw = typeof SEC.COMMUNE_DB === 'string' ? SEC.COMMUNE_DB : '';
    raw.split('\n').forEach(function (line) {
      var name = line.split('\t')[0];
      var k = SEC.Communes.key(name);
      if (k) CITY_DB[k] = true;
    });
    return CITY_DB;
  }

  /* True when the WHOLE sampled column reads as official commune names. */
  function looksLikeCommuneColumn(profile) {
    var samples = (profile && profile.samples) || [];
    if (samples.length < 3) return false;
    return samples.filter(function (v) { return isKnownCommuneName(v); }).length / samples.length >= 0.8;
  }

  function isKnownCommuneName(v) {
    var idx = cityIndex();
    if (!idx) return false;
    var found = SEC.Communes.lookupName(v);
    return !!(found && found.entry);
  }

  /* Share of the real values that are official commune names. */
  function cityRatio(profile) {
    var samples = profile.samples || [];
    if (!samples.length) return 0;
    var hits = samples.filter(function (v) {
      return N.trimClean(v).length <= 60 && isKnownCommuneName(v);
    }).length;
    return hits / samples.length;
  }

  /* Share of a column's sample values that are ages ("70+75", "63 ans",
   * "1950-1948"…). Ages are NOT an official column: they go to Commentaire. */
  function ageRatioOf(col) {
    return N.ageRatio ? N.ageRatio((col.profile && col.profile.samples) || []) : 0;
  }

  function resolveConflicts(columns) {
    /* -1) AGE columns ("70+75", "84+85", "63 ans", "1950-1948"…): never a
     * cotisation nor a postal code. If the header itself says age, the whole
     * column goes to Commentaire with certainty; otherwise a column whose
     * values are mostly ages (>= 60%) is moved to Commentaire too. */
    columns.forEach(function (c) {
      if (c.excluded || c.blocked || c.fullName) return;
      var headerAge = /^(age|ages|age du client|age adherent|annee de naissance|annees? de naissance)/.test(c.headerNorm);
      var ratio = ageRatioOf(c);
      if (!headerAge && ratio < 0.6) return;
      /* Ages are OUT OF SCOPE (no official column): the header or the values
       * prove it, so the column is EXCLUDED deliberately instead of flooding
       * « Commentaire » with "63 | 71 | 64" and pushing every row to review. */
      if (!c.best || c.best === 'commentaire') {
        c.best = null;
        c.excluded = true;
        c.profileExcluded = true;
        c.score = 0;
        c.reasons.push((headerAge ? 'colonne d ages (« ' + (c.header || 'sans entete') + ' »)'
          : 'valeurs d ages (' + Math.round(ratio * 100) + ' %)') +
          ' : aucune des 11 colonnes officielles ne la concerne — exclue (signalee dans le rapport)');
        return;
      }
      /* It carries another real value ("Titre", "Prime"…): keep that value. */
      c.reasons.push('colonne d ages : les valeurs ne sont pas des montants ni des codes postaux');
    });

    /* 0) Official commune names are a strong Ville signal: "Commune",
     * "Localité", "City"… already map through the dictionary; an unknown
     * header whose values are real communes is a Ville column too. */
    var villeTaken = columns.some(function (c) { return c.best === 'ville'; });
    if (!villeTaken) {
      var cityCol = null;
      columns.forEach(function (c) {
        if (c.excluded || c.blocked || c.best || c.fullName) return;
        if (cityRatio(c.profile) < 0.6) return;
        if (!cityCol || c.profile.ratios.city > cityCol.profile.ratios.city) cityCol = c;
      });
      if (cityCol) {
        cityCol.best = 'ville';
        cityCol.score = Math.max(cityCol.score, 74);
        cityCol.rescuedByValues = true;
        cityCol.headerMatched = false;   // inferred: the AI may still confirm it
        cityCol.reasons.push('en-tête non reconnu mais valeurs de communes françaises officielles détectées : placé dans « Ville »');
      }
    }

    /* 1) Phone pairing across all phone-capable columns. */
    var phoneCols = columns.filter(function (c) {
      if (c.excluded || c.blocked) return false;
      return c.best === 'tel1' || c.best === 'tel2' || c.profile.ratios.phone > 0.6;
    }).sort(function (a, b) { return b.score - a.score; });

    phoneCols.forEach(function (c) {
      c.phoneHint = phoneHint(c.headerNorm);
    });

    if (phoneCols.length === 1) {
      phoneCols[0].best = 'tel1';
      phoneCols[0].score = Math.max(phoneCols[0].score, 75);
      if (phoneCols[0].profile.ratios.phone < 0.7) phoneCols[0].score = Math.min(phoneCols[0].score, 85);
    } else if (phoneCols.length >= 2) {
      var primaries = phoneCols.filter(function (c) { return c.phoneHint === 'primary'; });
      var secondaries = phoneCols.filter(function (c) { return c.phoneHint === 'secondary'; });
      var ordered = phoneCols.slice().sort(function (a, b) { return a._index - b._index; });
      var p = primaries[0] || ordered[0];
      var s = secondaries[0] || ordered.filter(function (c) { return c !== p; })[0];
      if (p) { p.best = 'tel1'; p.score = Math.max(p.score, 80); }
      if (s) {
        s.best = 'tel2'; s.score = Math.max(s.score, 78);
        if (s.phoneHint === 'neutral') {
          s.score = Math.min(s.score, 85);
          s.reasons.push('deux numeros sans distinction primaire/secondaire : attribution par ordre, a verifier');
        }
      }
      phoneCols.filter(function (c) { return c !== p && c !== s; }).forEach(function (c) {
        c.best = null; c.score = Math.min(c.score, 60);
        c.reasons.push('plus de deux colonnes telephoniques : colonne supplementaire a confirmer ou exclure');
      });
    }

    /* 2) Other duplicate destinations: strongest wins, others flagged. */
    TARGETS.filter(function (t) { return t !== 'tel1' && t !== 'tel2'; }).forEach(function (target) {
      var holders = columns.filter(function (c) { return !c.excluded && c.best === target; })
        .sort(function (a, b) { return b.score - a.score; });
      if (holders.length > 1) {
        holders[0].reasons.push('plusieurs colonnes candidates pour « ' + SEC.TARGET_LABELS[target] + '» : meilleure correspondance retenue');
        holders.slice(1).forEach(function (c) {
          c.reasons.push('conflit : une autre colonne correspond mieux a « ' + SEC.TARGET_LABELS[target] + '»');
          if (target === 'commentaire') {
            /* Comments are safely mergeable instead of being dropped. */
            c.mergeInto = target;
            c.reasons.push('les commentaires seront fusionnes avec separateur clair');
          } else {
            c.score = Math.min(c.score, 68);
            c.best = null;
          }
        });
      }
    });

    /* 3) Merge-marked comment columns keep their destination for extraction. */
    columns.forEach(function (c) { if (c.mergeInto) c.best = c.mergeInto; });

    /* 4) RESCUE: a column with an unknown header ("fix", "gsm", "TEL 2"…) whose
     * values are clearly phone numbers must NOT be silently lost. If a phone
     * destination is still free, use it — this is how the second phone number
     * of real files ends up in « Téléphone 2 » instead of vanishing. */
    var used = {};
    columns.forEach(function (c) { if (c.best && !c.excluded) used[c.best] = true; });
    columns.forEach(function (c) {
      if (c.excluded || c.best || c.blocked) return;
      var r = c.profile.ratios;
      if (r.phone < 0.5) return;
      if (isKnownCommuneName((c.profile.samples || [])[0] || '')) return;
      var slot = !used.tel1 ? 'tel1' : (!used.tel2 ? 'tel2' : null);
      if (!slot) return;
      c.best = slot;
      c.score = 76;
      c.rescuedByValues = true;
      used[slot] = true;
      c.reasons.push('en-tête non reconnu mais valeurs de téléphone détectées : placé dans « ' +
        SEC.TARGET_LABELS[slot] + '»');
    });

    /* 5) RESCUE: the same for e-mail, postal code and amount columns whose
     * header is unknown — the values decide, never the header. */
    [['email', 'email', 0.6], ['cp', 'postal', 0.5], ['cotisation', 'amount', 0.6]].forEach(function (pair) {
      var target = pair[0], ratioKey = pair[1], min = pair[2];
      if (used[target]) return;
      var best = null;
      columns.forEach(function (c) {
        if (c.excluded || c.best || c.blocked || c.fullName) return;
        if (ratioKey === 'postal' && isKnownCommuneName((c.profile.samples || [])[0] || '')) return;
        if (c.profile.ratios[ratioKey] < min) return;
        if (!best || c.profile.ratios[ratioKey] > best.profile.ratios[ratioKey]) best = c;
      });
      if (!best) return;
      best.best = target;
      best.score = 72;
      best.rescuedByValues = true;
      used[target] = true;
      best.reasons.push('en-tête non reconnu mais valeurs typiques détectées : placé dans « ' +
        SEC.TARGET_LABELS[target] + '»');
    });

    /* 6) Name columns rescued by their values ("Titre client" is not a
     * city…): a column of real person names with an unknown header that is
     * still undecided or wrongly placed as Ville goes to the name slots. */
    var nomHolder = columns.some(function (c) { return c.best === 'nom'; });
    columns.forEach(function (c) {
      if (c.excluded || c.blocked || c.fullName || c.headerMatched) return;
      if (c.best === 'nom' || c.best === 'prenom') return;
      if (c.profile.ratios.person < 0.6) return;
      /* Real commune names are a city column, never names. */
      if (typeof isKnownCommuneName === 'function' &&
          (isKnownCommuneName((c.profile.samples || [])[0] || ''))) return;
      /* A doubtful header ("Titre client") guessed as a city, or a column left
       * undecided, whose REAL values are person names: « Nom » (to confirm). */
      var doubtful = c.best === 'ville' || !c.best || !c.headerMatched;
      if (doubtful && !nomHolder) {
        c.best = 'nom';
        c.score = 72;
        c.rescuedByValues = true;
        c.headerMatched = false;
        c.reasons.push('valeurs de noms de personnes : colonne placee dans « Nom » (a confirmer)');
        nomHolder = true;
      }
    });

    return columns;
  }

  function analyzeSheet(rows, headerRowIndex) {
    var hdr = detectHeaderRow(rows);
    var idx = (typeof headerRowIndex === 'number') ? headerRowIndex : hdr.index;
    var headerRow = rows[idx] || [];
    var dataRows = rows.slice(idx + 1);

    var columns = headerRow.map(function (header, colIndex) {
      var colValues = dataRows.map(function (r) { return r[colIndex]; });
      var col = analyzeColumn(header, colValues);
      col._index = colIndex;
      col.headerRowIndex = idx;
      return col;
    }).filter(function (c) {
      /* Fully empty columns (no header text and no values) are ignored quietly. */
      return N.trimClean(c.header) !== '' || c.profile.nonEmpty > 0;
    });

    resolveConflicts(columns);
    return { headerRowIndex: idx, detectedHeaderRow: hdr.index, columns: columns, dataRows: dataRows };
  }

  SEC.Agent = {
    detectHeaderRow: detectHeaderRow,
    analyzeColumn: analyzeColumn,
    analyzeSheet: analyzeSheet,
    profileColumn: profileColumn,
    validateForTarget: validateForTarget,
    lookupPhrase: lookupPhrase,
    scoreTokens: scoreTokens,
    hasBlockerTerm: hasBlockerTerm,
    isFullNameHeader: isFullNameHeader,
    looksLikeCommuneColumn: looksLikeCommuneColumn,
    resolveConflicts: resolveConflicts
  };
})(typeof self !== 'undefined' ? self : globalThis);

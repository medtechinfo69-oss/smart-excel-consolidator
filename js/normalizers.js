/* ============================================================
 * Smart Excel - Normalization & French data rules
 * Telephone, postal code, email, amount, name, address handling.
 * Every function is pure so it can be unit-tested in Node.
 * ============================================================ */
(function (root) {
  'use strict';
  var SEC = (root.SEC = root.SEC || {});
  var D = SEC.DICTS || {};

  /* ---------- Text helpers ---------- */

  function asString(v) {
    if (v === null || v === undefined) return '';
    return String(v);
  }

  /* Collapses spaces and trims. Returns the same string object when nothing
   * changes: on multi-KB notes, rebuilding a 32 000-char string for every
   * check was a top source of CPU time and garbage. */
  function trimClean(v) {
    var s = asString(v);
    if (!s) return s;
    var hasNbsp = s.indexOf('\u00A0') !== -1;
    if (!hasNbsp && s.indexOf('  ') === -1 && !/^\s|\s$/.test(s)) return s;
    if (hasNbsp) s = s.replace(/\u00A0/g, ' ');
    return s.replace(/ {2,}/g, ' ').trim();
  }

  /* Placeholder junk used by legacy exports to mean "empty":
   * ".", "....", "//", "-", "N/A", "NEANT", "?” …
   * It must be treated as an EMPTY cell, never as an invalid value. */
  var PLACEHOLDER_RE = /^(?:[.\-\/_*\s]*|\/{1,2}|n\/?a|néant|neant|x{1,3}|\?+|r\.?a\.?s\.?|aucun[e]?)$/i;
  function isPlaceholder(v) {
    var s = trimClean(v);
    if (!s) return true;
    return PLACEHOLDER_RE.test(s);
  }

  function stripAccents(s) {
    return asString(s).normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  /* Normalize a source header: lowercase, no accents, no punctuation,
   * abbreviations expanded, whitespace collapsed. */
  /* Memoised: the same few headers are normalised for EVERY row (a 100 000
   * row file re-normalised the same 11 strings 300 000 times). */
  var HEADER_CACHE = new Map();
  function normalizeHeader(header) {
    var raw = asString(header);
    var hit = HEADER_CACHE.get(raw);
    if (hit !== undefined) return hit;
    var s = stripAccents(raw).toLowerCase();
    s = s.replace(/['’´`]/g, ' ');
    s = s.replace(/[^a-z0-9]+/g, ' ');
    s = s.replace(/\s+/g, ' ').trim();
    var tokens = s.split(' ').filter(Boolean).map(function (t) {
      return SEC.ABBREVIATIONS && SEC.ABBREVIATIONS[t] ? SEC.ABBREVIATIONS[t] : t;
    });
    var out = tokens.join(' ');
    if (HEADER_CACHE.size < 4000) HEADER_CACHE.set(raw, out);
    return out;
  }

  /* Singularized variant for tolerant French/English plural matching. */
  function singularize(normHeader) {
    return normHeader.split(' ').map(function (t) {
      if (t.length > 3 && t.slice(-1) === 's' && t !== 'frais') return t.slice(0, -1);
      return t;
    }).join(' ');
  }

  function isBlank(v) {
    return trimClean(v) === '';
  }

  /* ---------- Telephone (French rules) ---------- */

  /* Returns { value, valid: true|false|'international', kind, reason }
   * value is always a string so leading zeros are preserved. */
  function analyzePhone(raw) {
    /* Same numbers repeat across files: the analysis is memoised (the result
     * is a fresh object each call, so callers can never share it by mistake). */
    var key = asString(raw);
    var hit = PHONE_CACHE.get(key);
    if (hit) return Object.assign({}, hit);
    var res = analyzePhoneUncached(raw);
    if (PHONE_CACHE.size >= 5000) PHONE_CACHE.clear();
    PHONE_CACHE.set(key, res);
    return Object.assign({}, res);
  }

  var PHONE_CACHE = new Map();

  function analyzePhoneUncached(raw) {
    var original = asString(raw).replace(/ /g, ' ').replace(/[ ]+/g, ' ').trim();
    if (!original) return { value: '', valid: null, kind: null, reason: '' };
    /* Placeholder junk (".", "//", …) means "no phone": not an error. */
    if (PLACEHOLDER_RE.test(original)) return { value: '', valid: null, kind: null, reason: '' };
    /* Trailing/leading dots and dashes are formatting noise: remove them. */
    if (/[.\-]\s*$/.test(original) && /\d/.test(original)) {
      original = original.replace(/[\s.\-]+$/, '').replace(/^[\s.\-]+/, '').trim();
      if (!original) return { value: '', valid: null, kind: null, reason: '' };
    }

    var compact = original.replace(/[ .\-()]/g, '');
    var isIntl33 = /^(\+|00)33/.test(compact);
    var isOtherIntl = /^(\+|00)\d{1,4}/.test(compact);
    var digits = original.replace(/\D/g, '');

    if (isIntl33) {
      /* Strip 0033 or 33 then rebuild the national 0-prefixed number. */
      var national = digits.replace(/^(00)?33/, '');
      var withZero = '0' + national;
      if (/^0[1-9]\d{8}$/.test(withZero)) {
        return {
          value: withZero,
          valid: true,
          kind: (withZero[1] === '6' || withZero[1] === '7') ? 'mobile' : 'fixe',
          reason: 'converti depuis le format international +33/0033'
        };
      }
      return { value: original, valid: false, kind: null, reason: 'prefixe 33 mais longueur invalide' };
    }

    /* French numbering plan: 01 Île-de-France, 02 North-West, 03 North-East
     * (03 85 = Saône-et-Loire / Le Creusot), 04 South-East, 05 South-West,
     * 06-07 mobile, 08 special, 09 VoIP. Every 0X prefix IS valid: rejecting
     * "03" threw away every landline of a whole region ("0385561532"). */
    if (/^0[1-9]\d{8}$/.test(digits) && !/^00/.test(digits)) {
      return {
        value: digits,
        valid: true,
        kind: (digits[1] === '6' || digits[1] === '7') ? 'mobile' : 'fixe',
        reason: ''
      };
    }

    /* Truncated numbers whose digits carry a COMPLETE 9-digit body after the
     * leading 3-prefix marker ("3243821836" handled below). Here: a short
     * number like "09 83 69 6" (7 digits) cannot be completed safely and is
     * reported instead of guessed. */

    /* Legacy exports store French numbers WITHOUT the leading 0:
     *   493806327 -> 0493806327 (9 digits, first digit 1-9).
     * That is a valid French number, not an error: restore the 0. */
    if (/^[1-9]\d{8}$/.test(digits)) {
      var restored = '0' + digits;
      return {
        value: restored,
        valid: true,
        kind: (restored[1] === '6' || restored[1] === '7') ? 'mobile' : 'fixe',
        reason: 'zero initial rajoute (format sans le 0)'
      };
    }

    /* 11 digits starting with 33: international without the + prefix. */
    if (/^33[1-9]\d{8}$/.test(digits)) {
      var nat = digits.slice(2);
      return {
        value: '0' + nat,
        valid: true,
        kind: (nat[0] === '6' || nat[0] === '7') ? 'mobile' : 'fixe',
        reason: 'converti depuis le format international 33'
      };
    }

    /* A valid 9-digit VoIP number stored with an extra trailing digit:
     *   "9847405322" -> 0984740532 (the last digit is an artefact).
     * Restricted to the 09 range, where the 9 digits after the prefix are
     * unmistakable — never applied to mobile/landline ranges. */
    if (/^9\d{9}$/.test(digits) && /^09\d{8}$/.test('0' + digits.slice(0, 9))) {
      return {
        value: '0' + digits.slice(0, 9),
        valid: true,
        kind: 'fixe',
        reason: 'chiffre final en trop supprime'
      };
    }

    /* 10 digits where the country code REPLACED the leading 0 instead of
     * preceding it: "3243821836" = 3 + 243821836 -> 0243821836.
     * Only when the remaining 9 digits form a valid national number. */
    if (/^3[1-9]\d{8}$/.test(digits)) {
      var body = digits.slice(1);
      if (/^[1-9]\d{8}$/.test(body)) {
        return {
          value: '0' + body,
          valid: true,
          kind: (body[0] === '6' || body[0] === '7') ? 'mobile' : 'fixe',
          reason: 'indicatif pays 33 sans le + : zero initial restaure'
        };
      }
    }

    if (isOtherIntl) {
      /* Non-French international number: preserve, do not rewrite. */
      return { value: original.replace(/[ .\-()]/g, ' ').replace(/\s+/g, ' ').trim(), valid: 'international', kind: 'international', reason: 'numero international conserve tel quel' };
    }

    /* Wrong digit count or unexpected shape. Keep the original text. */
    return { value: original, valid: false, kind: null, reason: 'format de telephone invalide (' + digits.length + ' chiffres)' };
  }

  /* A value that can NEVER be a phone number: too short to be a French
   * number ("1", "12", "123"), a 5-digit postal code, a date, or text.
   * Such a value must never be written into « Téléphone 1/2 ». */
  function isUnusablePhone(raw) {
    /* A phone value is SHORT. A long text normally contains digits, so the
     * old code scanned the whole string to count them. */
    if (raw != null && String(raw).length > 40) return true;
    var s = trimClean(raw);
    if (!s) return true;
    var digits = s.replace(/\D/g, '');
    if (digits.length < 9) return true;                       // 1..8 digits: never a phone
    if (digits.length === 5 && /^0[1-9]\d{3}$|^[1-8]\d{4}$|^9[0-8]\d{3}$/.test(digits)) return true; // postal code
    /* Reject date patterns that should never be in phone columns */
    var sClean = s.toLowerCase();
    /* Check for common date patterns: DD/MM/YYYY, MM/DD/YYYY, DD-MM-YYYY, etc. */
    if (/\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b/.test(s)) {
      /* If the entire value is just a date, reject it */
      if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}$/.test(s)) {
        return true;
      }
      /* If the value contains a date AND has time markers (h, heure, min) or French day names, reject it */
      var hasTime = /h|min|heure/.test(sClean);
      var hasDayNames = /lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche/.test(sClean);
      if (hasTime || hasDayNames) {
        return true;
      }
    }
    /* Also check for "à hhh" pattern (French: "Rappel mardi 02/04/2022 à 13h") */
    if (/à\s\d{1,2}h/.test(sClean)) {
      return true;
    }
    /* A REMINDER / appointment glued into one string
     * (« rappel lundi 08042022 a 10h »): the digits of the date and of the
     * hour add up to exactly 10 digits starting with 0, so the value looked
     * like a perfect French phone number and was written into « Téléphone ».
     * A day name, an hour, an 8-digit date or a reminder word means NOTE. */
    if (/\b(?:lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\b/.test(sClean)) return true;
    if (/\b\d{1,2}\s*h(?:\s*\d{1,2})?\b/.test(sClean)) return true;          // 10h, 10h30
    if (/\b(?:a|à|vers|entre|apres|après|de|jusqu)\s*\d{1,2}\s*h/.test(sClean)) return true;
    if (/\b\d{8}\b/.test(sClean)) return true;                               // 08042022 = 08/04/2022
    if (/\b(?:rappel|rappeler|rappelle|rappellee|rdv|rendez[- ]?vous|appel(?:er|e|le)?|repondeur|absent|occupe|injognable|injoignable|message|passera?|repasser|venir|disponible)\b/.test(sClean)) return true;
    return false;
  }

  function looksLikePhone(raw) {
    if (raw != null && String(raw).length > 40) return false; // never a phone
    var r = analyzePhone(raw);
    return r.valid === true || r.valid === 'international';
  }

  /* Remplit un numéro de téléphone : on récupère les chiffres négatifs ou
   * séparateurs (espace, tiret, slash…) pour ne garder que le flux de
   * chiffres. Une valeur comme « 0248636694 | 0678960143 » ou
   * « 02 48 63 66 94 | 06 78 96 01 43 » donne deux numéros. */
  function splitPhoneValues(raw) {
    var segments = trimClean(raw).split(/[|\/;]+/);
    var phones = [];
    for (var i = 0; i < segments.length; i++) {
      var s = trimClean(segments[i]);
      if (!s) continue;
      var digits = s.replace(/\s/g, '').replace(/[^0-9]/g, '');
      // Only accept 9 or 10 digit numbers as valid phone numbers
      if (digits.length === 9 || digits.length === 10) {
        phones.push(digits);
      }
    }
    return phones;
  }

  /* ---------- French postal code ---------- */

  function analyzePostal(raw) {
    var original = trimClean(raw);
    if (!original) return { value: '', valid: null, reason: '' };
    /* Placeholder junk (".", "....", "//") means "no postal code". */
    if (PLACEHOLDER_RE.test(original)) return { value: '', valid: null, reason: '' };
    /* A postal code contains ONLY digits (plus harmless separators). Anything
     * carrying a currency sign, decimals or letters is NEVER a postal code
     * (e.g. "120 €" must not be read as "120"). */
    if (/[^\d\s\-\/.]/.test(original) || /\d[.,]\d{1,2}\s*$/.test(original)) {
      return { value: original, valid: false, reason: 'code postal non reconnu (5 chiffres attendus)' };
    }
    /* Reject date-like values (M/D/YY, MM/DD/YYYY, D.M.YYYY, etc.) — these are
     * calendar dates, not postal codes, even if their compact digits look like 5. */
    if (/^\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}$/.test(original)) {
      return { value: original, valid: false, reason: 'format de date detecte, non un code postal' };
    }
    var digits = original.replace(/\D/g, '');
    /* A duplicated code repeats the same 5 digits ("22260\t22260",
     * "62118 62118 62118"): keep one copy. Any length multiple of 5 whose
     * 5-digit blocks are all identical is a duplication artefact. */
    if (digits.length > 5 && digits.length % 5 === 0) {
      var block = digits.slice(0, 5);
      var same = true;
      for (var bi = 0; bi < digits.length; bi += 5) {
        if (digits.slice(bi, bi + 5) !== block) { same = false; break; }
      }
      if (same && /^0[1-9]\d{3}$|^[1-8]\d{4}$|^9[0-8]\d{3}$/.test(block)) {
        return { value: block, valid: true, reason: 'code postal duplique corrige (' + block + ')' };
      }
    }
    /* French metropolitan 5-digit codes incl. leading zero; overseas 97/98xxx. */
    if (/^\d{5}$/.test(digits) && /^0[1-9]\d{3}$|^[1-8]\d{4}$|^9[0-8]\d{3}$/.test(digits)) {
      return { value: digits, valid: true, reason: '' };
    }
    /* Legacy exports drop the leading zero: 6000 -> 06000, 7501 -> 07501.
     * 4 digits starting 1-9 is a valid French postal code. */
    if (/^[1-9]\d{3}$/.test(digits)) {
      return { value: '0' + digits, valid: true, reason: 'zero initial rajoute (format sans le 0)' };
    }
    /* A postal code glued to a DATE ("4939015/02/1932" = 49390 + 15/02/1932):
     * keep the first 5-digit block. Restricted to values that look like
     * code+date (first part is ≥5 digits), never IBANs or account numbers
     * (those stay flagged). Pure short dates like "3/28/24" are excluded by
     * the date-rejection guard above and never reach this point. */
    if (digits.length > 5 && /[\/.]/.test(original) && !/^[A-Za-z]{2}\d/.test(original)) {
      /* The original must begin with at least 5 consecutive digits before
       * the first separator — ruling out M/D/YY patterns whose digit run
       * before the first "/" is just 1-2 characters. */
      var firstSep = original.search(/[\/.-]/);
      if (firstSep >= 5) {
        var head = digits.slice(0, 5);
        if (/^0[1-9]\d{3}$|^[1-8]\d{4}$|^9[0-8]\d{3}$/.test(head)) {
          return { value: head, valid: true, reason: 'code postal extrait du debut de la valeur (' + head + ')' };
        }
      }
    }

    /* 6 digits ending in 0: a trailing zero was appended by the export
     * ("161000" -> 16100, "741000" -> 74100). Only when the first 5 digits
     * already form a valid French code. */
    if (/^\d{6}$/.test(digits) && digits[5] === '0') {
      var five = digits.slice(0, 5);
      if (/^0[1-9]\d{3}$|^[1-8]\d{4}$|^9[0-8]\d{3}$/.test(five)) {
        return { value: five, valid: true, reason: 'zero final en trop supprime (' + five + ')' };
      }
    }
    /* 5 digits that are a dropped-zero code with an extra digit are NOT
     * valid; keep the original and flag it. */
    return { value: original, valid: false, reason: 'code postal non reconnu (5 chiffres attendus)' };
  }

  function looksLikePostal(raw) {
    if (raw != null && String(raw).length > 20) return false;
    return analyzePostal(raw).valid === true;
  }

  /* ---------- Email ---------- */

  var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

  /* Truncated e-mails from legacy exports: "x@orange.f", "y@gmail.c"
   * (the TLD lost letters). Complete the known French provider domains so the
   * address is kept usable instead of being flagged as a broken value. */
  var TRUNCATED_DOMAINS = {
    'orange.f': 'orange.fr', 'gmail.c': 'gmail.com', 'wanadoo': 'wanadoo.fr',
    'wanadoo.': 'wanadoo.fr', 'numericabl': 'numericable.fr', 'numericabl.': 'numericable.fr',
    'club-internet.f': 'club-internet.fr', 'free.f': 'free.fr', 'sfr.f': 'sfr.fr',
    'laposte.ne': 'laposte.net', 'hotmail.c': 'hotmail.com', 'yahoo.f': 'yahoo.fr',
    'bbox.f': 'bbox.fr', 'aliceadsl.f': 'aliceadsl.fr'
  };

  /* Domains that were CUT OFF mid-name by the source export (the cell simply
   * ended at the column width). Only completion of a KNOWN provider is safe:
   *    "...@oran"          -> @orange.fr
   *    "...@gmail"         -> @gmail.com
   *    "...@clubinternet"  -> @club-internet.fr
   *    "...@janville-sur-ju" -> left as is (unknown domain, cannot guess) */
  var PARTIAL_DOMAINS = [
    ['orange', 'orange.fr'], ['gmail', 'gmail.com'], ['hotmail', 'hotmail.com'],
    ['outlook', 'outlook.fr'], ['live.f', 'live.fr'], ['live', 'live.fr'],
    ['laposte', 'laposte.net'], ['wanadoo', 'wanadoo.fr'], ['sfr', 'sfr.fr'],
    ['free', 'free.fr'], ['bbox', 'bbox.fr'], ['numericable', 'numericable.fr'],
    ['clubinternet', 'club-internet.fr'], ['club-internet', 'club-internet.fr'],
    ['libertysurf', 'libertysurf.fr'], ['yahoo', 'yahoo.fr'], ['aliceadsl', 'aliceadsl.fr'],
    ['departement13', 'departement13.fr'], ['ac-grenoble', 'ac-grenoble.fr']
  ];

  function repairEmail(value) {
    var at = value.lastIndexOf('@');
    if (at === -1) return value;
    var localPart = value.slice(0, at);
    var domain = value.slice(at + 1).toLowerCase().replace(/\s+/g, '').replace(/\.+$/, '');
    if (!localPart) return value;

    if (TRUNCATED_DOMAINS[domain]) return localPart + '@' + TRUNCATED_DOMAINS[domain];

    /* Exact (or slight) MATCH against a known provider whose name was cut. */
    for (var i = 0; i < PARTIAL_DOMAINS.length; i++) {
      var short = PARTIAL_DOMAINS[i][0];
      var full = PARTIAL_DOMAINS[i][1];
      if (domain === short || (domain.length >= 4 && full.indexOf(domain) === 0)) {
        return localPart + '@' + full;
      }
    }
    /* A domain with NO dot at all ("...@inra", "...@dbm") is a truncated
     * e-mail: the site lost its TLD. Keep the address as-is and flag it weak
     * rather than broken — the local part + domain are still usable. */
    if (/^[a-z0-9\-]{2,}$/.test(domain) && domain.indexOf('.') === -1) {
      return localPart + '@' + domain;
    }

    /* "orange.f" style: single-letter TLD on a known provider -> ".fr". */
    var m = domain.match(/^([a-z\-]+)\.([a-z])$/);
    if (m) {
      var known = { orange: 'fr', gmail: 'com', hotmail: 'com', yahoo: 'fr', wanadoo: 'fr',
        sfr: 'fr', free: 'fr', laposte: 'net', numericable: 'fr', bbox: 'fr', aliceadsl: 'fr' };
      if (known[m[1]]) return localPart + '@' + m[1] + '.' + known[m[1]];
    }
    return value;
  }

  function analyzeEmail(raw) {
    var value = asString(raw).replace(/ /g, ' ').trim();
    if (!value) return { value: '', valid: null, reason: '' };
    if (PLACEHOLDER_RE.test(value)) return { value: '', valid: null, reason: '' };
    if (EMAIL_RE.test(value)) return { value: value, valid: true, reason: '' };
    /* Try repairing a truncated provider domain ("x@orange.f"). */
    var repaired = repairEmail(value);
    if (repaired !== value && EMAIL_RE.test(repaired)) {
      return { value: repaired, valid: true, reason: 'domaine email tronque complete (' + repaired + ')' };
    }
    return { value: value, valid: false, reason: 'adresse email au format suspect' };
  }

  function looksLikeEmail(raw) {
    if (raw != null && String(raw).length > 120) return false;
    return analyzeEmail(raw).valid === true;
  }

  /* ---------- Cotisation (French numeric formats) ---------- */

  /* Recognize 75,50 / 1 250,00 / "120 €" / 1.250,00 etc. without converting. */
  var AMOUNT_RE = /^-?\d{1,3}(?:[  .,]\d{3})*(?:[.,]\d{1,2})?$|^-?\d+(?:[.,]\d{1,2})?$/;

  function parseFrenchNumber(raw) {
    var s = trimClean(raw);
    if (!s) return null;
    var cleaned = s.replace(/[€$£%]/g, ' ')
      .replace(/euros?|eur|fcfa|francs?/gi, ' ')
      .replace(/ /g, ' ')
      .replace(/\s+/g, '')
      .replace(/,/g, '.');
    /* 1.250.30 (dot thousands) ambiguous with decimals: handle 1.250,00 style */
    if (/\.\d{3}\./.test(cleaned)) cleaned = cleaned.replace(/\./g, '');
    var n = Number(cleaned);
    return isFinite(n) ? n : null;
  }

  /* An amount whose meaning is CERTAIN on its own: it carries a currency or a
   * decimal part ("75,50", "80 €", "1 250,00 €"). A bare "70" or "80" is not
   * confirmed: it may be an age, a year or a count, so it is treated as
   * ambiguous unless the column context proves it is a price. */
  function isConfirmedAmount(raw) {
    var s = trimClean(raw);
    if (!s || !looksLikeAmount(s)) return false;
    if (/[€$£%]|\b(euros?|euro|eur|fcfa|francs?)\b/i.test(s)) return true;
    /* Decimal part written with a mark that cannot be mistaken for a letter. */
    if (/\d[.,]\d{1,2}(?:[^\d]|$)/.test(s)) return true;
    /* "1,250" / "2.500" (thousands) only with a 4+ digit integer part. */
    if (/^\d[\d\s]{2,}[.,]\d{3}$/.test(s)) return true;
    /* A 4+ digit integer with no separator ("361", "1250") is a price or a
     * count, never an age. */
    return /^\d{4,}$/.test(s.replace(/\s/g, ""));
  }

  /* The value is EXCLUSIVELY a money amount ("1 002,84 euro", "80 €"): used to
   * promote a price found in a free-text column to « Cotisation ». A text
   * that merely CONTAINS an amount ("Rappel mardi - 45 €") is not one. */
  function isPureAmount(raw) {
    var s = trimClean(raw);
    if (!s || s.length > 24) return false;
    if (!looksLikeAmount(s)) return false;
    /* Only a number + optional currency word/sign: nothing else. */
    var rest = s
      .replace(/\u00a0/g, ' ')
      .replace(/(?:^|[\s\d])(?:euros?|euro|eur|fcfa|francs?)(?:$|[\s.,;])/gi, ' ')
      .replace(/[€$£%]/g, ' ')
      .replace(/[\d\s.,+-]/g, '');
    return rest.trim() === '';
  }

  /* ---------- Packed free-text records ("test happ -- ... --")
   * Real exports hide a whole record in ONE cell, separated by "--":
   *   test happ -- 07*12*2020 ---- 2 69 02 13 055 625 36 -- GROUPAMA --
   *   16*06*1936 -- 139,6 -- FR76...
   * (date, n° sécurité sociale, organisme, naissance, COTISATION, IBAN).
   * The amount was buried in the middle of the text and never reached
   * « Cotisation ». Returns { amount, insurer } (empty strings when absent). */
  function parsePackedRecord(raw) {
    var out = { amount: '', insurer: '' };
    var s = trimClean(raw).replace(/\u00a0/g, ' ');
    if (!s || s.length < 12 || s.indexOf('--') === -1) return out;
    var parts = s.split(/\s*-{2,}\s*|\s*--+\s*/).map(function (p) { return trimClean(p); }).filter(Boolean);
    parts.forEach(function (p) {
      if (!out.amount) {
        /* "139,6" "109,33€" "104,05 €" : a short number, optionally with a
         * currency. Never a 5-digit postal code nor an age (<= 110 without
         * currency is ambiguous and stays out). */
        var m = p.match(/^(\d{1,3}(?:[\s\u00a0]\d{3})*|\d+)[.,](\d{1,2})\s*(?:€|euros?|eur)?$/i) ||
          p.match(/^(\d{1,3}(?:[\s\u00a0]\d{3})*|\d+)\s*\u20ac$/i);
        if (m && !/^\d{5}$/.test(p)) {
          out.amount = p;
          return;
        }
      }
      if (!out.insurer) {
        var ins = matchInsurer(p);
        if (ins && !/\d/.test(p)) out.insurer = ins;
      }
    });
    return out;
  }

  function looksLikeAmount(raw) {
    if (raw != null) {
      /* "109,33\u20ac" (no space before the sign) is a real Excel value: test the
       * currency FIRST so the decimal regex cannot be blocked by it. */
      var quick = trimClean(raw);
      if (quick && quick.length <= 40 && /\u20ac/.test(quick) && /\d/.test(quick) && !looksLikePhone(quick)) return true;
    }
    /* A cell value that is an amount is SHORT by nature: skip the (expensive)
     * whole-text regexes for long notes. Nothing is ever detected as an
     * amount only because of its 200 000th character. */
    if (raw != null && String(raw).length > 40) return false;
    var s = trimClean(raw);
    if (!s) return false;
    /* A phone number is NEVER an amount, even written with dots or spaces
     * ("06.12.34.56.78" would otherwise look like a decimal number). */
    if (looksLikePhone(s)) return false;
    /* A bare 5-digit French postal code is not an amount either. */
    if (/^\d{5}$/.test(s) && /^0[1-9]\d{3}$|^[1-8]\d{4}$|^9[0-8]\d{3}$/.test(s)) return false;
    /* A bare integer without a currency and without a separator ("70", "80")
     * is AMBIGUOUS: it may be an age, a year or a count. It is never stated as
     * an amount here; the context (header + other values) decides. A 4-digit
     * value ("1200", "1950") is never an age, so it keeps the old behaviour. */
    if (/^\d{1,3}$/.test(s)) return false;
    /* THE CURRENCY CHECK COMES FIRST: "361 euro" contains letters, so the
     * letter guard below used to reject it before the word "euro" was even
     * looked at (that is exactly why 361 euro was not a Cotisation). */
    /* \u00a0 = the NON-BREAKING space Excel writes as thousands separator
     * ("1 002,84 euro"): without it those amounts were not recognised. */
    var hasCurrency = /(?:^|[\s\u00a0\d])(?:euros?|euro|eur\b|fcfa|francs?)(?:$|[\s\u00a0.,;])/i.test(s) || /[€$£%]/.test(s);
    if (hasCurrency && /\d/.test(s)) {
      /* A currency marker next to a number is money, even with letters.
       * Only the currency WORDS are removed: "\u00a0" is kept in the test so a
       * non-breaking thousands separator never looks like a letter. */
      if (!/[A-Za-zÀ-ÿ]{4,}/.test(s.replace(/euros?|euro|eur|fcfa|francs?/gi, ""))) return true;
    }
    /* A value with real letters ("Neoliane", "MGEN") is never an amount.
     * Decimals are matched on the ORIGINAL text so a letter can never be a mark. */
    if (/[A-Za-zÀ-ÿ]{2,}/.test(s)) return false;
    if (!/\d/.test(s)) return false;
    var hasDecimalPart = /[.,]\d{1,2}$/.test(s) || /[.,]\d{3,}$/.test(s);
    return hasDecimalPart;
  }

  /* Pull the first numeric money-like value out of a compound text, e.g.
   *   "A· 45,60 €"      -> "45,60 €"
   *   "NV3 1 002,84 euro"-> "1 002,84 euro"
   *   "55 euros"         -> "55 euros"
   *   "120"              -> "120"
   * Returns '' when the value is a phone number, a postal code or has no
   * recoverable amount (never hijack those). */
  function extractAmount(raw) {
    var s = trimClean(raw);
    if (!s) return '';
    if (looksLikePhone(s) || /^\d{5}$/.test(s)) return '';
    /* Explicit currency marker wins and is kept verbatim. */
    var nb = s.replace(/\u00a0/g, ' ');
    var currency = nb.match(/-?\d[\d\s.]*(?:[.,]\d{1,2})?\s*(?:€|euros?|eur\b|fcfa|francs?)/i);
    if (currency) return trimClean(currency[0]);
    /* Otherwise a decimal amount anywhere in the text. */
    var decimal = nb.match(/-?\d[\d\s.]*[.,]\d{1,2}\b/);
    if (decimal) return trimClean(decimal[0]);
    return '';
  }

  function analyzeAmount(raw) {
    var original = trimClean(raw);
    if (!original) return { value: '', valid: null, reason: '' };
    var n = parseFrenchNumber(original);
    if (n === null) return { value: original, valid: false, reason: 'montant non numerique' };
    return { value: original, valid: true, numeric: n, reason: '' };
  }

  /* ---------- Age / age-range values (never Cotisation) ----------
   * Real exports carry columns of subscriber ages:
   *   "70+75"  "84+85"  "61+72"  "70"  "77"  "63 ans"  "76+76 ans"
   *   "87 ans"  "1950-1948"  (birth years of a couple)
   * Those are NEVER amounts and NEVER postal codes: they belong in
   * « Commentaire ». Recognized only as a WHOLE value, so a real amount
   * ("75,50") or a postal code ("75160") is never caught. */
  var AGE_RANGE_RE = /^\d{1,3}(?:\s*\+\s*\d{1,3})+(?:\s*ans?)?$/i;      // 70+75, 76+76 ans
  var AGE_SINGLE_RE = /^\d{1,3}\s*ans?$/i;                                // 63 ans, 87 ans
  var AGE_BARE_RE = /^\d{1,3}$/;                                           // 70, 77, 64
  var YEARS_RE = /^(?:19|20)\d{2}(?:\s*[-\/]\s*(?:19|20)\d{2})+$/;      // 1950-1948

  /* A bare number ("70", "80", "1950") whose meaning depends on context. */
  function isAmbiguousNumber(raw) {
    var s = trimClean(raw);
    if (!s || /[€$£%]|\b(euros?|eur|fcfa|francs?)\b/i.test(s)) return false;
    if (!/^\d{1,3}$/.test(s)) return false;
    /* A plausible human age (18..110). "1", "0" or "999" is not an age
     * anyone subscribes health cover at, so it is not "ambiguous" — the
     * column context decides it. */
    var n = parseInt(s, 10);
    return n >= 18 && n <= 110;
  }

  function looksLikeAge(raw) {
    if (raw != null && String(raw).length > 40) return false; // never a long note
    var s = trimClean(raw);
    if (!s) return false;
    if (AGE_RANGE_RE.test(s) || AGE_SINGLE_RE.test(s) || YEARS_RE.test(s)) return true;
    if (AGE_BARE_RE.test(s)) {
      var n = parseInt(s, 10);
      return n >= 1 && n <= 110;   // a plausible age, not 120 or 9999
    }
    return false;
  }

  /* Share of the values of a column that are ages (0..1). */
  function ageRatio(values) {
    var seen = 0, hits = 0;
    (values || []).forEach(function (v) {
      var s = trimClean(v);
      if (!s || PLACEHOLDER_RE.test(s)) return;
      seen++;
      if (looksLikeAge(s)) hits++;
    });
    return seen ? hits / seen : 0;
  }

  /* ---------- Mutuelle / insurer names ----------
   * "aesio", "alliance", "apivia" are health insurers (Mutuelle), never
   * amounts, even when they sit in a Cotisation column. */
  var INSURER_HINTS = [
    'neoliane', 'sursante', 'vitaneor', 'cegema', 'asaf', 'apreva', 'gerep',
    'aesio', 'alliance', 'apivia', 'ag2r', 'harmonie', 'malakoff', 'humanis',
    'groupama', 'maif', 'macif', 'matmut', 'gmf', 'mgen', 'mutuelle', 'mutualia',
    'klesia', 'klésia', 'pro btp', 'cqfd', 'apicil', 'swiss', 'gan', 'mma',
    'mnh', 'mnt', 'smacl', 'sma', 'caisse', 'assur', 'prevoyance', 'prévoyance',
    'complementaire', 'complémentaire', 'sante', 'santé', 'axa', 'allianz',
    'generali', 'aviva', 'cnp', 'mae', 'mgefi', 'vitale', 'radiance', 'solimut',
    'uni', 'via santé', 'via sante', 'asaf', 'ocean', 'océan', 'groupe',
    /* Abbreviations seen in real exports ("C.MUT", "C MUT", "MUT.").
     * They need a SEPARATOR, otherwise "communaute" would match "mut". */
    'c.mut', 'c mut', 'c-mut', 'mut.', 'mut ',
  ];

  /* Official insurer database (js/data/mutuelles.js): alias -> official name.
   * Built once, lazily, from "Nom|alias1,alias2" lines. */
  var INSURER_DB = null;
  function insurerDb() {
    if (INSURER_DB) return INSURER_DB;
    INSURER_DB = Object.create(null);
    var raw = typeof SEC.MUTUELLE_DB === 'string' ? SEC.MUTUELLE_DB : '';
    raw.split('\n').forEach(function (line) {
      if (!line) return;
      var bar = line.indexOf('|');
      if (bar === -1) return;
      var official = line.slice(0, bar).trim();
      line.slice(bar + 1).split(',').forEach(function (alias) {
        var key = stripAccents(alias).toLowerCase().replace(/[^a-z0-9]+/g, '').trim();
        if (!key) return;
        if (!INSURER_DB[key]) INSURER_DB[key] = official;
        /* "C.MUT" / "C MUT" / "MUT." : the punctuation IS the abbreviation.
         * The compact key above ("cmut") matches "CMUT"; the spaced and the
         * dotted forms are matched separately below so they never swallow an
         * ordinary word ("communaute"). */
        var spaced = stripAccents(alias).toLowerCase().replace(/[.\-_/]+/g, ' ').replace(/\s+/g, ' ').trim();
        if (spaced && spaced !== key && !INSURER_DB[spaced]) INSURER_DB[spaced] = official;
      });
    });
    return INSURER_DB;
  }

  /* Return the official insurer name for a value, or '' when it is not one.
   * Matching order: exact database alias, then a database alias contained in
   * the value as a whole word, then the built-in hint list. */
  /* Memoised: the same values repeat across rows and files, and the fuzzy
   * pass below is the most expensive check of the whole pipeline. */
  var INSURER_PROBE_RE = null;
  function insurerProbe() {
    if (INSURER_PROBE_RE) return INSURER_PROBE_RE;
    var db = insurerDb();
    var parts = [];
    for (var k in db) if (k.length >= 4) parts.push(k);
    (INSURER_HINTS || []).forEach(function (h) {
      if (h && parts.indexOf(h) === -1) parts.push(h.replace(/\s+/g, '\\s*'));
    });
    /* Fragments of 4+ letters only: the fuzzy pass handles typos. */
    INSURER_PROBE_RE = new RegExp('(' + parts.join('|') + ')', 'i');
    return INSURER_PROBE_RE;
  }

  var INSURER_CACHE = new Map();
  var INSURER_CACHE_MAX = 20000;

  function matchInsurer(raw) {
    /* Length guard BEFORE any normalisation: a long text is never an insurer
     * and normalising it (accents + case over 32 000 chars) is pure waste. */
    if (raw != null && String(raw).length > 60) return '';
    var s = stripAccents(trimClean(raw)).toLowerCase();
    if (!s) return '';
    /* A pure number is not a name. The test must NOT accept the letters
     * "euro" from a company name like "EUROCOURTAGE" (which used to be
     * rejected as if it were a currency): only digits, spaces and ". ," are
     * a number here. */
    if (/^[\d.,\s+-]+$/.test(s)) return '';
    /* "1 002,84 euro" is an amount, never an insurer. */
    if (/^\d[\d\s\u00a0.,]*\s*(?:€|euros?|eur)$/i.test(s)) return '';
    if (/^(19|20)\d{2}/.test(s)) return '';
    var cached = INSURER_CACHE.get(s);
    if (cached !== undefined) return cached;
    var result = matchInsurerUncached(raw, s);
    /* Only short values are worth memoising: caching multi-KB comments would
     * keep megabytes of unique text alive for nothing. */
    if (s.length <= 60) {
      if (INSURER_CACHE.size >= INSURER_CACHE_MAX) INSURER_CACHE.clear();
      INSURER_CACHE.set(s, result);
    }
    return result;
  }

  /* A first initial followed by an insurer pattern ("m g e n"): checked with
   * the SAME normalisation as the list, but only for the rarer cases. */
  var LETTER_GAP_RE = /[a-z](?:\s+[a-z])+/;

  function matchInsurerUncached(raw, s) {
    /* A long text (comment, note, address) is never an insurer NAME. Testing
     * megabytes of notes against the insurer patterns and the bigram
     * similarity was the hottest path AND the biggest source of garbage: the
     * normaliser lower/uppercases and copies the whole string first. */
    if (s.length > 60) return '';
    /* Cheapest possible rejection FIRST: a value without any known insurer
     * fragment cannot match, and building/scanning the whole alphabetically
     * sorted insurer DB for every row was the hottest path of the batch. */
    if (
      !insurerProbe().test(s) &&
      !LETTER_GAP_RE.test(s) &&
      !(INSURER_HINTS || []).some(function (h) { return h && s.indexOf(h) !== -1; })
    )
      return '';
    var db = insurerDb();
    var compact = s.replace(/[^a-z0-9]+/g, '');
    if (db[compact]) return db[compact];
    var words = s.split(/[^a-z0-9]+/).filter(Boolean);
    /* Abbreviations written with a separator ("c.mut", "c mut", "mut."):
     * matched as a whole value before the word loop. */
    if (db[s]) return db[s];
    var dotted = s.replace(/[.\-_/]+/g, ' ').replace(/\s+/g, ' ').trim();
    if (dotted !== s && db[dotted]) return db[dotted];
    /* Whole-word match, so "alliance" matches but "alliancing" does not. */
    for (var i = 0; i < words.length; i++) {
      if (db[words[i]]) return db[words[i]];
      if (i + 1 < words.length && db[words[i] + words[i + 1]]) return db[words[i] + words[i + 1]];
    }
    for (var k in db) {
      if (k.length >= 5 && compact.indexOf(k) !== -1) return db[k];
    }
    /* Tolerant match for frequent typos ("swisselife", "maciff"):
     * an alias of 5+ letters that differs by at most one inserted/removed
     * letter is accepted, provided the word is long enough to be unambiguous. */
    var cand = compact.length >= 6 ? compact : words.join('');
    /* A person name, an address or a comment is not a mistyped insurer: the
     * fuzzy comparison only runs on short, single-block values. */
    var nameShaped = words.length <= 3 && !/\s/.test(trimClean(raw)) && raw.length <= 40;
    if (cand.length >= 6 && cand.length <= 24 && nameShaped) {
      for (var k2 in db) {
        if (k2.length < 6) continue;
        var d = Math.abs(k2.length - cand.length);
        if (d > 2) continue;
        if (overlapScore(cand, k2) >= 0.82) return db[k2];
      }
    }
    if (INSURER_HINTS.some(function (h) { return s.indexOf(h) !== -1; })) {
      return trimClean(raw);
    }
    return '';
  }

  /* Jaccard similarity of the character bigrams of two strings (0..1). */
  var GRAM_CACHE = new Map();
  function grams(x) {
    var hit = GRAM_CACHE.get(x);
    if (hit) return hit;
    var set = Object.create(null), n = 0;
    for (var i = 0; i < x.length - 1; i++) {
      var g = x.substr(i, 2);
      if (!set[g]) { set[g] = true; n++; }
    }
    var out = { set: set, n: n };
    if (GRAM_CACHE.size < 5000) GRAM_CACHE.set(x, out);
    return out;
  }

  function overlapScore(a, b) {
    if (!a || !b) return 0;
    var A = grams(a), B = grams(b);
    var inter = 0;
    for (var g in A.set) if (B.set[g]) inter++;
    var uni = A.n + B.n - inter;
    return uni ? inter / uni : 0;
  }

  /* True when the value IS an insurer name (and not an amount). */
  function looksLikeInsurer(raw) {
    if (raw != null && String(raw).length > 60) return false;
    return !!matchInsurer(raw);
  }

  /* The value is EXCLUSIVELY an insurer name ("C.MUT", "MGEN", "APREVA"),
   * with no surrounding sentence: such a value found in « Commentaire » is a
   * Mutuelle, not a note. */
  function isPureInsurer(raw) {
    var s = trimClean(raw);
    if (!s || s.length > 40) return false;
    if (/\d/.test(s) || /[.!?]{2,}/.test(s)) return false;
    if (s.split(/\s+/).length > 4) return false;
    return !!matchInsurer(s);
  }

  /* ---------- Address / city value signals ---------- */

  var STREET_RE = null;
  function looksLikeAddress(raw) {
    /* An address is short; a long note is not. The street word may appear
     * far inside a huge text: ignore such values entirely (cheap + safer). */
    if (raw != null && String(raw).length > 200) return false;
    var s = stripAccents(asString(raw)).toLowerCase();
    if (!s.trim()) return false;
    if (!STREET_RE) {
      /* WHOLE-WORD match, mandatory for the short abbreviations ("pl" for
       * place, "av" for avenue): as a substring, "pl" matched inside "se
       * PLAINT" and made ordinary notes look like postal addresses. */
      var alts = (SEC.STREET_TERMS || [])
        .filter(Boolean)
        .sort(function (a, b) { return b.length - a.length; })
        .map(function (t) {
          var esc = t.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
          return t.length <= 4 ? "(?<![a-z])" + esc + "(?![a-z])" : esc;
        });
      STREET_RE = new RegExp("(" + alts.join("|") + ")", "i");
    }
    /* A sentence is a NOTE, not an address, whatever word it contains. */
    if (/\s(?:se|ne|il|elle|on|nous|vous|ils|elles|est|sont|etait|a ete|sera|pour|avec|mais|donc|car|que|qui|du prix|trop|tres|plus|deja)\s/.test(" " + s + " ")) return false;
    var hasStreetTerm = STREET_RE.test(s);
    var hasNumber = /^\d{1,4}[a-z]?(\s|,|$)/.test(s.trim());
    /* A street term alone is not enough: "Place" in a note is a word, an
     * address is a number + a street (or a known house term such as hameau). */
    if (hasStreetTerm && hasNumber) return true;
    if (hasStreetTerm) {
      var words = s.trim().split(/\s+/);
      return words.length >= 2 && words.length <= 8 && !/[.;!?]/.test(s);
    }
    return hasNumber && s.length > 12;
  }

  /* Short free text that reads like a call note / remark rather than a value
   * that belongs to one of the 11 columns ("appele", "ne repond pas",
   * "rappel mardi", "injoignable", "ok"). */
  var COMMENT_TEXT_RE = /\b(rappel|rappeler|appel[eé]|appele|repond|repondu|injoignable|absent|relance|recontact|messagerie|sans suite|ne repond pas|ok|ras|annul[eé]|refus[eé]|attente|en cours)\b/i;
  function looksLikeCommentText(raw) {
    var s = trimClean(raw);
    if (!s) return false;
    if (/@/.test(s) || /\d{5,}/.test(s.replace(/\s/g, ''))) return false;
    var plain = stripAccents(s).toLowerCase();
    if (COMMENT_TEXT_RE.test(plain)) return true;
    /* A SHORT value with NO address structure and no column meaning is a note:
     * "injognable", "elle se plaint du prix", "pas de reponse". Those used to
     * drift into Adresse / Nom because any two capitalised words looked like a
     * name or a place. */
    if (/[.!?]/.test(plain)) return true;
    var words = plain.split(/\s+/).filter(Boolean);
    var NOTE_ONLY = /\b(prix|plaint|cher|trop|chang|mutuelle|sante|malade|veuve|decede|mari|femme|fille|fils|enfant|travail|chomage|retraite|probleme|attente|reporte|compare|comparatif|devis|documents?|piece|justificatif)\b/;
    return words.length >= 2 && words.length <= 12 && NOTE_ONLY.test(plain);
  }

  /* A pure running sequence (1, 2, 3… or D-1, D-2, D-3…) is an identifier:
   * it never belongs to one of the 11 official columns. */
  function looksLikeSequenceSamples(samples) {
    var list = (samples || []).map(trimClean).filter(Boolean);
    if (list.length < 3) return false;
    var plain = list.filter(function (v) { return /^\d{1,9}$/.test(v); });
    if (plain.length === list.length) {
      var nums = plain.map(function (v) { return parseInt(v, 10); });
      for (var i = 1; i < nums.length; i++) if (nums[i] - nums[i - 1] !== 1) return false;
      return true;
    }
    var prefixed = list.filter(function (v) { return /^[A-Za-z]{1,6}[-_/ ]?\d{1,9}$/.test(v); });
    if (prefixed.length !== list.length) return false;
    var tails = prefixed.map(function (v) { return parseInt(String(v).replace(/^.*?[-_/ ]?(\d+)$/, '$1'), 10); });
    for (var j = 1; j < tails.length; j++) if (tails[j] - tails[j - 1] !== 1) return false;
    return true;
  }

  function looksLikeCity(raw, knownCities) {
    if (raw != null && String(raw).length > 120) return false;
    var s = trimClean(raw);
    if (!s || /\d/.test(s) || s.length > 45) return false;
    var accentless = stripAccents(s).toLowerCase();
    if (knownCities && knownCities.indexOf && knownCities.indexOf(accentless) !== -1) return true;
    /* A single French first name ("GEORGES", "PIERRE", "MICHEL") is a person,
     * not a city — those are also common town names but far more often used
     * as given names in these files. */
    if (isKnownPrenom(s)) return false;
    /* Two words where one is a known French first name -> it is a PERSON name
     * ("DUPONT Jean"), never a city. */
    var words = s.split(/\s+/).filter(Boolean);
    if (words.length >= 2 && words.length <= 3) {
      var hasPrenom = words.some(function (w) { return isKnownPrenom(w); });
      if (hasPrenom) return false;
    }
    /* A city is letters only (spaces, hyphens, apostrophes). Up to 5 words
     * covers "BERRE L ETANG", "SAINT VINCENT DE TYROSSE". */
    if (words.length > 5) return false;
    if (!/^[A-Za-zÀ-ÿ'’\-\s]+$/.test(s)) return false;
    /* Lowercase-only values are still cities when not a known first name
     * ("nevers", "lyon" in legacy exports). */
    if (s === s.toLowerCase()) return !isKnownPrenom(s);
    /* Capitalized (allowing hyphenated compound names), letters/spaces/hyphens/apostrophes only. */
    return /^[A-ZÉÈÀÂÏÔÜÇŒÆ][A-Za-zÀ-ÿ'’\-]*(?:[\s\-][A-ZÉÈÀÂÏÔÜÇŒÆa-zÀ-ÿ][A-Za-zÀ-ÿ'’\-]*)*$/.test(s);
  }

    /* Strict check for invalid person names:
   * A person name must NOT contain any digits (dates, hours, numbers),
   * nor reminder/workflow keywords like "rappel mardi 23042024 a 15h00". */
  /* Words that can never be part of a person name (precompiled ONCE: these
   * checks run on every row of every file). */
  var INVALID_NAME_WORDS = new RegExp(
    '(^|\\s)(' +
      [
        'rappel', 'rappeler', 'a rappeler', 'rdv', 'rendez vous', 'repondeur',
        'injoignable', 'injognable', 'relance', 'recontact', 'contacte', 'devis',
        'deja assure', 'pas interesse', 'faux numero', 'ne repond pas',
        'date et heure', 'commentaire', 'observation', 'retour', 'statut',
        'ehpad', 'maison de retraite', 'residence', 'hopital', 'clinique',
        'decede', 'decedee', 'dec', 'veuf', 'veuve', 'tutelle', 'curatelle',
        'mandataire', 'enfant', 'fille', 'fils', 'epoux', 'epouse', 'conjoint',
        'ans', 'age', 'ne le', 'nee le', 'sans suite', 'a jour', 'inactif',
        'contrat', 'adhesion', 'radiation', 'resilie',
      ]
        .map(function (t) { return t.replace(/\s+/g, '\\s+'); })
        .join('|') +
      ')(\\s|$)',
  );

  function isInvalidPersonName(raw) {
    var original = raw == null ? '' : String(raw);
    if (original.length > 120) return true; // a name never this long
    var s = trimClean(raw);
    if (!s) return false;
    // Any digits in the name -> invalid person name!
    if (/\d/.test(s)) return true;
    /* Fast path: the token check below is precompiled once (compiling a
     * RegExp per call cost 37 µs and capped the batch at ~850 rows/s). */
    if (INVALID_NAME_WORDS.test(stripAccents(s).toLowerCase())) return true;
    // Symbols like @, http, urls, currency, etc.
    if (/[@&_$%*+=<>{}\[\]\\]/.test(s)) return true;
    /* A comma is the classic "Nom, Prénom" separator (a sentence ending with a
     * comma is caught as a note by isNoteLike). */
    if (!/^[A-Za-zÀ-ÿ'’\-\s.,]+$/.test(s)) return true;
    /* A note is a SENTENCE; a person name has at most 4 words. */
    if (s.replace(/,/g, ' ').split(/\s+/).filter(Boolean).length > 4) return true;
    return false;
  }

  /* Does this value carry a reminder / note / unrelated information? Such a
   * value must never reach Nom or Prénom, whatever column it came from. */
  function isNoteLike(raw) {
    /* A text that is far longer than any name IS a note: decide on the length
     * alone, without normalising megabytes of text for a boolean. */
    var original = raw == null ? '' : String(raw);
    if (original.length > 120) return true;
    var s = trimClean(raw);
    if (!s) return false;
    if (/\d/.test(s)) return true; // dates, hours, ages, phone numbers...
    var norm = stripAccents(s).toLowerCase();
    /* "Dupont, Jean" is a name, not a sentence: only check a comma-separated
     * pair token by token. */
    if (/^[A-Za-zÀ-ÿ'’\-]+\s*[,;/]\s*[A-Za-zÀ-ÿ'’\-]+$/.test(s)) {
      return s.split(/\s*[,;/]\s*/).some(function (part) {
        return isNoteLike(part);
      });
    }
    return /(^|\s)(rappel|rappeler|rdv|rendez vous|relance|injoignable|repondeur|absent|ne repond|pas de reponse|mail envoye|courrier|ehpad|maison de retraite|decede|veuf|veuve|tutelle|curatelle|contrat|adhesion|radiation|resilie|inactif|actif)\b/.test(
      norm,
    );
  }

  function looksLikePersonName(raw) {
    if (raw != null && String(raw).length > 120) return false;
    var s = trimClean(raw);
    if (!s || /\d/.test(s) || s.length < 2 || s.length > 45) return false;
    if (isNoteLike(s)) return false;
    /* KEEP the comma: trimClean may drop it, so test the original too. */
    var orig = String(raw == null ? '' : raw).trim();
    if (/^[A-Za-zÀ-ÿ'’\-]+\s*[,;/]\s*[A-Za-zÀ-ÿ'’\- .]+$/.test(orig)) return true;
    /* "Dupont, Jean" / "Dupont; Jean" : the comma separates surname and first
     * name. */
    if (/^[A-Za-zÀ-ÿ'’\-]+\s*[,;/]\s*[A-Za-zÀ-ÿ'’\- .]+$/.test(s)) return true;
    if (!/^[A-ZÉÈÀÂÏÔÜÇŒÆa-zà-ÿ'’\-\s.]+$/.test(s)) return false;
    /* "DUPONT Jean", "Jean Dupont" : two words, at least one with a capital
     * letter and one being a known first name or an ALL-CAPS surname. */
    var words = s.split(/\s+/).filter(Boolean);
    /* A single known French first name IS a person name. */
    if (words.length === 1) return isKnownPrenom(s);
    var hasPrenom = words.some(function (w) { return isKnownPrenom(w); });
    var hasCaps = words.some(function (w) {
      var letters = w.replace(/[^A-Za-zÀ-ÿ]/g, '');
      return letters.length > 1 && letters === letters.toUpperCase();
    });
    var hasTitleCase = words.filter(function (w) {
      return /^[A-ZÉÈÀÂÏÔÜÇŒÆ][a-zà-ÿ'’\-]+$/.test(w) || /^[A-Z]\.?$/.test(w);
    }).length >= 2;
    return hasPrenom || hasCaps || hasTitleCase;
  }

  /* ---------- Full name splitting (Nom / Prénom) ---------- */

  /* Known French first names are lower-case in COMMON_PRENOMS. */
  var PRENOM_CACHE = new Map();
  var PRENOM_SET = null;
  function isKnownPrenom(token) {
    var raw = String(token == null ? '' : token);
    var hit = PRENOM_CACHE.get(raw);
    if (hit !== undefined) return hit;
    var n = stripAccents(trimClean(raw)).toLowerCase();
    if (!PRENOM_SET) {
      PRENOM_SET = Object.create(null);
      (SEC.COMMON_PRENOMS || []).forEach(function (p) { PRENOM_SET[p] = true; });
    }
    var out = !!n && PRENOM_SET[n] === true;
    if (PRENOM_CACHE.size < 8000) PRENOM_CACHE.set(raw, out);
    return out;
  }

  /* Split "DUPONT Jean", "Jean DUPONT", "Dupont Jean-Pierre", "MARTIN Sophie"
   * into { nom, prenom } WITHOUT ever losing text.
   * Strategy: the token that is a known first name becomes Prénom; when the
   * result is ambiguous the ALL-CAPS token wins as Nom and the other as Prénom.
   * If still ambiguous, keep the FIRST token as Nom and the rest as Prénom
   * (the conventional French "NOM Prénom" order). Merge variables are used
   * so names such as "DE LA VALLEE" keep their particles. */
  /* Split "DUPONT Jean", "Jean DUPONT", "Dupont Jean-Pierre", "MARTIN Sophie",
   * "beau noelle", "Rappenot marie claude", "Mme galan" into { nom, prenom }
   * WITHOUT ever losing text. Handles French compound first names (Jean-Marie,
   * Marie-Claude), civilities (M., Mme, etc.), particles (de, du, van), and
   * ALL-CAPS surnames accurately. */
  function splitFullName(raw) {
    var value = trimClean(raw);
    if (!value) return { nom: '', prenom: '' };

    var civility = '';
    var civMatch = value.match(/^(m\.|mr|mme|mlle|monsieur|madame|mademoiselle)\.?\s+/i);
    if (civMatch) {
      civility = civMatch[1];
      value = value.slice(civMatch[0].length).trim();
    }

    /* Compound particle names (de, du, la, le, van, von, di, da, del) stay
     * attached to the following word. */
    var PARTICLES = ['de', 'du', 'des', 'la', 'le', 'les', 'van', 'von', 'di', 'da', 'del', 'della', 'bin', 'ben'];
    var tokens = value.split(/\s+/).filter(Boolean);
    var groups = [];
    for (var i = 0; i < tokens.length; i++) {
      var t = tokens[i];
      var bare = stripAccents(t).toLowerCase().replace(/[^a-z'\-]/g, '');
      if (PARTICLES.indexOf(bare) !== -1 && i + 1 < tokens.length) {
        groups.push({ text: t + ' ' + tokens[i + 1], tokens: [t, tokens[i + 1]] });
        i++;
      } else {
        groups.push({ text: t, tokens: [t] });
      }
    }
    if (groups.length === 1) {
      var singleNom = civility ? (civility + ' ' + groups[0].text) : groups[0].text;
      return { nom: singleNom, prenom: '' };
    }

    var isFirstKnown = isKnownPrenom(groups[0].tokens[0]);
    var isLastKnown = isKnownPrenom(groups[groups.length - 1].tokens[0]);

    var nom = '', prenom = '';

    /* ALL-CAPS checks */
    var firstIsCaps = groups[0].tokens.join('').replace(/[^A-Za-zÀ-ÿ]/g, '') ===
      groups[0].tokens.join('').replace(/[^A-Za-zÀ-ÿ]/g, '').toUpperCase() &&
      groups[0].tokens.join('').replace(/[^A-Za-zÀ-ÿ]/g, '').length > 1;

    var lastIsCaps = groups[groups.length - 1].tokens.join('').replace(/[^A-Za-zÀ-ÿ]/g, '') ===
      groups[groups.length - 1].tokens.join('').replace(/[^A-Za-zÀ-ÿ]/g, '').toUpperCase() &&
      groups[groups.length - 1].tokens.join('').replace(/[^A-Za-zÀ-ÿ]/g, '').length > 1;

    if (firstIsCaps && !lastIsCaps) {
      /* Authoritative ALL-CAPS surname first: NOM Prénom */
      nom = groups[0].text;
      prenom = groups.slice(1).map(function(g) { return g.text; }).join(' ');
    } else if (lastIsCaps && !firstIsCaps) {
      /* Authoritative ALL-CAPS surname last: Prénom NOM */
      nom = groups[groups.length - 1].text;
      prenom = groups.slice(0, -1).map(function(g) { return g.text; }).join(' ');
    } else if (!isFirstKnown && isLastKnown) {
      /* Surname first, followed by one or more given names: "Rappenot marie claude", "maydat Jean marie" */
      var firstPrenomIdx = -1;
      for (var k = 1; k < groups.length; k++) {
        if (isKnownPrenom(groups[k].tokens[0])) {
          firstPrenomIdx = k;
          break;
        }
      }
      if (firstPrenomIdx !== -1) {
        nom = groups.slice(0, firstPrenomIdx).map(function(g) { return g.text; }).join(' ');
        prenom = groups.slice(firstPrenomIdx).map(function(g) { return g.text; }).join(' ');
      } else {
        nom = groups[0].text;
        prenom = groups.slice(1).map(function(g) { return g.text; }).join(' ');
      }
    } else if (isFirstKnown && !isLastKnown) {
      /* Given name first, followed by surname: "Jean Dupont", "Jean-Pierre Andrée", "Yves Fresquet" */
      var lastPrenomIdx = 0;
      for (var m = 0; m < groups.length - 1; m++) {
        if (isKnownPrenom(groups[m].tokens[0])) {
          lastPrenomIdx = m;
        } else {
          break;
        }
      }
      prenom = groups.slice(0, lastPrenomIdx + 1).map(function(g) { return g.text; }).join(' ');
      nom = groups.slice(lastPrenomIdx + 1).map(function(g) { return g.text; }).join(' ');
    } else if (isFirstKnown && isLastKnown) {
      /* Both first and last tokens are known prenoms: e.g. "jean pierre martin" */
      if (groups.length >= 3 && isKnownPrenom(groups[0].tokens[0]) && isKnownPrenom(groups[1].tokens[0])) {
        prenom = groups.slice(0, -1).map(function(g) { return g.text; }).join(' ');
        nom = groups[groups.length - 1].text;
      } else {
        /* Standard French commercial convention: NOM first, Prénom second */
        nom = groups[0].text;
        prenom = groups.slice(1).map(function(g) { return g.text; }).join(' ');
      }
    } else {
      /* Neither token is a known first name: standard French convention NOM Prénom */
      nom = groups[0].text;
      prenom = groups.slice(1).map(function(g) { return g.text; }).join(' ');
    }

    if (civility) nom = (civility + ' ' + nom).trim();
    if (!nom && prenom) { nom = prenom; prenom = ''; }
    return { nom: nom.trim(), prenom: prenom.trim() };
  }

  /* Split a value that holds BOTH a surname and a first name into the two
   * parts, WITHOUT losing a single character: "DUPONT Jean" -> {nom:"DUPONT",
   * prenom:"Jean"}, "Jean Dupont" -> {nom:"Dupont", prenom:"Jean"},
   * "Dupont, Jean" -> {nom:"Dupont", prenom:"Jean"}.
   * Returns null when the value is not a clear pair (single token, note, or
   * two unknown words whose order cannot be proven) — callers then keep the
   * value whole rather than guess. */
  function splitNamePair(raw) {
    if (raw != null && String(raw).length > 120) return null;
    var s = trimClean(raw);
    if (!s || isNoteLike(s)) return null;
    /* "Dupont, Jean" / "Dupont Jean" / "DUPONT-Jean" */
    var separated = s.split(/\s*[,;/]\s*/).filter(Boolean);
    var tokens = (separated.length === 2 ? separated : s.split(/\s+/))
      .map(function (t) { return t.replace(/^[\s\-]+|[\s\-]+$/g, ''); })
      .filter(Boolean);
    if (tokens.some(function (t) { return isNoteLike(t) || !/[A-Za-zÀ-ÿ]/.test(t); })) return null;

    /* A surname may start with particles ("O Brien", "De La Vallee", "Van Der
     * Berg"): these words can never END a split. The value is split at the
     * LAST prefix of non-particle words (the given names). */
    var wholeWords = s.replace(/\s*[,;/]\s*/g, ' ').split(/\s+/).filter(Boolean);
    if (wholeWords.length < 2) return null;
    if (PARTICLE_WORDS.test(wholeWords[0])) {
      var cutIdx = wholeWords.length - 1;
      while (cutIdx > 0 && PARTICLE_WORDS.test(wholeWords[cutIdx])) cutIdx--;
      if (cutIdx <= 0) return null;
      return {
        nom: wholeWords.slice(0, cutIdx).join(' '),
        prenom: wholeWords.slice(cutIdx).join(' '),
      };
    }
    var n = tokens.length;
    var a = tokens[0], b = tokens[n - 1];
    var aPrenom = isKnownPrenom(a), bPrenom = isKnownPrenom(b);
    var aCaps = isAllCapsWord(a), bCaps = isAllCapsWord(b);
    var surnameFirst = function () { return { nom: tokens[0], prenom: tokens.slice(1).join(' ') }; };
    var surnameLast = function () {
      return { nom: tokens[n - 1], prenom: tokens.slice(0, n - 1).join(' ') };
    };

    /* Exactly one END is a known first name: that end is the given name. */
    if (aPrenom && !bPrenom) return surnameLast();
    if (!aPrenom && bPrenom) return surnameFirst();
    /* Otherwise the ALL-CAPS end is the surname. */
    if (aCaps && !bCaps) return surnameFirst();
    if (!aCaps && bCaps) return surnameLast();
    if (aPrenom && bPrenom && n === 2) {
      /* Two first names ("Jean Pierre"): not a pair we can prove. */
      return null;
    }
    /* Lowercase tokens ("dupont jean"): French NOM Prénom order. */
    var anyLower = /^[a-zà-ÿ]/.test(a) || /^[a-zà-ÿ]/.test(b);
    if (!aPrenom && !bPrenom && anyLower) {
      return { nom: capitalizeWord(tokens[0]), prenom: tokens.slice(1).map(capitalizeWord).join(' ') };
    }
    /* Last resort, only for real Title Case words ("Dupont Jean"): the
     * convention of these files is NOM Prénom. Anything ambiguous (mixed or
     * unknown casing) is left WHOLE by the caller. */
    var titleWords = tokens.filter(function (t) { return /^[A-ZÉÈÀÂÏÔÜÇŒÆ][a-zà-ÿ]/.test(t); });
    if (!aPrenom && !bPrenom && titleWords.length === n) return surnameFirst();
    return null;
  }

  /* ["O","Brien","Marie-Claude"] -> ["O Brien", "Marie-Claude"] */
  function glueParticles(tokens) {
    var out = [];
    for (var i = 0; i < tokens.length; i++) {
      if (isNameParticle(tokens[i]) && i + 1 < tokens.length) {
        out.push(tokens[i] + ' ' + tokens[i + 1]);
        i++;
      } else {
        out.push(tokens[i]);
      }
    }
    return out;
  }

  var PARTICLE_WORDS = /^(o|de|du|des|la|le|les|van|von|di|da|del|della|bin|ben|st|sainte?|mac|mc)['’]?$/i;
  function isNameParticle(w) {
    return PARTICLE_WORDS.test(String(w || '').trim());
  }

  function isAllCapsWord(w) {
    var letters = String(w || '').replace(/[^A-Za-zÀ-ÿ]/g, '');
    return letters.length > 1 && letters === letters.toUpperCase();
  }

  /* A single-token surname keeps its particle: "O Brien" is the surname of
   * "Marie-Claude", never "O". */
  function hasLeadingParticle(raw) {
    var words = trimClean(raw).split(/\s+/);
    return words.length > 1 && PARTICLE_WORDS.test(words[0]);
  }

  function capitalizeWord(w) {
    return String(w || '')
      .toLowerCase()
      .replace(/(^|[\s\-'’])([a-zà-ÿ])/g, function (_m, sep, ch) {
        return sep + ch.toUpperCase();
      });
  }

  function countCommonPrenom(values) {
    var hits = 0;
    values.forEach(function (v) {
      var first = stripAccents(trimClean(v)).toLowerCase().split(/[\s\-]/)[0];
      if (first && SEC.COMMON_PRENOMS.indexOf(first) !== -1) hits++;
    });
    return hits;
  }

  /* The insurer list is compiled into ONE alternation: testing it word by
   * word for every row cost ~130 ms per 20 000 rows. The regex keeps the same
   * "alias contained in the value" semantics. */
  var MUTUELLE_RE = null;
  function looksLikeMutuelle(raw) {
    /* An insurer NAME is short. A 30 000-char note that happens to contain
     * the word "assurance" is not an insurer column. */
    if (raw != null && String(raw).length > 80) return false;
    var s = stripAccents(asString(raw)).toLowerCase();
    if (!s) return false;
    if (!MUTUELLE_RE) {
      var alts = (SEC.KNOWN_MUTUELLES || [])
        .filter(function (k) { return k && k.length >= 3; })
        .sort(function (a, b) { return b.length - a.length; })
        .map(function (k) { return k.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); });
      MUTUELLE_RE = new RegExp("(" + alts.join("|") + ")", "i");
    }
    return MUTUELLE_RE.test(s);
  }

  /* ---------- Packed commentaire / free-text record parsing ---------- */

  /* Real exports often pack several fields into one cell separated by "--",
   * e.g.:
   *   "test happ -- 07*12*2020------2 69 02 13 055 625 36----15*02*1969--48,83--"
   *   "test happ -- 05*12*2020----x@y.fr--1 56 10 77 131 252 29--APIVIA--06*10*1956--38,72--"
   * Each segment is inspected and routed to the right official column, and the
   * remaining free text stays in « Commentaire ». Nothing is dropped. */
  function parsePackedComment(raw) {
    var text = asString(raw);
    var out = { email: '', cotisation: '', mutuelle: '', tel: '', cp: '', date: '', text: '', found: false };
    if (!text || text.indexOf('--') === -1) return out;

    var segments = text.split(/--+/).map(trimClean).filter(Boolean);
    var leftovers = [];

    segments.forEach(function (seg) {
      /* e-mail */
      if (!out.email && looksLikeEmail(seg)) { out.email = seg; out.found = true; return; }
      /* French phone (mobile or landline) */
      if (!out.tel && looksLikePhone(seg) && /[0-9]/.test(seg)) {
        var ph = analyzePhone(seg);
        if (ph.valid === true) { out.tel = ph.value; out.found = true; return; }
      }
      /* postal code */
      if (!out.cp && looksLikePostal(seg)) { out.cp = analyzePostal(seg).value; out.found = true; return; }
      /* known health insurer (must be a real organisation, not free text) */
      if (!out.mutuelle && seg.length <= 45 && looksLikeMutuelle(seg)) { out.mutuelle = seg; out.found = true; return; }
      /* dd*mm*yyyy / dd-mm-yyyy / dd/mm/yyyy date -> not an official column */
      if (/^\d{1,2}[\/*.-]\d{1,2}[\/*.-]\d{2,4}$/.test(seg)) {
        if (!out.date) out.date = seg;
        out.found = true;
        return;
      }
      /* a decimal amount, alone, is the cotisation ("48,83", "135.67") */
      if (!out.cotisation && /^\d[\d\s.]*[.,]\d{1,2}$/.test(seg) && !looksLikePhone(seg)) {
        out.cotisation = seg;
        out.found = true;
        return;
      }
      /* a long digit run that is not a phone / postal code is a social
       * security number or account ref: not an official column */
      var digits = seg.replace(/\D/g, '');
      if (digits.length >= 13 && !looksLikePhone(seg)) { out.found = true; return; }

      /* short status words ("retr") are noise when they occupy their own
       * segment, but longer free text is meaningful -> keep it. */
      if (seg.replace(/[^A-Za-zÀ-ÿ]/g, '').length >= 3 || /\s/.test(seg)) leftovers.push(seg);
    });

    out.text = leftovers.join(' -- ').trim();
    return out;
  }

  /* ---------- Export sanitization ---------- */

  /* Prevent spreadsheet formula injection (OWASP). Prefix dangerous leading
   * characters with an apostrophe so every viewer treats the cell as text.
   * Values merged from several source cells ("a | b | =1+1") are neutralized
   * on EVERY segment, not only the first, because a cell content can be
   * re-split by the viewer or by a later export. */
  function sanitizeExportValue(v) {
    var s = asString(v);
    if (!s) return s;
    /* Fast path with a native search: most values are plain text and are
     * returned untouched (this runs on every cell of the export). */
    if (s.indexOf("=") === -1 && s.indexOf("|") === -1 && s.indexOf("+") === -1 &&
        s.indexOf("-") === -1 && s.indexOf("@") === -1 &&
        s.charAt(0) !== "\t" && s.charAt(0) !== "\r") return s;
    var parts = s.split(" | ");
    if (parts.length > 1) {
      return parts
        .map(function (p) {
          return /^[=+\-@\t\r]/.test(p) ? "'" + p : p;
        })
        .join(" | ");
    }
    return /^[=+\-@\t\r]/.test(s) ? "'" + s : s;
  }

  /* Excel limits ONE cell to 32 767 characters (a hard OOXML rule: SheetJS
   * throws "Text length must not exceed 32767 characters"). A long value is
   * therefore SPLIT into consecutive cells of the SAME column (continuation
   * rows, prefixed "[suite]") so not a single character is lost. Callers get
   * an array of strings; a short value returns a one-element array, so the
   * normal path costs nothing. */
  var EXCEL_MAX_CELL = 32767;
  var SUITE = "[suite] ";

  function safeFit(v) {
    var s = asString(v);
    if (s.length <= EXCEL_MAX_CELL) return [s];
    var parts = [];
    var rest = s;
    while (rest.length > 0) {
      var budget = parts.length === 0 ? EXCEL_MAX_CELL : EXCEL_MAX_CELL - SUITE.length;
      // Prefer a newline / space boundary so no word is broken when possible.
      var cut = Math.min(rest.length, budget);
      if (cut < rest.length) {
        var nl = rest.lastIndexOf("\n", cut);
        var sp = rest.lastIndexOf(" ", cut);
        var boundary = Math.max(nl, sp);
        if (boundary > budget * 0.6) cut = boundary + 1;
      }
      var piece = rest.slice(0, cut);
      parts.push(parts.length === 0 ? piece : SUITE + piece);
      rest = rest.slice(cut);
    }
    return parts;
  }

  /* Characters that a spreadsheet cannot store at all (control chars 0x00-
   * 0x08, 0x0B, 0x0C, 0x0E-0x1F): removed, and only those. Everything
   * printable is kept verbatim. */
  var CONTROL_RE = /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/;
  function stripControlChars(v) {
    var s = asString(v);
    /* Fast path: no control char means NO new string is built (long notes are
     * scanned once, not copied at every step of the export). */
    if (!CONTROL_RE.test(s)) return s;
    return s.replace(new RegExp(CONTROL_RE.source, "g"), "");
  }

  SEC.Normalizers = {
    asString: asString,
    EXCEL_MAX_CELL: EXCEL_MAX_CELL,
    safeFit: safeFit,
    stripControlChars: stripControlChars,
    trimClean: trimClean,
    stripAccents: stripAccents,
    normalizeHeader: normalizeHeader,
    singularize: singularize,
    isBlank: isBlank,
    analyzePhone: analyzePhone,
    looksLikePhone: looksLikePhone,
    isUnusablePhone: isUnusablePhone,
    analyzePostal: analyzePostal,
    looksLikePostal: looksLikePostal,
    analyzeEmail: analyzeEmail,
    looksLikeEmail: looksLikeEmail,
    parseFrenchNumber: parseFrenchNumber,
    looksLikeAmount: looksLikeAmount,
    isConfirmedAmount: isConfirmedAmount,
    isPureAmount: isPureAmount,
    parsePackedRecord: parsePackedRecord,
    isAmbiguousNumber: isAmbiguousNumber,
    analyzeAmount: analyzeAmount,
    extractAmount: extractAmount,
    looksLikeAddress: looksLikeAddress,
    looksLikeCity: looksLikeCity,
    looksLikeCommentText: looksLikeCommentText,
    looksLikeSequenceSamples: looksLikeSequenceSamples,
    looksLikePersonName: looksLikePersonName,
    isInvalidPersonName: isInvalidPersonName,
    isNoteLike: isNoteLike,
    hasLeadingParticle: hasLeadingParticle,
    isKnownPrenom: isKnownPrenom,
    splitFullName: splitFullName,
    splitNamePair: splitNamePair,
    countCommonPrenom: countCommonPrenom,
    looksLikeMutuelle: looksLikeMutuelle,
    looksLikeInsurer: looksLikeInsurer,
    isPureInsurer: isPureInsurer,
    matchInsurer: matchInsurer,
    insurerCount: function () { return Object.keys(insurerDb()).length; },
    looksLikeAge: looksLikeAge,
    ageRatio: ageRatio,
    sanitizeExportValue: sanitizeExportValue,
    parsePackedComment: parsePackedComment,
    isPlaceholder: isPlaceholder,
    repairEmail: repairEmail,
    splitPhoneValues: splitPhoneValues
  };
})(typeof self !== 'undefined' ? self : globalThis);

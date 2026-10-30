/* ============================================================
 * Smart Excel - Configuration & constants
 * Purely local application. No network, no storage of customer data.
 * ============================================================ */
(function (root) {
  'use strict';
  var SEC = (root.SEC = root.SEC || {});

  /* The 11 mandatory official French output headers, in exact order. */
  SEC.OFFICIAL_HEADERS = [
    'Nom',
    'Prénom',
    'Téléphone 1',
    'Téléphone 2',
    'Adresse',
    'Code postal',
    'Ville',
    'Email',
    'Cotisation',
    'Mutuelle',
    'Commentaire'
  ];

  /* Machine keys matching the official headers (index order is authoritative). */
  SEC.TARGET_KEYS = [
    'nom', 'prenom', 'tel1', 'tel2', 'adresse', 'cp',
    'ville', 'email', 'cotisation', 'mutuelle', 'commentaire'
  ];

  SEC.TARGET_LABELS = {
    nom: 'Nom',
    prenom: 'Prénom',
    tel1: 'Téléphone 1',
    tel2: 'Téléphone 2',
    adresse: 'Adresse',
    cp: 'Code postal',
    ville: 'Ville',
    email: 'Email',
    cotisation: 'Cotisation',
    mutuelle: 'Mutuelle',
    commentaire: 'Commentaire'
  };

  /* Processing statuses (shared by UI and tests). */
  SEC.STATUS = {
    WAITING: 'waiting',
    VALIDATING: 'validating',
    PROCESSING: 'processing',
    NEEDS_REVIEW: 'needs_review',
    COMPLETED: 'completed',
    COMPLETED_WARNINGS: 'completed_warnings',
    FAILED: 'failed',
    CANCELLED: 'cancelled'
  };

  /* Safety / performance limits. */
  SEC.LIMITS = {
    MAX_SCAN_HEADER_ROWS: 15,        // rows examined to locate the header row
    MAX_PROFILE_SAMPLE: 2000,        // max values profiled per column
    REPORT_SAMPLE_PER_COLUMN: 3,     // sample values kept in the warning report
    MEMORY_WARN_BYTES: 300 * 1024 * 1024, // aggregate selected size warning
    MEMORY_HARD_HINT_FILES: 150,     // file count at which we show a hint
    PREVIEW_ROWS: 100,
    YIELD_EVERY_ROWS: 1500           // event-loop yielding granularity
  };

  SEC.SUPPORTED_EXTENSIONS = ['csv', 'txt', 'xlsx', 'xlsm', 'xls'];

  /* Human-readable status labels (interface in French; data headers stay French). */
  SEC.STATUS_LABELS = {
    waiting: 'En attente',
    validating: 'Vérification',
    processing: 'Traitement',
    needs_review: 'À vérifier',
    completed: 'Succès',
    completed_warnings: 'Succès',
    failed: 'Échec',
    cancelled: 'Ignoré'
  };
})(typeof self !== 'undefined' ? self : globalThis);

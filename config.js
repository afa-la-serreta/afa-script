// ============================================================================
// AFA La Serreta — Gestió de socis
//
// Documentació completa i codi font:
//   https://github.com/afa-la-serreta/afa-script
//
// Llegeix el README.md per entendre l'arquitectura, el flux d'ús,
// la lògica de deduplicació, graduació, SEPA i enviament de correus.
//
// ⚠️  NO modifiqueu el codi directament a l'editor d'Apps Script!
//     El projecte està versionat amb Git i sincronitzat amb `clasp`.
//     Feu els canvis al repositori i desplegeu amb `clasp push`.
// ============================================================================
const SHEET_RESPONSES = 'Respuestas de formulario 1';
const SHEET_CANON = 'Famílies';
const SHEET_POTENTIAL = 'Possibles duplicats';
const SHEET_BIC = 'BIC Bancs';

// IBANAPI (clau guardada a Script Properties)
const IBANAPI_KEY_PROP = 'IBANAPI_KEY';

// Webapp
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbzqxvri6XcnTZLO5ZKhcVXkLZp96tOYa0wOtCzKUKGaEU7J6HaC4sXlEFEi5ZFhY9CSKQ/exec';

// Email
const EMAIL_SUBJECT = 'AFA La Serreta - Confirmaci\u00f3 de dades';
const EMAIL_SENDER_NAME = 'AFA La Serreta';

// SEPA Direct Debit (pain.008.001.02 / Cuaderno 19.44)
const SEPA_CREDITOR_NAME = 'AFA ESCOLA LA SERRETA';
const SEPA_CREDITOR_ID = 'ES91000G65305732';
const SEPA_CREDITOR_IBAN = 'ES3900810447060001265633';
const SEPA_CREDITOR_BIC = 'BSABESBBXXX';
const SEPA_CREDITOR_STREET = 'NURIA 26-28';
const SEPA_CREDITOR_POSTCODE = '08328';
const SEPA_CREDITOR_TOWN = 'ALELLA';
const SEPA_CREDITOR_COUNTRY = 'ES';
const SEPA_DRIVE_FOLDER_ID = '1sQYYaHHD96FQ6QAdPUnIZEgcYfwX9lW4';

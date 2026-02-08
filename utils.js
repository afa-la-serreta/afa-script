/****************
 * utils.gs
 *
 * Funcions pures d'utilitat: normalització, hashing,
 * lògica de graduació i helpers generals.
 *
 * Usat per: families.gs, dedup.gs, parsing.gs, webapp.gs
 ****************/

/* =========================
   Index de headers
   ========================= */

function headerIndex_(headers) {
  const m = {};
  headers.forEach((h, i) => m[String(h).trim()] = i);
  return m;
}

/* =========================
   Normalització
   ========================= */

function normEmail_(s) { return (s || '').toString().trim().toLowerCase(); }
function normPhone_(s) { return (s || '').toString().replace(/\D+/g, ''); }
function normIban_(s) { return (s || '').toString().toUpperCase().replace(/\s+/g, ''); }

/**
 * Valida que un valor sembli un document d'identitat real (NIF/NIE/passaport).
 * Requereix mínim 7 caràcters i almenys 5 dígits per descartar
 * placeholders (p.ex. "."), noms (p.ex. "Silvia"), etc.
 */
function isValidDni_(s) {
  const n = (s || '').toString().toUpperCase().replace(/[\s\-]/g, '');
  if (n.length < 7) return false;
  const digits = (n.match(/\d/g) || []).length;
  return digits >= 5;
}

function normDni_(s) {
  const n = (s || '').toString().toUpperCase().replace(/[\s\-]/g, '');
  if (!isValidDni_(n)) return '';
  return n;
}

function stripAccents_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function normName_(s) {
  return stripAccents_((s || '').toString().trim().toLowerCase())
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function firstSurname_(s) {
  const t = normName_(s);
  if (!t) return '';
  return t.split(' ')[0] || '';
}

function normAddress_(s) {
  let t = stripAccents_((s || '').toString().trim().toLowerCase());
  t = t.replace(/[,.;]/g, ' ');
  t = t.replace(/\b(pis|porta|escala|bloc)\b/g, ' ');
  t = t.replace(/\s+/g, ' ').trim();
  return t;
}

/* =========================
   Col·leccions
   ========================= */

function uniq_(arr) { return Array.from(new Set(arr)); }

function intersect_(a, b) {
  const setB = new Set(b);
  return a.filter(x => setB.has(x));
}

function jaccardTokens_(a, b) {
  if (!a || !b) return 0;
  const ta = new Set(String(a).split(' ').filter(Boolean));
  const tb = new Set(String(b).split(' ').filter(Boolean));
  if (!ta.size || !tb.size) return 0;
  let inter = 0;
  ta.forEach(x => { if (tb.has(x)) inter++; });
  const union = ta.size + tb.size - inter;
  return union ? inter / union : 0;
}

/* =========================
   IBAN
   ========================= */

/**
 * Emmascara un IBAN mostrant només els últims 4 dígits.
 * Ex: "ES12 3456 7890 1234 5678" → "****5678"
 * Retorna '' si l'IBAN és buit o massa curt.
 */
function maskIban_(iban) {
  const clean = (iban || '').toString().replace(/\s+/g, '');
  if (clean.length < 4) return '';
  return '****' + clean.slice(-4);
}

/* =========================
   BIC: cache i lectura del full
   ========================= */

var _bicCache = null;

/**
 * Llegeix el full "BIC Bancs" i retorna un objecte { codiEntitat: bic }.
 * Fa una sola lectura per execució (cached).
 * Si el full no existeix, retorna {}.
 */
function loadBicMap_() {
  if (_bicCache) return _bicCache;
  _bicCache = {};
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_BIC);
  if (!sh) return _bicCache;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var code = String(data[i][0] || '').trim();
    // Sheets pot treure zeros inicials (0049 → 49); re-padejar a 4 dígits
    if (/^\d+$/.test(code)) {
      while (code.length < 4) code = '0' + code;
    }
    var bic = String(data[i][1] || '').trim().toUpperCase();
    if (code && bic) _bicCache[code] = bic;
  }
  return _bicCache;
}

/* =========================
   BIC: lookup via IBANAPI
   ========================= */

/**
 * Crida l'API d'IBANAPI per obtenir el BIC a partir d'un IBAN complet.
 * Retorna { bic: string, bankName: string } o null si no es pot resoldre.
 *
 * Requereix la clau API a Script Properties (IBANAPI_KEY).
 * Retorna null silenciosament si: no hi ha clau, quota exhaurida, o error.
 */
function lookupBicApi_(iban) {
  var apiKey;
  try {
    apiKey = PropertiesService.getScriptProperties().getProperty(IBANAPI_KEY_PROP);
  } catch (e) {
    Logger.log('lookupBicApi_: no es pot accedir a Script Properties: ' + e.message);
    return null;
  }
  if (!apiKey) {
    Logger.log('lookupBicApi_: no hi ha clau API configurada (Script Properties > ' + IBANAPI_KEY_PROP + ')');
    return null;
  }

  var url = 'https://api.ibanapi.com/v1/validate/' + encodeURIComponent(iban)
          + '?api_key=' + encodeURIComponent(apiKey);
  try {
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    var httpCode = resp.getResponseCode();
    var body = resp.getContentText();
    if (httpCode !== 200) {
      Logger.log('lookupBicApi_(' + iban.substring(0, 8) + '...): HTTP ' + httpCode + ' — ' + body.substring(0, 200));
      return null;
    }
    var json = JSON.parse(body);
    if (!json || json.result !== 200) {
      Logger.log('lookupBicApi_(' + iban.substring(0, 8) + '...): API result=' + (json && json.result) + ' msg=' + (json && json.message));
      return null;
    }
    var bank = (json.data && json.data.bank) || {};
    var bic = (bank.bic || '').trim().toUpperCase();
    var bankName = (bank.bank_name || '').trim();
    if (!bic) {
      Logger.log('lookupBicApi_(' + iban.substring(0, 8) + '...): API ok però BIC buit (bank_name=' + bankName + ')');
      return null;
    }
    Logger.log('lookupBicApi_(' + iban.substring(0, 8) + '...): resolt → ' + bic + ' (' + bankName + ')');
    return { bic: bic, bankName: bankName };
  } catch (e) {
    Logger.log('lookupBicApi_ error: ' + e.message);
    return null;
  }
}

/* =========================
   BIC: escriure al full
   ========================= */

/**
 * Afegeix (o crea) una fila al full "BIC Bancs".
 * Si el full no existeix, el crea amb capçaleres.
 */
function saveBicToSheet_(bankCode, bic, bankName) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_BIC);
  if (!sh) {
    sh = ss.insertSheet(SHEET_BIC);
    sh.getRange(1, 1, 1, 3).setValues([['Codi Entitat', 'BIC', 'Nom Entitat']]);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold');
    // Formatar columna A com a text per preservar zeros inicials (0049)
    sh.getRange('A:A').setNumberFormat('@');
  }
  sh.appendRow([bankCode, bic, bankName || '']);
}

/* =========================
   BIC: derivació IBAN → BIC
   ========================= */

/**
 * Deriva el BIC/SWIFT a partir del codi d'entitat bancària de l'IBAN.
 * Per a IBANs espanyols (ES), el codi d'entitat són les posicions 4-7.
 *
 * 1. Busca al full "BIC Bancs" (cached per execució).
 * 2. Si no el troba, consulta IBANAPI i guarda el resultat al full.
 *
 * Retorna '' si l'IBAN no és espanyol o el BIC no es pot determinar.
 */
function ibanToBic_(iban) {
  var clean = normIban_(iban);
  if (!clean || clean.length < 8) return '';
  if (clean.substring(0, 2) !== 'ES') return '';
  var bankCode = clean.substring(4, 8);

  // 1. Buscar al full (cache)
  var map = loadBicMap_();
  if (map[bankCode]) return map[bankCode];

  // 2. Lookup via API per codis desconeguts
  var result = lookupBicApi_(clean);
  if (result && result.bic) {
    saveBicToSheet_(bankCode, result.bic, result.bankName);
    map[bankCode] = result.bic;
    return result.bic;
  }

  return '';
}

/* =========================
   Hashing
   ========================= */

function sha256_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

/* =========================
   Timestamps
   ========================= */

/**
 * Converteix un timestamp (Date, string o number) a millis per comparació numèrica.
 * Evita el bug de localeCompare amb strings de Date (que ordena per nom del dia).
 */
function toTime_(ts) {
  if (!ts) return 0;
  if (ts instanceof Date) return ts.getTime();
  const n = Number(ts);
  if (!isNaN(n) && n > 0) return n;
  const d = new Date(ts);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

/* =========================
   Graduació (lògica de curs escolar)
   ========================= */

/**
 * Retorna l'any de tall: els infants que acaben 6è al juny d'aquest any
 * o abans ja han marxat.
 *
 * Set-Des → retorna any actual (curs acabat al juny d'enguany)
 * Gen-Ago → retorna any - 1 (últim curs acabat al juny de l'any passat)
 *
 * Ex: feb 2026 → últim curs acabat = 2024-2025 → juny 2025 → retorna 2025.
 */
function graduationCutoffYear_() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth(); // 0-based (0=gen, 5=juny, 8=set)

  if (month >= 8) {
    return year;
  } else {
    return year - 1;
  }
}

/**
 * Extreu l'any de naixement d'un valor DOB (Date, string, o number).
 * Retorna null si no es pot parsejar.
 */
function extractBirthYear_(dob) {
  if (dob instanceof Date) {
    const y = dob.getFullYear();
    return isNaN(y) ? null : y;
  }
  if (typeof dob === 'number') {
    const d = new Date((dob - 25569) * 86400000);
    const y = d.getFullYear();
    return isNaN(y) ? null : y;
  }
  const s = String(dob).trim();
  if (!s) return null;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getFullYear();
  const m = s.match(/\b(19|20)\d{2}\b/);
  return m ? parseInt(m[0]) : null;
}

/**
 * Comprova si una família (fila canònica) hauria de ser INACTIVE
 * perquè tots els seus infants ja han acabat 6è.
 *
 * Retorna true si s'ha de desactivar, false en cas contrari.
 * Famílies sense DOBs o amb DOBs sospitosos retornen false (no desactivar).
 */
function shouldDeactivateFamily_(row, idx) {
  const cutoffYear = graduationCutoffYear_();
  const dobKeys = ['c1_dob', 'c2_dob', 'c3_dob', 'c4_dob'];
  const dobs = dobKeys
    .filter(k => idx[k] != null)
    .map(k => row[idx[k]])
    .filter(v => v != null && String(v).trim() !== '');

  if (dobs.length === 0) return false;

  const years = [];
  for (const dob of dobs) {
    const y = extractBirthYear_(dob);
    if (y == null || y < 1900) return false;
    years.push(y);
  }

  return years.every(y => y + 12 <= cutoffYear);
}

/**
 * Aplica l'status de graduació a una fila canònica.
 * @param {Array} row - fila canònica
 * @param {Object} idxC - índex de columnes
 * @param {string} mode - 'create' o 'update'
 * @returns {boolean} true si la família ha estat desactivada
 */
function applyGraduationStatus_(row, idxC, mode) {
  if (mode === 'create') {
    if (shouldDeactivateFamily_(row, idxC)) {
      row[idxC.status] = 'INACTIVE';
      if (idxC.inactive_reason != null) row[idxC.inactive_reason] = '6è acabat';
      if (idxC.inactive_at != null) row[idxC.inactive_at] = new Date();
      return true;
    }
    row[idxC.status] = 'ACTIVE';
    return false;
  }
  // mode === 'update': només canvia si estava ACTIVE
  if (String(row[idxC.status]).trim() === 'ACTIVE' && shouldDeactivateFamily_(row, idxC)) {
    row[idxC.status] = 'INACTIVE';
    if (idxC.inactive_reason != null) row[idxC.inactive_reason] = '6è acabat';
    if (idxC.inactive_at != null) row[idxC.inactive_at] = new Date();
    return true;
  }
  return false;
}

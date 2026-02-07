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

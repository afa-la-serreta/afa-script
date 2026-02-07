/****************
 * deduplicar.gs
 *
 * Requisits:
 * - Tens un config.gs amb:
 *   const SHEET_RESPONSES = 'Respuestas de formulario 1';
 *   const SHEET_CANON = 'Famílies'; // o 'Families' segons el teu nom real
 *   const SHEET_POTENTIAL = 'Possibles duplicats';
 *
 * - A SHEET_CANON (Famílies) tens aquests headers (incloent traça):
 *   ... ,source_rows,source_count,dedup_reasons
 *
 * - A SHEET_POTENTIAL tens els headers que escriu findPotentialDuplicatesAll()
 ****************/

/* =========================
   Entry points
   ========================= */

/**
 * Import històric:
 * - Llegeix totes les respostes
 * - Agrupa per claus fortes (email/tel/iban/dni) via Union-Find
 * - Escriu/actualitza 1 fila canònica per cluster a "Famílies"
 * - Escriu traça: source_rows/source_count/dedup_reasons
 */
function importAll() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RESPONSES);
  const canon = ss.getSheetByName(SHEET_CANON);

  if (!sh) throw new Error(`No trobo la pestanya: ${SHEET_RESPONSES}`);
  if (!canon) throw new Error(`No trobo la pestanya: ${SHEET_CANON}`);

  const respValues = sh.getDataRange().getValues();
  if (respValues.length < 2) return;

  // Headers únics (perquè a Respuestas hi ha headers duplicats)
  const rawHeaders = respValues[0].map(h => (h == null ? '' : String(h)));
  const headers = makeUniqueHeaders_(rawHeaders);

  const parsedRows = respValues.slice(1).map((r, i) => parseResponseRow_(headers, r, i + 2));
  const clusters = clusterByStrongKeys_(parsedRows);

  const canonValues = canon.getDataRange().getValues();
  if (canonValues.length < 1) throw new Error(`${SHEET_CANON} no té headers a la fila 1.`);
  const canonHeaders = canonValues[0].map(h => String(h).trim());
  const idxC = headerIndex_(canonHeaders);

  // Validació columnes mínimes (incloent traça)
  const requiredCanonCols = [
    'family_id','token_edit','status','created_at','updated_at','source_last_timestamp',
    'g1_nom','g1_cognoms','g1_email','g1_telefon',
    'bank_iban',
    'source_rows','source_count','dedup_reasons'
  ];
  requiredCanonCols.forEach(c => {
    if (idxC[c] == null) throw new Error(`Falta la columna ${c} a ${SHEET_CANON}`);
  });

  // Índex de Families existent per match fort (email/tel/iban)
  // Nota: DNI incremental el fem per hash en el futur; ara el clustering històric ja usa DNI en clar.
  const existingIndex = buildCanonStrongIndex_(canonValues, idxC);

  let created = 0;
  let updated = 0;

  for (const cluster of clusters) {
    const clusterKeys = strongKeysForCluster_(cluster);

    // Troba si aquest cluster ja correspon a una fila canònica existent
    let canonRowIdx = null; // 0-based dins canonValues
    for (const k of clusterKeys) {
      if (existingIndex.has(k)) {
        canonRowIdx = existingIndex.get(k);
        break;
      }
    }

    if (canonRowIdx == null) {
      // CREATE
      const newRow = new Array(canonHeaders.length).fill('');
      newRow[idxC.family_id] = Utilities.getUuid();
      newRow[idxC.token_edit] = Utilities.getUuid() + Utilities.getUuid();
      newRow[idxC.status] = 'ACTIVE';
      newRow[idxC.created_at] = new Date();
      newRow[idxC.updated_at] = new Date();

      const merged = mergeClusterToCanonRow_(cluster, newRow, idxC);

      // Traça
      const tr = clusterTrace_(cluster);
      setCanon_(merged, idxC, 'source_rows', tr.source_rows, true);
      setCanon_(merged, idxC, 'source_count', tr.source_count, true);
      setCanon_(merged, idxC, 'dedup_reasons', tr.dedup_reasons, true);

      canon.appendRow(merged);
      created++;

      // actualitza índex (aprox)
      addCanonRowToIndex_(existingIndex, merged, idxC, canonValues.length);
    } else {
      // UPDATE
      const current = canonValues[canonRowIdx].slice();
      const merged = mergeClusterToCanonRow_(cluster, current, idxC);

      // Traça (força update)
      const tr = clusterTrace_(cluster);
      setCanon_(merged, idxC, 'source_rows', tr.source_rows, true);
      setCanon_(merged, idxC, 'source_count', tr.source_count, true);
      setCanon_(merged, idxC, 'dedup_reasons', tr.dedup_reasons, true);

      merged[idxC.updated_at] = new Date();

      canon.getRange(canonRowIdx + 1, 1, 1, canonHeaders.length).setValues([merged]);
      updated++;

      addCanonRowToIndex_(existingIndex, merged, idxC, canonRowIdx);
    }
  }

  Logger.log(`importAll: clusters=${clusters.length}, created=${created}, updated=${updated}`);
}

/**
 * Genera la pestanya "Possibles duplicats"
 * - Compara clusters diferents
 * - Escriu només suggeriments (no fa merges)
 */
function findPotentialDuplicatesAll() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RESPONSES);
  if (!sh) throw new Error(`No trobo la pestanya: ${SHEET_RESPONSES}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  const rawHeaders = values[0].map(h => (h == null ? '' : String(h)));
  const headers = makeUniqueHeaders_(rawHeaders);

  const rows = values.slice(1).map((r, i) => parseResponseRow_(headers, r, i + 2));
  const clusters = clusterByStrongKeys_(rows);

  const out = [];
  for (let i = 0; i < clusters.length; i++) {
    for (let j = i + 1; j < clusters.length; j++) {
      const best = bestPairScoreBetweenClusters_(clusters[i], clusters[j]);
      if (best && best.score >= 70) out.push(best); // llindar (ajusta si vols)
    }
  }

  out.sort((a, b) => b.score - a.score);

  const pot = ss.getSheetByName(SHEET_POTENTIAL) || ss.insertSheet(SHEET_POTENTIAL);
  pot.clearContents();

  const headersOut = [
    'score','reason','rowA','rowB',
    'emailA','emailB','phoneA','phoneB','ibanA','ibanB',
    'guardianSurnamesA','guardianSurnamesB',
    'childSurnamesA','childSurnamesB',
    'addressA','addressB',
    'timestampA','timestampB',
    'strongMatch',
    'decision','winner','decision_by','decision_at','notes'
  ];
  pot.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);

  if (out.length) {
    const rowsOut = out.map(x => ([
      x.score, x.reason, x.a.sheetRow, x.b.sheetRow,
      x.a.emails.join(' | '), x.b.emails.join(' | '),
      x.a.phones.join(' | '), x.b.phones.join(' | '),
      x.a.ibans.join(' | '), x.b.ibans.join(' | '),
      x.a.guardianSurnames.join(' | '), x.b.guardianSurnames.join(' | '),
      x.a.childSurnames.join(' | '), x.b.childSurnames.join(' | '),
      x.a.addressNorm, x.b.addressNorm,
      x.a.timestamp || '', x.b.timestamp || '',
      x.strongMatch ? 'YES' : '',
      '', '', '', '', ''
    ]));

    pot.getRange(2, 1, rowsOut.length, headersOut.length).setValues(rowsOut);
  }
}

/**
 * Trigger instal·lable: From spreadsheet -> On form submit
 * Processa només la nova resposta i actualitza "Famílies" + traça.
 */
function onFormSubmit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  if (!sh || sh.getName() !== SHEET_RESPONSES) return;

  const rowNumber = e.range.getRow();
  if (rowNumber <= 1) return;

  const lastCol = sh.getLastColumn();

  const rawHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h == null ? '' : String(h)));
  const headers = makeUniqueHeaders_(rawHeaders);

  const row = sh.getRange(rowNumber, 1, 1, lastCol).getValues()[0];
  const parsed = parseResponseRow_(headers, row, rowNumber);

  upsertSingleResponse_(parsed);
}

/* =========================
   Incremental upsert
   ========================= */

function upsertSingleResponse_(parsed) {
  const ss = SpreadsheetApp.getActive();
  const canon = ss.getSheetByName(SHEET_CANON);
  if (!canon) throw new Error(`No trobo la pestanya: ${SHEET_CANON}`);

  const canonValues = canon.getDataRange().getValues();
  if (canonValues.length < 2) throw new Error(`${SHEET_CANON} no té dades (mínim headers + 1 fila).`);

  const canonHeaders = canonValues[0].map(h => String(h).trim());
  const idxC = headerIndex_(canonHeaders);

  ['family_id','token_edit','status','created_at','updated_at','source_last_timestamp','source_rows','source_count','dedup_reasons']
    .forEach(c => { if (idxC[c] == null) throw new Error(`Falta la columna ${c} a ${SHEET_CANON}`); });

  // claus fortes del nou registre
  const keys = [];
  parsed.emails.forEach(e => keys.push({k:'email', v:e}));
  parsed.phones.forEach(p => keys.push({k:'phone', v:p}));
  parsed.ibans.forEach(i => keys.push({k:'iban', v:i}));

  // busca match fort a Families
  let hit = -1;
  for (let r = 1; r < canonValues.length && hit === -1; r++) {
    const row = canonValues[r];

    const emails = [
      normEmail_(row[idxC.g1_email]),
      normEmail_(row[idxC.g2_email]),
      normEmail_(row[idxC.email_alternatiu]),
    ].filter(Boolean);

    const phones = [
      normPhone_(row[idxC.g1_telefon]),
      normPhone_(row[idxC.g2_telefon]),
    ].filter(Boolean);

    const iban = normIban_(row[idxC.bank_iban]);

    for (const kk of keys) {
      if (kk.k === 'email' && emails.includes(kk.v)) { hit = r; break; }
      if (kk.k === 'phone' && phones.includes(kk.v)) { hit = r; break; }
      if (kk.k === 'iban' && iban && iban === kk.v) { hit = r; break; }
    }
  }

  if (hit === -1) {
    // CREATE
    const newRow = new Array(canonHeaders.length).fill('');
    newRow[idxC.family_id] = Utilities.getUuid();
    newRow[idxC.token_edit] = Utilities.getUuid() + Utilities.getUuid();
    newRow[idxC.status] = 'ACTIVE';
    newRow[idxC.created_at] = new Date();
    newRow[idxC.updated_at] = new Date();

    const merged = mergeClusterToCanonRow_(
      [{ raw: parsed.raw, timestamp: parsed.timestamp, sheetRow: parsed.sheetRow }],
      newRow,
      idxC
    );

    // Traça
    merged[idxC.source_rows] = String(parsed.sheetRow);
    merged[idxC.source_count] = '1';
    merged[idxC.dedup_reasons] = ''; // no hi ha "duplicate" encara

    canon.appendRow(merged);
    return;
  }

  // UPDATE existent
  const rowIndex1 = hit + 1;
  const current = canonValues[hit].slice();

  const merged = mergeClusterToCanonRow_(
    [{ raw: parsed.raw, timestamp: parsed.timestamp, sheetRow: parsed.sheetRow }],
    current,
    idxC
  );

  // Traça: afegeix fila si no hi és
  const existing = (merged[idxC.source_rows] || '').toString().trim();
  const parts = existing ? existing.split('|') : [];
  const sr = String(parsed.sheetRow);
  if (!parts.includes(sr)) parts.push(sr);
  parts.sort((a,b)=>Number(a)-Number(b));

  merged[idxC.source_rows] = parts.join('|');
  merged[idxC.source_count] = String(parts.length);

  // dedup_reasons: marca la clau que ha fet match (aprox)
  // (no perfecte, però útil per auditar)
  merged[idxC.dedup_reasons] = merged[idxC.dedup_reasons] || '';

  merged[idxC.updated_at] = new Date();

  canon.getRange(rowIndex1, 1, 1, canonHeaders.length).setValues([merged]);
}

/* =========================
   Traça (audit)
   ========================= */

function clusterTrace_(cluster) {
  const sourceRows = cluster.map(x => x.sheetRow).filter(Boolean).sort((a,b)=>a-b);
  const reasons = [];
  if (hasAnyDuplicateKey_(cluster, 'emails')) reasons.push('email');
  if (hasAnyDuplicateKey_(cluster, 'phones')) reasons.push('phone');
  if (hasAnyDuplicateKey_(cluster, 'ibans')) reasons.push('iban');
  if (hasAnyDuplicateKey_(cluster, 'dnis')) reasons.push('dni');

  return {
    source_rows: sourceRows.join('|'),
    source_count: String(sourceRows.length),
    dedup_reasons: reasons.join('+')
  };
}

function hasAnyDuplicateKey_(cluster, field) {
  const seen = new Set();
  for (const r of cluster) {
    for (const v of (r[field] || [])) {
      if (!v) continue;
      if (seen.has(v)) return true;
      seen.add(v);
    }
  }
  return false;
}

/* =========================
   Parsing de respostes
   ========================= */

function parseResponseRow_(headers, row, sheetRowNumber) {
  const o = {};
  for (let i = 0; i < headers.length; i++) {
    o[headers[i]] = row[i];
  }

  const emails = uniq_([
    normEmail_(o['Correu electrònic:']),
    normEmail_(o['Correu electrònic:.1']),
    normEmail_(o['Correu alternatiu:']),
  ].filter(Boolean));

  const phones = uniq_([
    normPhone_(o['Telèfon:']),
    normPhone_(o['Telèfon:.1']),
  ].filter(Boolean));

  const ibans = uniq_([
    normIban_(o['Número de compte IBAN']),
  ].filter(Boolean));

  const dnis = uniq_([
    normDni_(o['DNI:']),
    normDni_(o['DNI:.1']),
    normDni_(o['DNI Titular compte:']),
  ].filter(Boolean));

  const guardianSurnames = uniq_([
    firstSurname_(o['Cognoms:']),
    firstSurname_(o['Cognoms:.1']),
  ].filter(Boolean));

  const childSurnames = uniq_([
    firstSurname_(o['Cognoms:.2']),
    firstSurname_(o['Cognoms:.3']),
    firstSurname_(o['Cognoms:.4']),
    firstSurname_(o['Cognoms:.5']),
  ].filter(Boolean));

  return {
    sheetRow: sheetRowNumber,
    timestamp: o['Marca temporal'] ? String(o['Marca temporal']) : '',
    raw: o,

    emails,
    phones,
    ibans,
    dnis,

    guardianSurnames,
    childSurnames,
    addressNorm: normAddress_(o['Adreça:'] || ''),
    holderName: normName_(o['Nom i cognoms del titular compte'] || ''),
  };
}

/**
 * Converteix headers duplicats en headers únics:
 * "Nom:" -> "Nom:", "Nom:.1", "Nom:.2", ...
 */
function makeUniqueHeaders_(headers) {
  const counts = {};
  return headers.map((h, i) => {
    let key = (h == null) ? '' : String(h).trim();
    if (!key) key = `__EMPTY_${i}`;

    if (counts[key] == null) {
      counts[key] = 0;
      return key;
    } else {
      counts[key] += 1;
      return `${key}.${counts[key]}`;
    }
  });
}

/* =========================
   Clustering per claus fortes (Union-Find)
   ========================= */

function clusterByStrongKeys_(rows) {
  const parent = rows.map((_, i) => i);
  const find = (x) => parent[x] === x ? x : (parent[x] = find(parent[x]));
  const union = (a, b) => { a = find(a); b = find(b); if (a !== b) parent[b] = a; };

  const emailMap = new Map();
  const phoneMap = new Map();
  const ibanMap = new Map();
  const dniMap = new Map();

  rows.forEach((r, i) => {
    r.emails.forEach(e => { if (emailMap.has(e)) union(i, emailMap.get(e)); else emailMap.set(e, i); });
    r.phones.forEach(p => { if (phoneMap.has(p)) union(i, phoneMap.get(p)); else phoneMap.set(p, i); });
    r.ibans.forEach(x => { if (ibanMap.has(x)) union(i, ibanMap.get(x)); else ibanMap.set(x, i); });
    r.dnis.forEach(d => { if (dniMap.has(d)) union(i, dniMap.get(d)); else dniMap.set(d, i); });
  });

  const groups = new Map();
  rows.forEach((r, i) => {
    const root = find(i);
    if (!groups.has(root)) groups.set(root, []);
    groups.get(root).push(r);
  });

  const clusters = Array.from(groups.values());
  // ordena per timestamp desc dins cluster
  clusters.forEach(c => c.sort((a, b) => (b.timestamp || '').localeCompare(a.timestamp || '')));
  return clusters;
}

function strongKeysForCluster_(cluster) {
  const keys = [];
  cluster.forEach(r => {
    r.emails.forEach(e => keys.push('email:' + e));
    r.phones.forEach(p => keys.push('phone:' + p));
    r.ibans.forEach(i => keys.push('iban:' + i));
    r.dnis.forEach(d => keys.push('dni:' + d));
  });
  return uniq_(keys);
}

/* =========================
   Merge cluster -> fila canònica
   ========================= */

function mergeClusterToCanonRow_(cluster, canonRow, idxC) {
  // Base = resposta més nova del cluster
  const newest = cluster[0].raw;

  // timestamp d'origen (per auditar)
  setCanon_(canonRow, idxC, 'source_last_timestamp', newest['Marca temporal'], true);

  // Tutor 1
  setCanon_(canonRow, idxC, 'g1_nom', newest['Nom:']);
  setCanon_(canonRow, idxC, 'g1_cognoms', newest['Cognoms:']);
  setCanon_(canonRow, idxC, 'g1_adreca', newest['Adreça:']);
  setCanon_(canonRow, idxC, 'g1_poblacio', newest['Població:']);
  setCanon_(canonRow, idxC, 'g1_cp', newest['Codi Postal:']);
  setCanon_(canonRow, idxC, 'g1_email', newest['Correu electrònic:']);
  setCanon_(canonRow, idxC, 'g1_telefon', newest['Telèfon:']);
  setCanon_(canonRow, idxC, 'g1_vincle', newest["Vincle amb l'infant:"]);

  const dni1 = (newest['DNI:'] || '').toString().trim();
  if (dni1) setCanon_(canonRow, idxC, 'g1_dni_hash', sha256_(dni1));

  // Tutor 2
  setCanon_(canonRow, idxC, 'g2_nom', newest['Nom:.1']);
  setCanon_(canonRow, idxC, 'g2_cognoms', newest['Cognoms:.1']);
  setCanon_(canonRow, idxC, 'g2_adreca', newest['Adreça:.1']);
  setCanon_(canonRow, idxC, 'g2_poblacio', newest['Població:.1']);
  setCanon_(canonRow, idxC, 'g2_cp', newest['Codi Postal:.1']);
  setCanon_(canonRow, idxC, 'g2_email', newest['Correu electrònic:.1']);
  setCanon_(canonRow, idxC, 'g2_telefon', newest['Telèfon:.1']);
  setCanon_(canonRow, idxC, 'g2_vincle', newest["Vincle amb l'infant:.1"]);

  const dni2 = (newest['DNI:.1'] || '').toString().trim();
  if (dni2) setCanon_(canonRow, idxC, 'g2_dni_hash', sha256_(dni2));

  // Correu alternatiu
  setCanon_(canonRow, idxC, 'email_alternatiu', newest['Correu alternatiu:']);

  // Infants (1..4)
  setCanon_(canonRow, idxC, 'c1_nom', newest['Nom:.2']);
  setCanon_(canonRow, idxC, 'c1_cognoms', newest['Cognoms:.2']);
  setCanon_(canonRow, idxC, 'c1_dob', newest['Data de naixement']);

  setCanon_(canonRow, idxC, 'c2_nom', newest['Nom:.3']);
  setCanon_(canonRow, idxC, 'c2_cognoms', newest['Cognoms:.3']);
  setCanon_(canonRow, idxC, 'c2_dob', newest['Data de naixement.1']);

  setCanon_(canonRow, idxC, 'c3_nom', newest['Nom:.4']);
  setCanon_(canonRow, idxC, 'c3_cognoms', newest['Cognoms:.4']);
  setCanon_(canonRow, idxC, 'c3_dob', newest['Data de naixement.2']);

  setCanon_(canonRow, idxC, 'c4_nom', newest['Nom:.5']);
  setCanon_(canonRow, idxC, 'c4_cognoms', newest['Cognoms:.5']);
  setCanon_(canonRow, idxC, 'c4_dob', newest['Data de naixement.3']);

  // Banc
  setCanon_(canonRow, idxC, 'bank_swift', newest['Entitat Bancària o SWIFT (Pot contenir 8 o 11 posicions):']);
  setCanon_(canonRow, idxC, 'bank_iban', newest['Número de compte IBAN']);
  const holderDni = (newest['DNI Titular compte:'] || '').toString().trim();
  if (holderDni) setCanon_(canonRow, idxC, 'bank_holder_dni_hash', sha256_(holderDni));
  setCanon_(canonRow, idxC, 'bank_holder_name', newest['Nom i cognoms del titular compte']);

  // Consentiments (si algun header canvia al futur, ho fem més robust després)
  setCanon_(canonRow, idxC, 'consent_a', newest["a) Que la imatge del meu/s fill/s o filla/es pugui/n aparèixer en fotografies o vídeos corresponents a les activitats extraescolars o de l'Associació organitzades per l'AFA La Serreta i publicades al blog o a la pàgina web, així com en material de difusió (presentacions, impersos, díptics, cartells, etc.)"]);
  setCanon_(canonRow, idxC, 'consent_b', newest["b) Que la imatge del meu/s fill/a pugui aparèixer en el calendari anual (o suport similar amb presència de fotografies o vídoes) realitzat per les famílies de 6è per recaptar fons pel viatge de fi de curs."]);
  setCanon_(canonRow, idxC, 'consent_c', newest["c) Autorizo a l'ús del telèfon móbil amb la finalitat d'estar en els grups de missatgeria instantània com whatsapp o similars. "]);

  // Omple buits des d'altres respostes del cluster
  for (let k = 1; k < cluster.length; k++) {
    const o = cluster[k].raw;

    setCanon_(canonRow, idxC, 'g1_email', o['Correu electrònic:']);
    setCanon_(canonRow, idxC, 'g1_telefon', o['Telèfon:']);
    setCanon_(canonRow, idxC, 'g2_email', o['Correu electrònic:.1']);
    setCanon_(canonRow, idxC, 'g2_telefon', o['Telèfon:.1']);
    setCanon_(canonRow, idxC, 'email_alternatiu', o['Correu alternatiu:']);

    // si algun dni no era al newest
    const d1 = (o['DNI:'] || '').toString().trim();
    if (d1) setCanon_(canonRow, idxC, 'g1_dni_hash', sha256_(d1));
    const d2 = (o['DNI:.1'] || '').toString().trim();
    if (d2) setCanon_(canonRow, idxC, 'g2_dni_hash', sha256_(d2));
  }

  return canonRow;
}

function setCanon_(row, idxC, key, value, overwriteEvenIfFilled) {
  const i = idxC[key];
  if (i == null) return;
  const v = (value == null) ? '' : String(value).trim();
  if (!v) return;
  if (overwriteEvenIfFilled || !row[i]) row[i] = v;
}

/* =========================
   Families index (match fort)
   ========================= */

function buildCanonStrongIndex_(canonValues, idxC) {
  const m = new Map(); // key -> rowIndex (0-based)
  for (let r = 1; r < canonValues.length; r++) {
    addCanonRowToIndex_(m, canonValues[r], idxC, r);
  }
  return m;
}

function addCanonRowToIndex_(map, canonRow, idxC, rowIndex0) {
  const email1 = normEmail_(canonRow[idxC.g1_email]);
  const email2 = normEmail_(canonRow[idxC.g2_email]);
  const emailAlt = normEmail_(canonRow[idxC.email_alternatiu]);

  const phone1 = normPhone_(canonRow[idxC.g1_telefon]);
  const phone2 = normPhone_(canonRow[idxC.g2_telefon]);

  const iban = normIban_(canonRow[idxC.bank_iban]);

  if (email1) map.set('email:' + email1, rowIndex0);
  if (email2) map.set('email:' + email2, rowIndex0);
  if (emailAlt) map.set('email:' + emailAlt, rowIndex0);

  if (phone1) map.set('phone:' + phone1, rowIndex0);
  if (phone2) map.set('phone:' + phone2, rowIndex0);

  if (iban) map.set('iban:' + iban, rowIndex0);
}

/* =========================
   Scoring “possible duplicate”
   ========================= */

function bestPairScoreBetweenClusters_(A, B) {
  let best = null;
  for (const a of A) {
    for (const b of B) {
      const s = scorePair_(a, b);
      if (!best || s.score > best.score) best = { ...s, a, b };
    }
  }
  return best;
}

function scorePair_(a, b) {
  let score = 0;
  const reasons = [];
  let strongMatch = false;

  const emailCommon = intersect_(a.emails, b.emails);
  if (emailCommon.length) { score += 70; reasons.push('email'); strongMatch = true; }

  const phoneCommon = intersect_(a.phones, b.phones);
  if (phoneCommon.length) { score += 60; reasons.push('phone'); strongMatch = true; }

  const ibanCommon = intersect_(a.ibans, b.ibans);
  if (ibanCommon.length) { score += 90; reasons.push('iban'); strongMatch = true; }

  const dniCommon = intersect_(a.dnis, b.dnis);
  if (dniCommon.length) { score += 90; reasons.push('dni'); strongMatch = true; }

  // Conflictes forts
  if (a.ibans.length && b.ibans.length && !ibanCommon.length) { score -= 80; reasons.push('iban_conflict'); }
  if (a.emails.length && b.emails.length && !emailCommon.length) { score -= 40; reasons.push('email_conflict'); }
  if (a.phones.length && b.phones.length && !phoneCommon.length) { score -= 30; reasons.push('phone_conflict'); }

  // Cognoms tutors
  const gCommon = intersect_(a.guardianSurnames, b.guardianSurnames);
  if (gCommon.length) { score += 25; reasons.push('guardian_surname'); }

  // Cognoms infants
  const cCommon = intersect_(a.childSurnames, b.childSurnames);
  if (cCommon.length) { score += 25; reasons.push('child_surname'); }

  // Adreça
  const addr = jaccardTokens_(a.addressNorm, b.addressNorm);
  if (addr >= 0.85) { score += 30; reasons.push('address_high'); }
  else if (addr >= 0.65) { score += 15; reasons.push('address_mid'); }

  // Titular compte
  if (a.holderName && b.holderName) {
    const hn = jaccardTokens_(a.holderName, b.holderName);
    if (hn >= 0.85) { score += 20; reasons.push('holder_high'); }
    else if (hn >= 0.65) { score += 10; reasons.push('holder_mid'); }
  }

  if (score < 0) score = 0;
  return { score, reason: reasons.join('+'), strongMatch };
}

/* =========================
   Helpers (normalització + utilitats)
   ========================= */

function headerIndex_(headers) {
  const m = {};
  headers.forEach((h, i) => m[String(h).trim()] = i);
  return m;
}

function normEmail_(s) { return (s || '').toString().trim().toLowerCase(); }
function normPhone_(s) { return (s || '').toString().replace(/\D+/g, ''); }
function normIban_(s) { return (s || '').toString().toUpperCase().replace(/\s+/g, ''); }
function normDni_(s) { return (s || '').toString().toUpperCase().replace(/\s+/g, ''); }

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

function sha256_(s) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return raw.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

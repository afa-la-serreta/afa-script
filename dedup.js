/****************
 * dedup.gs
 *
 * Lògica de deduplicació: clustering per claus fortes (Union-Find),
 * merge de clusters a files canòniques, scoring de possibles duplicats,
 * i upsert incremental.
 *
 * Depèn de: utils.gs, parsing.gs
 ****************/

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
  clusters.forEach(c => c.sort((a, b) => toTime_(b.timestamp) - toTime_(a.timestamp)));
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

function mergeClusterToCanonRow_(cluster, canonRow, idxC, overwrite) {
  // Base = resposta més nova del cluster
  // overwrite = true → sobreescriu camps existents (per a upsert d'edicions)
  var ow = !!overwrite;
  const newest = cluster[0].raw;

  // timestamp d'origen (per auditar)
  setCanon_(canonRow, idxC, 'source_last_timestamp', newest['Marca temporal'], true);

  // Tutor 1
  setCanon_(canonRow, idxC, 'g1_nom', newest['Nom:'], ow);
  setCanon_(canonRow, idxC, 'g1_cognoms', newest['Cognoms:'], ow);
  setCanon_(canonRow, idxC, 'g1_adreca', newest['Adreça:'], ow);
  setCanon_(canonRow, idxC, 'g1_poblacio', newest['Població:'], ow);
  setCanon_(canonRow, idxC, 'g1_cp', newest['Codi Postal:'], ow);
  setCanon_(canonRow, idxC, 'g1_email', newest['Correu electrònic:'], ow);
  setCanon_(canonRow, idxC, 'g1_telefon', newest['Telèfon:'], ow);
  setCanon_(canonRow, idxC, 'g1_vincle', newest["Vincle amb l'infant:"], ow);

  const dni1 = (newest['DNI:'] || '').toString().trim();
  if (isValidDni_(dni1)) setCanon_(canonRow, idxC, 'g1_dni_hash', sha256_(dni1), ow);

  // Tutor 2
  setCanon_(canonRow, idxC, 'g2_nom', newest['Nom:.1'], ow);
  setCanon_(canonRow, idxC, 'g2_cognoms', newest['Cognoms:.1'], ow);
  setCanon_(canonRow, idxC, 'g2_adreca', newest['Adreça:.1'], ow);
  setCanon_(canonRow, idxC, 'g2_poblacio', newest['Població:.1'], ow);
  setCanon_(canonRow, idxC, 'g2_cp', newest['Codi Postal:.1'], ow);
  setCanon_(canonRow, idxC, 'g2_email', newest['Correu electrònic:.1'], ow);
  setCanon_(canonRow, idxC, 'g2_telefon', newest['Telèfon:.1'], ow);
  setCanon_(canonRow, idxC, 'g2_vincle', newest["Vincle amb l'infant:.1"], ow);

  const dni2 = (newest['DNI:.1'] || '').toString().trim();
  if (isValidDni_(dni2)) setCanon_(canonRow, idxC, 'g2_dni_hash', sha256_(dni2), ow);

  // Correu alternatiu
  setCanon_(canonRow, idxC, 'email_alternatiu', newest['Correu alternatiu:'], ow);

  // Infants (1..4)
  setCanon_(canonRow, idxC, 'c1_nom', newest['Nom:.2'], ow);
  setCanon_(canonRow, idxC, 'c1_cognoms', newest['Cognoms:.2'], ow);
  setCanon_(canonRow, idxC, 'c1_dob', newest['Data de naixement'], ow);

  setCanon_(canonRow, idxC, 'c2_nom', newest['Nom:.3'], ow);
  setCanon_(canonRow, idxC, 'c2_cognoms', newest['Cognoms:.3'], ow);
  setCanon_(canonRow, idxC, 'c2_dob', newest['Data de naixement.1'], ow);

  setCanon_(canonRow, idxC, 'c3_nom', newest['Nom:.4'], ow);
  setCanon_(canonRow, idxC, 'c3_cognoms', newest['Cognoms:.4'], ow);
  setCanon_(canonRow, idxC, 'c3_dob', newest['Data de naixement.2'], ow);

  setCanon_(canonRow, idxC, 'c4_nom', newest['Nom:.5'], ow);
  setCanon_(canonRow, idxC, 'c4_cognoms', newest['Cognoms:.5'], ow);
  setCanon_(canonRow, idxC, 'c4_dob', newest['Data de naixement.3'], ow);

  // Banc
  setCanon_(canonRow, idxC, 'bank_swift', newest['Entitat Bancària o SWIFT (Pot contenir 8 o 11 posicions):'], ow);
  setCanon_(canonRow, idxC, 'bank_iban', newest['Número de compte IBAN'], ow);
  const holderDni = (newest['DNI Titular compte:'] || '').toString().trim();
  if (isValidDni_(holderDni)) setCanon_(canonRow, idxC, 'bank_holder_dni_hash', sha256_(holderDni), ow);
  setCanon_(canonRow, idxC, 'bank_holder_name', newest['Nom i cognoms del titular compte'], ow);

  // Consentiments
  setCanon_(canonRow, idxC, 'consent_a', newest["a) Que la imatge del meu/s fill/s o filla/es pugui/n aparèixer en fotografies o vídeos corresponents a les activitats extraescolars o de l'Associació organitzades per l'AFA La Serreta i publicades al blog o a la pàgina web, així com en material de difusió (presentacions, impersos, díptics, cartells, etc.)"], ow);
  setCanon_(canonRow, idxC, 'consent_b', newest["b) Que la imatge del meu/s fill/a pugui aparèixer en el calendari anual (o suport similar amb presència de fotografies o vídoes) realitzat per les famílies de 6è per recaptar fons pel viatge de fi de curs."], ow);
  setCanon_(canonRow, idxC, 'consent_c', newest["c) Autorizo a l'ús del telèfon móbil amb la finalitat d'estar en els grups de missatgeria instantània com whatsapp o similars."], ow);

  // Omple buits des d'altres respostes del cluster
  for (let k = 1; k < cluster.length; k++) {
    const o = cluster[k].raw;

    setCanon_(canonRow, idxC, 'g1_email', o['Correu electrònic:']);
    setCanon_(canonRow, idxC, 'g1_telefon', o['Telèfon:']);
    setCanon_(canonRow, idxC, 'g2_email', o['Correu electrònic:.1']);
    setCanon_(canonRow, idxC, 'g2_telefon', o['Telèfon:.1']);
    setCanon_(canonRow, idxC, 'email_alternatiu', o['Correu alternatiu:']);

    const d1 = (o['DNI:'] || '').toString().trim();
    if (isValidDni_(d1)) setCanon_(canonRow, idxC, 'g1_dni_hash', sha256_(d1));
    const d2 = (o['DNI:.1'] || '').toString().trim();
    if (isValidDni_(d2)) setCanon_(canonRow, idxC, 'g2_dni_hash', sha256_(d2));
  }

  // Auto-derivar BIC des de l'IBAN (les dades del formulari són poc fiables)
  if (idxC.bank_swift != null && idxC.bank_iban != null) {
    var derivedBic = ibanToBic_(canonRow[idxC.bank_iban]);
    if (derivedBic) canonRow[idxC.bank_swift] = derivedBic;
  }

  return canonRow;
}

function setCanon_(row, idxC, key, value, overwriteEvenIfFilled) {
  const i = idxC[key];
  if (i == null) return;
  // Preserva Date sense convertir a string
  if (value instanceof Date) {
    if (overwriteEvenIfFilled || !row[i]) row[i] = value;
    return;
  }
  const v = (value == null) ? '' : String(value).trim();
  if (!v) return;
  if (overwriteEvenIfFilled || !row[i]) row[i] = v;
}

/* =========================
   Families index (match fort)
   ========================= */

function buildCanonStrongIndex_(canonValues, idxC) {
  const m = new Map();
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
   Scoring "possible duplicate"
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
   Incremental upsert (form submit)
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
    newRow[idxC.created_at] = new Date();
    newRow[idxC.updated_at] = new Date();

    const merged = mergeClusterToCanonRow_(
      [{ raw: parsed.raw, timestamp: parsed.timestamp, sheetRow: parsed.sheetRow }],
      newRow,
      idxC
    );

    applyGraduationStatus_(merged, idxC, 'create');

    // Traça
    merged[idxC.source_rows] = String(parsed.sheetRow);
    merged[idxC.source_count] = '1';
    merged[idxC.dedup_reasons] = '';

    canon.appendRow(merged);
    return;
  }

  // UPDATE existent — overwrite=true per reflectir edicions
  const rowIndex1 = hit + 1;
  const current = canonValues[hit].slice();

  const merged = mergeClusterToCanonRow_(
    [{ raw: parsed.raw, timestamp: parsed.timestamp, sheetRow: parsed.sheetRow }],
    current,
    idxC,
    true
  );

  // Traça: afegeix fila si no hi és
  const existing = (merged[idxC.source_rows] || '').toString().trim();
  const parts = existing ? existing.split('|') : [];
  const sr = String(parsed.sheetRow);
  if (!parts.includes(sr)) parts.push(sr);
  parts.sort((a,b)=>Number(a)-Number(b));

  merged[idxC.source_rows] = parts.join('|');
  merged[idxC.source_count] = String(parts.length);
  merged[idxC.dedup_reasons] = merged[idxC.dedup_reasons] || '';

  applyGraduationStatus_(merged, idxC, 'update');

  merged[idxC.updated_at] = new Date();

  canon.getRange(rowIndex1, 1, 1, canonHeaders.length).setValues([merged]);
}

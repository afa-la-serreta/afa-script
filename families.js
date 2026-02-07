/****************
 * families.gs
 *
 * Entry points principals per a la gestió de famílies:
 * - onOpen: menú personalitzat
 * - importAll: import complet de respostes → Famílies
 * - findPotentialDuplicatesAll: cerca possibles duplicats
 * - onFormSubmit: trigger incremental
 * - deactivateGraduatedFamilies: desactiva famílies graduades
 *
 * Depèn de: config.gs, utils.gs, parsing.gs, dedup.gs
 ****************/

/* =========================
   Menú personalitzat
   ========================= */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AFA')
    .addItem('Importar totes les respostes', 'importAll')
    .addItem('Buscar possibles duplicats', 'findPotentialDuplicatesAll')
    .addSeparator()
    .addItem('Desactivar famílies graduades (6è)', 'deactivateGraduatedFamilies')
    .addToUi();
}

/* =========================
   Import complet
   ========================= */

/**
 * Import històric:
 * - Llegeix totes les respostes
 * - Agrupa per claus fortes (email/tel/iban/dni) via Union-Find
 * - Escriu/actualitza 1 fila canònica per cluster a "Famílies"
 * - Escriu traça: source_rows/source_count/dedup_reasons
 * - Aplica graduació automàticament (ACTIVE / INACTIVE)
 */
function importAll() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_RESPONSES);
  const canon = ss.getSheetByName(SHEET_CANON);

  if (!sh) throw new Error(`No trobo la pestanya: ${SHEET_RESPONSES}`);
  if (!canon) throw new Error(`No trobo la pestanya: ${SHEET_CANON}`);

  const respValues = sh.getDataRange().getValues();
  if (respValues.length < 2) return;

  const rawHeaders = respValues[0].map(h => (h == null ? '' : String(h)));
  const headers = makeUniqueHeaders_(rawHeaders);

  const parsedRows = respValues.slice(1).map((r, i) => parseResponseRow_(headers, r, i + 2));
  const clusters = clusterByStrongKeys_(parsedRows);

  const canonValues = canon.getDataRange().getValues();
  if (canonValues.length < 1) throw new Error(`${SHEET_CANON} no té headers a la fila 1.`);
  const canonHeaders = canonValues[0].map(h => String(h).trim());
  const idxC = headerIndex_(canonHeaders);

  const requiredCanonCols = [
    'family_id','token_edit','status','created_at','updated_at','source_last_timestamp',
    'g1_nom','g1_cognoms','g1_email','g1_telefon',
    'bank_iban',
    'source_rows','source_count','dedup_reasons'
  ];
  requiredCanonCols.forEach(c => {
    if (idxC[c] == null) throw new Error(`Falta la columna ${c} a ${SHEET_CANON}`);
  });

  const existingIndex = buildCanonStrongIndex_(canonValues, idxC);

  let created = 0;
  let updated = 0;

  for (const cluster of clusters) {
    const clusterKeys = strongKeysForCluster_(cluster);

    let canonRowIdx = null;
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
      newRow[idxC.created_at] = new Date();
      newRow[idxC.updated_at] = new Date();

      const merged = mergeClusterToCanonRow_(cluster, newRow, idxC);
      applyGraduationStatus_(merged, idxC, 'create');

      const tr = clusterTrace_(cluster);
      setCanon_(merged, idxC, 'source_rows', tr.source_rows, true);
      setCanon_(merged, idxC, 'source_count', tr.source_count, true);
      setCanon_(merged, idxC, 'dedup_reasons', tr.dedup_reasons, true);

      canon.appendRow(merged);
      created++;

      addCanonRowToIndex_(existingIndex, merged, idxC, canonValues.length);
    } else {
      // UPDATE
      const current = canonValues[canonRowIdx].slice();
      const merged = mergeClusterToCanonRow_(cluster, current, idxC);

      const tr = clusterTrace_(cluster);
      setCanon_(merged, idxC, 'source_rows', tr.source_rows, true);
      setCanon_(merged, idxC, 'source_count', tr.source_count, true);
      setCanon_(merged, idxC, 'dedup_reasons', tr.dedup_reasons, true);

      applyGraduationStatus_(merged, idxC, 'update');

      merged[idxC.updated_at] = new Date();

      canon.getRange(canonRowIdx + 1, 1, 1, canonHeaders.length).setValues([merged]);
      updated++;

      addCanonRowToIndex_(existingIndex, merged, idxC, canonRowIdx);
    }
  }

  Logger.log(`importAll: clusters=${clusters.length}, created=${created}, updated=${updated}`);
  SpreadsheetApp.getActive().toast(`${clusters.length} clusters, ${created} creades, ${updated} actualitzades`, 'Import completat ✓');
}

/* =========================
   Possibles duplicats
   ========================= */

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
      if (best && best.score >= 70) out.push(best);
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
      x.a.timestamp instanceof Date ? x.a.timestamp : (x.a.timestamp || ''),
      x.b.timestamp instanceof Date ? x.b.timestamp : (x.b.timestamp || ''),
      x.strongMatch ? 'YES' : '',
      '', '', '', '', ''
    ]));

    pot.getRange(2, 1, rowsOut.length, headersOut.length).setValues(rowsOut);
  }

  SpreadsheetApp.getActive().toast(`${out.length} possibles duplicats trobats`, 'Cerca completada ✓');
}

/* =========================
   Trigger incremental
   ========================= */

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
   Desactivació famílies graduades
   ========================= */

/**
 * Desactiva les famílies on TOTS els infants han acabat 6è de primària.
 * Botó manual per auditar i forçar desactivació.
 */
function deactivateGraduatedFamilies() {
  const ss = SpreadsheetApp.getActive();
  const canon = ss.getSheetByName(SHEET_CANON);
  if (!canon) throw new Error(`No trobo la pestanya: ${SHEET_CANON}`);

  const values = canon.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => String(h).trim());
  const idx = headerIndex_(headers);

  ['status', 'inactive_reason', 'inactive_at',
   'c1_dob', 'c2_dob', 'c3_dob', 'c4_dob'].forEach(c => {
    if (idx[c] == null) throw new Error(`Falta la columna ${c} a ${SHEET_CANON}`);
  });

  const now = new Date();
  let deactivated = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx.status]).trim() !== 'ACTIVE') continue;

    if (!shouldDeactivateFamily_(row, idx)) continue;

    row[idx.status] = 'INACTIVE';
    row[idx.inactive_reason] = '6è acabat';
    row[idx.inactive_at] = now;
    row[idx.updated_at] = now;

    canon.getRange(r + 1, 1, 1, headers.length).setValues([row]);
    deactivated++;

    Logger.log(`✓ Desactivada: ${row[idx.g1_nom]} ${row[idx.g1_cognoms]}`);
  }

  Logger.log(`deactivateGraduatedFamilies: desactivades=${deactivated}`);
  SpreadsheetApp.getActive().toast(`${deactivated} famílies desactivades`, 'Graduació ✓');
}

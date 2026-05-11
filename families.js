// Codi font i docs: https://github.com/afa-la-serreta/afa-script — No editeu aquí, useu clasp push.
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
  var ui = SpreadsheetApp.getUi();

  var advanced = ui.createMenu('\uD83D\uDD27 Avançat')
    .addItem('\uD83D\uDCE5 Importar totes les respostes', 'importAll')
    .addItem('\uD83D\uDD0D Buscar possibles duplicats', 'findPotentialDuplicatesAll')
    .addItem('\uD83D\uDD04 Refrescar enlla\u00e7os d\u0027edici\u00f3', 'refreshEditUrls');

  ui.createMenu('\uD83C\uDF4E AFA')
    .addItem('1\uFE0F\u20E3 Desactivar fam\u00edlies graduades (6\u00e8)', 'deactivateGraduatedFamilies')
    .addItem('2\uFE0F\u20E3 Enviar correu de confirmaci\u00f3', 'sendConfirmationEmails')
    .addItem('3\uFE0F\u20E3 Generar fitxer SEPA (rebuts)', 'generateSepaXml')
    .addSeparator()
    .addItem('\uD83D\uDD17 Enlla\u00e7 d\u0027edici\u00f3 (fila activa)', 'showEditUrlForActiveRow')
    .addSeparator()
    .addSubMenu(advanced)
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
   Sincronització incremental (edicions)
   ========================= */

/**
 * Sincronitza respostes editades.
 * Guarda la data de l'última sincronització a PropertiesService
 * i només re-processa respostes amb timestamp posterior.
 * Pensat per executar-se amb un trigger temporitzat (cada 15 min).
 */
function syncEdited() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_RESPONSES);
  var canon = ss.getSheetByName(SHEET_CANON);

  if (!sh || !canon) return;

  var props = PropertiesService.getScriptProperties();
  var lastSyncStr = props.getProperty('syncEdited_lastRun');
  var lastSyncMs;

  if (lastSyncStr) {
    lastSyncMs = Number(lastSyncStr);
  } else {
    // Primera execució: agafar el timestamp més recent de Famílies
    var canonValues = canon.getDataRange().getValues();
    if (canonValues.length < 2) return;
    var canonHeaders = canonValues[0].map(function(h) { return String(h).trim(); });
    var idxC = headerIndex_(canonHeaders);
    if (idxC.source_last_timestamp == null) return;

    lastSyncMs = 0;
    for (var r = 1; r < canonValues.length; r++) {
      var ts = canonValues[r][idxC.source_last_timestamp];
      var ms = (ts instanceof Date) ? ts.getTime() : new Date(ts).getTime();
      if (!isNaN(ms) && ms > lastSyncMs) lastSyncMs = ms;
    }
  }

  // Escanejar respostes amb timestamp posterior a l'última sincronització
  var respValues = sh.getDataRange().getValues();
  if (respValues.length < 2) return;

  var rawHeaders = respValues[0].map(function(h) { return h == null ? '' : String(h); });
  var headers = makeUniqueHeaders_(rawHeaders);
  var tsCol = headers.indexOf('Marca temporal');

  var changed = [];
  for (var i = 1; i < respValues.length; i++) {
    var rowTs = respValues[i][tsCol >= 0 ? tsCol : 0];
    var tsMs = (rowTs instanceof Date) ? rowTs.getTime() : new Date(rowTs).getTime();
    if (!isNaN(tsMs) && tsMs > lastSyncMs) {
      changed.push({ row: respValues[i], rowNumber: i + 1 });
    }
  }

  // Guardar marca temporal ABANS de processar (per evitar saltar-se res)
  props.setProperty('syncEdited_lastRun', String(Date.now()));

  if (changed.length === 0) return;

  // Re-processar només les files canviades
  var upserted = 0;
  for (var j = 0; j < changed.length; j++) {
    var parsed = parseResponseRow_(headers, changed[j].row, changed[j].rowNumber);
    upsertSingleResponse_(parsed);
    upserted++;
  }

  Logger.log('syncEdited: ' + upserted + ' respostes sincronitzades');
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

/* =========================
   Enviament de correus de confirmació
   ========================= */

/**
 * Envia un correu de confirmació a totes les famílies ACTIVE
 * amb les dades bancàries emmascarades i un enllaç per editar
 * la seva resposta al Google Form original.
 *
 * Requereix:
 * - El full de respostes estigui vinculat a un Google Form
 * - El form tingui activada l'opció "Permet editar respostes"
 */
function sendConfirmationEmails() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const canon = ss.getSheetByName(SHEET_CANON);
  const respSheet = ss.getSheetByName(SHEET_RESPONSES);

  if (!canon) throw new Error(`No trobo la pestanya: ${SHEET_CANON}`);
  if (!respSheet) throw new Error(`No trobo la pestanya: ${SHEET_RESPONSES}`);

  // Assegurar que existeix la columna cau de l'edit URL
  ensureCanonColumn_(canon, 'edit_url');

  // Detectar enviament en curs
  const allValues = canon.getDataRange().getValues();
  const allHeaders = allValues[0].map(function(h) { return String(h).trim(); });
  const idxPre = headerIndex_(allHeaders);

  if (idxPre.confirmation_sent_at == null) {
    throw new Error('Falta la columna confirmation_sent_at a ' + SHEET_CANON);
  }

  var earliestSent = null;
  var sentCount = 0;
  for (var r = 1; r < allValues.length; r++) {
    var v = allValues[r][idxPre.confirmation_sent_at];
    if (v) {
      sentCount++;
      var d = (v instanceof Date) ? v : new Date(v);
      if (!isNaN(d.getTime()) && (earliestSent === null || d < earliestSent)) {
        earliestSent = d;
      }
    }
  }

  if (sentCount > 0 && earliestSent) {
    var dateStr = earliestSent.toLocaleDateString('ca-ES', { day: 'numeric', month: 'long', year: 'numeric' });
    var resp = ui.alert(
      'Enviament en curs',
      'Hi ha un enviament iniciat el ' + dateStr + ' amb ' + sentCount + ' correus ja enviats.\n\n' +
      'S\u00ed = Continuar enviant els pendents\n' +
      'No = Comen\u00e7ar de nou (esborra les marques anteriors)',
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (resp === ui.Button.CANCEL) {
      ss.toast('Enviament cancel\u00b7lat.', 'Cancel\u00b7lat');
      return;
    }
    if (resp === ui.Button.NO) {
      // Esborrar la columna confirmation_sent_at
      var col = idxPre.confirmation_sent_at + 1;
      canon.getRange(2, col, allValues.length - 1, 1).clearContent();
      ss.toast('Marques esborrades. Preparant nou enviament...', 'AFA');
    }
    // YES → continuar amb les marques existents
  }

  ss.toast('Preparant enviament...', 'AFA');

  // Obtenir el form vinculat
  const formUrl = respSheet.getFormUrl();
  if (!formUrl) {
    ui.alert('Error', 'El full de respostes no est\u00e0 vinculat a cap Google Form.', ui.ButtonSet.OK);
    return;
  }

  const form = FormApp.openByUrl(formUrl);
  const formResponses = form.getResponses();

  // Construir mapa: timestamp (ms) → editResponseUrl
  const editUrlByTimestamp = new Map();
  for (const resp of formResponses) {
    const ts = resp.getTimestamp().getTime();
    const editUrl = resp.getEditResponseUrl();
    if (editUrl) editUrlByTimestamp.set(ts, editUrl);
  }

  // Construir mapa: email → editUrl (fallback)
  // Usa les dades del full de respostes (ràpid) en lloc de la Forms API (lent).
  const editUrlByEmail = new Map();
  const respValues = respSheet.getDataRange().getValues();
  if (respValues.length > 1) {
    const respHeaders = makeUniqueHeaders_(respValues[0].map(h => (h == null ? '' : String(h))));
    const emailColIdx = respHeaders.indexOf('Correu electr\u00f2nic:');
    const tsColIdx = respHeaders.indexOf('Marca temporal');

    if (emailColIdx >= 0 && tsColIdx >= 0) {
      for (let i = 1; i < respValues.length; i++) {
        const rowTs = respValues[i][tsColIdx];
        const email = normEmail_(respValues[i][emailColIdx]);
        if (!email || !rowTs) continue;

        const tsMs = (rowTs instanceof Date) ? rowTs.getTime() : new Date(rowTs).getTime();
        if (isNaN(tsMs)) continue;

        const editUrl = editUrlByTimestamp.get(tsMs);
        if (editUrl) editUrlByEmail.set(email, editUrl);
      }
    }
  }

  // Llegir famílies (re-llegir per si s'ha esborrat la columna)
  const values = canon.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => String(h).trim());
  const idx = headerIndex_(headers);

  ['status', 'g1_nom', 'g2_nom', 'g1_email', 'g2_email', 'bank_iban', 'source_last_timestamp', 'token_edit', 'confirmation_sent_at',
   'c1_nom', 'c1_dob', 'c2_nom', 'c2_dob', 'c3_nom', 'c3_dob', 'c4_nom', 'c4_dob',
   'g1_adreca', 'g1_poblacio', 'g1_cp'].forEach(c => {
    if (idx[c] == null) throw new Error(`Falta la columna ${c} a ${SHEET_CANON}`);
  });

  // URL base del webapp (per al link de baixa)
  const webappUrl = WEBAPP_URL;

  // Preparar llista de destinataris
  const toSend = [];
  const skipped = [];
  let alreadySent = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx.status]).trim() !== 'ACTIVE') continue;

    // Saltar famílies ja enviades
    const sentAt = row[idx.confirmation_sent_at];
    if (sentAt) {
      alreadySent++;
      continue;
    }

    const email = normEmail_(row[idx.g1_email]);
    const email2 = normEmail_(row[idx.g2_email]);
    const nom = (row[idx.g1_nom] || '').toString().trim();
    const nom2 = (row[idx.g2_nom] || '').toString().trim();
    const iban = (row[idx.bank_iban] || '').toString().trim();
    const tokenEdit = (row[idx.token_edit] || '').toString().trim();

    if (!email) {
      skipped.push(`Fila ${r + 1}: sense email`);
      continue;
    }

    // Combinar destinataris (g1 + g2 si és diferent)
    var toAddresses = email;
    if (email2 && email2 !== email) {
      toAddresses = email + ',' + email2;
    }

    // Trobar l'edit URL: primer per timestamp, després per email com a fallback
    let editUrl = null;
    const srcTs = row[idx.source_last_timestamp];
    if (srcTs instanceof Date) {
      editUrl = editUrlByTimestamp.get(srcTs.getTime()) || null;
    } else if (srcTs) {
      const d = new Date(srcTs);
      if (!isNaN(d.getTime())) editUrl = editUrlByTimestamp.get(d.getTime()) || null;
    }
    if (!editUrl) {
      editUrl = editUrlByEmail.get(email) || null;
    }

    if (!editUrl) {
      skipped.push(`Fila ${r + 1} (${email}): sense URL d'edici\u00f3 del form`);
      continue;
    }

    // Cau l'edit URL al full per fer lookups ràpids després
    if (idx.edit_url != null && row[idx.edit_url] !== editUrl) {
      canon.getRange(r + 1, idx.edit_url + 1).setValue(editUrl);
    }

    // URL de baixa voluntària via webapp
    const baixaUrl = tokenEdit ? `${webappUrl}?token=${encodeURIComponent(tokenEdit)}&action=baixa` : null;

    // Validar IBAN (ES ha de tenir 24 caràcters)
    var cleanIban = normIban_(iban);
    var ibanInvalid = !cleanIban || cleanIban.length !== 24 || cleanIban.substring(0, 2) !== 'ES';

    // Recollir infants actius (nom + curs calculat, excloent graduats i placeholders)
    var children = [];
    var childSlots = [
      { nom: 'c1_nom', dob: 'c1_dob' },
      { nom: 'c2_nom', dob: 'c2_dob' },
      { nom: 'c3_nom', dob: 'c3_dob' },
      { nom: 'c4_nom', dob: 'c4_dob' }
    ];
    for (var s = 0; s < childSlots.length; s++) {
      var childNom = (row[idx[childSlots[s].nom]] || '').toString().trim();
      if (!childNom || /^[.\-_]+$/.test(childNom) || childNom.length < 2) continue; // skip empty/placeholder names
      var birthYear = extractBirthYear_(row[idx[childSlots[s].dob]]);
      var grade = childGrade_(birthYear);
      if (!grade) continue; // skip graduated or not-yet-in-school children
      children.push({ nom: childNom, grade: grade });
    }

    // Recollir adreça
    var adr = (row[idx.g1_adreca] || '').toString().trim();
    var cp = (row[idx.g1_cp] || '').toString().trim();
    var pobl = (row[idx.g1_poblacio] || '').toString().trim();
    var address = [adr, cp, pobl].filter(function(x) { return x; }).join(', ');

    toSend.push({ email, toAddresses, nom, nom2, iban, ibanInvalid, editUrl, baixaUrl, children, address, sheetRow: r + 1, canonRow: r });
  }

  if (toSend.length === 0) {
    const parts = [];
    if (alreadySent) parts.push(`${alreadySent} ja enviades anteriorment`);
    if (skipped.length) parts.push(`${skipped.length} saltades (sense email o URL)`);
    ui.alert('Cap correu pendent',
      parts.length ? parts.join('\n') : 'No hi ha fam\u00edlies actives pendents.',
      ui.ButtonSet.OK);
    return;
  }

  // Comprovar quota restant
  const remaining = MailApp.getRemainingDailyQuota();

  // Confirmació
  const pendingMsg = `${toSend.length} correus pendents d'enviar.` +
    (alreadySent ? `\n${alreadySent} ja enviats anteriorment.` : '') +
    (skipped.length ? `\n${skipped.length} saltades (sense email o URL).` : '') +
    `\n\nQuota di\u00e0ria restant: ${remaining} correus.` +
    (toSend.length > remaining
      ? `\n\u26a0\ufe0f Nom\u00e9s s'enviaran ${remaining} avui. Torneu a executar dem\u00e0 per continuar.`
      : '') +
    '\n\nVols continuar?';

  const answer = ui.alert('Confirmar enviament', pendingMsg, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) {
    ss.toast('Enviament cancel\u00b7lat.', 'Cancel\u00b7lat');
    return;
  }

  // Enviar correus (respectant la quota)
  const maxToSend = Math.min(toSend.length, remaining);
  let sent = 0;
  let errors = 0;
  const now = new Date();

  for (let i = 0; i < maxToSend; i++) {
    const family = toSend[i];
    try {
      const htmlBody = buildConfirmationEmailHtml_(family.nom, family.nom2, family.iban, family.editUrl, family.baixaUrl, family.ibanInvalid, family.children, family.address);

      MailApp.sendEmail({
        to: family.toAddresses,
        subject: EMAIL_SUBJECT,
        htmlBody: htmlBody,
        name: EMAIL_SENDER_NAME,
      });

      sent++;

      // Marcar com a enviat
      canon.getRange(family.sheetRow, idx.confirmation_sent_at + 1).setValue(now);

      if (sent % 10 === 0 || sent === maxToSend) {
        ss.toast(`Enviant ${sent} / ${maxToSend}...`, 'AFA');
      }
      Logger.log(`\u2713 Enviat a: ${family.email} (fila ${family.sheetRow})`);
    } catch (err) {
      errors++;
      Logger.log(`\u2717 Error enviant a ${family.email}: ${err.message}`);
      // Si l'error és de quota, aturem
      if (err.message && err.message.indexOf('quota') >= 0) {
        Logger.log('Quota exhaurida, aturant enviament.');
        break;
      }
    }
  }

  const pendingRemaining = toSend.length - sent;
  const result = `${sent} correus enviats` +
    (errors ? `, ${errors} errors` : '') +
    (pendingRemaining > 0 ? `, ${pendingRemaining} pendents per dem\u00e0` : '') +
    (skipped.length ? `, ${skipped.length} saltades` : '');

  Logger.log(`sendConfirmationEmails: ${result}`);
  ss.toast(result, 'Enviament completat \u2713');
}

/**
 * Genera l'HTML del correu de confirmació.
 */
function buildConfirmationEmailHtml_(nom, nom2, iban, editUrl, baixaUrl, ibanInvalid, children, address) {
  const masked = maskIban_(iban);
  var ibanLine;
  if (ibanInvalid) {
    ibanLine = '&#x26a0;&#xfe0f; <b>L\'IBAN que tenim enregistrat no &eacute;s v&agrave;lid.</b> Si us plau, actualitzeu-lo amb el n&uacute;mero correcte.';
  } else if (masked) {
    ibanLine = `El compte bancari que tenim enregistrat acaba en <b>${masked}</b>.`;
  } else {
    ibanLine = 'No tenim cap compte bancari enregistrat per a la vostra fam&iacute;lia.';
  }

  const baixaLine = baixaUrl
    ? `<p>Si voleu donar-vos de baixa de l'AFA, podeu fer-ho <a href="${baixaUrl}" style="color: #1a73e8;">aqu&iacute;</a>.</p>`
    : '';

  // Construir salutació (un o dos tutors)
  var greeting = escapeHtml_(nom || 'fam\u00edlia');
  if (nom2) greeting += ' i ' + escapeHtml_(nom2);

  // Construir secció d'infants
  var childrenHtml = '';
  if (children && children.length > 0) {
    var items = children.map(function(c) {
      var label = escapeHtml_(c.nom);
      if (c.grade) {
        label += ' <span style="color: #666;">(' + escapeHtml_(c.grade) + ')</span>';
      }
      return '<li style="margin: 4px 0;">' + label + '</li>';
    }).join('\n          ');
    childrenHtml = '<b>Infants enregistrats:</b>\n        <ul style="margin: 6px 0; padding-left: 20px;">\n          ' + items + '\n        </ul>';
  } else {
    childrenHtml = '<b>Infants enregistrats:</b> no en tenim const&agrave;ncia';
  }

  // Construir secció d'adreça
  var addressHtml = '';
  if (address) {
    addressHtml = '<b>Adre&ccedil;a:</b> ' + escapeHtml_(address);
  } else {
    addressHtml = '<b>Adre&ccedil;a:</b> no enregistrada';
  }

  // Nota sobre dades d'infants (info o warning segons si en tenim)
  var childrenNote = '';
  if (children && children.length > 0) {
    childrenNote = '&#x2139;&#xfe0f; Manteniu les dades dels infants actualitzades. Quan tots acabin 6&egrave; de prim&agrave;ria, la fam&iacute;lia es dona de baixa autom&agrave;ticament i deixareu de pagar la quota.';
  } else {
    childrenNote = '&#x26a0;&#xfe0f; No tenim les dades dels vostres infants. Necessitem el nom i la data de naixement per poder donar de baixa la fam&iacute;lia autom&agrave;ticament quan acabin 6&egrave; de prim&agrave;ria.';
  }

  return `
<div style="font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; color: #333;">
  <p>Bona tarda ${greeting},</p>

  <p>Encantats de saludar-vos.</p>

  <p>Informar-vos que en breu es passar&agrave; la quota de l'AFA (<b>35,00 &euro;</b>) i volem confirmar que les vostres dades siguin correctes i vigents.</p>

  <p>${ibanLine}</p>

  <div style="background-color: #f5f5f5; border-radius: 8px; padding: 14px 18px; margin: 16px 0;">
    <p style="margin: 0 0 8px 0;">
      ${childrenHtml}
    </p>
    <p style="margin: 8px 0 0 0;">
      ${addressHtml}
    </p>
  </div>

  <p>Si les dades no s&oacute;n correctes o teniu nous infants que hagin comen&ccedil;at l'escola, actualitzeu-les fent clic al bot&oacute; seg&uuml;ent:</p>

  <p style="text-align: center; margin: 24px 0;">
    <a href="${editUrl}" style="background-color: #1b5e20; color: #ffffff; padding: 12px 24px; border-radius: 8px; text-decoration: none; font-weight: 600; display: inline-block;">
      <span style="color: #ffffff;">Revisar i actualitzar les meves dades</span>
    </a>
  </p>

  <p style="font-size: 13px; color: #666;">${childrenNote}</p>

  ${baixaLine}

  <p>Gr&agrave;cies per la vostra col&middot;laboraci&oacute; i participaci&oacute;!</p>

  <p style="margin-top: 32px; color: #666; border-top: 1px solid #eee; padding-top: 16px;">
    --<br>
    <b>AFA La Serreta</b>
  </p>
</div>`;
}

/**
 * Genera una previsualització HTML de tots els correus de confirmació
 * i la desa a Google Drive. No envia res.
 * Executar des de l'editor d'Apps Script (Executa > previewAllConfirmationEmails).
 */
function previewAllConfirmationEmails() {
  var ss = SpreadsheetApp.getActive();
  var canon = ss.getSheetByName(SHEET_CANON);
  if (!canon) throw new Error('No trobo la pestanya: ' + SHEET_CANON);

  var values = canon.getDataRange().getValues();
  if (values.length < 2) throw new Error('No hi ha dades');

  var headers = values[0].map(function(h) { return String(h).trim(); });
  var idx = headerIndex_(headers);

  ['status', 'g1_nom', 'g2_nom', 'g1_email', 'g2_email', 'bank_iban',
   'c1_nom', 'c1_dob', 'c2_nom', 'c2_dob', 'c3_nom', 'c3_dob', 'c4_nom', 'c4_dob',
   'g1_adreca', 'g1_poblacio', 'g1_cp'].forEach(function(c) {
    if (idx[c] == null) throw new Error('Falta la columna ' + c + ' a ' + SHEET_CANON);
  });

  var childSlots = [
    { nom: 'c1_nom', dob: 'c1_dob' },
    { nom: 'c2_nom', dob: 'c2_dob' },
    { nom: 'c3_nom', dob: 'c3_dob' },
    { nom: 'c4_nom', dob: 'c4_dob' }
  ];

  var parts = [];
  var count = 0;

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (String(row[idx.status]).trim() !== 'ACTIVE') continue;

    var nom = (row[idx.g1_nom] || '').toString().trim();
    var nom2 = (row[idx.g2_nom] || '').toString().trim();
    var email = (row[idx.g1_email] || '').toString().trim();
    var email2 = normEmail_(row[idx.g2_email]);
    var iban = (row[idx.bank_iban] || '').toString().trim();
    var cleanIban = normIban_(iban);
    var ibanInvalid = !cleanIban || cleanIban.length !== 24 || cleanIban.substring(0, 2) !== 'ES';

    var children = [];
    for (var s = 0; s < childSlots.length; s++) {
      var childNom = (row[idx[childSlots[s].nom]] || '').toString().trim();
      if (!childNom || /^[.\-_]+$/.test(childNom) || childNom.length < 2) continue; // skip empty/placeholder names
      var birthYear = extractBirthYear_(row[idx[childSlots[s].dob]]);
      var grade = childGrade_(birthYear);
      if (!grade) continue; // skip graduated or not-yet-in-school children
      children.push({ nom: childNom, grade: grade });
    }

    var adr = (row[idx.g1_adreca] || '').toString().trim();
    var cp = (row[idx.g1_cp] || '').toString().trim();
    var pobl = (row[idx.g1_poblacio] || '').toString().trim();
    var address = [adr, cp, pobl].filter(function(x) { return x; }).join(', ');

    var html = buildConfirmationEmailHtml_(nom, nom2, iban, '#', '#', ibanInvalid, children, address);

    var emailDisplay = escapeHtml_(email);
    if (email2 && email2 !== normEmail_(email)) emailDisplay += ', ' + escapeHtml_(email2);

    parts.push(
      '<div style="border: 2px solid #ccc; border-radius: 8px; margin: 20px auto; max-width: 640px; padding: 10px;">' +
      '<p style="background: #eee; padding: 8px 12px; margin: 0 0 10px 0; border-radius: 4px; font-size: 13px;">' +
      '<b>Fila ' + (r + 1) + '</b> &mdash; ' + emailDisplay + ' &mdash; ' + escapeHtml_(nom) +
      '</p>' + html + '</div>'
    );
    count++;
  }

  if (count === 0) throw new Error('No hi ha famílies ACTIVE');

  var fullHtml = '<!DOCTYPE html><html><head><meta charset="utf-8"><title>Preview correus AFA</title></head><body style="background:#fafafa; padding: 20px;">' +
    '<h1 style="text-align:center; font-family: sans-serif;">Preview de ' + count + ' correus de confirmació</h1>' +
    parts.join('\n') +
    '</body></html>';

  var file = DriveApp.createFile('preview_correus_afa.html', fullHtml, MimeType.HTML);
  Logger.log('Preview creada: ' + file.getUrl());
  SpreadsheetApp.getActive().toast('Preview creada a Google Drive.\n' + file.getUrl(), 'Preview ✓');
}

/* =========================
   Cau d'enllaços d'edició (column `edit_url` a Famílies)
   ========================= */

/**
 * Assegura que existeix una columna anomenada `name` al full canònic.
 * Si no hi és, l'afegeix al final. Retorna l'índex 0-based de la columna.
 */
function ensureCanonColumn_(canon, name) {
  const lastCol = canon.getLastColumn();
  const headers = canon.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const existing = headers.indexOf(name);
  if (existing >= 0) return existing;

  const newCol = lastCol + 1;
  canon.getRange(1, newCol).setValue(name);
  return newCol - 1;
}

/**
 * Construeix els mapes timestamp/email → editResponseUrl a partir del Form vinculat.
 * Extret de sendConfirmationEmails per reutilitzar-lo des d'altres punts.
 */
function buildEditUrlMaps_(respSheet) {
  const formUrl = respSheet.getFormUrl();
  if (!formUrl) throw new Error('El full de respostes no està vinculat a cap Google Form.');

  const form = FormApp.openByUrl(formUrl);
  const formResponses = form.getResponses();

  const byTs = new Map();
  for (const resp of formResponses) {
    const ts = resp.getTimestamp().getTime();
    const editUrl = resp.getEditResponseUrl();
    if (editUrl) byTs.set(ts, editUrl);
  }

  const byEmail = new Map();
  const respValues = respSheet.getDataRange().getValues();
  if (respValues.length > 1) {
    const respHeaders = makeUniqueHeaders_(respValues[0].map(h => (h == null ? '' : String(h))));
    const emailColIdx = respHeaders.indexOf('Correu electrònic:');
    const tsColIdx = respHeaders.indexOf('Marca temporal');
    if (emailColIdx >= 0 && tsColIdx >= 0) {
      for (let i = 1; i < respValues.length; i++) {
        const email = normEmail_(respValues[i][emailColIdx]);
        const rowTs = respValues[i][tsColIdx];
        if (!email || !rowTs) continue;
        const tsMs = (rowTs instanceof Date) ? rowTs.getTime() : new Date(rowTs).getTime();
        if (isNaN(tsMs)) continue;
        const editUrl = byTs.get(tsMs);
        if (editUrl) byEmail.set(email, editUrl);
      }
    }
  }
  return { byTs, byEmail };
}

/**
 * Donada una fila canònica (valors ja llegits) i els mapes, retorna l'editResponseUrl
 * fent servir primer la coincidència per source_last_timestamp i, com a fallback, per email.
 */
function resolveEditUrlFromMaps_(row, idx, maps) {
  const ts = row[idx.source_last_timestamp];
  let editUrl = null;
  if (ts instanceof Date) {
    editUrl = maps.byTs.get(ts.getTime()) || null;
  } else if (ts) {
    const d = new Date(ts);
    if (!isNaN(d.getTime())) editUrl = maps.byTs.get(d.getTime()) || null;
  }
  if (!editUrl) {
    const email = normEmail_(row[idx.g1_email]);
    if (email) editUrl = maps.byEmail.get(email) || null;
  }
  return editUrl;
}

/**
 * Refresca la columna `edit_url` de totes les famílies recalculant-la
 * a partir de les respostes del Form. Útil quan canvien respostes
 * o cal omplir la cau per primer cop.
 */
function refreshEditUrls() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const canon = ss.getSheetByName(SHEET_CANON);
  const respSheet = ss.getSheetByName(SHEET_RESPONSES);
  if (!canon) throw new Error(`No trobo la pestanya: ${SHEET_CANON}`);
  if (!respSheet) throw new Error(`No trobo la pestanya: ${SHEET_RESPONSES}`);

  ensureCanonColumn_(canon, 'edit_url');

  ss.toast('Recopilant respostes del Form...', 'Refresc');
  const maps = buildEditUrlMaps_(respSheet);

  const values = canon.getDataRange().getValues();
  if (values.length < 2) {
    ui.alert('No hi ha files a Famílies.');
    return;
  }
  const headers = values[0].map(h => String(h).trim());
  const idx = headerIndex_(headers);
  if (idx.edit_url == null) throw new Error('Falta la columna edit_url després d\'assegurar-la.');

  let updated = 0;
  let unchanged = 0;
  let missing = 0;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const editUrl = resolveEditUrlFromMaps_(row, idx, maps);
    if (!editUrl) { missing++; continue; }
    if (row[idx.edit_url] === editUrl) { unchanged++; continue; }
    canon.getRange(r + 1, idx.edit_url + 1).setValue(editUrl);
    updated++;
  }

  ss.toast(`${updated} actualitzades, ${unchanged} ja al dia, ${missing} sense URL`, 'Refresc ✓');
}

/**
 * Mostra l'enllaç d'edició del Google Form per a la fila seleccionada del full Famílies.
 * Llegeix la cau (`edit_url`) si està disponible; si no, la calcula i la desa.
 */
function showEditUrlForActiveRow() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const canon = ss.getSheetByName(SHEET_CANON);
  if (!canon) throw new Error(`No trobo la pestanya: ${SHEET_CANON}`);

  const cell = ss.getActiveCell();
  if (!cell || cell.getSheet().getName() !== SHEET_CANON) {
    ui.alert('Posa el cursor en una fila del full «' + SHEET_CANON + '».');
    return;
  }
  const r = cell.getRow();
  if (r < 2) {
    ui.alert('Selecciona una fila de dades (no la capçalera).');
    return;
  }

  ensureCanonColumn_(canon, 'edit_url');

  const lastCol = canon.getLastColumn();
  const headers = canon.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const idx = headerIndex_(headers);
  const row = canon.getRange(r, 1, 1, lastCol).getValues()[0];

  let editUrl = (row[idx.edit_url] || '').toString().trim() || null;

  if (!editUrl) {
    const respSheet = ss.getSheetByName(SHEET_RESPONSES);
    if (!respSheet) throw new Error(`No trobo la pestanya: ${SHEET_RESPONSES}`);
    const maps = buildEditUrlMaps_(respSheet);
    editUrl = resolveEditUrlFromMaps_(row, idx, maps);
    if (editUrl) {
      canon.getRange(r, idx.edit_url + 1).setValue(editUrl);
    }
  }

  if (!editUrl) {
    ui.alert('No hi ha cap resposta del Form que coincideixi amb aquesta fila (ni per timestamp ni per email).');
    return;
  }

  const nom = (row[idx.g1_nom] || '').toString().trim();
  const cognoms = (row[idx.g1_cognoms] || '').toString().trim();
  const familyId = (row[idx.family_id] || '').toString().trim();
  const titol = [nom, cognoms].filter(Boolean).join(' ') || familyId;

  const safeUrl = editUrl.replace(/"/g, '&quot;');
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; padding: 8px 4px;">' +
    '  <p style="margin: 0 0 12px 0;"><b>' + escapeHtml_(titol) + '</b></p>' +
    '  <p style="margin: 0 0 8px 0;"><a href="' + safeUrl + '" target="_blank">Obrir el formulari per editar</a></p>' +
    '  <textarea readonly style="width:100%; height:80px; font-family: monospace;" onclick="this.select()">' + safeUrl + '</textarea>' +
    '  <p style="margin: 12px 0 0 0; font-size: 12px; color: #666;">Fes clic al text per seleccionar-lo i copiar-lo.</p>' +
    '</div>'
  ).setWidth(560).setHeight(220);
  ui.showModalDialog(html, 'Enllaç d\'edició');
}

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
    .createMenu('\uD83D\uDC68\u200D\uD83D\uDC69\u200D\uD83D\uDC67 AFA')
    .addItem('\uD83D\uDCE5 Importar totes les respostes', 'importAll')
    .addItem('\uD83D\uDD0D Buscar possibles duplicats', 'findPotentialDuplicatesAll')
    .addItem('\uD83D\uDD04 Sincronitzar edicions', 'syncEdited')
    .addSeparator()
    .addItem('\uD83C\uDF93 Desactivar fam\u00edlies graduades (6\u00e8)', 'deactivateGraduatedFamilies')
    .addSeparator()
    .addItem('\u2709\uFE0F Enviar correu de confirmaci\u00f3', 'sendConfirmationEmails')
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

  ['status', 'g1_nom', 'g1_email', 'bank_iban', 'source_last_timestamp', 'token_edit', 'confirmation_sent_at'].forEach(c => {
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
    const nom = (row[idx.g1_nom] || '').toString().trim();
    const iban = (row[idx.bank_iban] || '').toString().trim();
    const tokenEdit = (row[idx.token_edit] || '').toString().trim();

    if (!email) {
      skipped.push(`Fila ${r + 1}: sense email`);
      continue;
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

    // URL de baixa voluntària via webapp
    const baixaUrl = tokenEdit ? `${webappUrl}?token=${encodeURIComponent(tokenEdit)}&action=baixa` : null;

    toSend.push({ email, nom, iban, editUrl, baixaUrl, sheetRow: r + 1, canonRow: r });
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
      const htmlBody = buildConfirmationEmailHtml_(family.nom, family.iban, family.editUrl, family.baixaUrl);

      MailApp.sendEmail({
        to: family.email,
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
function buildConfirmationEmailHtml_(nom, iban, editUrl, baixaUrl) {
  const masked = maskIban_(iban);
  const ibanLine = masked
    ? `El compte bancari que tenim enregistrat acaba en <b>${masked}</b>.`
    : 'No tenim cap compte bancari enregistrat per a la vostra fam&iacute;lia.';

  const baixaLine = baixaUrl
    ? `<p>Si voleu donar-vos de baixa de l'AFA, podeu fer-ho <a href="${baixaUrl}" style="color: #1a73e8;">aqu&iacute;</a>.</p>`
    : '';

  return `
<div style="font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; color: #333;">
  <p>Bona tarda ${escapeHtml_(nom || 'fam\u00edlia')},</p>

  <p>Encantats de saludar-vos.</p>

  <p>Informar-vos que en breu es passar&agrave; la quota de l'AFA (<b>35,00 &euro;</b>) i volem confirmar que les vostres dades siguin correctes i vigents.</p>

  <p>${ibanLine}</p>

  <p>Si us plau, reviseu les vostres dades i actualitzeu-les si cal fent clic al bot&oacute; seg&uuml;ent:</p>

  <p style="text-align: center; margin: 24px 0;">
    <a href="${editUrl}" style="background-color: #0b6; color: white; padding: 12px 24px; border-radius: 8px; text-decoration: none; font-weight: 600; display: inline-block;">
      Revisar i actualitzar les meves dades
    </a>
  </p>

  ${baixaLine}

  <p>Gr&agrave;cies per la vostra col&middot;laboraci&oacute; i participaci&oacute;!</p>

  <p style="margin-top: 32px; color: #666; border-top: 1px solid #eee; padding-top: 16px;">
    --<br>
    <b>AFA La Serreta</b>
  </p>
</div>`;
}

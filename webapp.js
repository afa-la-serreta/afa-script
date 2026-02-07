/****************
 * webapp.gs
 * Web App per a famílies (magic link)
 ****************/

/**
 * GET /?token=...
 * Renderitza el formulari d’edició
 */
function doGet(e) {
  const token = (e.parameter && e.parameter.token) ? String(e.parameter.token).trim() : '';
  if (!token) return htmlMessage_('Falta el token de l’enllaç.');

  const canon = SpreadsheetApp.getActive().getSheetByName(SHEET_CANON);
  if (!canon) return htmlMessage_('No existeix la base de dades de famílies.');

  const rows = canon.getDataRange().getValues();
  if (rows.length < 2) return htmlMessage_('Encara no hi ha cap família registrada.');

  const headers = rows[0];
  const h = headerIndex_(headers);

  const match = findRowByToken_(rows, h.token_edit, token);
  if (!match) return htmlMessage_('Enllaç invàlid o caducat.');

  const row = rows[match.r];

  // No mostrem famílies inactives
  if (row[h.status] === 'INACTIVE') {
    return htmlMessage_('Aquesta família ja no està activa a l’AFA.');
  }

  const dto = toEditableDto_(headers, row);

  const t = HtmlService.createTemplateFromFile('index');
  t.data = dto;
  t.token = token;

  return t.evaluate()
    .setTitle("Actualitza dades AFA")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Guarda canvis del formulari (magic link)
 * NO toca IBAN ni DNI
 */
function saveFamily(token, payload) {
  token = String(token || '').trim();
  if (!token) throw new Error('Falta token.');
  if (!payload || typeof payload !== 'object') throw new Error('Payload invàlid.');

  const canon = SpreadsheetApp.getActive().getSheetByName(SHEET_CANON);
  const rows = canon.getDataRange().getValues();
  const headers = rows[0];
  const h = headerIndex_(headers);

  const match = findRowByToken_(rows, h.token_edit, token);
  if (!match) throw new Error('Token invàlid.');

  const rowIndex1 = match.r + 1;
  const row = rows[match.r];

  // Camps editables
  apply_(row, h.g1_email, payload.g1_email);
  apply_(row, h.g1_telefon, payload.g1_telefon);
  apply_(row, h.g1_adreca, payload.g1_adreca);
  apply_(row, h.g1_poblacio, payload.g1_poblacio);
  apply_(row, h.g1_cp, payload.g1_cp);

  apply_(row, h.g2_email, payload.g2_email);
  apply_(row, h.g2_telefon, payload.g2_telefon);
  apply_(row, h.g2_adreca, payload.g2_adreca);
  apply_(row, h.g2_poblacio, payload.g2_poblacio);
  apply_(row, h.g2_cp, payload.g2_cp);

  apply_(row, h.email_alternatiu, payload.email_alternatiu);

  for (let i = 1; i <= 4; i++) {
    apply_(row, h[`c${i}_nom`], payload[`c${i}_nom`]);
    apply_(row, h[`c${i}_cognoms`], payload[`c${i}_cognoms`]);
    apply_(row, h[`c${i}_dob`], payload[`c${i}_dob`]);
  }

  apply_(row, h.consent_a, payload.consent_a);
  apply_(row, h.consent_b, payload.consent_b);
  apply_(row, h.consent_c, payload.consent_c);

  row[h.updated_at] = new Date();

  canon.getRange(rowIndex1, 1, 1, headers.length).setValues([row]);
  return { ok: true };
}

/**
 * Baixa soft: família ha acabat 6è
 */
function deactivateFamily(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Falta token.');

  const canon = SpreadsheetApp.getActive().getSheetByName(SHEET_CANON);
  const rows = canon.getDataRange().getValues();
  const headers = rows[0];
  const h = headerIndex_(headers);

  const match = findRowByToken_(rows, h.token_edit, token);
  if (!match) throw new Error('Token invàlid.');

  const rowIndex1 = match.r + 1;
  const row = rows[match.r];

  row[h.status] = 'INACTIVE';
  row[h.inactive_reason] = 'GRADUATED_6E';
  row[h.inactive_at] = new Date();
  row[h.updated_at] = new Date();

  canon.getRange(rowIndex1, 1, 1, headers.length).setValues([row]);
  return { ok: true };
}

/* =========================
   Helpers
   ========================= */

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function htmlMessage_(msg) {
  return HtmlService.createHtmlOutput(
    `<div style="font-family:system-ui;padding:24px;max-width:720px;margin:0 auto">
      <h2>Actualitza dades AFA</h2>
      <p>${escapeHtml_(msg)}</p>
     </div>`
  ).setTitle("Actualitza dades AFA");
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function headerIndex_(headers) {
  const m = {};
  headers.forEach((h, i) => m[String(h).trim()] = i);
  return m;
}

function findRowByToken_(rows, tokenColIndex, token) {
  for (let r = 1; r < rows.length; r++) {
    if (String(rows[r][tokenColIndex] || '').trim() === token) return { r };
  }
  return null;
}

function toEditableDto_(headers, row) {
  const h = headerIndex_(headers);
  const get = (k) => (row[h[k]] == null) ? '' : String(row[h[k]]);
  return {
    g1_nom: get('g1_nom'),
    g1_cognoms: get('g1_cognoms'),
    g2_nom: get('g2_nom'),
    g2_cognoms: get('g2_cognoms'),

    g1_email: get('g1_email'),
    g1_telefon: get('g1_telefon'),
    g1_adreca: get('g1_adreca'),
    g1_poblacio: get('g1_poblacio'),
    g1_cp: get('g1_cp'),

    g2_email: get('g2_email'),
    g2_telefon: get('g2_telefon'),
    g2_adreca: get('g2_adreca'),
    g2_poblacio: get('g2_poblacio'),
    g2_cp: get('g2_cp'),

    email_alternatiu: get('email_alternatiu'),

    c1_nom: get('c1_nom'), c1_cognoms: get('c1_cognoms'), c1_dob: get('c1_dob'),
    c2_nom: get('c2_nom'), c2_cognoms: get('c2_cognoms'), c2_dob: get('c2_dob'),
    c3_nom: get('c3_nom'), c3_cognoms: get('c3_cognoms'), c3_dob: get('c3_dob'),
    c4_nom: get('c4_nom'), c4_cognoms: get('c4_cognoms'), c4_dob: get('c4_dob'),

    consent_a: get('consent_a'),
    consent_b: get('consent_b'),
    consent_c: get('consent_c'),
  };
}

function apply_(row, idx, val) {
  if (idx == null) return;
  if (val == null) return;
  row[idx] = String(val).trim();
}

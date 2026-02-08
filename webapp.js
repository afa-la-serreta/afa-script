// Codi font i docs: https://github.com/afa-la-serreta/afa-script — No editeu aquí, useu clasp push.
/****************
 * webapp.gs
 * Web App per a famílies (magic link)
 ****************/

/**
 * GET /?token=...
 * GET /?token=...&action=baixa
 * Renderitza el formulari d'edició o la pàgina de baixa.
 */
function doGet(e) {
  const token = (e.parameter && e.parameter.token) ? String(e.parameter.token).trim() : '';
  const action = (e.parameter && e.parameter.action) ? String(e.parameter.action).trim() : '';

  if (!token) return htmlMessage_("Falta el token de l'enlla\u00e7.");

  const canon = SpreadsheetApp.getActive().getSheetByName(SHEET_CANON);
  if (!canon) return htmlMessage_("No existeix la base de dades de fam\u00edlies.");

  const rows = canon.getDataRange().getValues();
  if (rows.length < 2) return htmlMessage_("Encara no hi ha cap fam\u00edlia registrada.");

  const headers = rows[0];
  const h = headerIndex_(headers);

  const match = findRowByToken_(rows, h.token_edit, token);
  if (!match) return htmlMessage_("Enlla\u00e7 inv\u00e0lid o caducat.");

  const row = rows[match.r];

  // No mostrem famílies inactives
  if (row[h.status] === 'INACTIVE') {
    return htmlMessage_("Aquesta fam\u00edlia ja no est\u00e0 activa a l'AFA.");
  }

  // Pàgina de baixa voluntària
  if (action === 'baixa') {
    const nom = (row[h.g1_nom] || '').toString().trim();
    const cognoms = (row[h.g1_cognoms] || '').toString().trim();
    return htmlBaixaPage_(token, nom, cognoms);
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
  if (!payload || typeof payload !== 'object') throw new Error('Payload inv\u00e0lid.');

  const canon = SpreadsheetApp.getActive().getSheetByName(SHEET_CANON);
  const rows = canon.getDataRange().getValues();
  const headers = rows[0];
  const h = headerIndex_(headers);

  const match = findRowByToken_(rows, h.token_edit, token);
  if (!match) throw new Error('Token inv\u00e0lid.');

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
 * Baixa voluntària: la família demana donar-se de baixa.
 */
function deactivateFamily(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Falta token.');

  const canon = SpreadsheetApp.getActive().getSheetByName(SHEET_CANON);
  const rows = canon.getDataRange().getValues();
  const headers = rows[0];
  const h = headerIndex_(headers);

  const match = findRowByToken_(rows, h.token_edit, token);
  if (!match) throw new Error('Token inv\u00e0lid.');

  const rowIndex1 = match.r + 1;
  const row = rows[match.r];

  row[h.status] = 'INACTIVE';
  row[h.inactive_reason] = 'baixa volunt\u00e0ria';
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
      <h2>AFA La Serreta</h2>
      <p>${escapeHtml_(msg)}</p>
     </div>`
  ).setTitle("AFA La Serreta");
}

/**
 * Pàgina de confirmació de baixa voluntària.
 */
function htmlBaixaPage_(token, nom, cognoms) {
  const fullName = escapeHtml_([nom, cognoms].filter(Boolean).join(' ') || 'fam\u00edlia');

  const html = `<!doctype html>
<html>
<head>
  <base target="_top">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial; margin: 0; background: #f6f7f9; }
    .wrap { max-width: 600px; margin: 0 auto; padding: 20px; }
    .card { background: #fff; border-radius: 16px; padding: 24px; box-shadow: 0 6px 20px rgba(0,0,0,.06); }
    h1 { font-size: 20px; margin: 0 0 16px; }
    .warn { background: #fff3e0; border-left: 4px solid #e65100; padding: 12px 16px; border-radius: 8px; margin: 16px 0; }
    button { border: 0; border-radius: 12px; padding: 12px 20px; font-weight: 600; cursor: pointer; font-size: 14px; }
    .danger { background: #d32f2f; color: #fff; }
    .secondary { background: #eef1f5; color: #222; }
    .actions { display: flex; gap: 12px; align-items: center; margin-top: 20px; }
    .status { font-size: 14px; margin-top: 12px; }
    .done { text-align: center; }
    .done h2 { color: #2e7d32; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="card" id="confirmCard">
      <h1>Baixa de l'AFA La Serreta</h1>
      <p>Fitxa de: <b>${fullName}</b></p>

      <div class="warn">
        <b>Atenci&oacute;:</b> Si confirmeu la baixa, la vostra fam&iacute;lia deixar&agrave; de ser s&ograve;cia de l'AFA.
        Aquesta acci&oacute; no es pot desfer des d'aqu&iacute;.
      </div>

      <p>Esteu segurs que voleu donar-vos de baixa?</p>

      <div class="actions">
        <button class="danger" onclick="confirmarBaixa()">S&iacute;, vull donar-me de baixa</button>
        <button class="secondary" onclick="cancelar()">Cancel&middot;lar</button>
      </div>

      <div class="status" id="status"></div>
    </div>

    <div class="card done" id="doneCard" style="display:none;">
      <h2>Baixa confirmada</h2>
      <p>La vostra fam&iacute;lia ha estat donada de baixa de l'AFA La Serreta.</p>
      <p>Si ha estat un error, poseu-vos en contacte amb la junta de l'AFA.</p>
    </div>
  </div>

  <script>
    var TOKEN = "${token}";

    function cancelar() {
      document.getElementById('confirmCard').innerHTML =
        '<div style="text-align:center;padding:24px;">' +
        '<h2 style="color:#2e7d32;">Cap canvi realitzat</h2>' +
        '<p>No us heu donat de baixa. Podeu tancar aquesta pestanya.</p>' +
        '</div>';
      try { google.script.host.close(); } catch(e) {}
    }

    function confirmarBaixa() {
      document.getElementById('status').textContent = 'Processant...';
      document.querySelector('.danger').disabled = true;

      google.script.run
        .withSuccessHandler(function() {
          document.getElementById('confirmCard').style.display = 'none';
          document.getElementById('doneCard').style.display = 'block';
        })
        .withFailureHandler(function(err) {
          document.getElementById('status').textContent = 'Error: ' + (err.message || err);
          document.getElementById('status').style.color = '#d32f2f';
          document.querySelector('.danger').disabled = false;
        })
        .deactivateFamily(TOKEN);
    }
  </script>
</body>
</html>`;

  return HtmlService.createHtmlOutput(html)
    .setTitle("Baixa AFA La Serreta")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// headerIndex_ definit a utils.gs (compartit)

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

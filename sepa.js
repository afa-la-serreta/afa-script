/****************
 * sepa.gs
 *
 * Generació de fitxers SEPA Direct Debit (pain.008.001.02)
 * per al cobrament de la quota de l'AFA.
 *
 * Format: Cuaderno 19.44 XML ISO 20022
 *
 * Depèn de: config.gs, utils.gs
 ****************/

/**
 * Entry point: mostra el formulari HTML per configurar els paràmetres SEPA.
 * Es crida des del menú AFA.
 */
function generateSepaXml() {
  var ui = SpreadsheetApp.getUi();

  // Calcular valors per defecte
  var defaultDate = nextBusinessDay_(new Date(), 2);
  var defaultDateStr = formatDate_(defaultDate);

  var today = new Date();
  var month = today.getMonth() + 1;
  var year = today.getFullYear();
  var cursStart = month >= 9 ? year : year - 1;
  var defaultRemittance = 'QUOTA AFA CURS ' + cursStart + '-' + (cursStart + 1);

  // Crear dialog HTML amb valors per defecte injectats
  var template = HtmlService.createTemplateFromFile('sepa_dialog');
  template.defaultDate = defaultDateStr;
  template.defaultAmount = '35.00';
  template.defaultRemittance = defaultRemittance;

  var html = template.evaluate()
    .setWidth(420)
    .setHeight(340);

  ui.showModalDialog(html, '\uD83C\uDF4E Generar fitxer SEPA');
}

/**
 * Genera el fitxer SEPA XML amb els paràmetres del formulari HTML.
 * Es crida des de sepa_dialog.html via google.script.run.
 * (No pot acabar en _ perquè google.script.run no pot cridar funcions privades.)
 */
function generateSepaWithParams(collectionDate, amountStr, remittance) {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var canon = ss.getSheetByName(SHEET_CANON);

  if (!canon) throw new Error('No trobo la pestanya: ' + SHEET_CANON);

  // Validar paràmetres
  if (!/^\d{4}-\d{2}-\d{2}$/.test(collectionDate)) {
    throw new Error('Format de data inv\u00e0lid. Usa YYYY-MM-DD.');
  }
  var amount = parseFloat(String(amountStr).replace(',', '.'));
  if (isNaN(amount) || amount <= 0) {
    throw new Error('Import inv\u00e0lid.');
  }
  remittance = String(remittance || '').trim();
  if (!remittance) throw new Error('Cal indicar el concepte.');

  // Derivar curs escolar
  var pYear = parseInt(collectionDate.substring(0, 4), 10);
  var pMonth = parseInt(collectionDate.substring(5, 7), 10);
  var cursStart = pMonth >= 9 ? pYear : pYear - 1;

  var params = {
    collectionDate: collectionDate,
    amount: amount,
    remittance: remittance,
    year: cursStart
  };

  ss.toast('Resolent BICs des de l\'IBAN...', 'AFA');

  // Llegir famílies
  var values = canon.getDataRange().getValues();
  if (values.length < 2) {
    ui.alert('Error', 'No hi ha fam\u00edlies a la pestanya ' + SHEET_CANON, ui.ButtonSet.OK);
    return;
  }

  var headers = values[0].map(function(h) { return String(h).trim(); });
  var idx = headerIndex_(headers);

  var required = ['status', 'g1_nom', 'g1_cognoms', 'g1_adreca', 'g1_poblacio', 'g1_cp',
                   'bank_iban', 'created_at'];
  for (var i = 0; i < required.length; i++) {
    if (idx[required[i]] == null) throw new Error('Falta la columna ' + required[i] + ' a ' + SHEET_CANON);
  }

  // Filtrar famílies vàlides
  var families = [];
  var skipped = [];
  var missingBic = [];

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (String(row[idx.status]).trim() !== 'ACTIVE') continue;

    var cognoms = (row[idx.g1_cognoms] || '').toString().trim();
    var iban = normIban_(row[idx.bank_iban]);
    var sheetSwift = (row[idx.bank_swift] || '').toString().trim().toUpperCase().replace(/\s+/g, '');
    var nom = (row[idx.g1_nom] || '').toString().trim();
    var adreca = (row[idx.g1_adreca] || '').toString().trim();
    var poblacio = (row[idx.g1_poblacio] || '').toString().trim();
    var cp = (row[idx.g1_cp] || '').toString().trim();
    var createdAt = row[idx.created_at];

    if (!cognoms) {
      skipped.push('Fila ' + (r + 1) + ': sense cognoms');
      continue;
    }
    if (!iban || iban.length < 15) {
      skipped.push('Fila ' + (r + 1) + ' (' + cognoms + '): IBAN inv\u00e0lid o absent');
      continue;
    }

    // Derivar BIC des de l'IBAN (consulta full BIC Bancs + IBANAPI si cal);
    // fallback a la columna del full
    var swift = ibanToBic_(iban) || sheetSwift;
    if (!swift || !/^[A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?$/.test(swift)) {
      var bankCode = iban.length >= 8 ? iban.substring(4, 8) : '????';
      missingBic.push('Fila ' + (r + 1) + ' (' + cognoms + '): codi entitat ' + bankCode);
      continue;
    }

    // Normalitzar BIC a 11 caràcters
    if (swift.length === 8) swift += 'XXX';

    // Data de signatura del mandat
    var dtSign = '';
    if (createdAt instanceof Date) {
      dtSign = formatDate_(createdAt);
    } else if (createdAt) {
      var d = new Date(createdAt);
      if (!isNaN(d.getTime())) dtSign = formatDate_(d);
    }
    if (!dtSign) dtSign = '2009-10-31'; // fallback

    families.push({
      nom: nom,
      cognoms: cognoms,
      iban: iban,
      swift: swift,
      adreca: adreca,
      cp: cp,
      poblacio: poblacio,
      dtOfSgntr: dtSign
    });
  }

  // Assegurar que el full BIC Bancs existeix si hi ha BICs desconeguts
  if (missingBic.length > 0) {
    var bicSh = ss.getSheetByName(SHEET_BIC);
    if (!bicSh) {
      bicSh = ss.insertSheet(SHEET_BIC);
      bicSh.getRange(1, 1, 1, 3).setValues([['Codi Entitat', 'BIC', 'Nom Entitat']]);
      bicSh.getRange(1, 1, 1, 3).setFontWeight('bold');
      bicSh.getRange('A:A').setNumberFormat('@');
    }
  }

  // Combinar tots els problemes
  var allProblems = skipped.concat(missingBic);

  if (families.length === 0) {
    ui.alert('Cap fam\u00edlia v\u00e0lida',
      'No s\'han trobat fam\u00edlies actives amb dades v\u00e0lides.' +
      (allProblems.length ? '\n\nProblemes:\n' + allProblems.join('\n') : ''),
      ui.ButtonSet.OK);
    return;
  }

  // Si hi ha problemes, deixar que l'usuari decideixi
  if (allProblems.length > 0) {
    var problemMsg = families.length + ' fam\u00edlies v\u00e0lides, ' + allProblems.length + ' amb problemes:\n\n' +
      allProblems.join('\n') +
      (missingBic.length ? '\n\n(Per a BICs desconeguts, afegiu-los al full "' + SHEET_BIC + '")' : '') +
      '\n\nVoleu continuar sense les fam\u00edlies amb problemes?';
    var answer = ui.alert('Fam\u00edlies amb problemes', problemMsg, ui.ButtonSet.YES_NO);
    if (answer !== ui.Button.YES) {
      ss.toast('Generaci\u00f3 cancel\u00b7lada.', 'Cancel\u00b7lat');
      return;
    }
  }

  // Confirmació
  var msg = families.length + ' fam\u00edlies incloses al fitxer SEPA.' +
    (skipped.length ? '\n' + skipped.length + ' saltades (sense IBAN, SWIFT o cognoms).' : '') +
    '\n\nImport total: ' + (families.length * params.amount).toFixed(2) + ' \u20ac' +
    '\nData de cobrament: ' + params.collectionDate +
    '\nConcepte: ' + params.remittance +
    '\n\nGenerar el fitxer?';

  var answer = ui.alert('Confirmar SEPA', msg, ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) {
    ss.toast('Generaci\u00f3 cancel\u00b7lada.', 'Cancel\u00b7lat');
    return;
  }

  // Generar XML
  var xml = buildSepaXml_(families, params);

  // Desar a Google Drive
  var folder = DriveApp.getFolderById(SEPA_DRIVE_FOLDER_ID);
  var fileName = 'SEPA_AFA_' + params.collectionDate + '.xml';
  var file = folder.createFile(fileName, xml, 'application/xml');

  var fileUrl = file.getUrl();
  Logger.log('SEPA XML generat: ' + fileUrl);

  // Mostrar enllaç
  var resultHtml = HtmlService.createHtmlOutput(
    '<p>Fitxer generat correctament:</p>' +
    '<p><b>' + escapeXml_(fileName) + '</b></p>' +
    '<p>' + families.length + ' rebuts, ' + (families.length * params.amount).toFixed(2) + ' \u20ac</p>' +
    '<p><a href="' + fileUrl + '" target="_blank">Obrir fitxer a Google Drive</a></p>' +
    (skipped.length ? '<p style="color:#999;">' + skipped.length + ' fam\u00edlies saltades.</p>' : '')
  ).setWidth(400).setHeight(250);

  ui.showModalDialog(resultHtml, 'SEPA generat \u2713');
}

/* =========================
   Generació XML
   ========================= */

/**
 * Construeix el XML SEPA pain.008.001.02.
 */
function buildSepaXml_(families, params) {
  var now = new Date();
  var dateStr = formatDate_(now);
  var dateCompact = dateStr.replace(/-/g, '');
  var nbOfTxs = families.length;
  var ctrlSum = (nbOfTxs * params.amount).toFixed(2);

  // IDs únics
  var msgId = 'AFA' + dateCompact + 'T' + padZero_(now.getHours(), 2) + padZero_(now.getMinutes(), 2) + padZero_(now.getSeconds(), 2);
  var pmtInfId = 'AFA' + dateCompact + 'P001';

  var lines = [];
  lines.push('<?xml version="1.0" encoding="UTF-8" standalone="no"?>');
  lines.push('<Document xmlns="urn:iso:std:iso:20022:tech:xsd:pain.008.001.02" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');
  lines.push('\t<CstmrDrctDbtInitn>');

  // Group Header
  lines.push('\t\t<GrpHdr>');
  lines.push('\t\t\t<MsgId>' + escapeXml_(msgId) + '</MsgId>');
  lines.push('\t\t\t<CreDtTm>' + dateStr + 'T00:00:00</CreDtTm>');
  lines.push('\t\t\t<NbOfTxs>' + nbOfTxs + '</NbOfTxs>');
  lines.push('\t\t\t<CtrlSum>' + ctrlSum + '</CtrlSum>');
  lines.push('\t\t\t<InitgPty>');
  lines.push('\t\t\t\t<Nm>' + escapeXml_(SEPA_CREDITOR_NAME) + '</Nm>');
  lines.push('\t\t\t\t<Id>');
  lines.push('\t\t\t\t\t<OrgId>');
  lines.push('\t\t\t\t\t\t<Othr>');
  lines.push('\t\t\t\t\t\t\t<Id>' + escapeXml_(SEPA_CREDITOR_ID) + '</Id>');
  lines.push('\t\t\t\t\t\t\t<SchmeNm>');
  lines.push('\t\t\t\t\t\t\t\t<Cd>CORE</Cd>');
  lines.push('\t\t\t\t\t\t\t</SchmeNm>');
  lines.push('\t\t\t\t\t\t</Othr>');
  lines.push('\t\t\t\t\t</OrgId>');
  lines.push('\t\t\t\t</Id>');
  lines.push('\t\t\t</InitgPty>');
  lines.push('\t\t</GrpHdr>');

  // Payment Information (single block)
  lines.push('\t\t<PmtInf>');
  lines.push('\t\t\t<PmtInfId>' + escapeXml_(pmtInfId) + '</PmtInfId>');
  lines.push('\t\t\t<PmtMtd>DD</PmtMtd>');
  lines.push('\t\t\t<PmtTpInf>');
  lines.push('\t\t\t\t<SvcLvl>');
  lines.push('\t\t\t\t\t<Cd>SEPA</Cd>');
  lines.push('\t\t\t\t</SvcLvl>');
  lines.push('\t\t\t\t<LclInstrm>');
  lines.push('\t\t\t\t\t<Cd>CORE</Cd>');
  lines.push('\t\t\t\t</LclInstrm>');
  lines.push('\t\t\t\t<SeqTp>RCUR</SeqTp>');
  lines.push('\t\t\t</PmtTpInf>');
  lines.push('\t\t\t<ReqdColltnDt>' + params.collectionDate + '</ReqdColltnDt>');

  // Creditor
  lines.push('\t\t\t<Cdtr>');
  lines.push('\t\t\t\t<Nm>' + escapeXml_(SEPA_CREDITOR_NAME) + '</Nm>');
  lines.push('\t\t\t\t<PstlAdr>');
  lines.push('\t\t\t\t\t<StrtNm>' + escapeXml_(SEPA_CREDITOR_STREET) + '</StrtNm>');
  lines.push('\t\t\t\t\t<PstCd>' + escapeXml_(SEPA_CREDITOR_POSTCODE) + '</PstCd>');
  lines.push('\t\t\t\t\t<TwnNm>' + escapeXml_(SEPA_CREDITOR_TOWN) + '</TwnNm>');
  lines.push('\t\t\t\t\t<Ctry>' + SEPA_CREDITOR_COUNTRY + '</Ctry>');
  lines.push('\t\t\t\t</PstlAdr>');
  lines.push('\t\t\t</Cdtr>');

  // Creditor Account
  lines.push('\t\t\t<CdtrAcct>');
  lines.push('\t\t\t\t<Id>');
  lines.push('\t\t\t\t\t<IBAN>' + SEPA_CREDITOR_IBAN + '</IBAN>');
  lines.push('\t\t\t\t</Id>');
  lines.push('\t\t\t</CdtrAcct>');

  // Creditor Agent
  lines.push('\t\t\t<CdtrAgt>');
  lines.push('\t\t\t\t<FinInstnId>');
  lines.push('\t\t\t\t\t<BIC>' + SEPA_CREDITOR_BIC + '</BIC>');
  lines.push('\t\t\t\t</FinInstnId>');
  lines.push('\t\t\t</CdtrAgt>');

  // Charge Bearer
  lines.push('\t\t\t<ChrgBr>SLEV</ChrgBr>');

  // Creditor Scheme Id
  lines.push('\t\t\t<CdtrSchmeId>');
  lines.push('\t\t\t\t<Id>');
  lines.push('\t\t\t\t\t<PrvtId>');
  lines.push('\t\t\t\t\t\t<Othr>');
  lines.push('\t\t\t\t\t\t\t<Id>' + escapeXml_(SEPA_CREDITOR_ID) + '</Id>');
  lines.push('\t\t\t\t\t\t\t<SchmeNm>');
  lines.push('\t\t\t\t\t\t\t\t<Prtry>SEPA</Prtry>');
  lines.push('\t\t\t\t\t\t\t</SchmeNm>');
  lines.push('\t\t\t\t\t\t</Othr>');
  lines.push('\t\t\t\t\t</PrvtId>');
  lines.push('\t\t\t\t</Id>');
  lines.push('\t\t\t</CdtrSchmeId>');

  // Transactions
  for (var i = 0; i < families.length; i++) {
    var fam = families[i];
    var seq = padZero_(i + 1, 3);
    var mandateId = params.year + seq;
    // EndToEndId: max 35 chars, SEPA charset (a-zA-Z0-9/-?:().,'+space)
    var rawE2E = mandateId + '/' + (fam.cognoms || 'N') + '/' + dateCompact;
    var endToEndId = rawE2E.replace(/[^a-zA-Z0-9\/\-?:().,'+\s]/g, '').substring(0, 35);

    lines.push('\t\t\t<DrctDbtTxInf>');

    // Payment Id
    lines.push('\t\t\t\t<PmtId>');
    lines.push('\t\t\t\t\t<EndToEndId>' + escapeXml_(endToEndId) + '</EndToEndId>');
    lines.push('\t\t\t\t</PmtId>');

    // Amount
    lines.push('\t\t\t\t<InstdAmt Ccy="EUR">' + params.amount.toFixed(2) + '</InstdAmt>');

    // Mandate
    lines.push('\t\t\t\t<DrctDbtTx>');
    lines.push('\t\t\t\t\t<MndtRltdInf>');
    lines.push('\t\t\t\t\t\t<MndtId>' + mandateId + '</MndtId>');
    lines.push('\t\t\t\t\t\t<DtOfSgntr>' + fam.dtOfSgntr + '</DtOfSgntr>');
    lines.push('\t\t\t\t\t</MndtRltdInf>');
    lines.push('\t\t\t\t</DrctDbtTx>');

    // Debtor Agent
    lines.push('\t\t\t\t<DbtrAgt>');
    lines.push('\t\t\t\t\t<FinInstnId>');
    lines.push('\t\t\t\t\t\t<BIC>' + escapeXml_(fam.swift) + '</BIC>');
    lines.push('\t\t\t\t\t</FinInstnId>');
    lines.push('\t\t\t\t</DbtrAgt>');

    // Debtor — full name (cognoms + nom), structured address, no Id block (no DNI data)
    var debtorName = fam.cognoms;
    if (fam.nom) debtorName += ' ' + fam.nom;
    // Nm max 70 chars (pain.008.001.02)
    if (debtorName.length > 70) debtorName = debtorName.substring(0, 70);

    lines.push('\t\t\t\t<Dbtr>');
    lines.push('\t\t\t\t\t<Nm>' + escapeXml_(debtorName) + '</Nm>');
    lines.push('\t\t\t\t\t<PstlAdr>');
    if (fam.adreca) {
      // StrtNm max 70 chars
      lines.push('\t\t\t\t\t\t<StrtNm>' + escapeXml_(fam.adreca.substring(0, 70)) + '</StrtNm>');
    }
    if (fam.cp) {
      // PstCd max 16 chars
      lines.push('\t\t\t\t\t\t<PstCd>' + escapeXml_(fam.cp.substring(0, 16)) + '</PstCd>');
    }
    if (fam.poblacio) {
      // TwnNm max 35 chars
      lines.push('\t\t\t\t\t\t<TwnNm>' + escapeXml_(fam.poblacio.substring(0, 35)) + '</TwnNm>');
    }
    lines.push('\t\t\t\t\t\t<Ctry>ES</Ctry>');
    lines.push('\t\t\t\t\t</PstlAdr>');
    lines.push('\t\t\t\t</Dbtr>');

    // Debtor Account
    lines.push('\t\t\t\t<DbtrAcct>');
    lines.push('\t\t\t\t\t<Id>');
    lines.push('\t\t\t\t\t\t<IBAN>' + escapeXml_(fam.iban) + '</IBAN>');
    lines.push('\t\t\t\t\t</Id>');
    lines.push('\t\t\t\t</DbtrAcct>');

    // Remittance
    lines.push('\t\t\t\t<RmtInf>');
    lines.push('\t\t\t\t\t<Ustrd>' + escapeXml_(params.remittance) + '</Ustrd>');
    lines.push('\t\t\t\t</RmtInf>');

    lines.push('\t\t\t</DrctDbtTxInf>');
  }

  lines.push('\t\t</PmtInf>');
  lines.push('\t</CstmrDrctDbtInitn>');
  lines.push('</Document>');

  return lines.join('\n');
}

/* =========================
   Helpers
   ========================= */

/**
 * Escapa caràcters especials per a XML.
 */
function escapeXml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Formata una Date com a YYYY-MM-DD.
 */
function formatDate_(d) {
  return d.getFullYear() + '-' + padZero_(d.getMonth() + 1, 2) + '-' + padZero_(d.getDate(), 2);
}

/**
 * Retorna la data resultant d'avançar N dies hàbils (dilluns-divendres)
 * a partir d'una data base.
 */
function nextBusinessDay_(from, businessDays) {
  var d = new Date(from.getTime());
  var added = 0;
  while (added < businessDays) {
    d.setDate(d.getDate() + 1);
    var dow = d.getDay(); // 0=diumenge, 6=dissabte
    if (dow !== 0 && dow !== 6) added++;
  }
  return d;
}

/**
 * Afegeix zeros a l'esquerra.
 */
function padZero_(n, len) {
  var s = String(n);
  while (s.length < len) s = '0' + s;
  return s;
}

// Codi font i docs: https://github.com/afa-la-serreta/afa-script — No editeu aquí, useu clasp push.
/****************
 * parsing.gs
 *
 * Parseja les respostes del formulari de Google Forms
 * i les converteix en objectes normalitzats per al clustering.
 *
 * Depèn de: utils.gs
 ****************/

/**
 * Parseja una fila de respostes del formulari.
 * Retorna un objecte amb camps normalitzats per al clustering.
 */
function parseResponseRow_(headers, row, sheetRowNumber) {
  const o = {};
  for (let i = 0; i < headers.length; i++) {
    o[headers[i]] = row[i];
  }

  // Corregeix files antigues on la columna DNI contenia el Nom
  fixColumnShift_(o);

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
    timestamp: o['Marca temporal'] || null,
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

/**
 * Corregeix el desplaçament de columnes en respostes antigues (files 2-17 aprox.)
 * on el formulari no tenia camp DNI: la columna DNI conté el Nom.
 * Detecció: Nom buit + DNI sense cap dígit (sembla un nom, no un document).
 */
function fixColumnShift_(o) {
  // Tutor 1
  const nom1 = (o['Nom:'] == null) ? '' : String(o['Nom:']).trim();
  const dni1 = (o['DNI:'] == null) ? '' : String(o['DNI:']).trim();
  if (!nom1 && dni1 && !/\d/.test(dni1)) {
    o['Nom:'] = o['DNI:'];
    o['DNI:'] = '';
  }

  // Tutor 2
  const nom2 = (o['Nom:.1'] == null) ? '' : String(o['Nom:.1']).trim();
  const dni2 = (o['DNI:.1'] == null) ? '' : String(o['DNI:.1']).trim();
  if (!nom2 && dni2 && !/\d/.test(dni2)) {
    o['Nom:.1'] = o['DNI:.1'];
    o['DNI:.1'] = '';
  }
}

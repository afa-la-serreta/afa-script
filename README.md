# AFA Families – Apps Script

Aquest projecte és una aplicació basada en **Google Apps Script** per gestionar
les dades de les famílies sòcies de l'AFA a partir d'un **Google Form**.

L'objectiu és:
- tenir **una fitxa canònica per família**
- deduplicar automàticament respostes repetides
- mantenir **traçabilitat completa** de les respostes originals
- permetre edició posterior via *magic link*
- desactivar automàticament famílies que ja han acabat l'escola
- enviar correus de confirmació de dades (amb batching i quota)
- permetre baixa voluntària via webapp

---

## Arquitectura (resum mental)

**Font de veritat**
- `Respuestas de formulario 1` (no es modifica mai)

**Taula canònica**
- `Famílies`
  - 1 fila = 1 família
  - pot agrupar múltiples respostes del formulari
  - conté traça (`source_rows`, `source_count`, `dedup_reasons`)
  - `status`: ACTIVE / INACTIVE

**Cua de revisió**
- `Possibles duplicats`
  - suggeriments "soft" (cognoms, adreça, infants)
  - no s'aplica automàticament

---

## Fulls de Google Sheets

### `Respuestas de formulario 1`
- Respostes originals del Google Form
- **No modificar ni esborrar files**

### `Famílies`
- Vista canònica
- Camps importants:
  - `family_id`
  - `token_edit`
  - `status` (ACTIVE / INACTIVE)
  - `inactive_reason`, `inactive_at`
  - `source_rows` → files del form que s'han fusionat
  - `source_count`
  - `dedup_reasons`
  - `confirmation_sent_at` → data d'enviament del correu de confirmació

### `BIC Bancs`
- Taula de conversió codi d'entitat → BIC/SWIFT
- Columnes: `Codi Entitat`, `BIC`, `Nom Entitat`
- Es crea automàticament quan cal
- Es pot editar manualment per afegir o corregir BICs
- Nous codis es resolen automàticament via IBANAPI (si la clau API està configurada)

### `Possibles duplicats`
- Resultat de `findPotentialDuplicatesAll()`
- Camps a omplir manualment:
  - `decision` → MERGE / NOT_DUPLICATE
  - `winner` → A / B

---

## Fitxers del projecte

| Fitxer | Responsabilitat |
|---|---|
| `config.js` | Constants (noms de fulls, URL webapp, configuració email, dades creditor SEPA, IBANAPI) |
| `utils.js` | Helpers purs: normalització, hashing, timestamps, lògica de graduació, emmascarament IBAN, BIC lookup (full + IBANAPI) |
| `parsing.js` | Parseig de respostes del formulari, correcció de columnes (`fixColumnShift_`) |
| `dedup.js` | Clustering Union-Find, merge a fila canònica (amb suport overwrite), scoring, upsert incremental |
| `families.js` | Entry points: `onOpen`, `importAll`, `findPotentialDuplicatesAll`, `onFormSubmit`, `syncEdited`, `deactivateGraduatedFamilies`, `sendConfirmationEmails`, `previewAllConfirmationEmails` |
| `webapp.js` | Web app per edició via magic link i baixa voluntària |
| `sepa.js` | Generació XML SEPA Direct Debit (pain.008.001.02 / Cuaderno 19.44) |
| `sepa_dialog.html` | Formulari HTML per als paràmetres SEPA (data picker, import, concepte) |
| `index.html` | Frontend del formulari d'edició |
| `bic_bancs_seed.csv` | Dades inicials per al full `BIC Bancs` (16 bancs extrets del XML 2024-2025) |

---

## Menú AFA (Google Sheets)

El menú **AFA** apareix automàticament al obrir el full de càlcul (`onOpen`):

| Opció | Funció | Descripció |
|---|---|---|
| 1️⃣ Desactivar famílies graduades (6è) | `deactivateGraduatedFamilies()` | Desactiva famílies on tots els infants han acabat 6è. |
| 2️⃣ Enviar correu de confirmació | `sendConfirmationEmails()` | Envia correus de confirmació a famílies actives (amb batching i quota). |
| 3️⃣ Generar fitxer SEPA (rebuts) | `generateSepaXml()` | Genera XML SEPA pain.008 per al cobrament de quotes i el desa a Google Drive. |
| 🔧 Avançat → 📥 Importar totes les respostes | `importAll()` | Reimporta tot des de zero. Aplica graduació automàticament. |
| 🔧 Avançat → 🔍 Buscar possibles duplicats | `findPotentialDuplicatesAll()` | Escriu suggeriments a `Possibles duplicats`. |

---

## Flux normal d'ús

### Import inicial (només el primer cop)
1. Buidar `Famílies` (deixar headers)
2. Executar **AFA → 🔧 Avançat → Importar totes les respostes**
3. (Opcional) Executar **AFA → 🔧 Avançat → Buscar possibles duplicats**

### Operació normal (automàtica)
- Cada nova resposta del form:
  - dispara `onFormSubmit` (trigger instal·lable)
  - actualitza o crea una família
  - aplica graduació automàticament
  - manté traça
- Respostes editades (via edit URL del form):
  - `syncEdited` les detecta automàticament (trigger temporitzat cada 15 min)

### Procés anual de cobrament de quotes
1. **Desactivar graduats**: AFA → 1️⃣ Desactivar famílies graduades (6è)
2. **Enviar correus**: AFA → 2️⃣ Enviar correu de confirmació
3. **Generar SEPA**: AFA → 3️⃣ Generar fitxer SEPA (rebuts)

### Enviament de correus de confirmació (pas 2)
1. Si hi ha un enviament anterior en curs, demana: **Continuar** o **Començar de nou**
2. Mostra la quota diària restant i envia fins al límit
3. Marca cada família enviada a la columna `confirmation_sent_at`
4. L'endemà, tornar a executar per enviar els pendents

El correu inclou:
- IBAN emmascarar (últims 4 dígits), o avís si l'IBAN és invàlid
- Infants actius amb el curs calculat (I3..6è); els graduats no apareixen
- Adreça registrada (o avís si no n'hi ha)
- Suggeriment d'actualitzar dades si tenen nous infants o canvis
- Botó per editar la resposta al Google Form
- Enllaç per donar-se de baixa (baixa voluntària via webapp)

Per previsualitzar tots els correus sense enviar-los, executar `previewAllConfirmationEmails()` des de l'editor d'Apps Script. Genera un fitxer HTML a Google Drive.

### Generació de rebuts SEPA (pas 3)
1. S'obre un formulari amb data picker, import (per defecte 35 EUR) i concepte pre-omplerts
2. El BIC/SWIFT es resol automàticament per cada família:
   - Primer consulta el full `BIC Bancs` (cache)
   - Si no el troba, consulta IBANAPI (i guarda el resultat al full)
   - Fallback: columna `bank_swift` de `Famílies`
3. Si hi ha famílies amb problemes (IBAN invàlid, BIC desconegut, dades incompletes):
   - Es mostra la llista completa de problemes
   - L'usuari pot continuar sense elles o cancel·lar per corregir-les
4. El fitxer XML (pain.008.001.02) es desa a Google Drive
5. Es mostra un enllaç per descarregar-lo i enviar-lo al banc

---

## BIC/SWIFT (resolució automàtica)

El BIC es deriva automàticament a partir del codi d'entitat de l'IBAN (posicions 4-7 per IBANs espanyols).

- **Full `BIC Bancs`**: taula editable amb les conversions conegudes. Es pot importar `bic_bancs_seed.csv` com a punt de partida.
- **IBANAPI**: per a codis d'entitat desconeguts, es consulta l'API automàticament (20 consultes/mes gratuïtes). Els resultats es guarden al full.
- **Configuració**: clau API a Script Properties (`IBANAPI_KEY`). Sense clau, funciona només amb el full.

---

## Lògica de graduació

- Un infant nascut l'any Y acaba 6è al juny de Y+12
- La funció `graduationCutoffYear_()` calcula dinàmicament l'any de tall
- S'aplica automàticament durant `importAll()` i `onFormSubmit()`
- El botó manual serveix per forçar una revisió puntual

---

## Decisions de disseny importants

- **No esborrem dades**: només marquem INACTIVE
- **Dedup automàtic només amb claus fortes** (email, telèfon, IBAN, DNI)
- La traça és obligatòria per auditoria i confiança
- `Respuestas de formulario 1` és immutable
- DNIs es validen (mínim 7 chars, 5 dígits) abans de fer hash

---

## Versionat

Aquest projecte està versionat amb GitHub i sincronitzat amb Apps Script via `clasp`.

No versionar:
- `.clasp.json` (conté IDs del projecte)

---

## Notes importants

- No canvieu noms de columnes sense revisar el codi
- No executeu `importAll()` en producció si ja s'han enviat tokens
- Qualsevol canvi en deduplicació ha de preservar traçabilitat

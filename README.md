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

### `Possibles duplicats`
- Resultat de `findPotentialDuplicatesAll()`
- Camps a omplir manualment:
  - `decision` → MERGE / NOT_DUPLICATE
  - `winner` → A / B

---

## Fitxers del projecte

| Fitxer | Responsabilitat |
|---|---|
| `config.js` | Constants (noms de fulls, URL webapp, configuració email) |
| `utils.js` | Helpers purs: normalització, hashing, timestamps, lògica de graduació, emmascarament IBAN |
| `parsing.js` | Parseig de respostes del formulari, correcció de columnes (`fixColumnShift_`) |
| `dedup.js` | Clustering Union-Find, merge a fila canònica (amb suport overwrite), scoring, upsert incremental |
| `families.js` | Entry points: `onOpen`, `importAll`, `findPotentialDuplicatesAll`, `onFormSubmit`, `syncEdited`, `deactivateGraduatedFamilies`, `sendConfirmationEmails` |
| `webapp.js` | Web app per edició via magic link i baixa voluntària |
| `index.html` | Frontend del formulari d'edició |

---

## Menú AFA (Google Sheets)

El menú **AFA** apareix automàticament al obrir el full de càlcul (`onOpen`):

| Opció | Funció | Descripció |
|---|---|---|
| 📥 Importar totes les respostes | `importAll()` | Reimporta tot des de zero. Aplica graduació automàticament. |
| 🔍 Buscar possibles duplicats | `findPotentialDuplicatesAll()` | Escriu suggeriments a `Possibles duplicats`. |
| 🔄 Sincronitzar edicions | `syncEdited()` | Sincronitza respostes editades (també es pot executar automàticament cada 15 min). |
| 🎓 Desactivar famílies graduades (6è) | `deactivateGraduatedFamilies()` | Desactiva famílies on tots els infants han acabat 6è. |
| ✉️ Enviar correu de confirmació | `sendConfirmationEmails()` | Envia correus de confirmació a famílies actives (amb batching i quota). |

---

## Flux normal d'ús

### Import inicial
1. Buidar `Famílies` (deixar headers)
2. Executar **AFA → Importar totes les respostes**
3. (Opcional) Executar **AFA → Buscar possibles duplicats**

### Operació normal
- Cada nova resposta del form:
  - dispara `onFormSubmit` (trigger instal·lable)
  - actualitza o crea una família
  - aplica graduació automàticament
  - manté traça
- Respostes editades (via edit URL del form):
  - `syncEdited` les detecta automàticament (trigger temporitzat cada 15 min)
  - També es pot executar manualment des del menú

### Enviament de correus de confirmació
1. Executar **AFA → ✉️ Enviar correu de confirmació**
2. Si hi ha un enviament anterior en curs, demana: **Continuar** o **Començar de nou**
3. Mostra la quota diària restant i envia fins al límit
4. Marca cada família enviada a la columna `confirmation_sent_at`
5. L'endemà, tornar a executar per enviar els pendents

El correu inclou:
- IBAN emmascarar (últims 4 dígits)
- Botó per editar la resposta al Google Form
- Enllaç per donar-se de baixa (baixa voluntària via webapp)

### Inici de curs (setembre)
- Executar **AFA → 🎓 Desactivar famílies graduades (6è)**
- O simplement reimportar: `importAll()` ja aplica la graduació

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

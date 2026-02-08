# AGENTS.md – Instructions for AI agents

This repository contains **Google Apps Script code with sensitive data logic**.

Agents must follow these rules strictly.

---

## Project structure

| File | Purpose |
|---|---|
| `config.js` | Sheet name constants, webapp URL, email config, SEPA creditor details, IBANAPI config |
| `utils.js` | Pure helpers: normalization, hashing, timestamps, graduation logic, grade calculation (`childGrade_`), IBAN masking, BIC lookup (sheet + IBANAPI) |
| `parsing.js` | Form response parsing, column shift fix |
| `dedup.js` | Union-Find clustering, canonical row merging (with overwrite support), scoring, incremental upsert |
| `families.js` | Entry points: `onOpen`, `importAll`, `findPotentialDuplicatesAll`, `onFormSubmit`, `syncEdited`, `deactivateGraduatedFamilies`, `sendConfirmationEmails`, `previewAllConfirmationEmails` |
| `webapp.js` | Web app for magic link editing and self-service deactivation (baixa voluntària) |
| `sepa.js` | SEPA Direct Debit XML generation (pain.008.001.02 / Cuaderno 19.44), BIC resolution pre-step |
| `sepa_dialog.html` | HTML form for SEPA params (date picker, amount, concept); calls `generateSepaWithParams()` via `google.script.run` |
| `index.html` | Edit form frontend |
| `bic_bancs_seed.csv` | Seed data for `BIC Bancs` sheet (16 banks from 2024-2025 XML) |

All `.js` files share a **global namespace** (Google Apps Script). No imports/exports.

---

## Golden rules

1. NEVER delete rows from `Respuestas de formulario 1`
2. NEVER silently remove deduplication or trace logic
3. NEVER regenerate `token_edit` for existing families
4. NEVER auto-merge "soft" duplicates without human confirmation
5. ALWAYS preserve `source_rows` / `source_count` / `dedup_reasons`
6. NEVER change graduation logic without understanding the school year calendar (P3 to 6th grade, 9 years)

---

## Data model invariants

- One row in `Famílies` represents one canonical family
- A family may aggregate multiple form responses
- Deduplication levels:
  - Strong: email, phone, IBAN, DNI → automatic
  - Soft: surnames, address, children → suggestion only
- Status lifecycle: ACTIVE → INACTIVE (never deleted)
- Graduation: a child born in year Y finishes 6th grade in June Y+12

---

## Key functions (dependency graph)

```
families.js
  ├── importAll()                    → parsing.js, dedup.js, utils.js
  ├── findPotentialDuplicatesAll()   → parsing.js, dedup.js
  ├── onFormSubmit()                 → parsing.js, dedup.js
  ├── syncEdited()                   → parsing.js, dedup.js (time-driven trigger)
  ├── deactivateGraduatedFamilies()  → utils.js
  ├── sendConfirmationEmails()       → utils.js, config.js (quota-aware, batched, IBAN validation, children+address)
  └── previewAllConfirmationEmails() → utils.js (generates HTML preview to Drive, no emails sent)

dedup.js
  ├── clusterByStrongKeys_()    → utils.js
  ├── mergeClusterToCanonRow_(cluster, canonRow, idxC, overwrite) → utils.js
  ├── upsertSingleResponse_()  → utils.js, dedup.js (overwrite=true on update)
  └── scorePair_()              → utils.js

parsing.js
  └── parseResponseRow_()      → utils.js

utils.js (BIC resolution)
  ├── loadBicMap_()        → reads "BIC Bancs" sheet, caches per execution
  ├── lookupBicApi_(iban)  → calls IBANAPI, returns {bic, bankName} or null
  ├── saveBicToSheet_()    → appends row to "BIC Bancs" (creates sheet if needed)
  └── ibanToBic_(iban)     → sheet cache → API fallback → saves to sheet

webapp.js
  ├── doGet()              → edit page or baixa (deactivation) page
  ├── saveFamily()         → utils.js
  ├── deactivateFamily()   → sets inactive_reason='baixa voluntària'
  └── uses headerIndex_() from utils.js

sepa.js
  ├── generateSepaXml()        → shows sepa_dialog.html modal with defaults
  └── generateSepaWithParams() → config.js, utils.js (BIC resolution + XML generation, saves to Drive)
      (public name — called from HTML via google.script.run, cannot end in _)
```

## BIC resolution

- `ibanToBic_(iban)` extracts the 4-digit bank code from ES IBANs and resolves to BIC
- Source of truth: `BIC Bancs` sheet (editable by users)
- Auto-discovery: IBANAPI (free tier: 20 BIC lookups/month), results saved to sheet
- API key stored in Script Properties (`IBANAPI_KEY`); works without it (sheet-only mode)
- Bank codes with leading zeros: `loadBicMap_()` re-pads numeric codes to 4 digits (Sheets may strip zeros)
- SEPA generation resolves BICs as pre-step; shows all problems and lets user choose to continue or abort
- Debtor address uses structured XML fields (`StrtNm`, `PstCd`, `TwnNm`, `Ctry`) — NOT `AdrLine` (Sabadell rejects unstructured addresses)
- Debtor name: `cognoms + nom` (full name); no `<Id>` block (we don't have valid DNIs)

---

## Triggers

| Trigger | Type | Function | Purpose |
|---|---|---|---|
| `onOpen` | Simple | Menu setup | Adds AFA menu to spreadsheet |
| `onFormSubmit` | Installable (spreadsheet) | New submissions | Upserts new form responses instantly |
| `syncEdited` | Installable (time-driven, every 15 min) | Edited responses | Catches form edits (onFormSubmit doesn't fire on edits) |

`syncEdited` uses `PropertiesService` to track `syncEdited_lastRun` — only processes responses with timestamps newer than the last run.

---

## Email sending

- `sendConfirmationEmails()` is **quota-aware** (checks `MailApp.getRemainingDailyQuota()`)
- Stamps `confirmation_sent_at` on each family after successful send
- On re-run, detects previous batch and asks: **Continue** (send remaining) or **Start over** (clear marks, re-send all)
- Shows batch start date and count of already-sent emails in the dialog
- Stops early if quota is exhausted mid-send
- Free Gmail: 100/day; Google Workspace: 1,500/day
- Validates IBAN and shows warning in email if invalid (prompts family to update)
- Shows active children with calculated grade (I3..6è) via `childGrade_(birthYear)` in `utils.js`; graduated children are excluded
- Shows registered address; suggests updating if new children started school
- `previewAllConfirmationEmails()` generates all emails as an HTML file to Google Drive (no sending) for testing

## Yearly workflow

1. **Desactivar famílies graduades (6è)** — deactivate families where all children finished primary school
2. **Enviar correu de confirmació** — send data confirmation emails to active families
3. **Generar fitxer SEPA (rebuts)** — generate direct debit XML for the bank

---

## Allowed changes

- Refactor code for readability
- Improve performance (same semantics)
- Add logging or comments
- Make header matching more robust
- Add safeguards / validations
- Split files further if they grow too large

---

## Forbidden changes

- Removing audit/trace columns
- Changing sheet names without config update (includes `BIC Bancs`)
- Replacing deterministic dedup with probabilistic logic
- Writing back to the form responses sheet
- Auto-applying decisions from `Possibles duplicats`
- Changing `WEBAPP_URL` in config.js without redeploying the webapp first
- Changing `graduationCutoffYear_()` without understanding the Catalan school calendar

---

## When in doubt

If a change could:
- lose information
- reduce traceability
- change dedup behavior
- change graduation/deactivation behavior

→ STOP and ask the human.

---

## Tooling & MCP integrations

- Code is deployed via **clasp** (Google Apps Script CLI)
- MCP servers available:
  - `clasp-enhanced` – push, pull, deploy, version, and run Apps Script projects
  - `google-docs-mcp` – read/write Google Docs, Sheets, and Drive files
- Use `clasp push` (or the MCP tool) to deploy; never edit code directly in the Apps Script editor
- `clasp push` updates the project code; **redeploy** (`clasp deploy -i <ID>`) is only needed for webapp changes
- `.claspignore` excludes `node_modules/`, `.xlsx`, `.md`, and `.git/`

---

## Mental model reminder

This system prioritizes:
1. Trust
2. Auditability
3. Reversibility

Over:
- clever automation
- aggressive merging
- compact data

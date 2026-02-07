# AGENTS.md – Instructions for AI agents

This repository contains **Google Apps Script code with sensitive data logic**.

Agents must follow these rules strictly.

---

## Project structure

| File | Purpose |
|---|---|
| `config.js` | Sheet name constants |
| `utils.js` | Pure helpers: normalization, hashing, timestamps, graduation logic |
| `parsing.js` | Form response parsing, column shift fix |
| `dedup.js` | Union-Find clustering, canonical row merging, scoring, incremental upsert |
| `families.js` | Entry points: `onOpen`, `importAll`, `findPotentialDuplicatesAll`, `onFormSubmit`, `deactivateGraduatedFamilies` |
| `webapp.js` | Web app for magic link editing |
| `index.html` | Edit form frontend |

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
  ├── importAll()         → parsing.js, dedup.js, utils.js
  ├── findPotentialDuplicatesAll() → parsing.js, dedup.js
  ├── onFormSubmit()      → parsing.js, dedup.js
  └── deactivateGraduatedFamilies() → utils.js

dedup.js
  ├── clusterByStrongKeys_()    → utils.js
  ├── mergeClusterToCanonRow_() → utils.js
  ├── upsertSingleResponse_()  → utils.js, dedup.js (self)
  └── scorePair_()              → utils.js

parsing.js
  └── parseResponseRow_()      → utils.js

webapp.js
  └── uses headerIndex_() from utils.js
```

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
- Changing sheet names without config update
- Replacing deterministic dedup with probabilistic logic
- Writing back to the form responses sheet
- Auto-applying decisions from `Possibles duplicats`
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

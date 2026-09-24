# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A single-page, client-side-only web app that audits Odessa Separator Inc. Sales Order PDFs against deterministic business rules. No backend, no build step, no package manager — three static files (`index.html`, `app.js`, `styles.css`) served as-is (e.g. via GitHub Pages).

## Running it

There is no dev server or build. Open `index.html` directly in a browser, or serve the directory with any static file server (needed for the PDF.js worker to load correctly via `file://` in some browsers):

```bash
python3 -m http.server 8000
```

There are no tests, linter, or package.json in this repo.

## Architecture

Everything lives in `app.js` (vanilla JS, ES modules, no framework, no bundler):

- **PDF text extraction** (`extractPDFText`): Uses PDF.js (loaded from CDN in `index.html` and imported in `app.js`) to pull text items with x/y coordinates per page, then reconstructs visual lines by clustering items within 3px of the same y-coordinate and sorting by x. This line reconstruction is foundational — all downstream parsing operates on these joined lines, not raw PDF.js text order.
- **SO field parsing** (`parseSO`): Regex/positional parsing tuned specifically to the Odessa Separator PDF layout (Sold To/Ship To blocks, stacked Total/Tax/Subtotal values, phone/contact lines, delivery date/time, STAMP note phrases). This has gone through many incremental fixes (see version history in README.md) because PDF.js line/column splitting is inconsistent across PDF exports — expect fragile edge cases here.
- **Item/line parsing**: Finds every unique Odessa part number (pattern `\d{2,4}-\d{4}-\d{2}(?:-\d{2})?`) in the item-table area of the document, then for each PN parses the first `Qty UnitPrice/EA Amount` pattern that follows it in the text segment up to the next PN. Deliberately does not assume PN and its data are on the same PDF.js line, since column layout varies.
- **Price List**: An OSI-provided Excel workbook loaded client-side via SheetJS (CDN), matched using the `PN LIST` sheet's `PartNumber`/`Pricing_UnitPrice0` columns. Persisted to `localStorage` (`soPriceList`) so it survives reloads.
- **Audit rules** (`runAudit`): Deterministic checks — Sold To vs Ship To consistency, phone presence, delivery charge tier (END USER/AFTER HOURS/no-charge based on day/time and customer), Inspection Fee requirement (PN `1014-0142-00` triggers requirement for `1004-0009-00`), item price vs. Price List, tax calculation (8.25% for taxable customers, $0 for Coterra/Diamondback), STAMP phrase requirement (XTO), all-items-priced check, plus soft suggestions (GRS/Reverse Flow observation, Diamondback AFE & GL note).
- **State/persistence**: All state is in-memory (`state` object) plus `localStorage` for the price list, a single draft (`soDraft`), and audit history (`soAuditHistory`, capped at 100 entries). No server-side storage anywhere — this is a stated privacy requirement (see README.md "Privacidad").

### Rules configuration gotcha

`rules.json` exists at the repo root and documents the intended rule set, but **`app.js` does not read it** — the same values are hardcoded in the `RULES` const near the top of `app.js` (app.js:9-21). If you change audit rules, update `RULES` in `app.js`; `rules.json` is currently just documentation/reference and will silently drift out of sync unless someone wires it up as the actual source of truth.

### Customer matching

Customer names from the PDF are normalized via `customerKey()` (app.js:49-59), which does substring matching against known customer names (DIAMONDBACK, COTERRA, XTO, BTA, SUMMIT, BURLESON, APACHE) and falls back to the raw normalized string for unrecognized customers (tax/stamp rules then default to "manual review required").

## External dependencies

Both loaded from CDN, no local vendoring (noted in README.md as a known V1 limitation/future improvement):
- PDF.js `4.10.38` (cdnjs) — PDF parsing
- SheetJS `xlsx-0.20.3` (cdn.sheetjs.com) — Excel Price List parsing

## Working in this codebase

- This app parses one specific, evolving PDF vendor format (Odessa Separator SOs). When fixing parser issues, check the README.md version history first (V1.1–V1.9) — many fixes were narrow corrections for specific PDF.js text-layout quirks, and a "fix" for one PDF layout can regress another. Prefer testing against multiple real sample SO PDFs before assuming a parser change is correct.
- UI strings and status messages are in Spanish; keep new user-facing strings consistent with that.
- No secrets, API keys, or backend calls exist in this app by design — do not introduce any (see README.md "Privacidad").

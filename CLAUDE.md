# LedgerLink — project brief for Claude Code

Read this file fully before every task. It is the single source of truth.

## What we are building
A portal for an Indian CA firm (R.K. Tantry & Co) that turns a client's Tally trial balance (TB)
into finished financial statements: schedules/notes, Balance Sheet, Profit & Loss, PPE, in Excel.
The firm has ~100 clients (proprietorships, partnerships, LLPs, private companies), each with a
finalised Excel workbook for the last year. Every year a new TB arrives and the workbook must be
rebuilt for the new year automatically, in that client's own layout.

## Stack
- frontend/ (this folder's React + Vite app; move it into frontend/ in Phase 1): React 18, react-router, plain CSS.
- backend/: Python 3.11+, FastAPI, openpyxl, pandas, SQLAlchemy with SQLite (file: backend/ledgerlink.db).
  Keep the DB layer simple so it can switch to PostgreSQL later.
- Spreadsheet in the browser: Univer (open source, @univerjs packages). Must feel like Excel:
  formulas, formula bar, sheet tabs, keyboard shortcuts, editing.
- AI layer (Phase 7 only): rapidfuzz + sentence-transformers locally; optional Ollama (Qwen) for explanations.
  No paid or cloud AI APIs. Client data must never leave the machine.

## Design rules (do not change the look)
- The current UI in src/ is the approved design. Keep fonts (Geist, Geist Mono), colours, spacing, layout exactly.
- All colours come from the tokens at the top of src/styles.css. Black and white, plum accent #5B2C83.
- NEVER use blue anywhere.
- Review colours: green = linked as last year, amber = AI suggestion, red = needs CA, plum = edited by CA.
- Indian number format everywhere (1,36,74,167.85). Dates like 12 Sep 2026.

## Screens and flow
1. Clients (/): table of clients: name + entity type, financial year, TB last updated, review status.
   Hover a row shows a delete icon; delete requires typing DELETE. Search box. "Add client" button.
2. Add client (/add): name, entity type, financial year; two choice boxes:
   a) "Finished workbook": upload a completed year's Excel → LEARN MODE.
   b) "Trial balance only": upload a Tally TB, pick sheets to create → BUILD MODE.
3. Processing (/processing/:jobId): real progress steps streamed from the backend
   (Read TB → Find sheets → Trace links → Check totals), with a live log of formulas traced.
4. Workbook (/client/:id): the Excel embedded (Univer) with sheet tabs, formula bar, and under it a
   plain-words trace ("Note 9 Cash & Bank from TB: Bank of Baroda, …"). Colour-coded cells.
   Right panel: checks + review items (Approve / Change) + legend + "Finalise FY" (disabled until clear).
   Buttons: "Download Excel", "Update TB for FY <next>" → ROLL-FORWARD MODE.

## Domain knowledge
### Tally TB export (see samples/Venus_Agencies_2025-26.xlsx, sheet "TB")
- Sheet name varies: "TB", "Trial Balance", "TB 25-26", etc. Detect by name first, then by content
  (header row containing "Particulars" with Debit/Credit columns).
- Company name, address, GSTIN, period on the rows above the header.
- Row types by formatting: GROUP = bold, indent 0. SUB-GROUP = not italic, indent > 0.
  LEDGER = italic, indented. Group rows carry subtotals; never double count them.
- Debit in one column, Credit in the next. Last row "Grand Total". Debit total must equal credit total.
- Special Tally lines: "Closing Stock" (Cr), "Opening Stock", "Profit & Loss A/c", "Inventory Difference".

### How a finished workbook links (Venus example)
- Schedules/notes pull from TB by cell reference (=TB!B48) or arithmetic (=TB!C24-TB!B24).
- Balance Sheet and P&L pull only from notes. P&L profit → capital note → Balance Sheet.
- PPE sheet computes depreciation (half rate for additions after 3 Oct); depreciation → P&L notes; closing WDV → BS.
- Proprietorship: personal income (FD interest, gold bond gain, SB interest, LIC) is credited to capital, not P&L.
- Some cells are typed numbers or constants like =3540+6080 (manual). Treat as MANUAL inputs.
- Prior-year column is mostly typed values.
- Known issues in the Venus file (use as review test cases):
  BS row 21 current year links 'BS Schedules'!C74 but prior year links D76 (total) → red.
  Insurance in P&L Notes is hardcoded =3540+6080 instead of TB ledgers → red/manual.
  TB ledger "GODS A/C" (3,384 Cr) is not used anywhere → red.
  Closing stock = TB Closing Stock + Inventory Difference 6,23,000 → amber, confirm.
- Expected Venus FY 2025-26 figures (tests must match to the paisa):
  Total Capital & Liabilities = Total Assets = 17,328,403.28; Profit = 1,669,525.88;
  Owners' funds = 13,674,167.85; Cash & bank = 2,221,798.16; Depreciation = 35,425.45.

## Core engine
### LEARN MODE (finished workbook)
1. Load twice with openpyxl (formulas, and data_only=True for values).
2. Find the TB sheet and parse it into ledgers with group path, Dr, Cr, row number.
3. For every formula cell in every sheet: parse references (sheet!cell, ranges, SUM), resolve any TB
   reference to the ledger NAME at that row and column (Dr/Cr). Build a dependency graph.
4. Classify every numeric cell: TB-linked, internal-linked (other sheet), manual/typed, prior-year.
5. Store per client: a MAPPING keyed by ledger name/group path (never by row number), the workbook
   as the template, and manual inputs. Save the original file too.
6. Report: unused TB ledgers, manual cells, inconsistencies (like row 21 above).

### ROLL-FORWARD MODE (new year TB)
1. Clone the client's latest finalised workbook (keep layout, formatting, sheets).
2. Current-year values become the prior-year column as fixed numbers; update headings and dates.
3. Replace the TB sheet with the new TB; rewrite every TB formula to point at the ledger's NEW row.
4. Carry forward: closing capital → "As per last Balance Sheet"; PPE closing → opening; closing stock → opening stock check.
5. Status per cell: green (same ledger found), amber (AI/fuzzy match for new or renamed ledger, with a reason),
   red (unmapped, missing ledger, or manual input needed). New TB ledgers not used anywhere → red.
6. Checks: TB Dr = Cr, BS difference = 0, every TB ledger used, opening capital = last closing.
7. CA corrections are saved to the mapping so next year repeats them.

### BUILD MODE (TB only)
Default templates by entity type: Non-corporate (ICAI guidance note format) and Company (Schedule III, Division I).
Default mapping from Tally primary groups to BS/P&L lines. Generate sheets the user picked, with live formulas to TB.

### Excel output
Generated with openpyxl with LIVE formulas (never pasted values for current year), same sheets and format as the
client's workbook. Font and formatting copied from the template.

## Working rules
- Work in the phase you are given. Do not start later phases.
- Only make changes the phase asks for. No extra features, abstractions or files beyond what is needed.
- Stop and ask me before: deleting files, adding a dependency not named in this file, changing the DB schema
  after Phase 2, or changing the UI design.
- Write pytest tests for backend logic using samples/Venus_Agencies_2025-26.xlsx. Run them before saying done.
- After each step print: ✅ <what was completed>. At the end list every file created or changed and how to run it.
- Keep code readable for a small team: clear names, short functions, comments where the accounting logic lives.

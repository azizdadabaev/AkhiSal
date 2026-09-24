# Donabay

A weekly crew-pay workbook for a filler-block plant: production log, cash
advances, and payday. Everything lives in **`Donabay.xlsx`**, which holds real
business data.

## Read first

- `docs/Donabay_project_notes.md` covers the business rules, the layout of
  every sheet, the design system, the change log and known quirks. It is
  exported from the workbook's own **Project Notes** sheet. Update both
  together.

## Rules for changing the workbook

- **Edit `Donabay.xlsx` in place with openpyxl.** Never regenerate it:
  `builder/build_donabay.py` only reproduces the old v1 layout. v2 shifted
  every row down by one and added Home, a nav bar and Guide. The builder
  writes to `build/` so it can't overwrite the real file.
- **openpyxl drops the Office add-in links on save**, including the Claude
  for Excel add-in's. After every save, run
  `python3 builder/keep_addins.py ORIGINAL.xlsx Donabay.xlsx` with a copy
  taken before editing.
- The calendar starts **Mon 31 Aug 2026** (moved back from 21 Sep to hold
  the old hand-kept sheet; see `builder/import_history_2026_09.py`). Amber
  cells with comments are waiting for the owner to confirm them. Don't
  change their values unless asked.
- Never save a workbook that was opened with `data_only=True`. That replaces
  every formula with its value.
- Use classic Excel functions only: SUMIFS, SUMPRODUCT, INDEX/MATCH,
  COUNTIFS, IFERROR. No LAMBDA, XLOOKUP or dynamic arrays.
- Number format is `#,##0`. `# ##0` displays wrongly in English Excel.
- Pay rules: only good blocks are paid, the pot is split by days worked
  (1 / 0.5 / empty), pay never goes negative, and debt carries forward.
  Mistakes are fixed with a Correction row, never by deleting.

## Verifying a change

1. Recalculate a **copy**; the recalc script rewrites the file in place. This
   needs `libreoffice-calc`. With only `libreoffice-core` installed,
   recalculation hangs forever:
   `apt-get install -y --no-install-recommends libreoffice-calc`
2. Expect `total_errors: 0` and `Checks!C5` reading `ALL CHECKS OK`.
3. Check the pay figures by hand against the week's raw Daily Log and Cash
   Ledger rows. A clean recalculation only proves the formulas run.

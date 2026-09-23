# Donabay

Daily filler-block production and weekly crew pay for beam-and-block
production — in one Excel file. No macros, no add-ins, nothing to install.

**[`Donabay.xlsx`](Donabay.xlsx)** — open it and type in the yellow cells.

---

## First time (2 minutes)

1. **Settings** — type your pay rate per block next to the start date.
   The workbook flags this until you do; it will not guess your rate.
2. **Workers** — replace `Worker 1 … Worker 4` with the real names.

## Then

| When | Where | What you type |
|---|---|---|
| **Every day** | Daily Log | Blocks moulded, blocks broken, and `1` under each worker who came (`0.5` for half a day, empty if absent) |
| **Someone asks for money** | Cash Ledger | Date, worker, `Advance`, amount, reason |
| **Payday** | Weekly Pay | Pick the week, pay each worker the green **TO PAY** amount |
| | Cash Ledger | Record each payment as `Weekly pay` |
| | Payslip | Print one per worker; they sign. It lists every advance they took |
| **Someone joins** | Workers | Next empty slot, status `Active` |
| **Someone leaves** | Workers | Status `Left` — never delete the row, their history stays |
| **Rate changes** | Settings | Add a **new** row with the start date — old weeks keep the old rate |
| **A mistake in the ledger** | Cash Ledger | Add a `Correction` with the opposite amount — never delete |

## How the pay is worked out

```
paid blocks   = moulded − broken                     (only good blocks are paid)
pot           = Σ  paid blocks × rate on that day     (for the week)
worker share  = pot × worker's days ÷ crew's days     (1 = full day, 0.5 = half)
to pay        = brought forward + earned − advances   (never below zero)
```

Each worker has **one running account**: earnings go in, every som handed over
comes out. If advances run ahead of earnings, pay is `0` and the rest carries
into next week automatically — shown on the payslip every week until it's paid
back. Nothing is forgotten and nothing goes negative.

## What's built in from accounting and production practice

Three research passes (payroll accounting, block-plant production management,
spreadsheet design) informed the design. Adopted:

- **Advances are a receivable, not wages.** Gross earnings are never touched;
  advances sit below the line. That's what keeps the totals honest.
- **Net pay floored at zero, debt carried forward** — the standard treatment
  when advances exceed a period's earnings.
- **Crew pay split by days worked.** Standard for a crew sharing one line; the
  post-Soviet brigade system (КТУ) reduces to exactly this when every worker
  counts equally.
- **Dated rate table** — a rate change never rewrites past weeks.
- **Stable worker IDs, never reused** — renaming updates everywhere; removing
  keeps history.
- **Corrections, not deletions** — the ledger is an audit trail.
- **Self-checks** — shares must add up to the pot, every production day needs
  attendance, unknown names in the ledger are caught.
- **Reject-rate alarm** (normal breakage is 2–5%; alarm above 8%) and an
  optional **cement bags per 1000 blocks** check — the usual sign that counts
  are off.

Two options are in **Settings** but off by default, because they change your
stated rules:

- **Breakage allowance** — pay for normal breakage up to a set %. Default `0%`
  (pay only good blocks).
- **Old debt taken back per week, max** — Default `100%`. For formally employed
  workers, Uzbek labour law caps wage deductions at 50%. Confirm on
  [lex.uz](https://lex.uz) before relying on it.

Deliberately left out as overkill for a four-person, owner-run crew: dual
sign-off on advances, a separate curing-inspection sheet, per-worker skill
coefficients (they would silently rewrite past weeks), minimum-wage floors.
Details and sources: [`docs/design-notes.md`](docs/design-notes.md).

## Verified

`tests/verify_donabay.py` recomputes the pay from raw inputs in plain Python —
never reading the workbook's formulas — and compares every figure with what
the workbook calculated. **197 / 197 checks pass** across three scenarios: a
mid-week rate rise, half days and absences, an advance larger than earnings
carried into the next week, an overpayment, a correction, the 50% cap, the
payslip's advance list and the ledger's warnings.

```
python3 builder/build_donabay.py                      # Donabay.xlsx
python3 builder/build_donabay.py --demo T1.xlsx       # a filled test copy
```

---

## Also in this repo

An earlier general Excel toolkit: a client matcher that pairs Latin and
Cyrillic spellings, a phone normaliser, receivables aging, and amounts in
Uzbek words. It lives in [`AkhiSal-Tools.xlsx`](AkhiSal-Tools.xlsx), with VBA,
Office Scripts and LAMBDA versions under `vba/`, `office-scripts/` and
`lambdas/`. See [`docs/excel-toolkit.md`](docs/excel-toolkit.md).

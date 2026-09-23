# Donabay design notes

Why the workbook is built the way it is: what was taken from research, what was
changed to fit, and what was left out.

## Decisions from the owner

- **Output is a crew total, split by days worked.** Not per-worker counts.
- **Only good blocks are paid.** Broken blocks are recorded and not paid.

## Adopted

| Practice | Where | Why |
|---|---|---|
| Advance = receivable from the employee, recovered from net pay | Weekly Pay, Payslip | Keeps gross earnings clean; advances never inflate or distort wage totals |
| Net pay floored at 0, remainder carried forward | Weekly Pay `TO PAY` | Standard treatment of negative net pay; debt stays visible instead of vanishing |
| One running account per worker | Cash Ledger + Weeks | Carry-forward falls out of the arithmetic; no manual transfer between weeks |
| Split by days present | Weeks | Collective piecework is distributed by time worked; КТУ at 1.0 reduces to this |
| Half-day attendance | Daily Log | Time actually worked is the defensible basis for a crew split |
| Dated rate history, applied per day | Settings, Daily Log | Rate changes (even mid-week) never recalculate past weeks |
| Stable IDs, deactivate instead of delete | Workers | Name as key breaks on duplicates and renames; deleting destroys history |
| Flat ledger for cash, grid for attendance | Cash Ledger, Daily Log | Cash can happen several times a day with per-event detail, so it needs a ledger. Attendance is exactly one value per worker per day, where a grid is faster to fill and loses nothing |
| Corrections instead of deletions | Cash Ledger | Audit trail |
| Totals above data, over-provisioned fixed blocks | all | Appended rows never fall outside totals; no formula depends on how many rows are used |
| Named ranges behind dropdowns | Settings | Excel 2007 will not take a dropdown source from another sheet otherwise |
| Self-checks and a master status | Checks | Shares = pot, days = crew-days, no unattended production, no orphan ledger rows |
| Reject-rate alarm, cement per 1000 blocks | Weeks | The two metrics that catch real problems at this scale |
| Signed payslip listing each advance | Payslip | In a cash business the signature is the proof of payment and settles disputes |

## Changed to fit

- **Debt cap.** Research recommends capping recovery (Uzbek labour law: 50% of
  wages for employer deductions). The owner's rule is full recovery, so it is a
  setting defaulting to 100%. It only limits recovery of *old* debt; an advance
  taken this week is a part-payment of this week's wages, not a deduction.
- **Breakage allowance.** Research suggests absorbing normal breakage (~4%).
  The owner's rule is good blocks only, so it is a setting defaulting to 0%.

## Left out

- **Dual authorisation, authoriser ≠ payer.** Segregation of duties assumes
  staff; here the owner is both.
- **Separate QC / curing-inspection sheet.** Broken counts can be filled in on
  the day's row whenever they are known.
- **Per-worker skill coefficients (КТУ ≠ 1).** A coefficient stored on the
  Workers sheet would retroactively change every past week's split. Doing it
  safely needs a per-week coefficient grid; add it only if pay disputes arise.
- **Minimum daily wage floor, cost per block, downtime logging.** Not asked
  for; easy to add later.

## Sources

Research was done from search-engine extracts; direct page fetches were blocked
by the network proxy, so figures should be checked against the originals.

- Payroll advances as employee receivables — double-entry-bookkeeping.com,
  accountingtools.com, accountingcoach.com
- Negative net pay and carry-forward — docs.checkhq.com
- Uzbek deduction limits — gov.uz, kadrovik.uz, kadry.uz (verify on lex.uz)
- КТУ / brigade pay split — assistentus.ru, glavkniga.ru, nitt.by
- Block reject rates — reitmachine.com; ScienceDirect precast study
- Cement per 1000 blocks — Columbia Machine materials guide
- Spreadsheet structure — FAST Standard, SSRB, Microsoft data-organisation guidelines

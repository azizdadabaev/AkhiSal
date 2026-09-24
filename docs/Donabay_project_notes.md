<!-- Exported from the 'Project Notes' sheet of Donabay.xlsx. The sheet is the
     original; update both together. -->

# Donabay – Project Notes (handover)

_Last updated: 24 Sep 2026. File: **Donabay.xlsx**_

Weekly pay workbook for a filler-block plant. The crew is paid **per good block** and the weekly pot is **shared by days worked**. Advances are logged in a cash ledger and taken off the next pay. Plain Excel only: no macros, no add-ins.

---

## 1. How the business rules work

| Rule | How it works |
|---|---|
| Pay base | Only good blocks are paid: `Good = Moulded − Broken`. `Paid blocks = Moulded − MAX(0, Broken − BreakAllow × Moulded)`. BreakAllow is currently 0%, so paid = good. |
| Pay rate | Settings rate table: so'm per paid block, each rate with a start date. The rate in force on each day is looked up, so old weeks keep their old rate. Current rate: **500 so'm from 2026-08-31**. |
| Weekly pot | Sum of the week's daily pay values (paid blocks × rate). |
| Split | Worker share = pot × (worker's days ÷ total crew-days). Attendance: `1` = full day, `0.5` = half day, empty = absent. |
| Advances | Logged in the Cash Ledger as type **Advance** and taken off that week's pay. |
| Debt | Pay never goes negative. If a worker owes money, TO PAY is 0 and the debt carries forward. At most **DebtCap** of a week's earnings (currently 100%) goes to paying back old debt. Note: Uzbek labour law caps wage deductions at 50% for formally employed staff. |
| Weeks | Monday to Sunday. The first week starts **2026-08-31**. |
| Mistakes | Never delete a ledger row. Add a **Correction** row with the opposite amount. |

## 2. Sheets (in tab order)

| # | Sheet | Purpose | Frozen | Notes |
|---|---|---|---|---|
| 1 | **Home** | Dashboard, the file opens here | – | KPI cards (rows 6–8), Workbook health (B11:I23), Quick actions (K11:R19), Crew table (rows 29–41). Hidden helper T3 = row of the current week in Weeks. The week shown is the current calendar week (the Monday of today). |
| 2 | **Daily Log** | One row per day (rows 7:377) | A1:B6 | A Date · B Day · C Moulded · D Broken · E Good · F Paid blocks · **G:R attendance for worker slots W01–W12** · S Crew-days · T Rate · U Pay value · V Week of · W Check · X Notes. Today's row is highlighted green. Attendance only accepts values from 0 to 1. Some columns are grouped (collapsed). |
| 3 | **Cash Ledger** | Every so'm handed out (rows 7:1006) | rows 1:6 | A # · B Date · C Worker · D Type (Advance / Weekly pay / Correction) · E Amount · F Reason · G Given by · H Signed? · I ID · J Week of · K key · L Worker balance after · M Check. The totals row is row 4. I:K are grouped helper columns. |
| 4 | **Weekly Pay** | Payday screen | A1:B11 | Week picker in **C5**. Worker rows 12:23 (slots W01–W12), crew total in row 24. F Earned · G Brought forward · H Advances · I Corrections · J Due · **K TO PAY** · L Paid (in ledger) · M Still to pay · N Carried forward · O Note. **Q:T "Ready for the Cash Ledger"** lists the rows to copy and Paste Values into the ledger. U is a hidden helper column. |
| 5 | **Payslip** | Printable slip | row 1 | Week in C5, worker in C6. Lists that week's advances and has signature lines. |
| 6 | Worker History | One worker's account week by week | rows 1:7 | Worker picker in row 5. |
| 7 | Workers | 12 fixed slots W01–W12 (rows 6:17) | rows 1:5 | ID · Name · Status (Active/Left) · Joined · Left on · Phone · Notes · Check. Never delete or reuse a slot. |
| 8 | Settings | Rules and pay rates | row 1 | C6 BreakAllow · C7 DebtCap · C8 RejectAlarm (8%). Rate table B13:C32. Dropdown lists F6:H8. |
| 9 | Guide | User instructions (was "Start here") | row 1 | |
| 10 | Weeks | Calculation engine, one row per week (rows 7:59) | A1:B6 | A Week of … J Good per crew-day · K Check · **L:W days per worker** · **X:AI share per worker** · AJ Shares total · AK Shares − pot. Worker IDs are in row 5 and matched by Weekly Pay, Payslip and Worker History. |
| 11 | Checks | Self-checks | row 1 | C5 overall result ("ALL CHECKS OK"), 9 checks in B8:C16, warnings B19:C20. |
| 12 | Project Notes | These notes | – | |

### Named ranges
`BreakAllow=Settings!C6` · `DebtCap=Settings!C7` · `RejectAlarm=Settings!C8` · `RateDates=Settings!B13:B32` · `RateValues=Settings!C13:C32` · `EntryTypes=Settings!F6:F8` · `StatusList=Settings!G6:G7` · `YesNo=Settings!H6:H7` · `WeekStarts=Weeks!A7:A59` · `WorkerIDs=Workers!A6:A17` · `WorkerNames=Workers!B6:B17`

## 3. Current data (at handover)

- **Workers:** W01 Davlatbek, W02 Xusanboy, W03 Oybek, W04 Nomonjon (was mistyped as Abdurashid). All Active and working since at least 4 Sep; real joining dates still to be entered. W05–W12 are empty.
- **Week of 2026-09-21** (as of 23 Sep): moulded 11,978, broken 121, good 11,857, pot 5,928,500 so'm, 24 crew-days, reject rate 1.0%.
- **Advances:** Davlatbek 150,000 (21 Sep) and Xusanboy 500,000 (22 Sep), 650,000 in total. TO PAY for the crew was 5,278,500.
- Checks: **ALL CHECKS OK**.
- More Daily Log rows were entered after this; check the file for the latest figures.

## 4. Design system (Light & clean)

- Font **Segoe UI**. White background, no gridlines.
- Colours:
  - ink `#0F172A`, text `#334155`, muted `#64748B`, border `#E2E8F0`, surface `#F8FAFC`
  - accent `#4F46E5` / `#4338CA`, accent soft `#EEF2FF`
  - success `#16A34A` / `#DCFCE7`, warning `#B45309` / `#FEF3C7`, danger `#B91C1C` / `#FEE2E2`
- **Input cells:** fill `#EEF2FF`, text `#3730A3` (light blue = you type here). Calculated cells are white.
- Table headers: fill `#F8FAFC`, bold 9pt `#475569`, bottom border `#CBD5E1`. Rows have thin horizontal lines only.
- Number format `#,##0` (the old `# ##0` displayed wrongly, e.g. "5928 500"). Zeros show as `–`. Percentages `0.0%`.
- **Nav bar (row 1, every sheet):** "◆ Donabay" brand, then links Home · Daily Log · Cash Ledger · Weekly Pay · Payslip · Workers · Settings · Guide, with History · Weeks · Checks in lighter grey. The current page has fill `#EEF2FF`, text `#4338CA` and a thick indigo bottom border. Row height 30. Some nav cells are **merged**.
- Tab colours: indigo for Home, Daily Log, Cash Ledger, Weekly Pay and Payslip. Grey for the rest.

## 5. Change log

1. Built the core workbook: Daily Log, Cash Ledger, Weekly Pay, Payslip, Workers, Settings, Weeks, Checks and Start here.
2. Tried checkboxes for attendance. The user couldn't tick them, so this was **rolled back** to 1 / 0.5 / empty.
3. Speed-ups:
   - Weekly Pay "Ready for the Cash Ledger" copy block.
   - Green highlight on today's row in the Daily Log.
   - Grouped helper columns.
   - Tips lines.
4. Redesign:
   - Inserted nav row 1 on all sheets, which moved every row down by 1.
   - New Home dashboard.
   - Restyled every sheet.
   - Renamed Start here to Guide.
   - Fixed the Cash Ledger # numbering.
   - Moved the Corrections total.
   - Amounts inside message text now show as `#,##0`.
5. **Removed cement tracking.** The plant uses batching weight controls, so these were dropped:
   - Daily Log Cement column.
   - Weeks Cement bags and Bags per 1000 columns.
   - The cement check in the Weeks Check column.
   - The Settings "expected bags" row and the `CementNorm` name.
   - The cement wording in the Checks warning.
6. Added this Project Notes tab (24 Sep 2026).
7. **Imported the old hand-kept sheet (24 Sep 2026).** The calendar now starts Mon 31 Aug: 4-18 Sep production and the Avans/Berdim cash went in as written. W04 was renamed to Nomonjon. Cells the owner must confirm are amber with a comment: unknown cash dates, three payments that disagree with the old sheet's Qoldi, joining dates, and 25-26 Sep entered in advance.
   - The old 4-9 Sep pay period spans two Monday-Sunday weeks here, so it shows as two weeks on Weekly Pay. The running balances are the same.

## 6. Known quirks and lessons

- Link formulas are **blocked** in this environment. Nav links are static cell links (Insert → Link → Place in this document).
- Nav bar cells can be merged. Unmerge them before inserting or deleting columns near the nav links, then rebuild.
- Inserting or deleting columns on Daily Log or Weeks moves the nav bar links. Check row 1 afterwards.
- Payslip printing has **not been test-printed** since the redesign.
- Capacity: Daily Log runs to row 377 (about 1 year), the Cash Ledger has 1,000 rows, Weeks has 53 rows, and there are 12 worker slots. Checks warns when these are nearly full.
- These notes are kept on this tab so they travel with the workbook.

## 7. Daily use (quick reference)

- **Every day (Daily Log):** enter moulded, broken, and 1 / 0.5 / empty for each worker.
  - Fast fill: select the cells, type `1`, then press **Ctrl+Enter** (Mac: Control+Return).
  - Ctrl+D fills down and Ctrl+R fills right.
- **Advance:** add a Cash Ledger row with date, worker, `Advance`, amount and reason. Ctrl+; enters today's date.
- **Payday:** on Weekly Pay, pick the week, check Checks says OK, and pay the **TO PAY** amounts. Then copy the "Ready for the Cash Ledger" list and Paste Values into the next empty ledger row. Print the Payslips and have each worker sign.
- **New worker:** fill the next empty slot on Workers and set it Active. **Leaver:** set Left and never delete the row.
- **Rate change:** add a new row to the Settings rate table with the start date.

## 8. Possible next steps

- Test-print the Payslip and set the print area and page setup.
- Optional: monthly summary or chart on Home.
- Optional: extend capacity (more worker slots or ledger rows) when Checks warns.

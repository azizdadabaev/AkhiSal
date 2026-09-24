# -*- coding: utf-8 -*-
"""
One-off, run 24 Sep 2026: bring the owner's old hand-kept sheet (4-18 Sep 2026)
into Donabay.xlsx, so the books run from the start of the month.

What it does
- Moves the calendar back three weeks: the Daily Log now starts Mon 31 Aug and
  the Weeks engine with it. Dates in both are plain values and nothing in the
  formulas hardcodes a date, so rewriting them is safe; the days already
  entered are moved to their new rows.
- The 500 so'm rate now starts 31 Aug - the old sheet used 500 throughout.
- Renames W04 from Abdurashid to Nomonjon.
- Adds the old sheet's production and cash to the Daily Log and Cash Ledger.
- Paints amber, with a comment, every cell the owner still has to confirm.

The old sheet only recorded blocks that were paid for, so those go in as
moulded with broken left empty. It split pay equally four ways, which here
means all four present on every production day. Cash amounts go in exactly
as written; where they disagree with the old sheet's own Qoldi column the
cells are painted for the owner to decide, rather than guessed at.

    python3 builder/import_history_2026_09.py Donabay.xlsx
"""
import sys, datetime as dt, shutil, tempfile, os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from keep_addins import keep_addins

D = dt.date
NEW_START = D(2026, 8, 31)                     # Monday of the week holding 1 Sep
LOG_ROWS = range(7, 378)                       # Daily Log rows 7..377
WEEK_ROWS = range(7, 60)                       # Weeks rows 7..59
LOG_INPUTS = ["C", "D"] + [chr(c) for c in range(ord("G"), ord("R") + 1)] + ["X"]
LEDGER_INPUTS = "BCDEFGH"
ATT = {"Davlatbek": "G", "Xusanboy": "H", "Oybek": "I", "Nomonjon": "J"}

CHECK = PatternFill("solid", fgColor="FEF3C7")  # the workbook's own warning colour
BY = "Donabay import"

# ------------------------------------------------ the old sheet, as written
PRODUCTION = {                                  # date: blocks paid for ("Dona")
    D(2026, 9, 4): 266, D(2026, 9, 5): 1967, D(2026, 9, 6): 2331,
    D(2026, 9, 7): 1771, D(2026, 9, 8): 1260, D(2026, 9, 9): 2380,
    D(2026, 9, 14): 2422, D(2026, 9, 15): 1750, D(2026, 9, 16): 2107,
    D(2026, 9, 18): 2905,
}
P1, P2 = "old sheet, 4-9 Sep", "old sheet, 14-18 Sep"
DATE_4_9 = ("Date not in the old sheet: the Avans was given sometime during 4-9 Sep. "
            "Only changes which week's statement shows it - balances are the same. "
            "Put the real date if you know it, then clear the colour.")
PAY_4_9 = ("Payday date not in the old sheet (its block is headed 9/4, the first day). "
           "Put the real date, then clear the colour.")
ADV_14_18 = ("Date not in the old sheet: given sometime during 14-18 Sep. "
             "Put the real date if you know it, then clear the colour.")
CASH = [  # date, worker, type, amount, reason, [(column, comment), ...]
    (D(2026, 9, 4), "Oybek", "Advance", 500000, f"Avans ({P1})", [("B", DATE_4_9)]),
    (D(2026, 9, 4), "Davlatbek", "Advance", 200000, f"Avans ({P1})", [("B", DATE_4_9)]),
    (D(2026, 9, 4), "Xusanboy", "Advance", 200000, f"Avans ({P1})", [("B", DATE_4_9)]),
    (D(2026, 9, 4), "Nomonjon", "Advance", 500000, f"Avans ({P1})", [("B", DATE_4_9)]),
    (D(2026, 9, 9), "Oybek", "Weekly pay", 1247000, f"Berdim ({P1})", [("B", PAY_4_9)]),
    (D(2026, 9, 9), "Davlatbek", "Weekly pay", 1247000, f"Berdim ({P1})", [("B", PAY_4_9), ("E",
        "Old sheet: Avans 200,000 + Berdim 1,247,000, yet Qoldi 0. That only adds up if this "
        "1,247,000 already INCLUDED the 200,000 advance - then change it to 1,047,000. "
        "As written, he got 200,000 more than he earned and now owes it.")]),
    (D(2026, 9, 9), "Xusanboy", "Weekly pay", 1247000, f"Berdim ({P1})", [("B", PAY_4_9), ("E",
        "Old sheet: Avans 200,000 + Berdim 1,247,000, yet Qoldi 0. That only adds up if this "
        "1,247,000 already INCLUDED the 200,000 advance - then change it to 1,047,000. "
        "As written, he got 200,000 more than he earned and now owes it.")]),
    (D(2026, 9, 9), "Nomonjon", "Weekly pay", 1247000, f"Berdim ({P1})", [("B", PAY_4_9)]),
    (D(2026, 9, 18), "Xusanboy", "Advance", 500000, f"Avans ({P2})", [("B", ADV_14_18)]),
    (D(2026, 9, 18), "Oybek", "Weekly pay", 1000000, f"Berdim ({P2})", []),
    (D(2026, 9, 18), "Davlatbek", "Weekly pay", 848000, f"Berdim ({P2})", [("E",
        "Old sheet: no advance, Berdim 848,000, Qoldi 0 - but he earned 1,148,000 that week. "
        "300,000 is unaccounted for: an advance missing from the old sheet (add it as an "
        "Advance row), or a different amount paid? As written, you still owe him about 300,000.")]),
    (D(2026, 9, 18), "Xusanboy", "Weekly pay", 648000, f"Berdim ({P2})", []),
    (D(2026, 9, 18), "Nomonjon", "Weekly pay", 1000000, f"Berdim ({P2})", [("E",
        "Old sheet: Qoldi 0 - but his numbers are the same as Oybek's (owed 500,000 from 4-9 Sep, "
        "earned 1,148,000, got 1,000,000), and Oybek's Qoldi is 352. As written he still owes "
        "about 352,000. Was the 500,000 settled some other way, or is his Qoldi wrong?")]),
]


def run(path):
    original = tempfile.mktemp(suffix=".xlsx")
    shutil.copy(path, original)
    wb = load_workbook(path)
    dl, wk, led = wb["Daily Log"], wb["Weeks"], wb["Cash Ledger"]

    # ---- Daily Log: lift out what's entered, move the calendar back, put it back
    entered = {}
    for r in LOG_ROWS:
        vals = {c: dl[f"{c}{r}"].value for c in LOG_INPUTS}
        if any(v not in (None, "") for v in vals.values()):
            entered[dl[f"A{r}"].value.date()] = vals
        for c in LOG_INPUTS:
            dl[f"{c}{r}"].value = None
    row_of = {}
    for i, r in enumerate(LOG_ROWS):
        day = NEW_START + dt.timedelta(days=i)
        dl[f"A{r}"].value = dt.datetime.combine(day, dt.time())
        row_of[day] = r
    for day, vals in entered.items():
        for c, v in vals.items():
            dl[f"{c}{row_of[day]}"].value = v
    for day, blocks in PRODUCTION.items():
        r = row_of[day]
        assert dl[f"C{r}"].value is None, f"{day} already has production"
        dl[f"C{r}"].value = blocks
        for col in ATT.values():
            dl[f"{col}{r}"].value = 1
        dl[f"X{r}"].value = "From the old sheet (paid blocks only; all 4 counted, as it split equally)"
    dl[f"X{row_of[D(2026, 9, 17)]}"].value = "No production (blank in the old sheet)"
    today = dt.date.today()
    for day in entered:
        if day > today and dl[f"C{row_of[day]}"].value not in (None, ""):
            c = dl[f"C{row_of[day]}"]
            c.fill = CHECK
            c.comment = Comment(f"Entered before {day:%d %b} had happened (it was on the sheet by "
                                f"23 Sep). If this is a plan or a test, clear the row - it is "
                                f"already counted in this week's pay. Then clear the colour.", BY)

    # ---- Weeks engine follows the calendar
    for i, r in enumerate(WEEK_ROWS):
        wk[f"A{r}"].value = dt.datetime.combine(NEW_START + dt.timedelta(days=7 * i), dt.time())

    # ---- Settings: the old sheet paid 500 throughout
    s = wb["Settings"]
    assert s["C13"].value == 500
    s["B13"].value = dt.datetime.combine(NEW_START, dt.time())

    # ---- Workers: W04 is Nomonjon; joining dates predate the workbook
    w = wb["Workers"]
    assert w["A9"].value == "W04"
    w["B9"].value = "Nomonjon"
    for r in range(6, 10):
        w[f"D{r}"].fill = CHECK
    w["D6"].comment = Comment("All four were working by 4 Sep, before this date. Type each "
                              "person's real start date, then clear the colour.", BY)

    # ---- Cash Ledger: history first, then what was already there, oldest first
    existing = []
    for r in range(7, 1007):
        vals = [led[f"{c}{r}"].value for c in LEDGER_INPUTS]
        if any(v not in (None, "") for v in vals):
            existing.append(vals)
            for c in LEDGER_INPUTS:
                led[f"{c}{r}"].value = None
    rows = [[dt.datetime.combine(d, dt.time()), who, typ, amt, why, None, None]
            for d, who, typ, amt, why, _ in CASH] + existing
    for i, vals in enumerate(rows):
        for c, v in zip(LEDGER_INPUTS, vals):
            led[f"{c}{7 + i}"].value = v
    for i, (*_, marks) in enumerate(CASH):
        for col, note in marks:
            cell = led[f"{col}{7 + i}"]
            cell.fill = CHECK
            cell.comment = Comment(note, BY)

    # ---- Project Notes: keep the handover notes true
    pn = wb["Project Notes"]
    edits = {
        "Current rate: **500 so'm from 2026-09-21**": "Current rate: **500 so'm from 2026-08-31**",
        "The first week starts **2026-09-21**": "The first week starts **2026-08-31**",
        "W04 Abdurashid. All Active, joined 2026-09-21.":
            "W04 Nomonjon (was mistyped as Abdurashid). All Active and working since at least 4 Sep; "
            "real joining dates still to be entered.",
    }
    for r in range(1, pn.max_row + 1):
        v = pn[f"A{r}"].value
        if isinstance(v, str):
            for old, new in edits.items():
                v = v.replace(old, new)
            pn[f"A{r}"].value = v
    anchor = next(r for r in range(1, pn.max_row + 1)
                  if str(pn[f"A{r}"].value or "").startswith("6. Added this Project Notes tab"))
    pn.insert_rows(anchor + 1, 2)
    pn[f"A{anchor + 1}"].value = (
        "7. **Imported the old hand-kept sheet (24 Sep 2026).** The calendar now starts Mon 31 Aug: "
        "4-18 Sep production and the Avans/Berdim cash went in as written. W04 was renamed to Nomonjon. "
        "Cells the owner must confirm are amber with a comment: unknown cash dates, three payments that "
        "disagree with the old sheet's Qoldi, joining dates, and 25-26 Sep entered in advance.")
    pn[f"A{anchor + 2}"].value = ("   - The old 4-9 Sep pay period spans two Monday-Sunday weeks here, "
                                  "so it shows as two weeks on Weekly Pay. The running balances are the same.")

    wb.calculation.fullCalcOnLoad = True
    wb.save(path)
    restored = keep_addins(original, path)
    os.remove(original)
    print(f"moved {len(entered)} existing day(s), added {len(PRODUCTION)} days and {len(CASH)} "
          f"cash entries, kept {len(existing)} existing entries, restored {restored} add-in part(s)")


if __name__ == "__main__":
    run(sys.argv[1])

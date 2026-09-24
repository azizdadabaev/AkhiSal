# -*- coding: utf-8 -*-
"""
Check the 24 Sep 2026 history import against the old hand-kept sheet.

Works out each worker's statement for every week from the raw numbers, never
reading a formula, and compares with LibreOffice's calculation of copies of
Donabay.xlsx that have Weekly Pay set to each week in turn:

    python3 tests/verify_history_import.py imp_2026-08-31.xlsx imp_2026-09-07.xlsx ...
"""
import sys, datetime as dt
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import load_workbook

D = dt.date
def xround(x): return float(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
def monday(d): return d - dt.timedelta(days=d.weekday())

WORKERS = ["Davlatbek", "Xusanboy", "Oybek", "Nomonjon"]            # W01..W04
PAID = {D(2026, 9, 4): 266, D(2026, 9, 5): 1967, D(2026, 9, 6): 2331, D(2026, 9, 7): 1771,
        D(2026, 9, 8): 1260, D(2026, 9, 9): 2380, D(2026, 9, 14): 2422, D(2026, 9, 15): 1750,
        D(2026, 9, 16): 2107, D(2026, 9, 18): 2905,
        # already in the workbook: moulded - broken
        D(2026, 9, 21): 1500 - 10, D(2026, 9, 22): 2000 - 21, D(2026, 9, 23): 2300 - 25,
        D(2026, 9, 24): 1900 - 20, D(2026, 9, 25): 2278 - 25, D(2026, 9, 26): 2000 - 20}
RATE = 500
CASH = [(D(2026, 9, 4), "Oybek", "Advance", 500000), (D(2026, 9, 4), "Davlatbek", "Advance", 200000),
        (D(2026, 9, 4), "Xusanboy", "Advance", 200000), (D(2026, 9, 4), "Nomonjon", "Advance", 500000),
        (D(2026, 9, 9), "Oybek", "Weekly pay", 1247000), (D(2026, 9, 9), "Davlatbek", "Weekly pay", 1247000),
        (D(2026, 9, 9), "Xusanboy", "Weekly pay", 1247000), (D(2026, 9, 9), "Nomonjon", "Weekly pay", 1247000),
        (D(2026, 9, 18), "Xusanboy", "Advance", 500000), (D(2026, 9, 18), "Oybek", "Weekly pay", 1000000),
        (D(2026, 9, 18), "Davlatbek", "Weekly pay", 848000), (D(2026, 9, 18), "Xusanboy", "Weekly pay", 648000),
        (D(2026, 9, 18), "Nomonjon", "Weekly pay", 1000000),
        (D(2026, 9, 21), "Davlatbek", "Advance", 150000), (D(2026, 9, 22), "Xusanboy", "Advance", 500000)]

def earned(name, wk):          # all four present every production day -> equal quarters
    pot = sum(b * RATE for d, b in PAID.items() if monday(d) == wk)
    return xround(pot / 4) if pot else 0.0

def cash(name, typ=None, lo=None, hi=None, before=None):
    return sum(a for d, n, t, a in CASH if n == name and (typ is None or t == typ)
               and (lo is None or d >= lo) and (hi is None or d <= hi) and (before is None or d < before))

def statement(name, wk):
    weeks = sorted({monday(d) for d in PAID})
    e = earned(name, wk)
    bf = sum(earned(name, w) for w in weeks if w < wk) - cash(name, before=wk)
    end = wk + dt.timedelta(days=6)
    a, c, p = cash(name, "Advance", wk, end), cash(name, "Correction", wk, end), cash(name, "Weekly pay", wk, end)
    due = bf + e - a - c
    pay = xround(max(0, due if bf >= 0 else e - a - c - min(-bf, e)))
    days = 4 * 0 + sum(1 for d in PAID if monday(d) == wk)
    return dict(days=days, earned=e, bf=bf, adv=a, corr=c, due=due, to_pay=pay, paid=p, carried=due - p)

fails = n = 0
for path in sys.argv[1:]:
    wb = load_workbook(path, data_only=True)
    wp = wb["Weekly Pay"]
    wk = wp["C5"].value.date()
    print(f"week of {wk}:")
    for i, name in enumerate(WORKERS):
        r = 12 + i
        assert wp[f"B{r}"].value == name, (wp[f"B{r}"].value, name)
        st = statement(name, wk)
        for col, key in zip("DFGHIJKLN", ["days", "earned", "bf", "adv", "corr", "due", "to_pay", "paid", "carried"]):
            got, want = wp[f"{col}{r}"].value, st[key]
            n += 1
            if got is None or abs(float(got) - want) > 0.5:
                fails += 1
                print(f"  FAIL {name} {key}: workbook={got} expected={want}")
        print(f"  {name:10} earned {st['earned']:>11,.0f}  carried {st['carried']:>12,.0f}   "
              f"(workbook {wp[f'N{r}'].value:>12,.0f})")
    chk = wb["Checks"]["C5"].value
    print("  checks:", chk)
print(f"\n{n - fails}/{n} checks passed.")
sys.exit(1 if fails else 0)

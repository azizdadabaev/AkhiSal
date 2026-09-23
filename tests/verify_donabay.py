# -*- coding: utf-8 -*-
"""
Independent check of Donabay.xlsx's pay maths.

Recomputes every figure from the raw demo inputs in plain Python - it never
reads a formula from the workbook - and compares the result with what
LibreOffice calculated in the demo copies. A clean recalculation only proves
the formulas run; this proves they give the right money.

    python3 builder/build_donabay.py --demo T1.xlsx
    python3 builder/build_donabay.py --demo T2.xlsx 2026-09-28
    python3 builder/build_donabay.py --demo T3.xlsx 2026-09-28 0.5
    (recalculate all three with LibreOffice, then)
    python3 tests/verify_donabay.py T1.xlsx T2.xlsx T3.xlsx
"""
import sys, datetime as dt
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import load_workbook

D = dt.date
def xround(x):                     # Excel ROUND: half away from zero, not banker's
    return float(Decimal(str(x)).quantize(Decimal("1"), rounding=ROUND_HALF_UP))

# ---- the same raw inputs fill_demo() writes, restated independently
RATES = [(D(2026, 9, 21), 200), (D(2026, 9, 26), 250)]
IDS = {"Ali": "W01", "Bobur": "W02", "Davron": "W03", "Eldor": "W04"}
DAYS = {
    D(2026, 9, 21): (820, 20, {"Ali": 1, "Bobur": 1, "Davron": 1, "Eldor": 1}),
    D(2026, 9, 22): (800, 10, {"Ali": 1, "Bobur": 1, "Eldor": 1}),
    D(2026, 9, 23): (840, 40, {"Ali": 1, "Bobur": 0.5, "Davron": 1, "Eldor": 1}),
    D(2026, 9, 24): (810, 5, {"Ali": 1, "Bobur": 1, "Davron": 1, "Eldor": 1}),
    D(2026, 9, 25): (790, 15, {"Ali": 1, "Bobur": 1, "Davron": 1, "Eldor": 1}),
    D(2026, 9, 26): (700, 30, {"Ali": 1, "Bobur": 1, "Davron": 1, "Eldor": 1}),
    D(2026, 9, 28): (800, 10, {"Ali": 1, "Bobur": 1, "Davron": 1, "Eldor": 1}),
}
LEDGER = [
    (D(2026, 9, 23), "Bobur", "Advance", 150000),
    (D(2026, 9, 24), "Ali", "Advance", 50000),
    (D(2026, 9, 25), "Bobur", "Advance", 100000),
    (D(2026, 9, 26), "Davron", "Advance", 2000000),
    (D(2026, 9, 24), "Eldor", "Advance", 30000),
    (D(2026, 9, 24), "Eldor", "Correction", -30000),
    (D(2026, 9, 26), "Ali", "Weekly pay", 400000),
]
BREAK_ALLOW = 0.0

def monday(d): return d - dt.timedelta(days=d.weekday())
def rate_on(d): return max((r for r in RATES if r[0] <= d), key=lambda r: r[0])[1]

def week_figures(wk):
    pot = crew = 0.0
    days = {n: 0.0 for n in IDS}
    for d, (m, b, att) in DAYS.items():
        if monday(d) != wk:
            continue
        paid = m - max(0, b - BREAK_ALLOW * m)
        pot += paid * rate_on(d)
        crew += sum(att.values())
        for n, v in att.items():
            days[n] += v
    earned = {n: (0.0 if crew == 0 else xround(pot * days[n] / crew)) for n in IDS}
    return pot, crew, days, earned

def cash(name, typ=None, lo=None, hi=None, before=None):
    t = 0
    for d, n, ty, a in LEDGER:
        if n != name or (typ and ty != typ): continue
        if lo and d < lo: continue
        if hi and d > hi: continue
        if before and d >= before: continue
        t += a
    return t

def earned_before(name, wk):
    return sum(week_figures(w)[3][name] for w in sorted({monday(d) for d in DAYS}) if w < wk)

def statement(name, wk, cap):
    pot, crew, days, earned = week_figures(wk)
    end = wk + dt.timedelta(days=6)
    E = earned[name]
    BF = earned_before(name, wk) - cash(name, before=wk)
    A = cash(name, "Advance", wk, end)
    C = cash(name, "Correction", wk, end)
    P = cash(name, "Weekly pay", wk, end)
    due = BF + E - A - C
    pay = xround(max(0, due if BF >= 0 else E - A - C - min(-BF, cap * E)))
    return dict(days=days[name], earned=E, bf=BF, adv=A, corr=C, due=due,
                to_pay=pay, paid=P, still=max(0, pay - P), carried=due - P)

# ------------------------------------------------------------------ compare
failures = checks = 0
def same(label, got, want, tol=0.5):
    global failures, checks
    checks += 1
    ok = (got == want) if isinstance(want, str) else (got is not None and abs(float(got) - want) <= tol)
    if not ok:
        failures += 1
        print(f"  FAIL {label}: workbook={got!r} expected={want!r}")

SCENARIOS = [(D(2026, 9, 21), 1.0), (D(2026, 9, 28), 1.0), (D(2026, 9, 28), 0.5)]

for path, (wk, cap) in zip(sys.argv[1:], SCENARIOS):
    print(f"{path}: week {wk}, debt cap {cap:.0%}")
    wb = load_workbook(path, data_only=True)
    pay = wb["Weekly Pay"]
    pot, crew, _, earned = week_figures(wk)
    same("pot", pay["F7"].value, pot)
    same("crew-days", pay["H7"].value, crew)
    same("shares sum to pot", sum(earned.values()), pot)
    for i, name in enumerate(IDS):
        r = 11 + i
        st = statement(name, wk, cap)
        same(f"{name} name", pay[f"B{r}"].value, name)
        for col, key in zip("DFGHIJKLMN", ["days", "earned", "bf", "adv", "corr", "due",
                                            "to_pay", "paid", "still", "carried"]):
            same(f"{name} {key}", pay[f"{col}{r}"].value, st[key])
    same("left worker earns nothing", pay["F15"].value, 0)
    same("overall checks", wb["Checks"]["C4"].value, "ALL CHECKS OK")

    # payslip for Bobur must match his Weekly Pay row exactly
    slip, st = wb["Payslip"], statement("Bobur", wk, cap)
    for cell, key, sign in (("C13", "days", 1), ("C16", "earned", 1), ("C17", "bf", 1), ("C18", "adv", -1),
                            ("C19", "corr", -1), ("C20", "due", 1), ("C22", "to_pay", 1), ("C24", "carried", 1)):
        same(f"payslip {key}", slip[cell].value, sign * st[key])
    advs = sorted((d, a) for d, n, t, a in LEDGER
                  if n == "Bobur" and t == "Advance" and wk <= d <= wk + dt.timedelta(days=6))
    for j in range(3):
        got = slip[f"C{28+j}"].value
        # a formula returning "" reads back from the file as None; both mean blank
        if j < len(advs):
            same(f"payslip advance line {j+1}", got, advs[j][1])
        else:
            same(f"payslip advance line {j+1} blank", got in (None, ""), True, tol=0)

    if wk == D(2026, 9, 21):
        # the ledger's running balance and the over-advance warnings
        led = wb["Cash Ledger"]
        wk_e = week_figures(wk)[3]
        for r, (d, n, t, a) in enumerate(LEDGER, start=6):
            bal = wk_e[n] - cash(n, hi=d)
            same(f"ledger row {r} balance", led[f"L{r}"].value, bal)
            want_warn = t == "Advance" and bal < 0
            got = led[f"M{r}"].value or ""
            same(f"ledger row {r} warning", got.startswith("Warning"), want_warn, tol=0)
        # Davron's history row: the account after his 2m advance
        h = wb["Worker History"]
        same("history earned", h["C7"].value, wk_e["Davron"])
        same("history balance", h["G7"].value, wk_e["Davron"] - 2000000)
        same("history meaning", h["H7"].value, "owes you")

print(f"\n{checks - failures}/{checks} checks passed.")
sys.exit(1 if failures else 0)

# -*- coding: utf-8 -*-
"""
Builds AkhiSal-Tools.xlsx.

Everything here uses Excel-2007-era functions only (INDEX, MATCH, SUBSTITUTE,
COUNTIF, SUMIFS, IFERROR). No LAMBDA, no XLOOKUP, no dynamic arrays. That is
deliberate on two counts: those work in every Excel version including the old
ones, and they can be verified by recalculating the file here before it ships.

    python3 builder/build_workbook.py
"""
import sys, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from crm_snapshot import CLIENTS, RECEIVABLES, PHONE_SAMPLES

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "AkhiSal-Tools.xlsx")

# ---------------------------------------------------------------- styling
FONT = "Arial"
H1     = Font(name=FONT, size=16, bold=True, color="1F3864")
H2     = Font(name=FONT, size=11, bold=True, color="FFFFFF")
BODY   = Font(name=FONT, size=10)
NOTE   = Font(name=FONT, size=9, color="595959", italic=True)
INPUT  = Font(name=FONT, size=10, color="0000FF")          # blue = you type here
CALC   = Font(name=FONT, size=10, color="000000")          # black = formula
BOLD   = Font(name=FONT, size=10, bold=True)

HDR_FILL   = PatternFill("solid", fgColor="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")        # yellow = editable
CALC_FILL  = PatternFill("solid", fgColor="F2F2F2")
WARN_FILL  = PatternFill("solid", fgColor="FCE4E4")
THIN = Side(style="thin", color="BFBFBF")
BOX  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Uzbek Cyrillic -> Latin. Inputs Cyrillic, outputs Latin, so no rule can
# re-consume another's output and the order does not matter.
TRANSLIT_MAP = [
    ("ё","yo"),("ю","yu"),("я","ya"),("ц","ts"),("ч","ch"),("ш","sh"),("щ","sh"),
    ("ў","o"),("қ","q"),("ғ","g"),("ҳ","h"),
    ("а","a"),("б","b"),("в","v"),("г","g"),("д","d"),("е","e"),("ж","j"),
    ("з","z"),("и","i"),("й","y"),("к","k"),("л","l"),("м","m"),("н","n"),
    ("о","o"),("п","p"),("р","r"),("с","s"),("т","t"),("у","u"),("ф","f"),
    ("х","x"),("ъ",""),("ы","i"),("ь",""),("э","e"),
]

def translit_formula(ref):
    """Nested SUBSTITUTE chain folding Cyrillic to Latin. Classic functions only."""
    expr = f"LOWER({ref})"
    for src, dst in TRANSLIT_MAP:
        expr = f'SUBSTITUTE({expr},"{src}","{dst}")'
    # apostrophes and the Uzbek turned comma vary by typist; drop them
    for ch in ("'", "‘", "’", "ʻ", "ʼ", "`"):
        expr = f'SUBSTITUTE({expr},"{ch}","")'
    return f"TRIM({expr})"

def strip_punct_formula(ref):
    """Reduce a phone to digits by removing the characters people actually type."""
    expr = ref
    for ch in (" ", "+", "-", "(", ")", ".", "/", " "):
        lit = 'CHAR(160)' if ch == " " else f'"{ch}"'
        expr = f'SUBSTITUTE({expr},{lit},"")'
    return expr

def header(ws, row, labels, widths):
    for i, (label, width) in enumerate(zip(labels, widths), start=1):
        c = ws.cell(row=row, column=i, value=label)
        c.font, c.fill, c.border = H2, HDR_FILL, BOX
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = 28

def title(ws, text, subtitle, span="A1:F1"):
    ws["A1"] = text
    ws["A1"].font = H1
    ws.merge_cells(span)
    ws["A2"] = subtitle
    ws["A2"].font = NOTE
    ws.merge_cells(span.replace("1", "2"))
    ws.row_dimensions[1].height = 22

wb = Workbook()
N_CLIENTS = len(CLIENTS)
CLI_LAST = N_CLIENTS + 1               # clients occupy rows 2..51
MATCH_ROWS = 40                        # editable rows on the matcher sheet

# ============================================================ 1. Start here
ws = wb.active
ws.title = "Start here"
ws.sheet_view.showGridLines = False
title(ws, "AkhiSal Excel toolkit", "Nothing to install. Open the file, type in the yellow cells, read the grey ones.", "A1:H1")

rows = [
    ("", ""),
    ("Sheet", "What it is for"),
    ("Client Matcher",
     "Paste a list of client names in any spelling - Latin or Cyrillic - and get the matching CRM client, "
     "their phone, and a warning when several clients share that name."),
    ("CRM Clients",
     "The reference list the matcher looks against: 50 clients copied from the CRM on 2026-09-19. "
     "Replace it with a fresh export whenever you like; the formulas follow."),
    ("Phone Cleaner",
     "Paste phone numbers written any which way and get them back as 998XXXXXXXXX, ready for the CRM."),
    ("Receivables",
     "Who owes what and for how long, bucketed 0-30 / 31-60 / 61-90 / 90+ days. Pre-filled with real "
     "unpaid orders."),
    ("Money in Words",
     "Turn an amount into Uzbek words for invoices and contracts: 13 500 000 -> o'n uch million besh yuz ming so'm."),
    ("", ""),
    ("Colour", "Meaning"),
    ("Yellow cell, blue text", "You type here. Everything else recalculates on its own."),
    ("Grey cell, black text", "A formula. Leave it alone - or copy it down if you add rows."),
    ("Pink cell", "Something needs a human eye."),
]
r = 4
for a, b in rows:
    ca, cb = ws.cell(row=r, column=1, value=a), ws.cell(row=r, column=2, value=b)
    if b in ("What it is for", "Meaning"):
        ca.font = cb.font = BOLD
    else:
        ca.font, cb.font = BOLD if a else BODY, BODY
    cb.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 30 if len(b) > 90 else 15
    r += 1
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 95

ws.cell(row=r + 1, column=1, value="Built with classic Excel functions only, so it works in any Excel version "
        "and needs no macros, no add-in and no admin rights.").font = NOTE
ws.merge_cells(start_row=r + 1, start_column=1, end_row=r + 1, end_column=2)

# ============================================================ 2. CRM Clients
ws = wb.create_sheet("CRM Clients")
ws.sheet_view.showGridLines = False
title(ws, "CRM Clients", "Reference list for the matcher. Snapshot taken 2026-09-19 - paste a fresh export over rows 4 and below any time.", "A1:E1")
header(ws, 3, ["Client name", "Phone", "Address", "Name, normalised", "Clients sharing this name"],
       [30, 16, 46, 26, 16])

for i, (name, phone, addr) in enumerate(CLIENTS):
    row = 4 + i
    ws.cell(row=row, column=1, value=name).font = BODY
    p = ws.cell(row=row, column=2, value=phone)
    p.font, p.number_format = BODY, "@"
    ws.cell(row=row, column=3, value=addr).font = BODY
    d = ws.cell(row=row, column=4, value=f"={translit_formula('$A%d' % row)}")
    d.font, d.fill = CALC, CALC_FILL
    e = ws.cell(row=row, column=5,
                value=f"=COUNTIF($D$4:$D${3+N_CLIENTS},$D{row})")
    e.font, e.fill = CALC, CALC_FILL
    e.alignment = Alignment(horizontal="center")
    for col in range(1, 6):
        ws.cell(row=row, column=col).border = BOX

ws.cell(row=4, column=4).comment = Comment(
    "Cyrillic folded to lowercase Latin so that a client written 'Xalimjon' in "
    "one sheet and 'Халимжон' in another lands on the same text.", "AkhiSal toolkit")
ws.cell(row=4, column=5).comment = Comment(
    "Anything above 1 means the name alone cannot identify the client - "
    "three different people are called Abdulloh. Match on phone in that case.", "AkhiSal toolkit")
ws.freeze_panes = "A4"

# ============================================================ 3. Client Matcher
ws = wb.create_sheet("Client Matcher")
ws.sheet_view.showGridLines = False
title(ws, "Client Matcher", "Type or paste names into the yellow column. Latin or Cyrillic, either works.", "A1:E1")
header(ws, 3, ["Name as you have it", "Normalised", "Matching CRM client", "Phone",
                "Result", "Matches"],
       [30, 24, 30, 16, 42, 9])

CR = f"'CRM Clients'!$D$4:$D${3+N_CLIENTS}"
CN = f"'CRM Clients'!$A$4:$A${3+N_CLIENTS}"
CP = f"'CRM Clients'!$B$4:$B${3+N_CLIENTS}"

for i in range(MATCH_ROWS):
    row = 4 + i
    a = ws.cell(row=row, column=1)
    a.font, a.fill, a.border = INPUT, INPUT_FILL, BOX

    b = ws.cell(row=row, column=2, value=f'=IF($A{row}="","",{translit_formula("$A%d" % row)})')
    c = ws.cell(row=row, column=3, value=(
        f'=IF($B{row}="","",IFERROR(INDEX({CN},MATCH($B{row},{CR},0)),"not found"))'))
    d = ws.cell(row=row, column=4, value=(
        f'=IF($B{row}="","",IFERROR(INDEX({CP},MATCH($B{row},{CR},0)),""))'))
    # count once in its own column, then read it back three times - COUNTIF
    # across the client list is what makes this sheet expensive
    f_ = ws.cell(row=row, column=6, value=f'=IF($B{row}="","",COUNTIF({CR},$B{row}))')
    e = ws.cell(row=row, column=5, value=(
        f'=IF($B{row}="","",'
        f'IF($F{row}=0,"no match - new client, or a typo",'
        f'IF($F{row}>1,"CHECK PHONE - "&$F{row}&" clients share this name",'
        f'"matched")))'))
    d.number_format = "@"
    for cell in (b, c, d, e, f_):
        cell.font, cell.fill, cell.border = CALC, CALC_FILL, BOX
    f_.alignment = Alignment(horizontal="center")

# one worked example, so the expected format is obvious
examples = ["Бекмурод", "Алишер ака", "Abdulloh", "Баходир", "Даврон Ваҳобов", "Халимжон"]
for i, ex in enumerate(examples):
    ws.cell(row=4 + i, column=1, value=ex)

ws.cell(row=4, column=1).comment = Comment(
    "Example rows - type over them. Most are Cyrillic spellings of clients stored in "
    "Latin, to show the two alphabets matching. 'Abdulloh' shows the ambiguous case "
    "(three clients share it) and the last row shows a name not in the list.", "AkhiSal toolkit")
ws.freeze_panes = "A4"
ws.cell(row=MATCH_ROWS + 6, column=1,
        value="Matching is exact after normalising, not fuzzy: it forgives alphabet and capitalisation, "
              "not misspellings. 'no match' on a name you expect usually means a real spelling difference - "
              "compare the Normalised column against the CRM Clients sheet.").font = NOTE
ws.merge_cells(start_row=MATCH_ROWS + 6, start_column=1, end_row=MATCH_ROWS + 6, end_column=5)

# ============================================================ 4. Phone Cleaner
ws = wb.create_sheet("Phone Cleaner")
ws.sheet_view.showGridLines = False
title(ws, "Phone Cleaner", "Paste numbers however they are written. Out comes 998XXXXXXXXX.", "A1:D1")
header(ws, 3, ["Number as you have it", "Digits only", "CRM format", "Result"], [30, 22, 22, 34])

PHONE_ROWS = 30
for i in range(PHONE_ROWS):
    row = 4 + i
    a = ws.cell(row=row, column=1)
    a.font, a.fill, a.border = INPUT, INPUT_FILL, BOX

    b = ws.cell(row=row, column=2, value=f'=IF($A{row}="","",{strip_punct_formula("$A%d" % row)})')
    # 12 digits starting 998 -> keep; 9 digits -> prefix; 10 starting 8 -> swap;
    # 00998... -> drop the international prefix; anything else is flagged.
    c = ws.cell(row=row, column=3, value=(
        f'=IF($B{row}="","",'
        f'IF(AND(LEN($B{row})=12,LEFT($B{row},3)="998"),$B{row},'
        f'IF(LEN($B{row})=9,"998"&$B{row},'
        f'IF(AND(LEN($B{row})=10,LEFT($B{row},1)="8"),"998"&MID($B{row},2,9),'
        f'IF(AND(LEN($B{row})=14,LEFT($B{row},5)="00998"),MID($B{row},3,12),'
        f'"?")))))'))
    d = ws.cell(row=row, column=4, value=(
        f'=IF($C{row}="","",IF($C{row}="?","cannot read - fix by hand","ok"))'))
    b.number_format = c.number_format = "@"
    for cell in (b, c, d):
        cell.font, cell.fill, cell.border = CALC, CALC_FILL, BOX
    d.fill = CALC_FILL

for i, s in enumerate(PHONE_SAMPLES):
    ws.cell(row=4 + i, column=1, value=s)
ws.cell(row=4, column=1).comment = Comment(
    "Example rows covering the spellings that actually turn up, including one "
    "('не указан') that cannot be read and is flagged rather than guessed.", "AkhiSal toolkit")
ws.freeze_panes = "A4"

# ============================================================ 5. Receivables
ws = wb.create_sheet("Receivables")
ws.sheet_view.showGridLines = False
title(ws, "Receivables", "Who owes what, and for how long. Ages against today's date every time the file opens.", "A1:I1")
header(ws, 3, ["Order", "Client", "Phone", "Due date", "Order total", "Paid",
                "Outstanding", "Days overdue", "Bucket"],
       [15, 18, 15, 12, 15, 15, 15, 12, 12])

REC_FIRST = 4
REC_LAST = REC_FIRST + len(RECEIVABLES) - 1
for i, (num, client, phone, due, total, paid) in enumerate(RECEIVABLES):
    row = REC_FIRST + i
    ws.cell(row=row, column=1, value=num).font = BODY
    ws.cell(row=row, column=2, value=client).font = BODY
    p = ws.cell(row=row, column=3, value=phone); p.font, p.number_format = BODY, "@"
    y, m, d_ = (int(x) for x in due.split("-"))
    dd = ws.cell(row=row, column=4)
    dd.value = __import__("datetime").date(y, m, d_)
    dd.number_format, dd.font = "yyyy-mm-dd", BODY
    for col, val in ((5, total), (6, paid)):
        c = ws.cell(row=row, column=col, value=val)
        c.font, c.number_format = BODY, "# ##0"
    g = ws.cell(row=row, column=7, value=f"=$E{row}-$F{row}")
    g.font, g.number_format, g.fill = CALC, "# ##0", CALC_FILL
    h = ws.cell(row=row, column=8, value=f'=IF($G{row}<=0,"",TODAY()-$D{row})')
    h.font, h.fill, h.alignment = CALC, CALC_FILL, Alignment(horizontal="center")
    ibkt = ws.cell(row=row, column=9, value=(
        f'=IF($G{row}<=0,"paid",'
        f'IF(TODAY()-$D{row}<=0,"not due",'
        f'IF(TODAY()-$D{row}<=30,"0-30",'
        f'IF(TODAY()-$D{row}<=60,"31-60",'
        f'IF(TODAY()-$D{row}<=90,"61-90","90+")))))'))
    ibkt.font, ibkt.fill, ibkt.alignment = CALC, CALC_FILL, Alignment(horizontal="center")
    for col in range(1, 10):
        ws.cell(row=row, column=col).border = BOX

tot = REC_LAST + 1
ws.cell(row=tot, column=6, value="Total").font = BOLD
t = ws.cell(row=tot, column=7, value=f"=SUM($G${REC_FIRST}:$G${REC_LAST})")
t.font, t.number_format = BOLD, "# ##0"
t.border = BOX

# summary by bucket
s0 = tot + 3
ws.cell(row=s0, column=1, value="Aging summary").font = Font(name=FONT, size=12, bold=True, color="1F3864")
header(ws, s0 + 1, ["Bucket", "Orders", "Outstanding"], [15, 18, 18])
BUCKETS = ["not due", "0-30", "31-60", "61-90", "90+"]
for i, b in enumerate(BUCKETS):
    row = s0 + 2 + i
    ws.cell(row=row, column=1, value=b).font = BODY
    c2 = ws.cell(row=row, column=2,
                 value=f'=COUNTIFS($I${REC_FIRST}:$I${REC_LAST},$A{row})')
    c3 = ws.cell(row=row, column=3,
                 value=f'=SUMIFS($G${REC_FIRST}:$G${REC_LAST},$I${REC_FIRST}:$I${REC_LAST},$A{row})')
    c2.font, c3.font = CALC, CALC
    c2.fill = c3.fill = CALC_FILL
    c3.number_format = "# ##0"
    for col in (1, 2, 3):
        ws.cell(row=row, column=col).border = BOX
srow = s0 + 2 + len(BUCKETS)
ws.cell(row=srow, column=1, value="Total").font = BOLD
for col, f in ((2, f"=SUM($B${s0+2}:$B${srow-1})"), (3, f"=SUM($C${s0+2}:$C${srow-1})")):
    c = ws.cell(row=srow, column=col, value=f)
    c.font, c.border = BOLD, BOX
    if col == 3:
        c.number_format = "# ##0"
ws.freeze_panes = "A4"

# ============================================================ 6. Money in Words
ws = wb.create_sheet("Money in Words")
ws.sheet_view.showGridLines = False
title(ws, "Money in Words", "For invoices and contracts. Type an amount in the yellow column.", "A1:C1")
header(ws, 3, ["Amount (UZS)", "In words"], [18, 92])

# lookup tables, parked to the right and greyed out
ONES = ["", "bir", "ikki", "uch", "to'rt", "besh", "olti", "yetti", "sakkiz", "to'qqiz"]
TENS = ["", "", "yigirma", "o'ttiz", "qirq", "ellik", "oltmish", "yetmish", "sakson", "to'qson"]
ws.cell(row=3, column=13, value="lookup tables - leave alone").font = NOTE
for i in range(10):
    ws.cell(row=4 + i, column=13, value=i).font = NOTE
    ws.cell(row=4 + i, column=14, value=ONES[i]).font = NOTE
    ws.cell(row=4 + i, column=15, value=TENS[i]).font = NOTE
ONES_R, TENS_R = "$N$4:$N$13", "$O$4:$O$13"
for _c, _w in (("D", 7), ("E", 7), ("F", 7), ("G", 7),
               ("H", 14), ("I", 14), ("J", 14), ("K", 14),
               ("M", 5), ("N", 10), ("O", 10)):
    ws.column_dimensions[_c].width = _w
ws.cell(row=3, column=4, value="working - the four digit groups, then their words").font = NOTE
ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=11)

def group_words(g):
    """Words for a 0-999 group expression `g`, using the ones/tens tables."""
    hund = f"INT({g}/100)"
    rem  = f"MOD({g},100)"
    return (
        f'IF({g}=0,"",TRIM('
        f'IF({hund}>0,INDEX({ONES_R},{hund}+1)&" yuz ","")&'
        f'IF({rem}>=20,INDEX({TENS_R},INT({rem}/10)+1)&" "&INDEX({ONES_R},MOD({rem},10)+1),'
        f'IF({rem}>=10,"o\'n "&INDEX({ONES_R},{rem}-10+1),'
        f'INDEX({ONES_R},{rem}+1)))))'
    )

MONEY_ROWS = 20
for i in range(MONEY_ROWS):
    row = 4 + i
    a = ws.cell(row=row, column=1)
    a.font, a.fill, a.border, a.number_format = INPUT, INPUT_FILL, BOX, "# ##0"

    n = f"INT($A{row})"
    # D..G hold the four digit groups, H..K their words. Splitting them out
    # keeps each formula small and makes a wrong answer easy to trace.
    for col, expr in ((4, f"INT({n}/1000000000)"),
                      (5, f"INT(MOD({n},1000000000)/1000000)"),
                      (6, f"INT(MOD({n},1000000)/1000)"),
                      (7, f"MOD({n},1000)")):
        c = ws.cell(row=row, column=col, value=f'=IF($A{row}="","",{expr})')
        c.font, c.fill = NOTE, CALC_FILL
    for col, src in ((8, "D"), (9, "E"), (10, "F"), (11, "G")):
        c = ws.cell(row=row, column=col,
                    value=f'=IF($A{row}="","",{group_words("$%s%d" % (src, row))})')
        c.font, c.fill = NOTE, CALC_FILL
    body = (
        f'TRIM('
        f'IF($D{row}>0,$H{row}&" milliard ","")&'
        f'IF($E{row}>0,$I{row}&" million ","")&'
        f'IF($F{row}>0,$J{row}&" ming ","")&'
        f'$K{row})&" so\'m"'
    )
    b = ws.cell(row=row, column=2, value=(
        f'=IF($A{row}="","",IF({n}=0,"nol so\'m",{body}))'))
    b.font, b.fill, b.border = CALC, CALC_FILL, BOX

for i, v in enumerate([13500000, 3005040, 22500000, 245508000, 1500, 9262000]):
    ws.cell(row=4 + i, column=1, value=v)
ws.cell(row=4, column=1).comment = Comment(
    "Example amounts taken from real orders - type over them.", "AkhiSal toolkit")
ws.freeze_panes = "A4"

wb.save(OUT)
print("wrote", OUT)

"""
Reference ports of the trickier VBA algorithms, with their expected output.

Excel is not available here, so these mirror the VBA line for line and pin the
behaviour that is easiest to get quietly wrong: Uzbek place-value grammar and
the Cyrillic/Latin fold used for client-name matching. Change the VBA, change
these to match, and re-run:

    python3 tests/verify_algorithms.py
"""
import sys

# ---------------------------------------------------------------- UZSWORDS
# Direct port of the VBA GroupWords/UZSWORDS walk, to check it before shipping.
ONES = {1:"bir",2:"ikki",3:"uch",4:"to'rt",5:"besh",6:"olti",7:"yetti",8:"sakkiz",9:"to'qqiz"}
TENS = {2:"yigirma",3:"o'ttiz",4:"qirq",5:"ellik",6:"oltmish",7:"yetmish",8:"sakson",9:"to'qson"}

def ones(d): return ONES.get(d, "")
def tens(d): return TENS.get(d, "")
def teens(d): return "o'n" if d == 10 else "o'n " + ones(d - 10)

def group(n):
    out = ""
    if n >= 1_000_000_000:
        out += group(n // 1_000_000_000) + " milliard "
        n -= (n // 1_000_000_000) * 1_000_000_000
    if n >= 1_000_000:
        out += group(n // 1_000_000) + " million "
        n -= (n // 1_000_000) * 1_000_000
    if n >= 1_000:
        out += group(n // 1_000) + " ming "
        n -= (n // 1_000) * 1_000
    if n >= 100:
        out += ones(n // 100) + " yuz "
        n -= (n // 100) * 100
    if n >= 20:
        out += tens(n // 10) + " "
        n -= (n // 10) * 10
    elif n >= 10:
        out += teens(n) + " "
        n = 0
    if n > 0:
        out += ones(n) + " "
    return out

def uzswords(a):
    n = int(abs(a))
    if n == 0: return "nol so'm"
    # WorksheetFunction.Trim = strip + collapse internal runs
    out = " ".join(group(n).split())
    if a < 0: out = "minus " + out
    return out + " so'm"

# ------------------------------------------------- TRANSLIT / SIMILARITY
# Port of the VBA TRANSLIT/SIMILARITY pair, to sanity-check the mapping.
M = {1072:"a",1073:"b",1074:"v",1075:"g",1076:"d",1077:"e",1105:"yo",1078:"j",
     1079:"z",1080:"i",1081:"y",1082:"k",1083:"l",1084:"m",1085:"n",1086:"o",
     1087:"p",1088:"r",1089:"s",1090:"t",1091:"u",1092:"f",1093:"x",1094:"ts",
     1095:"ch",1096:"sh",1097:"sh",1098:"",1099:"i",1100:"",1101:"e",1102:"yu",
     1103:"ya",1118:"o",1179:"q",1171:"g",1203:"h"}

def translit(s):
    out = []
    for ch in s:
        w = ord(ch)
        if 1040 <= w <= 1071: w += 32
        elif 65 <= w <= 90:   w += 32
        elif w == 1025: w = 1105
        elif w == 1038: w = 1118
        elif w == 1178: w = 1179
        elif w == 1170: w = 1171
        elif w == 1202: w = 1203
        if w in M: out.append(M[w])
        elif 97 <= w <= 122 or 48 <= w <= 57: out.append(chr(w))
        elif w in (32, 45): out.append(" ")
    return " ".join("".join(out).split())

def lev(s, t):
    if not s: return len(t)
    if not t: return len(s)
    prev = list(range(len(t) + 1))
    for i in range(1, len(s) + 1):
        cur = [i] + [0] * len(t)
        for j in range(1, len(t) + 1):
            cur[j] = min(cur[j-1]+1, prev[j]+1, prev[j-1] + (s[i-1] != t[j-1]))
        prev = cur
    return prev[len(t)]

def sim(a, b):
    x, y = translit(a), translit(b)
    n = max(len(x), len(y))
    return 1.0 if n == 0 else 1 - lev(x, y) / n

# ------------------------------------------------------------------ cases

WORDS = [
    (0, "nol so'm"),
    (1, "bir so'm"),
    (10, "o'n so'm"),
    (11, "o'n bir so'm"),
    (19, "o'n to'qqiz so'm"),
    (20, "yigirma so'm"),
    (25, "yigirma besh so'm"),
    (99, "to'qson to'qqiz so'm"),
    (100, "bir yuz so'm"),
    (115, "bir yuz o'n besh so'm"),
    (1000, "bir ming so'm"),
    (1500, "bir ming besh yuz so'm"),
    (1000000, "bir million so'm"),
    # real order totals out of the CRM
    (3005040, "uch million besh ming qirq so'm"),
    (9262000, "to'qqiz million ikki yuz oltmish ikki ming so'm"),
    (13500000, "o'n uch million besh yuz ming so'm"),
    (245508000, "ikki yuz qirq besh million besh yuz sakkiz ming so'm"),
    (3122510712,
     "uch milliard bir yuz yigirma ikki million besh yuz o'n ming "
     "yetti yuz o'n ikki so'm"),
    (-50000, "minus ellik ming so'm"),
]

TRANSLITS = [
    ("\u041d\u0430\u043c\u0430\u043d\u0433\u0430\u043d \u0432\u0438\u043b\u043e\u044f\u0442\u0438", "namangan viloyati"),
    ("\u0428\u0423\u04b2\u0420\u0410\u0422\u0416\u041e\u041d \u0410\u041a\u0410", "shuhratjon aka"),
    ("\u042f\u043d\u0433\u0438\u049b\u045e\u0440\u0493\u043e\u043d \u0442\u0443\u043c\u0430\u043d\u0438", "yangiqorgon tumani"),
    ("\u0416\u0443\u0440\u0430\u0432\u043e\u0439", "juravoy"),
]

# Latin spelling vs Cyrillic spelling of the same client must match; two
# different clients must not.
PAIRS = [
    ("Xalimjon", "\u0425\u0430\u043b\u0438\u043c\u0436\u043e\u043d", True),
    ("Shuhratjon aka", "\u0428\u0423\u04b2\u0420\u0410\u0422\u0416\u041e\u041d \u0410\u041a\u0410", True),
    ("Nematjon", "\u041d\u0435\u043c\u0430\u0442\u0436\u043e\u043d", True),
    ("Baxodir", "\u0411\u0430\u0445\u043e\u0434\u0438\u0440", True),
    ("Xalimjon", "\u041a\u043e\u0440\u043e\u0441\u043a\u043e\u043d", False),
]

CUTOFF = 0.85


def main():
    failures = 0
    for value, expected in WORDS:
        got = uzswords(value)
        if got != expected:
            print(f"FAIL uzswords({value}): {got!r} != {expected!r}")
            failures += 1
    for src, expected in TRANSLITS:
        got = translit(src)
        if got != expected:
            print(f"FAIL translit({src!r}): {got!r} != {expected!r}")
            failures += 1
    for a, b, should_match in PAIRS:
        score = sim(a, b)
        if (score >= CUTOFF) != should_match:
            print(f"FAIL similarity({a!r}, {b!r}) = {score:.3f}, "
                  f"expected {'match' if should_match else 'no match'}")
            failures += 1

    total = len(WORDS) + len(TRANSLITS) + len(PAIRS)
    if failures:
        print(f"\n{failures} of {total} checks failed.")
        return 1
    print(f"All {total} checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

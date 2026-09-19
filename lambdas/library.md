# LAMBDA library — no macros, no add-in, no install

The last-resort tier: these are plain formulas, so they work even where
company policy blocks both add-ins and VBA macros. They need Excel from
Microsoft 365 (LAMBDA is not in Excel 2019 or earlier).

## Installing one

1. **Formulas** tab → **Name Manager** → **New**
2. **Name**: the name below, exactly (the dot is part of it)
3. **Refers to**: paste the formula, `=` and all
4. **OK**, then use it in any cell like a built-in function

They are stored in the workbook, so anyone you send the file to gets them
too, with nothing to install on their side.

`AGE.BUCKET` calls `AGE.DAYS`, so add `AGE.DAYS` first.

The formulas below are shown across several lines for readability. Pasting them
that way into **Refers to** is fine, but do not press Enter *while typing* one —
Enter closes the dialog. Paste the whole block at once, or put it on a single
line by deleting the breaks; Excel treats both identically.

---

## `TRANSLIT`

Uzbek Cyrillic -> lowercase Latin, so the same client spelled either way compares equal.

```
=TRANSLIT("ШУҲРАТЖОН АКА")  ->  shuhratjon aka
```

**Refers to:**

```excel
=LAMBDA(txt, TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(LOWER(txt),"ё","yo"),"ю","yu"),"я","ya"),"ц","ts"),"ч","ch"),"ш","sh"),"щ","sh"),"ў","o"),"қ","q"),"ғ","g"),"ҳ","h"),"а","a"),"б","b"),"в","v"),"г","g"),"д","d"),"е","e"),"ж","j"),"з","z"),"и","i"),"й","y"),"к","k"),"л","l"),"м","m"),"н","n"),"о","o"),"п","p"),"р","r"),"с","s"),"т","t"),"у","u"),"ф","f"),"х","x"),"ъ",""),"ы","i"),"ь",""),"э","e")))
```

## `PHONE.UZ`

Any Uzbek phone spelling -> 998XXXXXXXXX. Returns "" when the digits do not add up.

```
=PHONE.UZ("+998 90 123-45-67")  ->  998901234567
```

**Refers to:**

```excel
=LAMBDA(txt, LET(
    t, txt & "",
    IF(LEN(t)=0, "", LET(
        d, CONCAT(IF(ISNUMBER(--MID(t, SEQUENCE(LEN(t)), 1)),
                     MID(t, SEQUENCE(LEN(t)), 1), "")),
        e, IF(LEFT(d,2)="00", MID(d,3,99), d),
        IFS(AND(LEN(e)=12, LEFT(e,3)="998"), e,
            LEN(e)=9, "998"&e,
            AND(LEN(e)=10, LEFT(e,1)="8"), "998"&MID(e,2,99),
            TRUE, "")))))
```

## `AGE.DAYS`

Whole days from a due date until today. Negative while still in the future.

```
=AGE.DAYS(D2)  ->  14
```

**Refers to:**

```excel
=LAMBDA(due, IF(N(due)=0, "", TODAY()-INT(due)))
```

## `AGE.BUCKET`

Receivables aging band for a due date.

```
=AGE.BUCKET(D2)  ->  31-60
```

**Refers to:**

```excel
=LAMBDA(due, LET(n, AGE.DAYS(due),
    IFS(n="", "",
        n<=0,  "not due",
        n<=30, "0-30",
        n<=60, "31-60",
        n<=90, "61-90",
        TRUE,  "90+")))
```

## `M2`

Millimetre dimensions -> square metres.

```
=M2(1200, 600, 40)  ->  28.8
```

**Refers to:**

```excel
=LAMBDA(lengthMM, widthMM, [qty],
    (lengthMM/1000) * (widthMM/1000) * IF(ISOMITTED(qty), 1, qty))
```

## `PCT.OF`

Share of a total, blank-safe -- no #DIV/0! when the total is still empty.

```
=PCT.OF(C2, $C$99)  ->  0.18
```

**Refers to:**

```excel
=LAMBDA(part, whole, IF(N(whole)=0, "", part/whole))
```

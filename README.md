# AkhiSal Excel toolkit

Handy Excel tools for the slab/block business — **none of which require
installing anything.** No add-in, no admin rights, no software request ticket.

---

## First: you don't need the Claude add-in to build Excel tools

The screenshot shows Excel's generic *"One or more add-ins failed to load"*
banner. Worth separating two things that are easy to conflate:

- **Claude *inside* Excel** (the add-in) is a convenience. It lets you chat with
  a sheet in a side panel.
- **Claude *building* Excel tools for you** needs no add-in at all. That is what
  produced this repository. You describe the tool, the code gets written here,
  you paste it into Excel once, and it is yours — it keeps working offline,
  after you close this session, and for colleagues you send the file to.

So the add-in being blocked doesn't block the actual goal. It only changes
where the code gets written.

### If you still want the add-in working

That banner is Excel's catch-all and rarely names the real cause. Check, in
order of how often it's the answer:

1. **File → Options → Add-ins → Manage: COM Add-ins → Go…** — an unrelated
   broken add-in (an old accounting or PDF one) often throws this banner and
   has nothing to do with Claude. Untick anything you don't recognise.
2. **Manage: Disabled Items → Go…** — Excel permanently disables add-ins that
   crashed once. Re-enable and restart Excel.
3. **Corporate policy.** If your organisation runs Microsoft 365 centrally, a
   tenant admin controls which add-ins may load. This is the usual cause on a
   work laptop, and nothing you can do locally will override it — it needs an
   admin to allow the add-in.
4. **Plan and platform.** Claude for Excel has its own plan and Excel-version
   requirements, and they change; check the current ones in Anthropic's docs
   rather than trusting anything cached here.

If 3 is the answer, the tiers below are the way through — none of them are
add-ins, so none of them are subject to that policy.

---

## Just open the workbook

**[`AkhiSal-Tools.xlsx`](AkhiSal-Tools.xlsx)** is finished and needs no setup at all —
no import, no macros to enable, no add-in. Open it and type in the yellow cells.

| Sheet | What it does |
|---|---|
| **Start here** | The legend: yellow = you type, grey = a formula |
| **Client Matcher** | Paste names in any spelling, Latin or Cyrillic, and get the matching CRM client, their phone, and a warning when several clients share the name |
| **CRM Clients** | The reference list — 50 real clients, snapshot 2026-09-19. Paste a fresh export over it any time |
| **Phone Cleaner** | Messy phone spellings in, `998XXXXXXXXX` out |
| **Receivables** | Real unpaid orders, aged against today's date, bucketed 0-30 / 31-60 / 61-90 / 90+ |
| **Money in Words** | `13 500 000` → `o'n uch million besh yuz ming so'm`, for invoices |

It is built entirely from Excel-2007-era functions (`INDEX`, `MATCH`, `SUBSTITUTE`,
`COUNTIF`, `SUMIFS`), so it works in any Excel version and every formula in it was
verified by recalculation before shipping. Rebuild it with
`python3 builder/build_workbook.py`.

**The matcher matches on phone, not just name** — deliberately. The CRM has three
clients called `Abdulloh`, three called `Davron` and two called `Baxodir`, so a
name alone cannot identify anyone. The sheet says `CHECK PHONE` instead of
guessing.

The tiers below are for building *further* tools on top.

---

## The three tiers, most to least capable

Pick the first one your machine allows.

| | Tier | Needs | Blocked by policy when… |
|---|---|---|---|
| 1 | **VBA macros** (`vba/`) | Any desktop Excel | macros are disabled org-wide |
| 2 | **Office Scripts** (`office-scripts/`) | Microsoft 365 licence | rarely — these aren't macros |
| 3 | **LAMBDA formulas** (`lambdas/`) | Microsoft 365 Excel | essentially never |

Tier 3 is plain formulas saved in the workbook, so it survives the strictest
lockdown and travels with the file to anyone you send it to.

---

## Tier 1 — VBA macros (start here)

VBA ships inside Excel. There is nothing to install and it works offline.

**One-time setup, about 60 seconds:**

1. Press **Alt + F11** (the VBA editor opens)
2. **File → Import File…**
3. Pick `vba/AkhiSalTools.bas`, then repeat for `vba/AkhiSalFuncs.bas`
4. **Alt + Q** to go back to Excel
5. Save the workbook as **`.xlsm`** — a plain `.xlsx` silently drops the code

**Running a macro:** **Alt + F8**, pick it, **Run**.

If the ribbon shows a yellow *"Macros have been disabled"* bar, click **Enable
Content**. For a file that came from outside the company, Windows may also need
**right-click the file → Properties → Unblock** before Excel will trust it.

### What's in `AkhiSalTools.bas`

| Macro | What it does |
|---|---|
| `MacroBackupSheet` | Duplicates the sheet first. **VBA edits can't be undone with Ctrl+Z** — run this before anything below. |
| `CleanImport` | Trims stray spaces, strips non-breaking spaces, turns numbers-stored-as-text into real numbers |
| `NormalizePhones` | Rewrites a column to `998XXXXXXXXX`; highlights anything it couldn't read |
| `FormatUZS` / `FormatThousands` | `13 500 000 UZS` number formats |
| `HighlightDuplicates` | Colours repeated values in the selection |
| `SplitByColumn` | Fans one table out into a sheet per region / status / driver |
| `ExportSheetToCSV_UTF8` | CSV that **keeps Cyrillic intact** — Excel's own CSV export mangles it |
| `RemoveEmptyRows` | Deletes fully blank rows |
| `MakePrintReady` | Landscape, fit-to-width, header repeated on every page, page numbers |

`SUMBYCOLOR` / `COUNTBYCOLOR` also live here and are used in cells:
`=SUMBYCOLOR($H$1, F2:F200)` totals every cell matching `H1`'s fill colour.
Excel doesn't recalculate when a colour changes, so press **Ctrl+Alt+F9** after
recolouring.

### What's in `AkhiSalFuncs.bas`

Custom functions you type into cells like any built-in:

| Function | Example | Result |
|---|---|---|
| `TRANSLIT` | `=TRANSLIT("ШУҲРАТЖОН АКА")` | `shuhratjon aka` |
| `SIMILARITY` | `=SIMILARITY("Xalimjon","Халимжон")` | `1.00` |
| `MATCHNAME` | `=MATCHNAME(A2, Clients!A:A)` | closest client name, or blank |
| `UZSWORDS` | `=UZSWORDS(13500000)` | `o'n uch million besh yuz ming so'm` |
| `M2` | `=M2(1200, 600, 40)` | `28.8` |
| `AGEDAYS` | `=AGEDAYS(D2)` | days overdue (negative = not yet due) |
| `AGEBUCKET` | `=AGEBUCKET(D2)` | `not due` / `0-30` / `31-60` / `61-90` / `90+` |

**`TRANSLIT` / `SIMILARITY` / `MATCHNAME` are the ones worth knowing about.**
Client names are written both ways — `Xalimjon` in one sheet, `Халимжон` in the
next — and a normal `VLOOKUP` sees two unrelated strings. These transliterate
both sides first, so the same person matches at `1.00`. `MATCHNAME` scans a
column and returns the closest name above a cutoff (default `0.85`), which makes
reconciling an Excel list against the CRM's client list a one-column job.

`UZSWORDS` writes an amount out in Uzbek for invoices and contracts.

> Cyrillic in these files is written as `ChrW()` codes rather than literal
> letters. The VBA editor imports `.bas` files using the machine's ANSI
> codepage, which turns pasted Cyrillic into garbage on a Latin-locale
> Windows — `ChrW()` survives that trip intact.

---

## Tier 2 — Office Scripts (when macros are blocked)

Office Scripts are **not** macros, so they usually survive a policy that kills
VBA. They run in Excel on the web and from the Automate tab of desktop Excel on
a Microsoft 365 licence.

1. **Automate** tab → **New Script**
2. Delete the starter code, paste a file from `office-scripts/`
3. **Save**, then **Run**

- `CleanImport.ts` — the Tier 1 cleanup, as a script
- `AgingReport.ts` — builds a receivables aging summary on a new `Aging` sheet.
  Set the two column letters at the top of the file to match your sheet first.

---

## Tier 3 — LAMBDA formulas (works under any lockdown)

Plain formulas, saved into the workbook via **Formulas → Name Manager**. No
macros, no scripts, nothing to enable — and they travel with the file, so
colleagues get them with nothing to install.

See **[`lambdas/library.md`](lambdas/library.md)** for `TRANSLIT`, `PHONE.UZ`,
`AGE.DAYS`, `AGE.BUCKET`, `M2` and `PCT.OF`, each with paste-ready text.

Needs Excel from Microsoft 365 — `LAMBDA` doesn't exist in Excel 2019 or earlier.

---

## Adding more tools

Everything here is plain text, so describe the tool you want and it gets written
into this repo the same way. Worth knowing what's *worth* building: a lot of
what people script by hand is already a built-in — **Power Query** (Data → Get
Data) handles repeated imports and merges better than any macro, and pivot
tables beat a hand-rolled summary macro. The tools above are deliberately the
ones Excel genuinely lacks.

---

## Verifying the tricky parts

Excel isn't available in CI, so `tests/verify_algorithms.py` mirrors the two
algorithms most likely to break quietly — Uzbek place-value grammar in
`UZSWORDS`, and the Cyrillic/Latin fold behind `TRANSLIT` / `SIMILARITY` — and
pins their expected output against real order totals and client names.

```
python3 tests/verify_algorithms.py
```

Change the VBA, change these to match.

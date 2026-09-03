# xlsx_forms_to_csv

Convert selected tabs ("forms") out of an Excel CRF/EDC export into
individual CSV files, bundled into one zip — without changing any of the
underlying data.

You pick the forms by their **real-world name** (the heading shown in
cell A1 of each tab, e.g. "DEMOGRAPHY"), not by the short internal tab
code (e.g. `DM`). The script re-reads the tab list fresh from whatever
workbook you point it at, so it works on any future export that follows
the same layout — nothing about a specific study is hardcoded.

---

## 1. What it assumes about the workbook

Each worksheet is expected to look like this:

| Row | Contents                                              |
|-----|--------------------------------------------------------|
| 1   | Form title (e.g. `DEMOGRAPHY`) — read from cell **A1**  |
| 2   | Human-readable field labels                            |
| 3   | Machine field codes                                     |
| 4+  | One data row per subject / record                      |

The worksheet's own tab name (e.g. `DM`, `TUN`, `AE`) is the short
internal code — that's what you'll see if you click through the tabs in
Excel, but it's *not* what you type into this script. You type the
title from row 1 instead (or the code, if you already know it — see
§5.4).

If a future export uses a different cell for the title, override it with
`--title-cell` (see §6).

---

## 2. Requirements & install

- Python 3.8+
- One dependency: [`openpyxl`](https://pypi.org/project/openpyxl/)

```bash
pip install -r requirements.txt
# or, without a venv:
pip install openpyxl --break-system-packages
```

No other setup. The script is a single file, `xlsx_forms_to_csv.py`.

---

## 3. Where does the Excel file go?

Nowhere in particular — the script just needs a path to it. Your `.xlsx`
doesn't need to live inside this folder or be renamed; you point the
script at it as the first argument on the command line, and it's only
ever read, never modified or moved.

**Option A — drop it into this folder (simplest):**

```
xlsx_forms_to_csv/
├── xlsx_forms_to_csv.py
├── requirements.txt
├── README.md
├── forms_this_study.txt
└── study_export.xlsx          <- your file, placed here
```

```bash
cd xlsx_forms_to_csv
python xlsx_forms_to_csv.py study_export.xlsx --forms-file forms_this_study.txt
```

**Option B — leave it wherever it already is, and give the full path:**

```bash
# macOS / Linux
python xlsx_forms_to_csv.py /path/to/somewhere/study_export.xlsx --forms-file forms_this_study.txt

# Windows
python xlsx_forms_to_csv.py "C:\Users\you\Downloads\study_export.xlsx" --forms-file forms_this_study.txt
```

Both work identically — Option A is just fewer characters to type. Since
the whole point of this script is to be reused on new exports over
time, Option B (pointing at wherever the file naturally lands, e.g. a
downloads folder or a shared drive) is usually the more sustainable
habit.

The output **zip** goes to your current folder by default, or wherever
you pass to `--output-dir` (see §6) — that's
the only file location choice that matters ongoing.

---

## 4. Quick start

```bash
# 1. See what forms exist in the workbook (no conversion happens yet)
python xlsx_forms_to_csv.py study_export.xlsx --list

# 2. Convert the ones you need
python xlsx_forms_to_csv.py study_export.xlsx --forms \
    "Demography" "Adverse Event" "Body Measurement"
```

This writes a zip (e.g. `csv_export_20260902_120000.zip`) to the current
folder, containing one CSV per matched form, named after the form's
title.

---

## 5. Ways to tell the script which forms you want

Pick whichever fits your workflow — you only need one.

### 5.1 `--forms` — type the names directly

```bash
python xlsx_forms_to_csv.py study_export.xlsx --forms \
    "Demography" \
    "Reproductive Status" \
    "Body Measurement" \
    "Adverse Event"
```

Case doesn't matter (`"demography"`, `"DEMOGRAPHY"`, `"Demography"` all
match the same tab), and exact wording doesn't have to match the title
character-for-character — see §5.6 for how
the fuzzy matching behaves.

### 5.2 `--forms-file` — a text file, one form per line

```
# forms.txt
# Lines starting with # are ignored, as are blank lines.
Demography
Reproductive Status
Body Measurement
Adverse Event
```

```bash
python xlsx_forms_to_csv.py study_export.xlsx --forms-file forms.txt
```

Best when the list is long, reused often, or handed to a
non-technical teammate to edit.

### 5.3 `--interactive` — numbered picker, no typing of names at all

```bash
python xlsx_forms_to_csv.py study_export.xlsx --interactive
```

```
Tabs found in this workbook:

    1) ADVERSE EVENT FORM  [AE, 4 data row(s)]
    2) ADVERSE EVENTS ASSESSMENT  [AES, 1 data row(s)]
    3) ARCHIVAL TISSUE  [AT, 2 data row(s)]
    ...

Enter the numbers of the forms you want, comma-separated (e.g. 1,3,7), or 'all':
>
```

No risk of a mismatch here since you're selecting by number, not by
name — good for a one-off run or when you're not sure of the exact
wording.

### 5.4 Forcing an exact tab

If you already know a tab's short code (from `--list`), you can type
that instead of the form name — in `--forms`, `--forms-file`, or at the
`--interactive` prompt indirectly via its list position. An exact code
match always wins immediately, bypassing fuzzy matching entirely:

```bash
python xlsx_forms_to_csv.py study_export.xlsx --forms CM BCR AE
```

### 5.5 `--all` (or just writing "all")

Skip selection and export every tab in the workbook:

```bash
python xlsx_forms_to_csv.py study_export.xlsx --all
```

You can also just write `all` as a form name instead — it's recognised
in any capitalisation (`all`, `ALL`, `All`, ...) and works the same way
whether it's passed via `--forms` or written as a line in
`--forms-file`:

```bash
python xlsx_forms_to_csv.py study_export.xlsx --forms all
```

### 5.6 How matching works (for `--forms` / `--forms-file`)

For each name you type, the script tries, in order:

1. **Exact tab code** (e.g. you typed `TU`) → picked immediately, and
   only that one tab.
2. **Exact title match**, ignoring case/spacing/punctuation.
3. **Fuzzy content match** — compares the words in your phrase against
   the words in every tab's title (case-insensitive, plurals normalized,
   filler words like "and"/"the"/"of" ignored). If one tab is clearly
   the best fit, it's picked with no fuss.
4. **Numbered continuation tabs are pulled together automatically.**
   Some forms are too wide for one Excel tab and get split into several
   (e.g. "Study drug administrations" → `EX` / `EX1` / `EX2` / `EX3`,
   covering columns 1-15 / 16-30 / 31-45 / 46-55). The script detects
   these because they share the same title once the trailing `(...)`
   range is stripped off, so typing the plain form name once — e.g.
   `Study drug administrations` — exports all of its continuation tabs
   in one go. Typing an exact tab code (e.g. `EX1`) still selects just
   that single tab, per point 1.
5. If the top candidates score close together and none of the above
   resolved it, the script assumes your phrase might actually be
   **naming two different forms joined by "and" / "&" / a comma** (e.g.
   "Biopsy Collection and Tissue Archival") and tries matching each half
   separately. If that resolves cleanly to two different tabs, both are
   exported.
6. If it's still ambiguous, the script picks its best guess **and prints
   a warning naming the close runner-up**, with the exact flag to force
   that one instead:

   ```
   ambiguous match: "Concomitant Medications" -> "CONCOMITANT MEDICATION
   ASSESSMENT" (CMA, score 0.80), but "PRIOR & CONCOMITANT MEDICATION
   FORM" (CM, score 0.70) is close behind. If that's the one you
   actually meant, rerun with --forms CM for that entry.
   ```

   This happens for genuinely similarly-named tabs where no amount of
   text matching can read your intent for certain (e.g. a medication log
   vs. a "was an assessment done" checkbox with a near-identical name).
   When you see this warning, just re-run with the suggested `--forms
   <CODE>` for that one entry, or use `--list` / `--interactive` to pick
   precisely.
7. If nothing scores reasonably, the script reports no match for that
   name rather than guessing wildly — run `--list` to see what's
   actually in the workbook.

---

## 6. Full option reference

```
python xlsx_forms_to_csv.py WORKBOOK [WORKBOOK ...] [options]

positional:
  WORKBOOK              One or more .xlsx files to read.

selection (pick exactly one approach):
  --forms NAME [NAME ...]   Form names to export (case-insensitive, fuzzy-matched)
  --forms-file PATH         Text file, one form name per line
  --interactive              Numbered checklist prompt
  --all                      Export every tab

inspection:
  --list                     Print every tab's form name + code, then exit
                              (ignores the selection options above)

other:
  --title-cell CELL          Cell holding the form title (default: A1)
  --keep-title-row            Include the form-title row (row 1) in the CSV output.
                              By default it's left out on every tab (see §7 below).
  --output-dir DIR           Where to write the zip (default: current folder)
  --output-name NAME         Zip filename (default: auto-generated with a timestamp)
```

### Multiple workbooks in one run

```bash
python xlsx_forms_to_csv.py site_a.xlsx site_b.xlsx --forms "Demography" "Adverse Event"
```

Each workbook's CSVs are kept in their own sub-folder inside the zip
(named after the workbook file), so identically-named forms from
different files never collide.

---

## 7. The form-title row is left out of every CSV, on purpose

Each tab's row 1 is just a heading used to identify the form (e.g.
`ADVERSE EVENT FORM`, `DEMOGRAPHY`) — it isn't actual field data, so by
default the script drops it from the exported CSV. This applies to
**every tab it exports, not just one specific form** — every CSV you get
starts straight from the field-label row (row 2) and the field-code row
(row 3), then the data.

If you specifically want that title row included — e.g. to keep a visual
label at the top of each file — add `--keep-title-row`:

```bash
python xlsx_forms_to_csv.py study_export.xlsx --forms "Adverse Event" --keep-title-row
```

This is a single switch for the whole run — it's either included on
every exported tab or excluded from every exported tab, there's no
per-tab option.

---

## 8. What's preserved, and one thing to know about opening CSVs in Excel

- Cells are copied to CSV exactly as stored — text stays text. If a
  field is stored as `"001"`, the CSV contains `001`.
- If a cell is a genuine *number* that carries a zero-padding display
  format (e.g. Excel format `000`), the script reproduces that padding
  in the CSV text too, since plain CSV has no concept of Excel's display
  formatting.
- Output is UTF-8 with a BOM, so special characters (e.g. `µ`) and
  Excel both display correctly.
- Blank spacer rows some forms use are kept, so row structure matches
  the source tab.

**One caveat that's about Excel, not the file itself:** if you
double-click a CSV to open it, Excel's own auto-detection can still
*display* `001` as `1` — that's Excel reformatting on open, the file on
disk is untouched. To see it exactly as written, use Excel's
**Data → Get Data → From Text/CSV** and set that column's type to Text
during import, instead of double-clicking the file.

---

## 9. Files in this folder

| File                      | Purpose                                              |
|----------------------------|-------------------------------------------------------|
| `xlsx_forms_to_csv.py`     | The script — the only thing you need to run           |
| `requirements.txt`         | The one dependency (`openpyxl`)                       |
| `forms_this_study.txt`     | Example `--forms-file` input, pre-filled for the 13 forms discussed for this study (two entries use exact tab codes — `BCR`, `CM` — to sidestep the ambiguous-name cases noted above) |
| `README.md`                | This file                                             |

`forms_this_study.txt` is just a convenience example — feel free to
copy it, rename it, and edit the list for a different set of forms or a
different workbook.

---

## 10. Troubleshooting

**"No good match found for ..."**
Run `--list` on that exact workbook and check the spelling/wording
against what's actually there — tab titles can vary slightly between
export versions.

**A form matched, but not the one I meant**
Check the console output — a genuinely close call always prints a
warning with the alternative's exact code. Re-run with `--forms
<CODE>` for that one entry, or use `--interactive` / `--list` instead.

**"File not found, skipping: ..."**
Check the path. The script skips missing files and continues with the
rest rather than stopping entirely.

**Nothing was exported**
Either no forms matched, or none were selected — the script won't
create an empty zip; it will exit with an error message instead.
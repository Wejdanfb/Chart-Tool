# MCH Monthly Consolidation Tool

Consolidates each month's unit-level Nursing Chart Review Checklist files
(Madinah Health Cluster) into one formatted `.xlsx` report, with a live
Summary sheet — replacing the manual copy/paste process.

## Running it locally

You need Python 3.9+ installed. Then, from this folder:

```
pip install -r requirements.txt
python app.py
```

Open **http://localhost:5050** in your browser. Leave the terminal window
open while you use it — closing it stops the tool. To stop on purpose,
press `Ctrl+C` in that terminal.

Nothing is uploaded anywhere — the browser page talks only to the Python
process running on your own machine, and all files stay in this folder's
`uploads/` and `outputs/` subfolders (safe to delete between months).

## How to use it each month

1. Drag in (or choose) all of this month's unit report files.
2. Check the review table — each row shows the detected unit, patient
   count, and status. Anything that isn't recognized is flagged with why,
   and blocks report generation until removed or replaced. A file with
   multiple Unit/Ward values, or an unusual merged cell, shows up as a
   warning underneath its status — not blocking, just worth a glance.
3. The column order in the final report follows the order files appear
   in that table. Use the ↑/↓ buttons to rearrange, or × to remove one.
4. Optionally type a name in **Output file name** — just the name, no
   `.xlsx` needed. Leave it blank to get the default name. This only
   changes what the downloaded file is called; it has no effect on the
   report's content.
5. Click **Generate Final Report**, then **Download Final Excel Report**.

## How consolidation works (so this keeps making sense next month)

- The **first file in the list becomes the master** for formatting,
  section structure, and checklist wording — its columns A:K, fonts,
  borders, merges, and row layout carry through to the output unchanged.
  Put your cleanest file first if you notice a unit's copy has typos in
  the item text (the tool will warn you about this when it detects a
  mismatch, but won't block generation over it).
- A file is recognized as valid **based only on its `Nursing Chart
  Review Checklist` sheet and structure** — the right section headers,
  the right patient-column layout, and a `Score` column. A unit file's
  own `SUMMSRY` sheet, if it has one, is completely ignored: not
  required, not read, not copied. Plenty of real monthly files don't
  have one at all, and that's fine.
- Every file's patient columns are placed side by side, **in list
  order**, with each file's own "Score" column dropped.
- **One new Score column is added at the very end** — its position is
  never fixed to a particular letter, it's always recalculated as
  `column L + total patient count` for that month's batch.
- Every per-item score and section-average formula is **recalculated
  fresh** against the full combined patient range for that run — not
  copied from any individual file.
- If every uploaded patient is "NA" for a given checklist item, that
  item's score cell is intentionally left blank (avoids a `#DIV/0!`
  error) — matching how the original manual reports have always handled
  it. If at least one patient anywhere in the batch has a real 0/1/2 for
  that item, the formula is included.
- **The number of sections is never assumed.** How many sections a
  checklist has — 9, 12, or any other count — is detected from the
  checklist's actual CONTENT: section headers in column A (rows like
  "A. INITIAL ASSESSMENT"), never from whether a section-average formula
  happens to still exist there. Real monthly files regularly have that
  one cell overwritten with "NA", a hard-typed number, or left blank by
  manual editing — that's a data-quality blemish in a single cell, not
  evidence the section is missing. Confirmed directly against two real
  files (P2CD, PICU): identical headers and item text at all 12
  sections, despite 1–3 average cells each having been individually
  damaged — both correctly recognized as the same 12-section template.
- **Every other uploaded file must use the exact same checklist
  structure as the master file** — same section headers, at the same
  row positions. Since patient columns are placed by row position, a
  file whose own section headers sit at different rows (or say
  something different) than the master's is a genuinely different
  checklist template, and combining it would silently attach its
  patients' data to the wrong checklist items. A damaged average
  formula never factors into this comparison — only the actual header
  text and position do, so a file is never wrongly rejected just
  because one of its cells got overwritten. The tool checks this for
  every file before generating anything, and refuses with a specific
  explanation (both files' section positions and headers, spelled out)
  rather than guessing or misaligning data. If you see this, either
  remove the mismatched file or reorder so a file matching the others
  is first — this is also the correct, working result when two files
  are genuinely different versions of the checklist (verified against
  real files: an older August-dated file has one fewer item in section
  A than the current September template, shifting everything after it
  — combining those two would misalign real data, and the tool
  correctly refuses).
- **A section's average is always calculated fresh from the final
  consolidated Score column, never copied from any source file.** This
  falls directly out of the point above: since a section is recognized
  by its header/items regardless of whether its own average formula
  survived, the final report always gets a real, live `AVERAGE(...)`
  formula for every detected section — including ones where every
  contributing source file's own average cell was damaged.
- **The Summary sheet is generated from scratch every run**, directly
  from the consolidated checklist — never copied or read from any
  source file's own Summary. It always has exactly as many rows as
  the master file has sections, and its labels come straight from the
  consolidated checklist's own section-header text, so they're never
  wrong or out of sync with the actual number of sections.
- **The dropdown (0/1/2/NA) validation on patient cells is rebuilt from
  the master file's own actual list source**, wherever it happens to
  live in that file — never assumed to be a fixed cell range. Every
  other property of that validation (blank-allowed, error messages,
  etc.) is copied from the master file too, and the whole thing is
  reapplied to span the full consolidated patient width. If the master
  file has no dropdown at all, none is invented, and a warning says so.
  This is checked, not assumed: every run actually resolves the source
  (whether it's a plain range, a defined name, or a defined name
  pointing at a hidden sheet) and confirms it produces real values,
  logging the result either way.
- **Merged cells inside the patient/Score-column area never crash a
  run, in ANY uploaded file** — not just the master. Some units format
  their file a little differently (e.g. merging one Unit/Ward label
  across a few patient columns instead of repeating it in every cell).
  The tool finds every merge in that area, in every file, before
  reading from it, and dissolves it — but for anything other than the
  two known structural merges (the title banner, and the Score
  column's blank metadata cell), it first copies the merged value into
  every cell the merge used to cover, so unmerging never silently
  blanks data. A file having multiple Unit/Ward values, or an unusual
  merge, is logged (see "Where the diagnostic detail goes" below) and
  never blocks generation.
- **A patient column with no actual response data anywhere in the
  checklist is left out of the final report entirely** — not based on
  its Unit/Ward cell, and not just because some of its cells are
  blank. Every column between L and the file's own Score column is
  checked against that file's real item rows; if there's a 0/1/2, an
  "NA", or any other real response anywhere in that column, it's kept
  in full. The kept columns from every file are laid out with no gaps
  and renumbered 1, 2, 3… in the final sheet, and the Score formulas
  only ever reference that final, gap-free range.
- **The output filename you type in the browser only ever changes the
  downloaded file's name** — invalid Windows characters (`\ / : * ? "
  < > |`) are stripped automatically, `.xlsx` is added for you, and an
  empty field falls back to the default name. It's handled entirely in
  `app.py`, completely separate from `consolidate.py`'s actual report
  generation.
- **A source file's own damaged average formula is logged, never
  blocked.** If a section's average cell in a source file has been
  overwritten with "NA", a hard-typed number, or left blank, that's
  noted as a data-quality observation about that specific file (see
  "Where the diagnostic detail goes") — it doesn't stop that file's
  section from being recognized, and it doesn't stop generation.

## Where the diagnostic detail goes

A successful run shows a clean confirmation in the browser: total
patients, the Score column, units included, the file name, and the
download button — nothing else. All the underlying checks (unusual
merges, checklist wording differences, multiple Unit/Ward values,
excluded empty columns, dropdown verification) still run on every
single file, every time — they just report to the terminal window
running `python app.py`, not the browser, so a normal month ends with
a simple "done" screen instead of a wall of internal detail. Worth
checking that terminal if a report ever looks off.

## Files

- `app.py` — local web server (Flask): serves the page, handles
  upload/validate/reorder/generate/download, the output-filename
  sanitizing, and logs the diagnostic detail mentioned above.
- `consolidate.py` — all the actual Excel logic (openpyxl). This is
  where to look if a future month's template changes and something
  needs adjusting.
- `templates/`, `static/` — the browser UI.
- `uploads/`, `outputs/` — working folders, created automatically.

## If something gets flagged as "not recognized"

The tool checks that a file has a sheet named exactly `Nursing Chart
Review Checklist`, that row 14 reads `A. INITIAL ASSESSMENT`, and that
it can find a column literally labeled `Score` in row 9. That's the
entire recognition rule — a missing `SUMMSRY` sheet is never a reason
for rejection. If a unit's file structure has drifted from the
checklist requirements above, you'll see the specific reason in the
status column rather than a silently wrong report.

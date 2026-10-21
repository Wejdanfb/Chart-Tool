"""
MCH Nursing Chart Review Checklist - monthly consolidation engine.

Everything here was reverse-engineered directly from real source files and a
real final example (never hand-designed), per the rules below:

- Every uploaded workbook holds one or more Nursing Chart Review Checklist
  worksheets. A worksheet is recognized by its CONTENT - never by its name,
  position, the file name, or the presence of a SUMMSRY sheet: it must have a
  "Score" column header (row 9 in the standard layout) to the right of the
  master columns, section headers in column A ("A. INITIAL ASSESSMENT" ...),
  numbered checklist items under them, and patient columns:
    * columns A:K = fixed master structure (item numbers, descriptions,
      section headers) - identical across compatible sheets
    * columns L onward = one column per patient reviewed that month
    * immediately after the last patient column, one column literally
      labeled "Score" in row 9, holding per-item formulas
      =((SUM(<patient range>))/(COUNT(<patient range>)*2))
      and per-section AVERAGE formulas at "average rows"
  Blank sheets, cover pages, summary sheets and anything else that doesn't look
  like a checklist are simply not processed. Every recognized worksheet of every
  file contributes its patient columns (file by file, sheet by sheet, in upload
  order). A single sheet may also mix several units/wards.

- Duplicate protection: each kept patient column gets an identity fingerprint
  (MRN/dates/unit rows + every response). A column identical to one already taken
  from a DIFFERENT sheet or file is skipped, so a workbook that holds duplicate
  copies of the same data never counts a patient twice. Columns with no
  identifying metadata at all are never treated as duplicates.

- The FIRST recognized checklist sheet of the FIRST uploaded file supplies the A:K
  structure, row heights, legend, and overall formatting - the master template
  for the output (no separate blank master file exists). The output sheet is
  always named 'Nursing Chart Review Checklist', whatever the source sheet was
  called.

- Consolidation = concatenate every file's patient columns in upload order
  (dropping each file's own Score column), then add ONE new Score column
  at the very end with freshly recalculated formulas across the FULL
  consolidated patient range. Nothing is copied from any file's own Score
  column - it's fully recomputed.

- Row/column position of the Score column is NEVER hardcoded - it's always
  (12 + total patient count) at the time of generation.

- Section count is NEVER assumed. A checklist can have any number of
  sections - 9, 12, or any other count.

- Section structure is detected from the checklist's actual CONTENT -
  section header text and item positions in columns A/B - NEVER from
  whether a section-average formula happens to still exist (architecture
  revised 2026-09-15, see below). Real monthly files regularly have their
  average-row formula overwritten with "NA", a hard-typed number, or left
  blank by manual editing; that is a data-quality blemish in one cell, not
  evidence the section doesn't exist. A section header is any row in
  column A matching the established "X. SECTION NAME" pattern; each
  section's extent runs from just after its header to the next row where
  columns A and B are BOTH blank (the conventional average-row position,
  regardless of what - if anything - actually occupies it).

- Every OTHER recognized sheet's section headers (row positions AND header
  text) and numbered items are compared against the master's before
  consolidating. If they don't match, that sheet uses a genuinely different
  checklist template and is rejected with a specific explanation (which headers
  or items differ) rather than silently misaligning its patient columns to the
  wrong checklist items. A damaged/missing average formula never enters
  into this comparison - only the actual section structure does.

- SUMMSRY: a source workbook's own 'SUMMSRY' sheet, if it has one, is
  *completely irrelevant* to this tool. It is never read, never required,
  and never copied, and it never determines how many sections are expected.
  Recognizing a file as a valid source depends only on its checklist
  worksheet(s) - some real monthly files don't have a SUMMSRY sheet at all,
  and that's fine.

      SOURCE FILES (checklist + patient data)
            |
      CONSOLIDATION ENGINE  <- section headers/positions detected fresh
            |                  from the base file's checklist CONTENT
            |                  every run - never from formula survival
      Combined checklist + freshly recalculated scores  <- every section
            |                  average is calculated here, from the final
            |                  consolidated Score column, always - a
            |                  source file's own (possibly damaged)
            |                  average value is never copied
      GENERATE NEW SUMMARY  <- built here, sized to however many sections
                                were actually detected, from the
                                consolidated checklist, every time. Labels
                                come from the consolidated checklist's own
                                section-header text.
"""

import copy
import datetime
import gc
import hashlib
import os
import pickle
import re
from dataclasses import asdict, dataclass, field
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SHEET_CHECKLIST = "Nursing Chart Review Checklist"  # canonical name given to the OUTPUT sheet -
                                                     # source sheets may be named anything
SHEET_SUMMARY = "SUMMSRY"
PATIENT_START_COL = 12  # column L - first patient column (columns A:K are the fixed master structure)
HEADER_SEARCH_ROWS = 40  # a sheet's "Score" header is looked for in its first N rows
SCORE_SEARCH_COLS = 200  # ...and in up to this many columns right of column L
MIN_ITEM_ROWS = 3  # a sheet needs at least this many numbered checklist items to count as a checklist
STRUCTURE_SCAN_LIMIT = 2000  # generous safety ceiling when scanning for section headers
TRAILING_ROW_PADDING = 12  # rows carried past the last section-average row (blank/
                           # signature rows below the last section) - not a structural
                           # assumption, just how far past the last real content to copy
SUMMARY_HEADER_ROW = 17  # layout choice for the tool's OWN generated sheet - not derived
                         # from any source file, so this is fine to fix
SUMMARY_LABEL_COL = 3  # C
SUMMARY_FORMULA_COL = 4  # D

# Section heading rows are normally titled "A. SOME NAME" / "F.SOME NAME". That wording is
# only a convention - it is NEVER required: sections are found from the checklist's
# structure (numbered items, average rows - see _find_section_headers) and a heading may be
# blank or worded differently. This pattern is only used as an extra hint (see there).
_SECTION_HEADER_RE = re.compile(r"^[A-Z]\.")

# Rows of the standard checklist template that carry a section heading (row 73 as supplied;
# the heading of section H sits on row 79 in the current template and is found
# structurally). The TEXT of these rows - and of every heading row a sheet is found to have
# - is never compared or validated: blank, reworded or reformatted headings are all
# accepted. Everything else about the structure is still checked.
TITLE_EXEMPT_ROWS = frozenset({14, 32, 36, 50, 55, 63, 68, 73, 89, 93, 104, 111})
_ITEM_NUMBER_RE = re.compile(r"^\d+\.?$")
_UNIT_LABEL_RE = re.compile(r"unit\s*/?\s*ward", re.IGNORECASE)


class ValidationError(Exception):
    pass


@dataclass
class SheetAnalysis:
    """One worksheet that was recognized, from its CONTENT alone, as a Nursing Chart
    Review Checklist (never from its name, position, or a fixed cell)."""
    name: str
    header_row: int  # row holding the patient index numbers + the "Score" label
    score_col: int  # 1-indexed column of this sheet's OWN "Score" column
    section_headers: list  # row positions of this sheet's section headers
    section_header_text: list  # normalized header text, same order
    item_numbers: list  # [(row, item number text)] for every numbered item row
    avg_rows: list  # section-average rows (derived from the headers, never from formulas)
    kept_offsets: list  # 0-indexed offsets from column L that have real response data
    excluded_empty_count: int = 0  # candidate columns dropped for having no data at all
    unit_name: Optional[str] = None
    column_keys: list = field(default_factory=list)  # per kept column: identity fingerprint
                                                     # (or None when it can't be told apart)
    warnings: list = field(default_factory=list)

    @property
    def patient_count(self):
        return len(self.kept_offsets)


@dataclass
class FileAnalysis:
    path: str
    filename: str
    valid: bool
    error: Optional[str] = None
    warnings: list = field(default_factory=list)
    unit_name: Optional[str] = None
    patient_count: int = 0  # across all recognized sheets (before cross-file duplicate removal)
    sheets: list = field(default_factory=list)  # [SheetAnalysis] - every recognized checklist sheet
    skipped_sheets: list = field(default_factory=list)  # [(sheet name, reason)] - internal only

    def to_dict(self):
        return asdict(self)

    @classmethod
    def from_dict(cls, d):
        d = dict(d)
        d["sheets"] = [SheetAnalysis(**s) for s in d.get("sheets", [])]
        d["skipped_sheets"] = [tuple(x) for x in d.get("skipped_sheets", [])]
        return cls(**d)


def _norm(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def _norm_value(v):
    """Stable text form of a cell value for fingerprinting (dates, 2 vs 2.0, stray spaces)."""
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.isoformat()
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return _norm(v)


def _find_score_header(ws):
    """Locate the sheet's own 'Score' column header: a cell that reads exactly 'Score' in
    one of the first rows, to the right of the master columns. Returns (row, col) or None.
    This - together with section headers and numbered items - is what identifies a
    checklist sheet; the sheet's NAME is never consulted."""
    last_row = min(ws.max_row, HEADER_SEARCH_ROWS)
    last_col = min(ws.max_column, PATIENT_START_COL + SCORE_SEARCH_COLS)
    for r in range(1, last_row + 1):
        for c in range(PATIENT_START_COL, last_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().lower() == "score":
                return r, c
    return None


def _unit_row(ws, header_row, first_header_row):
    """Row holding each patient's Unit/Ward - found by its label in column A (between the
    header row and the first section), falling back to header_row + 3 (row 12 in the
    standard layout) if no label matches."""
    for r in range(header_row + 1, first_header_row):
        if _UNIT_LABEL_RE.search(_norm(ws.cell(row=r, column=1).value)):
            return r
    return header_row + 3


def _analyze_sheet(ws, extract_out=None):
    """Decide from CONTENT whether `ws` is a Nursing Chart Review Checklist. Returns
    (SheetAnalysis, None) if it is, or (None, reason) if not - blank sheets, cover pages,
    summary sheets and anything else land in the second case. If `extract_out` is a dict,
    the sheet's patient-column data is also captured into it (see _extract_sheet) while the
    sheet is already loaded, so generating the report never has to re-open this workbook."""
    if ws.max_row < 5 or ws.max_column < PATIENT_START_COL:
        return None, "too small to be a checklist"

    found = _find_score_header(ws)
    if found is None:
        return None, "no 'Score' column header found"
    header_row, score_col = found

    patient_count = score_col - PATIENT_START_COL
    if patient_count <= 0:
        return None, "no patient columns between column L and the Score column"

    # Dissolve merges in the patient/Score region IN MEMORY ONLY (nothing here is saved) so
    # every cell read below - including patients whose cells were visually merged - holds
    # its real value. See _sweep_merges_in_region.
    _, _, merge_other_count = _sweep_merges_in_region(ws, PATIENT_START_COL, score_col)

    section_headers, header_problem = _find_section_headers(ws, header_row, score_col)
    if header_problem:
        return None, header_problem
    if not section_headers:
        return None, ("no checklist sections found (expected blocks of numbered items "
                      "separated by a blank average row)")
    boundaries = _compute_boundaries(ws, section_headers)
    avg_rows = [b[0] for b in boundaries]
    section_header_text = [_norm(ws.cell(row=h, column=1).value) for h in section_headers]
    item_rows = _item_rows_for(ws, boundaries)

    item_numbers = []
    for avg_row, start, end in boundaries:
        for r in range(start, end + 1):
            a = _norm(ws.cell(row=r, column=1).value)
            if _ITEM_NUMBER_RE.match(a):
                item_numbers.append((r, a.rstrip(".")))
    if len(item_numbers) < MIN_ITEM_ROWS:
        return None, "no numbered checklist items under the section headers"

    # A candidate patient column is only excluded if it has NO response data anywhere in
    # the actual checklist response area (the item rows - never based on the Unit/Ward
    # row, and never just because some individual cells are blank). Any real content at
    # all - a 0/1/2 score, "NA"/"N/A", any other text response - counts; formatting alone
    # never does, since only cell values are inspected here.
    kept_offsets, excluded_offsets = [], []
    for i in range(patient_count):
        col = PATIENT_START_COL + i
        has_response = any(_norm(ws.cell(row=r, column=col).value) != "" for r in item_rows)
        (kept_offsets if has_response else excluded_offsets).append(i)

    if not kept_offsets:
        return None, (f"all {patient_count} candidate patient column(s) are empty - no "
                      f"response data found in the checklist area")

    # Unit/ward name(s), over the KEPT columns only. Usually one value repeated across the
    # sheet, but a sheet can legitimately mix several units - that is fine and not an
    # error: every kept patient column is still included regardless of its label.
    unit_row = _unit_row(ws, header_row, section_headers[0])
    unit_values = [_norm(ws.cell(row=unit_row, column=PATIENT_START_COL + i).value)
                   for i in kept_offsets]
    unique_units = sorted(set(u for u in unit_values if u))
    warnings = []
    if len(unique_units) > 1:
        unit_name = " + ".join(unique_units)
        warnings.append(
            f"Multiple different Unit/Ward values found: {unique_units}. All "
            f"{len(kept_offsets)} patient columns with response data are still included."
        )
    else:
        unit_name = unique_units[0] if unique_units else None

    if excluded_offsets:
        warnings.append(
            f"{len(excluded_offsets)} of {patient_count} candidate patient column(s) had "
            f"no response data anywhere in the checklist and were excluded (not counted "
            f"as patients, no gap left in the numbering)."
        )

    # Data-quality note only, never a structural one (see _section_has_working_average).
    damaged = [section_header_text[i] or f"the section at row {section_headers[i]}"
               for i, avg_row in enumerate(avg_rows)
               if not _section_has_working_average(ws, score_col, avg_row)]
    if damaged:
        warnings.append(
            f"{len(damaged)} section(s) had a damaged/overwritten average formula in this "
            f"source ({', '.join(damaged)}) - a data-quality issue in the source only. The "
            f"section is still fully recognized from its header and items, and its average "
            f"in the final report is always calculated fresh, never copied from here."
        )

    # Identity fingerprint per kept patient column: everything the reviewer entered about
    # that patient (the metadata rows - MRN, dates, unit - plus every response). Used only
    # to avoid counting the same patient twice when a workbook, or two uploaded files,
    # contain copies of the same data. A column with NO metadata at all (no MRN/dates/unit)
    # gets None: without any identifying detail, two identical-looking columns could be
    # two different patients, so such a column is never treated as a duplicate.
    meta_rows = range(header_row + 1, section_headers[0])
    column_keys = []
    for i in kept_offsets:
        col = PATIENT_START_COL + i
        meta = [_norm_value(ws.cell(row=r, column=col).value) for r in meta_rows]
        if not any(meta):
            column_keys.append(None)
            continue
        resp = [_norm_value(ws.cell(row=r, column=col).value) for r in item_rows]
        column_keys.append(hashlib.sha1(repr((meta, resp)).encode("utf-8")).hexdigest())

    if extract_out is not None:
        extract_out[ws.title] = _extract_sheet(
            ws, header_row, kept_offsets, max(avg_rows), merge_other_count
        )

    return SheetAnalysis(
        name=ws.title, header_row=header_row, score_col=score_col,
        section_headers=section_headers, section_header_text=section_header_text,
        item_numbers=item_numbers, avg_rows=avg_rows, kept_offsets=kept_offsets,
        excluded_empty_count=len(excluded_offsets), unit_name=unit_name,
        column_keys=column_keys, warnings=warnings,
    ), None


def _analyze_workbook(path: str, filename: str, want_extract: bool):
    """Open the workbook ONCE, recognize its checklist worksheet(s), and (optionally)
    capture each one's patient-column data. Returns (FileAnalysis, extracts-or-None). The
    workbook is closed and released before returning, so at most one source workbook is
    ever held in memory at a time."""
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception as e:
        return FileAnalysis(path, filename, valid=False, error=f"Could not open file: {e}"), None

    extracts = {} if want_extract else None
    try:
        sheets, skipped = [], []
        for ws in wb.worksheets:
            try:
                sa, reason = _analyze_sheet(ws, extracts)
            except Exception as e:  # one odd sheet must never take the whole workbook down
                sa, reason = None, f"could not be read ({e})"
            if sa is None:
                skipped.append((ws.title, reason))
                if extracts is not None:
                    extracts.pop(ws.title, None)
            else:
                sheets.append(sa)
    finally:
        try:
            wb.close()
        except Exception:
            pass
        del wb
        gc.collect()  # openpyxl workbooks are reference cycles - free them now, not "eventually"

    if not sheets:
        detail = "; ".join(f"'{n}': {r}" for n, r in skipped) or "the workbook has no worksheets"
        return FileAnalysis(
            path, filename, valid=False, skipped_sheets=skipped,
            error=f"No Nursing Chart Review Checklist worksheet was found in this "
                  f"workbook. Every worksheet was checked by its content - {detail}.",
        ), None

    # File-level summary, with each sheet's own notes prefixed by its name when the
    # workbook holds more than one checklist sheet.
    multi = len(sheets) > 1
    warnings = []
    for sa in sheets:
        prefix = f"Sheet '{sa.name}': " if multi else ""
        warnings.extend(prefix + w for w in sa.warnings)
    units = sorted(set(u for sa in sheets for u in re.split(r" \+ ", sa.unit_name or "") if u))
    return FileAnalysis(
        path, filename, valid=True, warnings=warnings,
        unit_name=" + ".join(units) if units else None,
        patient_count=sum(sa.patient_count for sa in sheets),
        sheets=sheets, skipped_sheets=skipped,
    ), extracts


def _extract_path(path: str) -> str:
    return path + ".extract"


def analyze_file(path: str, filename: str, extract: bool = False) -> FileAnalysis:
    """Inspect EVERY worksheet of an uploaded workbook and recognize the checklist
    worksheet(s) by their content (section headers, numbered items, patient columns and a
    Score column) - never by sheet name, sheet position, filename, or the presence of a
    SUMMSRY sheet. A workbook may hold one checklist sheet or several (e.g. one per
    department); unrelated, blank and summary/cover sheets are simply not processed.

    With extract=True (what the web app uses at upload time) the recognized sheets'
    patient-column data is saved next to the file (`<path>.extract`), so that generating
    the report later does not have to parse this workbook a second time - parsing is by far
    the most expensive (CPU and memory) step on a small server."""
    info, extracts = _analyze_workbook(path, filename, extract)
    if extract and info.valid and extracts:
        tmp = _extract_path(path) + f".{os.getpid()}.tmp"
        with open(tmp, "wb") as fh:
            pickle.dump(extracts, fh, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, _extract_path(path))
    return info


def _load_extracts(info: "FileAnalysis") -> dict:
    """Patient-column data for every recognized sheet of `info`'s file: read from the file
    saved at upload time if present, otherwise (e.g. consolidate() called directly) captured
    now by opening the workbook once - exactly the same code path either way."""
    try:
        with open(_extract_path(info.path), "rb") as fh:
            data = pickle.load(fh)
        if all(sa.name in data for sa in info.sheets):
            return data
    except Exception:
        pass
    _, extracts = _analyze_workbook(info.path, info.filename, True)
    if not extracts or not all(sa.name in extracts for sa in info.sheets):
        raise ValidationError(f"{info.filename}: could not re-read the file's checklist data.")
    return extracts


# ---------------------------------------------------------------------------
# Batch planning: compatibility with the master structure + duplicate protection.
# Used both by the review screen (from stored analyses) and by consolidate().
# ---------------------------------------------------------------------------

@dataclass
class SheetPlan:
    sheet: SheetAnalysis
    use_offsets: list  # offsets actually consolidated (kept columns minus duplicates)
    notes: list = field(default_factory=list)


@dataclass
class FilePlan:
    analysis: FileAnalysis
    sheets: list = field(default_factory=list)  # [SheetPlan]
    error: Optional[str] = None  # blocking incompatibility with the master structure
    notes: list = field(default_factory=list)  # non-blocking: duplicates skipped, etc.

    @property
    def patient_count(self):
        return sum(len(sp.use_offsets) for sp in self.sheets)

    @property
    def unit_name(self):
        """Unit/Ward label(s) of the sheets that actually contribute patients (a sheet
        whose patients were all skipped as duplicates doesn't add its unit)."""
        units = sorted(set(
            u for sp in self.sheets if sp.use_offsets
            for u in re.split(r" \+ ", sp.sheet.unit_name or "") if u
        ))
        return " + ".join(units) if units else self.analysis.unit_name


def _title_exempt_rows(*sheets) -> set:
    """Rows whose heading text/number is never validated: the fixed template heading rows
    plus every heading row actually found in the sheet(s) being compared."""
    rows = set(TITLE_EXEMPT_ROWS)
    for sa in sheets:
        rows.update(sa.section_headers)
    return rows


def _structure_difference(base: SheetAnalysis, other: SheetAnalysis):
    """None if `other` can safely be combined with `base`; otherwise a specific,
    human-readable explanation of how the two checklist structures differ. Compared on the
    checklist's actual structure - header row, section positions and item numbers (never
    the wording of a section heading) - because patient answers are copied by row position, so a mismatch would put
    answers under the wrong checklist items."""
    if base.header_row != other.header_row:
        return (f"its patient/Score header row is row {other.header_row}, but the master "
                f"checklist's is row {base.header_row}")
    # Heading WORDING is deliberately not compared (blank / reworded headings are fine); the
    # sections themselves - how many there are and where each starts - are.
    if len(other.section_headers) != len(base.section_headers):
        return (f"it has {len(other.section_headers)} checklist sections (starting at rows "
                f"{other.section_headers}), but the master checklist has "
                f"{len(base.section_headers)} (starting at rows {base.section_headers})")
    if other.section_headers != base.section_headers:
        pairs = [f"section {i + 1} starts at row {y} (master: row {x})" for i, (x, y) in
                 enumerate(zip(base.section_headers, other.section_headers)) if x != y]
        return ("its sections are at different rows, so the items would not line up - "
                + "; ".join(pairs[:4]) + (" ..." if len(pairs) > 4 else ""))
    exempt = _title_exempt_rows(base, other)
    base_map = {r: n for r, n in base.item_numbers if r not in exempt}
    other_map = {r: n for r, n in other.item_numbers if r not in exempt}
    if other_map != base_map:
        rows = sorted(set(base_map) | set(other_map))
        diffs = [r for r in rows if base_map.get(r) != other_map.get(r)]
        shown = ", ".join(
            f"row {r} (item {other_map.get(r, 'none')} vs master {base_map.get(r, 'none')})"
            for r in diffs[:4]
        )
        return (f"its numbered checklist items differ from the master at {len(diffs)} "
                f"row(s): {shown}" + (" ..." if len(diffs) > 4 else ""))
    return None


def plan_batch(analyses: list) -> list:
    """Given the FileAnalysis of every uploaded file, in upload order, work out for each
    file (a) whether its checklist sheet(s) match the master structure - the first
    recognized sheet of the first valid file - and (b) which patient columns to actually
    take, skipping any that duplicate a column already taken from a DIFFERENT sheet or
    file (first occurrence in upload order wins) so no patient is counted twice.
    Returns one FilePlan per input, same order. Pure function of the analyses."""
    plans = [FilePlan(analysis=a) for a in analyses]
    base = None
    for a in analyses:
        if a.valid and a.sheets:
            base = a.sheets[0]
            break

    seen = {}  # column fingerprint -> (file index, sheet name) it was first taken from
    for fi, (a, plan) in enumerate(zip(analyses, plans)):
        if not (a.valid and a.sheets):
            continue
        for sa in a.sheets:
            diff = _structure_difference(base, sa) if sa is not base else None
            if diff:
                where = f"sheet '{sa.name}' in {a.filename}" if len(a.sheets) > 1 else a.filename
                plan.error = (
                    f"{where}: {diff}. These are different checklist templates and cannot "
                    f"be safely combined - patient data would misalign to the wrong "
                    f"checklist items. Remove this file, or upload a file that matches the "
                    f"master structure first (the master is the first uploaded file)."
                )
                break
            use, dup_from = [], {}
            for off, key in zip(sa.kept_offsets, sa.column_keys):
                origin = seen.get(key) if key is not None else None
                if origin is not None and origin != (fi, sa.name):
                    dup_from.setdefault(origin, 0)
                    dup_from[origin] += 1
                    continue
                if key is not None:
                    seen.setdefault(key, (fi, sa.name))
                use.append(off)
            sp = SheetPlan(sheet=sa, use_offsets=use)
            for (ofi, oname), n in dup_from.items():
                src = analyses[ofi].filename
                note = (f"{n} patient column(s) in "
                        f"{'sheet ' + repr(sa.name) if len(a.sheets) > 1 else 'this file'} "
                        f"duplicate patient data already taken from "
                        f"{'sheet ' + repr(oname) + ' of ' if len(analyses[ofi].sheets) > 1 else ''}"
                        f"{src} and were skipped so no patient is counted twice.")
                sp.notes.append(note)
                plan.notes.append(note)
            plan.sheets.append(sp)
    return plans


def _extract_sheet(ws, header_row, kept_offsets, last_avg_row, merge_other_count):
    """Capture, from an already-loaded and merge-swept sheet, everything consolidate() needs
    from it: each kept patient column's cell values and cell styles (rows 1..the copy
    extent, exactly what _copy_patient_column would read), its explicit column width, the
    A/B checklist wording (for the non-blocking wording comparison), and how many unexpected
    merges had to be dissolved. A few KB per sheet - the heavy openpyxl workbook is not
    needed again. Identical styles are stored once and referenced by number."""
    last_row = last_avg_row + TRAILING_ROW_PADDING
    styles, style_ids, columns = [], {}, {}
    for offset in kept_offsets:
        col = PATIENT_START_COL + offset
        cells = []
        for r in range(1, last_row + 1):
            cell = ws.cell(row=r, column=col)
            sid = None
            if cell.has_style:
                key = tuple(cell._style)
                sid = style_ids.get(key)
                if sid is None:
                    sid = style_ids[key] = len(styles)
                    styles.append((
                        copy.copy(cell.font), copy.copy(cell.border), copy.copy(cell.fill),
                        copy.copy(cell.alignment), cell.number_format, copy.copy(cell.protection),
                    ))
            cells.append((cell.value, sid))
        dim = ws.column_dimensions.get(get_column_letter(col))
        columns[offset] = {"cells": cells, "width": dim.width if (dim and dim.width) else None}
    text = {r: (_norm(ws.cell(row=r, column=1).value), _norm(ws.cell(row=r, column=2).value))
            for r in range(header_row, last_avg_row + 1)}
    return {"styles": styles, "columns": columns, "text": text, "merge_other": merge_other_count}


def _apply_extracted_column(extract, offset, dst_ws, dst_col, last_row, style_cache):
    """Write one captured patient column into the consolidated sheet - the same result as
    _copy_patient_column reading from the live source sheet. The first time a given captured
    style is applied it is registered in the output workbook the normal way; the resulting
    internal style record is then reused for every other cell with that style (a patient
    column has only a handful of distinct styles, repeated hundreds of times), which avoids
    re-hashing/re-copying full style objects per cell - the other big CPU cost after parsing."""
    col = extract["columns"][offset]
    styles = extract["styles"]
    for r, (value, sid) in enumerate(col["cells"], start=1):
        if r > last_row:
            break
        dst_cell = _ensure_writable(dst_ws, r, dst_col)
        dst_cell.value = value
        if sid is None:
            continue
        cache_key = (id(extract), sid)
        cached = style_cache.get(cache_key)
        if cached is not None:
            dst_cell._style = copy.copy(cached)
            continue
        font, border, fill, alignment, number_format, protection = styles[sid]
        dst_cell.font = copy.copy(font)
        dst_cell.border = copy.copy(border)
        dst_cell.fill = copy.copy(fill)
        dst_cell.alignment = copy.copy(alignment)
        dst_cell.number_format = number_format
        dst_cell.protection = copy.copy(protection)
        style_cache[cache_key] = copy.copy(dst_cell._style)


def compare_extract_text(master_ws, extract, first_row, last_row, master_headers=None) -> list:
    """Same check as compare_to_master_text, against the A/B wording captured when the
    other sheet was read."""
    diffs = []
    skip = TITLE_EXEMPT_ROWS | set(master_headers or ())  # heading wording is never compared
    for r in range(first_row, last_row + 1):
        if r in skip:
            continue
        a_m, b_m = _norm(master_ws.cell(row=r, column=1).value), _norm(master_ws.cell(row=r, column=2).value)
        a_o, b_o = extract["text"].get(r, ("", ""))
        if a_m != a_o or b_m != b_o:
            diffs.append(r)
    if diffs:
        return [
            f"Checklist text differs from the master sheet at row(s) "
            f"{', '.join(str(r) for r in diffs[:6])}"
            + (f" and {len(diffs) - 6} more" if len(diffs) > 6 else "")
            + ". Only the master sheet's wording is used in the output."
        ]
    return []


def compare_to_master_text(master_ws, other_ws, first_row, last_row, master_headers=None) -> list:
    """Light consistency check: flag if a sheet's own A/B checklist wording differs from
    the master sheet's. Non-blocking - just a warning (typo fixes etc. are common), since
    only the master's A:K wording ends up in the output. Structural differences (headers,
    item numbers, positions) are blocking and handled by _structure_difference."""
    diffs = []
    skip = TITLE_EXEMPT_ROWS | set(master_headers or ())  # heading wording is never compared
    for r in range(first_row, last_row + 1):
        if r in skip:
            continue
        a_m, b_m = master_ws.cell(row=r, column=1).value, master_ws.cell(row=r, column=2).value
        a_o, b_o = other_ws.cell(row=r, column=1).value, other_ws.cell(row=r, column=2).value
        if _norm(a_m) != _norm(a_o) or _norm(b_m) != _norm(b_o):
            diffs.append(r)
    if diffs:
        return [
            f"Checklist text differs from the master sheet at row(s) "
            f"{', '.join(str(r) for r in diffs[:6])}"
            + (f" and {len(diffs) - 6} more" if len(diffs) > 6 else "")
            + ". Only the master sheet's wording is used in the output."
        ]
    return []




def _copy_cell_style(src_cell, dst_cell):
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.alignment = copy.copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy.copy(src_cell.protection)


def _copy_patient_column(src_ws, src_col, dst_ws, dst_col, last_row):
    """Copy one full patient column's values + cell styles from src to dst, through
    `last_row` (computed per-run from the actual detected checklist extent - see
    consolidate(), never a fixed constant). Column width is handled separately by
    _set_column_width (applies to every patient column, including the first file's own,
    which is never "copied"). Every destination cell is passed through _ensure_writable
    first: the upfront region sweep should already have cleared every merge here, but
    this is a cheap per-cell safety net against any merge the sweep didn't anticipate."""
    same_workbook = src_ws.parent is dst_ws.parent
    for r in range(1, last_row + 1):
        src_cell = src_ws.cell(row=r, column=src_col)
        dst_cell = _ensure_writable(dst_ws, r, dst_col)
        dst_cell.value = src_cell.value
        if same_workbook:
            # Same workbook = same style tables, so the internal style record can be copied
            # as-is (identical result to copying font/border/fill/... one by one, far cheaper).
            if src_cell.has_style:
                dst_cell._style = copy.copy(src_cell._style)
        else:
            _copy_cell_style(src_cell, dst_cell)


def _set_column_width(src_ws, src_col, dst_ws, dst_col, fallback_width):
    """If the source column has no explicit width set (some source files leave one or
    two patient columns at Excel's un-set default), that default is much narrower than
    its neighbors and visibly squeezes/wraps the cell content. In that case a fallback
    width (the master template's own first patient column width) is used instead,
    rather than faithfully reproducing what is really just a formatting oversight in
    that file. Applied to every patient column, including the base file's own."""
    src_letter = get_column_letter(src_col)
    dst_letter = get_column_letter(dst_col)
    src_dim = src_ws.column_dimensions.get(src_letter)
    width = src_dim.width if (src_dim and src_dim.width) else fallback_width
    if width:
        dst_ws.column_dimensions[dst_letter].width = width


def _sweep_merges_in_region(ws, min_col: int, base_score_col: int):
    """Finds every merged range in `ws` that touches column `min_col` or beyond, and
    unmerges ALL of them - nothing outside that column range (the untouched A:K master
    columns, other sheets) is ever looked at, so section-header merges like B17:K17
    are never affected.

    Every cell we're about to write a value into (via _copy_patient_column, the new
    Score column, or the row-9 index renumbering) must be a real, normal cell first -
    writing to a non-anchor MergedCell raises AttributeError. Rather than guessing which
    merge patterns a given month's file might use, every merge in the region is found by
    inspection and dissolved before any writing starts; the two patterns known to be
    load-bearing are recorded here so they can be rebuilt at the final width once it's
    known:

      - "wide" merges: start at column A (e.g. the title banner, the scoring-legend row)
        and are rebuilt as A{row}:{final_score_col}{row} at the end.
      - "score-column" merges: confined to a single column at exactly the CURRENT file's
        own Score column (e.g. the blank rows-10-13 cell under a Score header, since that
        column has no per-patient metadata) - rebuilt at the NEW Score column position,
        same row span.

    Any other merge found in the region (e.g. a unit spreading one Unit/Ward label
    across several patient columns instead of repeating it in every cell) is dissolved
    and NOT rebuilt - a merge spanning several patients no longer has a single consistent
    meaning once those patients are individually laid out at the new consolidated width.
    But before dissolving one of these, the merged (anchor) value and style are copied to
    every cell the merge covered, so unmerging doesn't silently blank out data that was
    only visually displayed via the merge (e.g. a Unit/Ward label spanning 3 patient
    columns must not leave 2 of those 3 patients with a blank Unit/Ward after this)."""
    wide_spans = []
    score_col_spans = []
    other_count = 0

    for m in list(ws.merged_cells.ranges):
        if m.max_col < min_col:
            continue  # entirely outside the region we're about to modify - leave alone
        if m.min_col == 1:
            wide_spans.append((m.min_row, m.max_row))
        elif m.min_col == m.max_col == base_score_col:
            score_col_spans.append((m.min_row, m.max_row))
        else:
            other_count += 1
            anchor = ws.cell(row=m.min_row, column=m.min_col)
            anchor_value = anchor.value
            anchor_style = {
                "font": copy.copy(anchor.font), "border": copy.copy(anchor.border),
                "fill": copy.copy(anchor.fill), "alignment": copy.copy(anchor.alignment),
                "number_format": anchor.number_format, "protection": copy.copy(anchor.protection),
            }
            ws.unmerge_cells(str(m))
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    if r == m.min_row and c == m.min_col:
                        continue
                    cell = ws.cell(row=r, column=c)
                    cell.value = anchor_value
                    cell.font = anchor_style["font"]
                    cell.border = anchor_style["border"]
                    cell.fill = anchor_style["fill"]
                    cell.alignment = anchor_style["alignment"]
                    cell.number_format = anchor_style["number_format"]
                    cell.protection = anchor_style["protection"]
            continue
        ws.unmerge_cells(str(m))

    return wide_spans, score_col_spans, other_count


def _ensure_writable(ws, row: int, col: int):
    """Defensive belt-and-braces check: if, despite the upfront sweep, the target cell
    is still part of a merge (e.g. a merge that only appeared after the sweep ran, or one
    whose top-left anchor sits outside the swept region but tail overlaps into it),
    unmerge it before writing. Cheap no-op in the overwhelmingly common case."""
    cell = ws.cell(row=row, column=col)
    if cell.__class__.__name__ == "MergedCell":
        for m in list(ws.merged_cells.ranges):
            if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
                ws.unmerge_cells(str(m))
                break
    return ws.cell(row=row, column=col)


def _capture_source_validation(ws):
    """Capture the base file's OWN data-validation "template" - what values are allowed
    and how the validation behaves - before anything gets cleared/rebuilt.

    The exact referenced cells (formula1) MUST come from this specific file, never be
    assumed: a different month's master file could position its legend/allowed-values
    list somewhere other than B117:B120, or reference it differently, and hardcoding a
    specific cell reference would silently point the rebuilt dropdown at the wrong
    (possibly blank) cells for that file - producing an empty-looking dropdown despite
    the DataValidation object technically existing. Returns None if the base file has no
    list-type validation to model the rebuild on; callers must handle that by not
    rebuilding a dropdown, since there is nothing to know it should look like."""
    for dv in ws.data_validations.dataValidation:
        if dv.type == "list" and dv.formula1:
            return {
                "type": dv.type, "formula1": dv.formula1, "formula2": dv.formula2,
                "operator": dv.operator, "allowBlank": dv.allowBlank,
                "showDropDown": dv.showDropDown, "showErrorMessage": dv.showErrorMessage,
                "showInputMessage": dv.showInputMessage, "errorTitle": dv.errorTitle,
                "error": dv.error, "promptTitle": dv.promptTitle, "prompt": dv.prompt,
                "errorStyle": dv.errorStyle,
            }
    return None


_CELL_RANGE_RE = re.compile(
    r"^\$?(?P<c1>[A-Z]+)\$?(?P<r1>\d+)(:\$?(?P<c2>[A-Z]+)\$?(?P<r2>\d+))?$"
)
_SHEET_QUALIFIED_RE = re.compile(
    r"^(?:'(?P<sheet_q>[^']+)'|(?P<sheet_u>[A-Za-z_][\w.]*))!(?P<rest>.+)$"
)


def _resolve_dropdown_values(wb, ws, formula1: str):
    """Actually resolve `formula1` to the list of values it points at - the same
    resolution Excel/Calc performs to populate the dropdown - so a broken source can be
    caught and reported at generation time instead of only discovered by a person
    manually opening the file and clicking a cell. Handles the three forms a source file
    could plausibly use:

      1. A direct literal list, e.g. formula1='"0,1,2,NA"' - values split out directly.
      2. A same-sheet cell range, e.g. '$B$117:$B$120' - resolved against `ws`.
      3. A defined name (workbook- or sheet-scoped), which may itself point at a range on
         ANY sheet, hidden or not - resolved via wb.defined_names, then read from
         whichever sheet it actually names.

    Returns (values, note): `values` is the list of non-blank strings found (empty list
    if resolution genuinely found nothing), `note` is None on success or a short string
    explaining why resolution couldn't be completed (unknown name, missing sheet, etc.).
    Never raises - a formula this function can't parse is reported via `note`, not an
    exception, since an unusual-but-valid Excel construct shouldn't block a whole run."""
    formula1 = (formula1 or "").strip()
    if not formula1:
        return [], "empty formula1"

    # Form 1: literal quoted list, e.g. "0,1,2,NA"
    if formula1.startswith('"') and formula1.endswith('"'):
        values = [v.strip() for v in formula1[1:-1].split(",") if v.strip()]
        return values, None

    target_ws, rest = ws, formula1
    m = _SHEET_QUALIFIED_RE.match(formula1)
    if m:
        sheet_name = m.group("sheet_q") or m.group("sheet_u")
        rest = m.group("rest")
        if sheet_name not in wb.sheetnames:
            return [], f"references sheet '{sheet_name}' which doesn't exist in this workbook"
        target_ws = wb[sheet_name]
    else:
        # Not sheet-qualified - could be a same-sheet range, or a defined name (which
        # openpyxl reports without a leading '='). Try a defined name first if it isn't
        # shaped like a cell reference.
        if not _CELL_RANGE_RE.match(rest):
            defined = wb.defined_names.get(rest)
            if defined is None:
                return [], f"'{rest}' is not a cell range on this sheet and no defined " \
                           f"name by that name exists in the workbook"
            dest = list(defined.destinations)
            if not dest:
                return [], f"defined name '{rest}' has no target range"
            sheet_name, ref = dest[0]
            if sheet_name not in wb.sheetnames:
                return [], f"defined name '{rest}' points at sheet '{sheet_name}' which " \
                           f"doesn't exist in this workbook"
            target_ws = wb[sheet_name]
            rest = ref

    m2 = _CELL_RANGE_RE.match(rest.replace("$", ""))
    if not m2:
        return [], f"could not parse '{rest}' as a cell range"
    c1, r1 = m2.group("c1"), int(m2.group("r1"))
    c2, r2 = (m2.group("c2") or c1), int(m2.group("r2") or r1)
    from openpyxl.utils import column_index_from_string
    col1, col2 = column_index_from_string(c1), column_index_from_string(c2)

    values = []
    for r in range(min(r1, r2), max(r1, r2) + 1):
        for c in range(min(col1, col2), max(col1, col2) + 1):
            v = target_ws.cell(row=r, column=c).value
            if v is not None and str(v).strip() != "":
                values.append(str(v))
    return values, None


def _find_section_headers(ws, header_row: int, score_col: int):
    """Find every section's heading row from the checklist's STRUCTURE, not from what the
    heading says. Returns (rows, problem): `problem` is None, or a message when a section's
    heading row cannot be placed (the sheet is then not treated as a checklist).

    Structure used: a section is a run of consecutive rows (columns A/B not both blank)
    that contains numbered items; it is closed by a row where A and B are both blank (its
    average row). Within a run the heading row is
      * the row directly above its first numbered item, if it has no patient responses
        (whatever it says - "A. INITIAL ASSESSMENT", other wording, anything), or
      * the blank row just above the run, if the run starts straight away with an item
        (i.e. the heading cell was left blank).
    Runs without numbered items (the Medical Record Number / Date rows, the legend, signature
    lines) are not sections. A row whose column A matches the traditional 'X. NAME' pattern is
    additionally accepted as a heading, as before, so layouts that put sections back to back
    without a blank row keep working. Heading wording is never validated.
    Scans generously (bounded by the sheet's used range, capped only as a safety ceiling)
    so this works whether a checklist has 5 sections or 20."""
    last_row = min(ws.max_row, STRUCTURE_SCAN_LIMIT)

    def blank_ab(r):
        return _norm(ws.cell(row=r, column=1).value) == "" and _norm(ws.cell(row=r, column=2).value) == ""

    def numbered(r):
        return bool(_ITEM_NUMBER_RE.match(_norm(ws.cell(row=r, column=1).value)))

    def has_response(r):
        return any(_norm(ws.cell(row=r, column=c).value) != ""
                   for c in range(PATIENT_START_COL, score_col))

    # runs of consecutive non-blank rows
    runs, start = [], None
    for r in range(header_row + 1, last_row + 2):
        if r <= last_row and not blank_ab(r):
            if start is None:
                start = r
        elif start is not None:
            runs.append((start, r - 1))
            start = None

    rows = set()
    seen_section = False
    for first, last in runs:
        if not any(numbered(r) for r in range(first, last + 1)):
            continue  # not a section (metadata rows, legend, signature, ...)
        first_item = next(r for r in range(first, last + 1) if numbered(r))
        # Step up over any sub-heading lines (column A empty, a description in column B)
        # sitting between the heading and the first item.
        r = first_item - 1
        while r >= first and _norm(ws.cell(row=r, column=1).value) == "":
            r -= 1
        if r >= first:
            # a text row inside the run directly above the (sub-heading and) first item
            if has_response(r):
                return [], (f"the checklist section whose first item is on row {first_item} has "
                            f"no heading row above it (row {r} holds patient answers)")
            rows.add(r)  # heading row, whatever it says
        else:
            above = first - 1  # the run opens with the items: heading cell was left blank
            before_above_ok = (
                above - 1 < header_row + 1          # just under the patient header rows
                or blank_ab(above - 1)              # two blank rows: [average row][blank heading]
                or not seen_section                  # first section: only metadata rows above
            )
            if above > header_row and blank_ab(above) and before_above_ok:
                rows.add(above)
            else:
                return [], (f"the checklist section whose first item is on row {first_item} has "
                            f"no heading row above it (expected a heading - or a blank row - "
                            f"between the previous section's average row and its items)")
        seen_section = True

    # traditional 'X. NAME' headings (also inside runs), accepted in addition
    for r in range(header_row + 1, last_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and _SECTION_HEADER_RE.match(v.strip()):
            rows.add(r)
    return sorted(rows), None


def _is_data_row(ws, row: int, avg_rows: set) -> bool:
    """A row is a data-eligible ("item") row unless it's an average row or a
    section-header row (column A holding non-empty text, e.g. 'A. INITIAL ASSESSMENT').
    Most item rows have an integer in column A, but at least one (row 62 in the
    reference files, a sub-heading like 'Additional assessment includes...') has a
    blank column A yet still carries a real per-row score - so column-A type alone
    can't be trusted; only "is this a section header" is checked here."""
    if row in avg_rows:
        return False
    a_val = ws.cell(row=row, column=1).value
    if isinstance(a_val, str) and a_val.strip():
        return False
    return True


def _compute_boundaries(ws, header_rows: list) -> list:
    """For each section header, find its average-row position by scanning forward
    until the first row where BOTH column A and column B are blank - the established
    convention for the row immediately after a section's last item, before the next
    header. This never looks at the Score column: whether a section is recognized,
    and where its boundaries fall, does not depend on whether its average formula
    survived intact - only on the actual header/item content in columns A and B.
    Returns (avg_row, start, end) tuples, one per detected section."""
    boundaries = []
    for i, header_row in enumerate(header_rows):
        next_header = header_rows[i + 1] if i + 1 < len(header_rows) else None
        scan_limit = (next_header - 1) if next_header else min(ws.max_row, header_row + 200)
        start = header_row + 1
        avg_row = None
        for r in range(start, scan_limit + 1):
            if _norm(ws.cell(row=r, column=1).value) == "" and _norm(ws.cell(row=r, column=2).value) == "":
                avg_row = r
                break
        if avg_row is None:
            # No blank row found in range - unusual, but never crash: fall back to
            # the row right before the next header, or the scan limit for the last
            # section (whatever the actual layout turns out to be).
            avg_row = (next_header - 1) if next_header else scan_limit
        boundaries.append((avg_row, start, avg_row - 1))
    return boundaries


def _item_rows_for(ws, boundaries: list) -> list:
    """Every data-eligible ("item") row across every detected section - i.e. the actual
    checklist response area. Used both to decide which rows get a per-item Score formula,
    and to decide whether a candidate patient column has any real response data in it."""
    avg_row_set = {b[0] for b in boundaries}
    rows = []
    for avg_row, start, end in boundaries:
        for r in range(start, end + 1):
            if _is_data_row(ws, r, avg_row_set):
                rows.append(r)
    return rows


def _section_has_working_average(ws, score_col: int, avg_row: int) -> bool:
    """Purely a data-quality check, never a structural one: does this section's
    average-row cell in the Score column actually hold a working AVERAGE(...) formula
    right now? Used only to log an internal note when a source file's own average was
    overwritten (NA, a hard-typed number, blank) - never to decide whether the
    section exists, and never to block consolidation."""
    v = ws.cell(row=avg_row, column=score_col).value
    return isinstance(v, str) and v.strip().upper().startswith("=AVERAGE(")


def _numeric_count(ws, row: int, start_col: int, end_col: int) -> int:
    """Mirrors Excel's COUNT(): counts only numeric cells (the 0/1/2 scores),
    ignoring 'NA' text and blanks - used to decide whether a per-item formula would
    divide by zero for this row across the ACTUAL consolidated patient data (not
    inherited from any single source file's own, possibly-smaller, patient set)."""
    n = 0
    for c in range(start_col, end_col + 1):
        v = ws.cell(row=row, column=c).value
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            n += 1
    return n


# --- Fresh Summary sheet, styled independently of any source workbook (see module
#     docstring: a source file's own SUMMSRY, if any, is never used as a reference here,
#     not even for styling). Colors approximate the standard blue-header look these
#     reports have always used. ---
_SUMMARY_HEADER_FILL = PatternFill("solid", fgColor="4472C4")
_SUMMARY_HEADER_FONT = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
_SUMMARY_LABEL_FONT = Font(name="Calibri", size=11, bold=False, color="000000")
_SUMMARY_SCORE_FONT = Font(name="Calibri", size=11, bold=True, color="000000")
_SUMMARY_THIN = Side(style="thin", color="000000")
_SUMMARY_MEDIUM = Side(style="medium", color="000000")


def _build_summary_sheet(wb, checklist_ws, boundaries, score_col_letter):
    """Builds the SUMMSRY sheet from scratch every run, using only the consolidated
    checklist just built (never any source workbook's own Summary sheet - see module
    docstring). `boundaries` is the list of (avg_row, start, end) tuples already derived
    from the checklist - however many sections that turned out to be (never assumed to
    be any fixed count); section labels come from the checklist's own section-header
    text at each `start - 1` row, so labels are always correct automatically - they're
    simply what that row of the checklist itself says. The sheet's row range grows or
    shrinks to fit however many sections were actually detected."""
    if SHEET_SUMMARY in wb.sheetnames:
        del wb[SHEET_SUMMARY]
    summary_ws = wb.create_sheet(SHEET_SUMMARY, index=wb.index(checklist_ws) + 1)

    summary_first_row = SUMMARY_HEADER_ROW + 1
    summary_last_row = summary_first_row + len(boundaries) - 1

    summary_ws.column_dimensions["A"].width = 9.14
    summary_ws.column_dimensions["C"].width = 46.4
    summary_ws.column_dimensions["D"].width = 11.7
    summary_ws.row_dimensions[SUMMARY_HEADER_ROW].height = 27

    header_border = Border(left=_SUMMARY_MEDIUM, right=_SUMMARY_MEDIUM, top=_SUMMARY_MEDIUM, bottom=_SUMMARY_MEDIUM)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    items_cell = summary_ws.cell(row=SUMMARY_HEADER_ROW, column=SUMMARY_LABEL_COL, value="ITEMS")
    score_cell = summary_ws.cell(row=SUMMARY_HEADER_ROW, column=SUMMARY_FORMULA_COL, value="SCORE")
    for cell in (items_cell, score_cell):
        cell.fill = _SUMMARY_HEADER_FILL
        cell.font = _SUMMARY_HEADER_FONT
        cell.alignment = center
        cell.border = header_border

    for summary_row, (avg_row, start, _end) in zip(range(summary_first_row, summary_last_row + 1), boundaries):
        header_row = start - 1
        label = _norm(checklist_ws.cell(row=header_row, column=1).value)

        label_cell = summary_ws.cell(row=summary_row, column=SUMMARY_LABEL_COL, value=label)
        label_cell.font = _SUMMARY_LABEL_FONT
        label_cell.alignment = left
        label_cell.border = Border(left=_SUMMARY_MEDIUM, right=_SUMMARY_THIN, bottom=_SUMMARY_THIN)

        score_cell = summary_ws.cell(
            row=summary_row, column=SUMMARY_FORMULA_COL,
            value=f"='{checklist_ws.title}'!{score_col_letter}{avg_row}",
        )
        score_cell.font = _SUMMARY_SCORE_FONT
        score_cell.alignment = center
        score_cell.number_format = "0%"
        score_cell.border = Border(left=_SUMMARY_THIN, right=_SUMMARY_MEDIUM, bottom=_SUMMARY_THIN)

    # Close off the bottom border of the table.
    for col in (SUMMARY_LABEL_COL, SUMMARY_FORMULA_COL):
        c = summary_ws.cell(row=summary_last_row, column=col)
        b = c.border
        c.border = Border(left=b.left, right=b.right, top=b.top, bottom=_SUMMARY_MEDIUM)

    return summary_ws


def _retarget_sheet_refs(text, old, new):
    """Rewrite references to sheet `old` as `new` inside a formula/defined-name string
    (quoted or bare). Used when the consolidated sheet is renamed to its canonical name."""
    if not text or old == new:
        return text
    quoted_new = "'" + new.replace("'", "''") + "'"
    text = text.replace("'" + old.replace("'", "''") + "'!", quoted_new + "!")
    if re.fullmatch(r"[A-Za-z_][\w.]*", old):
        text = re.sub(rf"(?<![\w'.]){re.escape(old)}!", quoted_new + "!", text)
    return text


def _rename_checklist_sheet(wb, ws, new_name):
    """Give the consolidated sheet the canonical checklist name, whatever the source
    sheet was called - and keep every reference that pointed at the old name (defined
    names, e.g. a dropdown list's source) pointing at the renamed sheet. Returns the
    (old, new) names actually used. If another sheet already owns the canonical name the
    source's own name is kept, so nothing is ever overwritten."""
    old = ws.title
    if old == new_name or new_name in wb.sheetnames:
        return old, old
    ws.title = new_name
    for dn in list(wb.defined_names.values()):
        if dn.attr_text:
            dn.attr_text = _retarget_sheet_refs(dn.attr_text, old, new_name)
    for sheet in wb.worksheets:
        for dn in list(sheet.defined_names.values()):
            if dn.attr_text:
                dn.attr_text = _retarget_sheet_refs(dn.attr_text, old, new_name)
    return old, new_name


def consolidate(ordered_paths_and_names: list, output_path: str, analyses: list = None) -> dict:
    """
    ordered_paths_and_names: list of (filepath, original_filename) tuples, in the exact
    order patient columns should appear (upload order).
    Every recognized checklist worksheet of every file contributes its patient columns.
    analyses: optional FileAnalysis list, aligned with ordered_paths_and_names - the findings
    already produced when each file was uploaded. Passing them avoids parsing every workbook
    a second time here; omitted, each file is analyzed now (same function either way).

    Memory/CPU design: only the FIRST file's workbook (the output's master) is ever held
    open. Every other source contributes from the compact patient-column extract captured
    when it was analyzed (see _extract_sheet), so no second workbook is loaded alongside it.
    Returns a dict report of what was done (for display / audit).
    """
    if not ordered_paths_and_names:
        raise ValidationError("No files provided.")

    infos = []
    for i, (path, name) in enumerate(ordered_paths_and_names):
        info = analyses[i] if analyses else analyze_file(path, name)
        if not info.valid:
            raise ValidationError(f"{name}: {info.error}")
        infos.append(info)

    # Compatibility with the master structure + duplicate-patient protection (see
    # plan_batch). Anything that would put answers under the wrong checklist items stops
    # the run here with a specific explanation.
    plans = plan_batch(infos)
    for plan in plans:
        if plan.error:
            raise ValidationError(plan.error)

    # --- Base workbook = the FIRST recognized checklist sheet of the FIRST uploaded file,
    #     used as-is for its A:K master structure ---
    base_info = infos[0]
    base_sa = base_info.sheets[0]
    wb = openpyxl.load_workbook(base_info.path, data_only=False)
    ws = wb[base_sa.name]
    base_score_col = base_sa.score_col
    header_row = base_sa.header_row

    # Capture the base sheet's OWN dropdown/list-validation template (allowed values +
    # behavior) before anything is touched - see _capture_source_validation for why this
    # must never be a hardcoded assumption about where the allowed-values list lives.
    source_validation = _capture_source_validation(ws)

    # The section structure, from the base sheet's actual header/item CONTENT - never from
    # whether a section-average formula happens to still work. Every other sheet has
    # already been verified (plan_batch) to have exactly this structure.
    section_headers = base_sa.section_headers
    boundaries = _compute_boundaries(ws, section_headers)
    avg_rows = [b[0] for b in boundaries]
    avg_row_set = set(avg_rows)

    # How far into the patient columns to copy: past the last detected section-average
    # row, plus a little padding for any trailing blank/signature rows below the last
    # section - not a fixed row ceiling, so a shorter or taller checklist both work.
    copy_last_row = max(avg_rows) + TRAILING_ROW_PADDING

    # --- Capture the base sheet's own Score-column style (font/border/fill/etc per row)
    #     and width, BEFORE that column gets overwritten by the next source's data. This is
    #     reused for the brand-new final Score column at the end. ---
    score_style_by_row = {}
    for r in range(1, copy_last_row + 1):
        c = ws.cell(row=r, column=base_score_col)
        score_style_by_row[r] = {
            "font": copy.copy(c.font), "border": copy.copy(c.border),
            "fill": copy.copy(c.fill), "alignment": copy.copy(c.alignment),
            "number_format": c.number_format, "protection": copy.copy(c.protection),
        }
    base_score_letter = get_column_letter(base_score_col)
    score_col_width = None
    dim = ws.column_dimensions.get(base_score_letter)
    if dim and dim.width:
        score_col_width = dim.width

    # Fallback width for any patient column whose source never set one explicitly (see
    # _set_column_width) - use the master sheet's own first patient column (L).
    l_dim = ws.column_dimensions.get("L")
    fallback_patient_width = l_dim.width if (l_dim and l_dim.width) else 10.0

    # --- Find and dissolve every merge in the patient/Score-column region (>= column L)
    #     before any writes happen there - see _sweep_merges_in_region for why this must
    #     be a dynamic sweep rather than a fixed list of expected merge patterns. Columns
    #     A:K (and every other sheet) are never touched. ---
    wide_spans, score_col_spans, other_merge_count = _sweep_merges_in_region(
        ws, PATIENT_START_COL, base_score_col
    )

    # --- Place every recognized sheet's patient columns, in upload order (file by file,
    #     and sheet by sheet within a file) ---
    report_files = []
    current_col = PATIENT_START_COL
    running_index = 0
    style_cache = {}  # (extract, style number) -> output-workbook style record; see _apply_extracted_column
    for fi, (info, plan) in enumerate(zip(infos, plans)):
        extracts = None  # this file's captured patient-column data, loaded only if needed
        warnings = list(info.warnings) + list(plan.notes)
        start_col = current_col
        for sp in plan.sheets:
            sa = sp.sheet
            is_base_sheet = fi == 0 and sa is base_sa
            if is_base_sheet:
                src_ws = ws  # base sheet: data's already in place, nothing to copy
            else:
                if extracts is None:
                    extracts = _load_extracts(info)
                ex = extracts[sa.name]
                # (This sheet's merges in the patient-column area were already dissolved,
                # anchor value propagated to every covered cell, before its data was
                # captured - same fix the base sheet gets just above - so every value
                # captured is a real, independently-readable one.)
                warnings += compare_extract_text(ws, ex, header_row, max(avg_rows), section_headers)
                if ex["merge_other"]:
                    warnings.append(
                        f"{ex['merge_other']} unexpected merged cell(s) found in the "
                        f"patient columns of '{sa.name}' - resolved (value copied to every "
                        f"cell the merge covered) before reading its patient data."
                    )

            for offset in sp.use_offsets:
                src_col = PATIENT_START_COL + offset
                dst_col = current_col
                running_index += 1
                if is_base_sheet:
                    if dst_col != src_col:
                        _copy_patient_column(src_ws, src_col, ws, dst_col, copy_last_row)
                    _set_column_width(src_ws, src_col, ws, dst_col, fallback_patient_width)
                else:
                    _apply_extracted_column(ex, offset, ws, dst_col, copy_last_row, style_cache)
                    width = ex["columns"][offset]["width"] or fallback_patient_width
                    if width:
                        ws.column_dimensions[get_column_letter(dst_col)].width = width
                _ensure_writable(ws, header_row, dst_col)
                ws.cell(row=header_row, column=dst_col).value = running_index
                current_col += 1
        extracts = None

        got = current_col - start_col
        report_files.append({
            "filename": info.filename, "unit": plan.unit_name,
            "patient_count": got,
            "excluded_empty_count": sum(sp.sheet.excluded_empty_count for sp in plan.sheets),
            "sheets": [sp.sheet.name for sp in plan.sheets],
            "columns": (f"{get_column_letter(start_col)}:{get_column_letter(current_col - 1)}"
                        if got else "—"),
            "warnings": warnings,
        })

    final_score_col = current_col
    final_score_letter = get_column_letter(final_score_col)
    last_patient_letter = get_column_letter(final_score_col - 1)
    total_patients = running_index

    # --- Write the new consolidated Score column. Whatever previously sat in this column
    #     (an old patient column, or the old Score column) is cleared first so no stale
    #     value from a source file can survive in it. ---
    for r in range(1, copy_last_row + 1):
        style = score_style_by_row[r]
        cell = _ensure_writable(ws, r, final_score_col)
        if r > header_row:
            cell.value = None
        cell.font = style["font"]
        cell.border = style["border"]
        cell.fill = style["fill"]
        cell.alignment = style["alignment"]
        cell.number_format = style["number_format"]
        cell.protection = style["protection"]

    ws.cell(row=header_row, column=final_score_col).value = "Score"

    # Decide, per item row, whether a per-item formula is safe to write - based on the
    # ACTUAL consolidated patient data we just placed, not on any single source's own
    # (possibly smaller / all-"NA") patient set. A row where every patient is "NA" would
    # divide by zero, so it's deliberately left blank, exactly matching the master
    # template's own convention.
    formula_item_rows, blank_item_rows = [], []
    for avg_row, start, end in boundaries:
        for r in range(start, end + 1):
            if not _is_data_row(ws, r, avg_row_set):
                continue
            if _numeric_count(ws, r, PATIENT_START_COL, final_score_col - 1) > 0:
                formula_item_rows.append(r)
            else:
                blank_item_rows.append(r)

    for r in formula_item_rows:
        ws.cell(row=r, column=final_score_col).value = (
            f"=((SUM(L{r}:{last_patient_letter}{r}))/(COUNT(L{r}:{last_patient_letter}{r})*2))"
        )
    for r in blank_item_rows:
        ws.cell(row=r, column=final_score_col).value = None

    for avg_row, start, end in boundaries:
        ws.cell(row=avg_row, column=final_score_col).value = (
            f"=AVERAGE({final_score_letter}{start}:{final_score_letter}{end})"
        )

    if score_col_width:
        ws.column_dimensions[final_score_letter].width = score_col_width

    # --- Rebuild the merges found during the upfront sweep, at the final width/position.
    #     "Wide" merges (started at column A) are rebuilt spanning A:{final_score_col} at
    #     their original row span. "Score-column" merges are rebuilt at the same row span
    #     but in the NEW Score column. Any other merge found in the region during the
    #     sweep was a one-off and is intentionally left dissolved (see
    #     _sweep_merges_in_region for why). ---
    for min_row, max_row in wide_spans:
        ws.merge_cells(f"A{min_row}:{final_score_letter}{max_row}")
    for min_row, max_row in score_col_spans:
        ws.merge_cells(f"{final_score_letter}{min_row}:{final_score_letter}{max_row}")

    # --- Trim any leftover columns beyond the final Score column ---
    if ws.max_column > final_score_col:
        ws.delete_cols(final_score_col + 1, ws.max_column - final_score_col)

    # --- The other checklist sheets of the base workbook have been merged into this one;
    #     drop them so the output holds a single consolidated checklist. Unrelated
    #     VISIBLE sheets from the base workbook (cover pages, blank sheets) are not part of
    #     the report and are dropped too. Anything a dropdown or defined name may depend on
    #     is kept: hidden / very-hidden sheets, and any sheet whose name is referenced by a
    #     defined name or by the validation source. ---
    checklist_names = {sa.name for sa in base_info.sheets}
    reference_text = " ".join(
        [dn.attr_text or "" for dn in wb.defined_names.values()]
        + [(source_validation or {}).get("formula1") or "", (source_validation or {}).get("formula2") or ""]
    )
    for sheet in list(wb.worksheets):
        if sheet is ws or sheet.title == SHEET_SUMMARY:
            continue
        needed = (
            sheet.sheet_state != "visible"
            or (sheet.title not in checklist_names and sheet.title in reference_text)
        )
        if not needed:
            del wb[sheet.title]

    # --- Canonical sheet name, whatever the source sheet was called ---
    old_title, new_title = _rename_checklist_sheet(wb, ws, SHEET_CHECKLIST)
    if source_validation:
        for key in ("formula1", "formula2"):
            source_validation[key] = _retarget_sheet_refs(source_validation[key], old_title, new_title)

    # --- Rebuild data validation (the 0/1/2/NA dropdown) across the new patient column
    #     width, using the base sheet's OWN validation template captured at the very start
    #     (source_validation) - never a hardcoded reference. If the base sheet didn't have
    #     a list validation to model this on, no dropdown is fabricated from nothing. ---
    row_blocks = [(start, end) for _avg_row, start, end in boundaries]
    ws.data_validations.dataValidation = []
    dropdown_rebuilt = False
    if row_blocks and source_validation:
        sqref = " ".join(f"L{s}:{last_patient_letter}{e}" for s, e in row_blocks)
        dv = DataValidation(
            type=source_validation["type"],
            formula1=source_validation["formula1"],
            formula2=source_validation["formula2"],
            operator=source_validation["operator"],
            allowBlank=source_validation["allowBlank"],
            showDropDown=source_validation["showDropDown"],
            showErrorMessage=source_validation["showErrorMessage"],
            showInputMessage=source_validation["showInputMessage"],
            errorTitle=source_validation["errorTitle"],
            error=source_validation["error"],
            promptTitle=source_validation["promptTitle"],
            prompt=source_validation["prompt"],
            errorStyle=source_validation["errorStyle"],
        )
        dv.sqref = sqref
        ws.add_data_validation(dv)
        dropdown_rebuilt = True

    # Actually resolve the dropdown source to real values against the FINAL workbook - the
    # same resolution Excel does to populate the dropdown - rather than just trusting the
    # DataValidation object exists. Handles a literal list, a same-sheet range, or a
    # defined name pointing anywhere (including a hidden sheet).
    dropdown_values, dropdown_resolution_note = [], None
    if source_validation:
        dropdown_values, dropdown_resolution_note = _resolve_dropdown_values(
            wb, ws, source_validation["formula1"]
        )
        if not dropdown_values:
            dropdown_resolution_note = (
                dropdown_resolution_note
                or f"'{source_validation['formula1']}' resolved to zero non-blank values"
            )

    # --- Build a brand-new SUMMSRY sheet from the consolidated checklist only.
    #     Never reads or copies any source workbook's own SUMMSRY sheet (some don't have
    #     one at all - that's fine, this doesn't depend on it either way). ---
    _build_summary_sheet(wb, ws, boundaries, final_score_letter)

    # Only the consolidated checklist is selected/active (deleting sheets can otherwise
    # leave several tabs "selected", which makes Excel open in group-edit mode).
    wb.active = wb.index(ws)
    for sheet in wb.worksheets:
        sheet.sheet_view.tabSelected = sheet is ws

    checklist_title = ws.title
    wb.save(output_path)
    wb.close()
    del wb, ws
    gc.collect()  # release the output workbook now rather than whenever the GC next runs

    if other_merge_count:
        report_files[0]["warnings"] = report_files[0]["warnings"] + [
            f"{other_merge_count} unexpected merged cell(s) found in the base file's patient/Score "
            f"columns (not the usual banner/legend pattern) - they were unmerged to allow "
            f"consolidation; nothing outside the patient columns was touched."
        ]
    if not dropdown_rebuilt:
        report_files[0]["warnings"] = report_files[0]["warnings"] + [
            "The base file had no list-type dropdown validation to model the consolidated "
            "checklist's dropdowns on, so none were added. If patient columns should have a "
            "0/1/2/NA dropdown, check that the first uploaded file has one."
        ]
    elif dropdown_resolution_note:
        report_files[0]["warnings"] = report_files[0]["warnings"] + [
            f"Could not confirm the dropdown source resolves to real values: "
            f"{dropdown_resolution_note}. The DataValidation was still carried over "
            f"as-is, but it may not show selectable options when opened in Excel - "
            f"worth checking the base file's dropdown setup directly."
        ]

    return {
        "output_path": output_path,
        "total_patients": total_patients,
        "score_column": final_score_letter,
        "files": report_files,
        "checklist_sheet": checklist_title,
        "section_average_rows": avg_rows,
        "blank_item_rows": blank_item_rows,
        "formula_item_rows_count": len(formula_item_rows),
        "dropdown_source": source_validation["formula1"] if source_validation else None,
        "dropdown_values": dropdown_values,
        "dropdown_verified": bool(dropdown_values) and not dropdown_resolution_note,
        "dissolved_unexpected_merges": other_merge_count,
    }

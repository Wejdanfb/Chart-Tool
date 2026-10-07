"""
MCH Nursing Chart Review Checklist - monthly consolidation engine.

Everything here was reverse-engineered directly from real source files and a
real final example (never hand-designed), per the rules below:

- Every uploaded file is one unit/ward's monthly report. It has a sheet named
  'Nursing Chart Review Checklist' with:
    * columns A:K = fixed master structure (item numbers, descriptions,
      section headers) - identical across all files
    * columns L onward = one column per patient reviewed that month
    * immediately after the last patient column, one column literally
      labeled "Score" in row 9, holding per-item formulas
      =((SUM(<patient range>))/(COUNT(<patient range>)*2))
      and per-section AVERAGE formulas at "average rows"

- The FIRST uploaded file's A:K structure, row heights, legend, and overall
  formatting become the master template for the output (no separate blank
  master file exists - confirmed identical structure across real files).

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

- Every OTHER uploaded file's own section headers (row positions AND
  header text) are compared against the base file's before consolidating.
  If they don't match, that file uses a genuinely different checklist
  template and is rejected with a specific explanation (which headers
  differ) rather than silently misaligning its patient columns to the
  wrong checklist items. A damaged/missing average formula never enters
  into this comparison - only the actual section structure does.

- SUMMSRY: a source workbook's own 'SUMMSRY' sheet, if it has one, is
  *completely irrelevant* to this tool. It is never read, never required,
  and never copied, and it never determines how many sections are expected.
  Recognizing a file as a valid source depends only on its 'Nursing Chart
  Review Checklist' sheet and structure - some real monthly files don't
  have a SUMMSRY sheet at all, and that's fine.

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
import re
from dataclasses import dataclass, field
from typing import Optional

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

SHEET_CHECKLIST = "Nursing Chart Review Checklist"
SHEET_SUMMARY = "SUMMSRY"
PATIENT_START_COL = 12  # column L
HEADER_ROW = 9  # row with sequential patient index numbers + "Score" label
STRUCTURE_SCAN_LIMIT = 2000  # generous safety ceiling when scanning for section headers
TRAILING_ROW_PADDING = 12  # rows carried past the last section-average row (blank/
                           # signature rows below the last section) - not a structural
                           # assumption, just how far past the last real content to copy
SUMMARY_HEADER_ROW = 17  # layout choice for the tool's OWN generated sheet - not derived
                         # from any source file, so this is fine to fix
SUMMARY_LABEL_COL = 3  # C
SUMMARY_FORMULA_COL = 4  # D

# A section header row always looks like "A. SOME NAME" / "F.SOME NAME" (a single
# capital letter, a period, then the section name) in every real file seen so far -
# never a formula or number. This is the actual evidence a section exists.
_SECTION_HEADER_RE = re.compile(r"^[A-Z]\.")


class ValidationError(Exception):
    pass


@dataclass
class FileAnalysis:
    path: str
    filename: str
    valid: bool
    error: Optional[str] = None
    warnings: list = field(default_factory=list)
    unit_name: Optional[str] = None
    patient_count: int = 0
    score_col: Optional[int] = None  # 1-indexed column of this file's OWN "Score" column
    avg_rows: list = field(default_factory=list)  # this file's OWN average-row positions
                                                   # (derived from section headers, never
                                                   # from formula presence - see below)
    section_headers: list = field(default_factory=list)  # row positions of this file's
                                                           # OWN section headers
    section_header_text: list = field(default_factory=list)  # normalized header text,
                                                               # same order as above -
                                                               # the actual evidence used
                                                               # to compare templates
    kept_offsets: list = field(default_factory=list)  # 0-indexed offsets from column L
                                                       # that have real response data
    excluded_empty_count: int = 0  # candidate columns dropped for having no data at all


def _norm(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def analyze_file(path: str, filename: str) -> FileAnalysis:
    """Validate structure and detect patient columns / unit name for one uploaded file."""
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception as e:
        return FileAnalysis(path, filename, valid=False, error=f"Could not open file: {e}")

    if SHEET_CHECKLIST not in wb.sheetnames:
        return FileAnalysis(
            path, filename, valid=False,
            error=f"Sheet '{SHEET_CHECKLIST}' not found. Sheets present: {wb.sheetnames}",
        )
    ws = wb[SHEET_CHECKLIST]

    # Signature check: this is the anchor that confirms it's really this template.
    if not _norm(ws["A14"].value).upper().startswith("A. INITIAL ASSESSMENT"):
        return FileAnalysis(
            path, filename, valid=False,
            error="Row 14 does not read 'A. INITIAL ASSESSMENT' - this file doesn't "
                  "match the expected Nursing Chart Review Checklist structure.",
        )

    # Find this file's own "Score" column, scanning row 9 rightward from column L.
    score_col = None
    for c in range(PATIENT_START_COL, PATIENT_START_COL + 200):
        v = ws.cell(row=HEADER_ROW, column=c).value
        if isinstance(v, str) and v.strip().lower() == "score":
            score_col = c
            break
    if score_col is None:
        return FileAnalysis(
            path, filename, valid=False,
            error="Could not find a 'Score' column header in row 9 - cannot determine "
                  "where patient columns end.",
        )

    patient_count = score_col - PATIENT_START_COL
    if patient_count <= 0:
        return FileAnalysis(
            path, filename, valid=False,
            error="No patient columns found between column L and the Score column.",
        )

    # Detect this file's OWN section headers - the actual evidence its checklist
    # structure exists, never whether a section-average formula happens to still
    # work (real files regularly have that one cell overwritten by manual editing;
    # the header text is far more durable evidence of what checklist this is). A
    # file with 9 sections and a file with 12 are both fine on their own; whether
    # they're the SAME template as each other is checked later, once the upload
    # order (and so which file is the master) is known - see consolidate().
    section_headers = _find_section_headers(ws)
    if not section_headers:
        return FileAnalysis(
            path, filename, valid=False,
            error="No section headers found in this file's checklist (column A should "
                  "have rows like 'A. INITIAL ASSESSMENT') - can't determine its "
                  "checklist structure.",
        )
    boundaries = _compute_boundaries(ws, section_headers)
    avg_rows = [b[0] for b in boundaries]
    section_header_text = [_norm(ws.cell(row=h, column=1).value) for h in section_headers]

    # Data-quality note only, never a structural one (see _section_has_working_average) -
    # logged for anyone reviewing the source files, never blocks recognition.
    damaged_sections = [
        section_header_text[i] for i, avg_row in enumerate(avg_rows)
        if not _section_has_working_average(ws, score_col, avg_row)
    ]

    # Note: a source file's own 'SUMMSRY' sheet (if it has one at all) is never checked,
    # read, or required. Recognition depends only on the checklist sheet above. The final
    # Summary is always generated fresh by this tool from the consolidated checklist - see
    # module docstring.

    # A candidate patient column is only excluded if it has NO response data anywhere in
    # the actual checklist response area (the item rows - never based on the Unit/Ward
    # row, and never just because some individual cells are blank). Any real content at
    # all - a 0/1/2 score, "NA"/"N/A", any other text response - counts; formatting alone
    # never does, since only cell values are inspected here.
    item_rows = _item_rows_for(ws, boundaries)
    kept_offsets, excluded_offsets = [], []
    for i in range(patient_count):
        col = PATIENT_START_COL + i
        has_response = any(
            ws.cell(row=r, column=col).value is not None
            and str(ws.cell(row=r, column=col).value).strip() != ""
            for r in item_rows
        )
        (kept_offsets if has_response else excluded_offsets).append(i)

    if not kept_offsets:
        return FileAnalysis(
            path, filename, valid=False,
            error=f"All {patient_count} candidate patient column(s) in this file are "
                  f"empty - no response data found in the checklist area for any of them.",
        )

    # Detect unit/ward name(s) (row 12), over the KEPT columns only - an excluded empty
    # column's unit label (if any) doesn't end up in the output, so it shouldn't factor
    # into what unit this file is reported as. Usually one value repeated across every
    # patient column in a file, but a file can legitimately mix units (e.g. one
    # department's combined monthly submission covering two wards) - that's fine and not
    # an error; every kept patient column is still included regardless of its label.
    unit_values = [
        _norm(ws.cell(row=12, column=PATIENT_START_COL + i).value)
        for i in kept_offsets
    ]
    unique_units = sorted(set(u for u in unit_values if u))
    warnings = []
    if len(unique_units) > 1:
        unit_name = " + ".join(unique_units)
        warnings.append(
            f"Multiple different Unit/Ward values found within this file: {unique_units}. "
            f"All {len(kept_offsets)} patient columns with response data are still "
            f"included in the consolidated report."
        )
    else:
        unit_name = unique_units[0] if unique_units else None

    if excluded_offsets:
        warnings.append(
            f"{len(excluded_offsets)} of {patient_count} candidate patient column(s) had "
            f"no response data anywhere in the checklist and were excluded (not counted "
            f"as patients, no gap left in the numbering)."
        )

    if damaged_sections:
        warnings.append(
            f"{len(damaged_sections)} section(s) had a damaged/overwritten average "
            f"formula in this source file ({', '.join(damaged_sections)}) - a "
            f"data-quality issue in the source file only. The section itself is still "
            f"fully recognized from its header and items, and its average in the final "
            f"report is always calculated fresh, never copied from here."
        )

    return FileAnalysis(
        path, filename, valid=True, warnings=warnings,
        unit_name=unit_name, patient_count=len(kept_offsets), score_col=score_col,
        avg_rows=avg_rows, section_headers=section_headers,
        section_header_text=section_header_text,
        kept_offsets=kept_offsets,
        excluded_empty_count=len(excluded_offsets),
    )


def compare_to_master_text(master_ws, other_analysis: FileAnalysis, other_ws) -> list:
    """Light consistency check: flag if a file's own A/B checklist text differs from
    the file that will serve as the master (first uploaded). Non-blocking - just a warning,
    since only the FIRST file's A:K actually ends up in the output."""
    diffs = []
    for r in range(9, 109):
        a_m, b_m = master_ws.cell(row=r, column=1).value, master_ws.cell(row=r, column=2).value
        a_o, b_o = other_ws.cell(row=r, column=1).value, other_ws.cell(row=r, column=2).value
        if _norm(a_m) != _norm(a_o) or _norm(b_m) != _norm(b_o):
            diffs.append(r)
    if diffs:
        return [
            f"Checklist text differs from the first uploaded file at row(s) "
            f"{', '.join(str(r) for r in diffs[:6])}"
            + (f" and {len(diffs) - 6} more" if len(diffs) > 6 else "")
            + ". Only the first uploaded file's wording is used in the output."
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
    for r in range(1, last_row + 1):
        src_cell = src_ws.cell(row=r, column=src_col)
        dst_cell = _ensure_writable(dst_ws, r, dst_col)
        dst_cell.value = src_cell.value
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


def _find_section_headers(ws) -> list:
    """Find every section-header row: column A holds text matching the established
    'X. SECTION NAME' pattern (a single capital letter, a period, then the section
    name) - e.g. 'A. INITIAL ASSESSMENT', 'F.NURSING PROGRESS NOTES'. This is the
    actual evidence a section exists in this checklist - never whether a nearby
    average formula happens to still work, since in real files that cell gets
    overwritten by manual editing (NA, a hard-typed number, left blank) far more
    often than the header text itself changes. Scans generously (bounded by the
    sheet's own used range, capped only as a safety ceiling) so this works whether a
    checklist has 5 sections or 20."""
    last_row = min(ws.max_row, STRUCTURE_SCAN_LIMIT)
    rows = []
    for r in range(9, last_row + 1):
        v = ws.cell(row=r, column=1).value
        if isinstance(v, str) and _SECTION_HEADER_RE.match(v.strip()):
            rows.append(r)
    return rows


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
    summary_ws = wb.create_sheet(SHEET_SUMMARY, index=1)

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
            value=f"='{SHEET_CHECKLIST}'!{score_col_letter}{avg_row}",
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


def consolidate(ordered_paths_and_names: list, output_path: str) -> dict:
    """
    ordered_paths_and_names: list of (filepath, original_filename) tuples, in the exact
    order patient columns should appear (upload order).
    Returns a dict report of what was done (for display / audit).
    """
    if not ordered_paths_and_names:
        raise ValidationError("No files provided.")

    infos = []
    for path, name in ordered_paths_and_names:
        info = analyze_file(path, name)
        if not info.valid:
            raise ValidationError(f"{name}: {info.error}")
        infos.append(info)

    # --- Base workbook = the FIRST uploaded file, used as-is for its A:K master structure ---
    base_path = infos[0].path
    wb = openpyxl.load_workbook(base_path, data_only=False)
    ws = wb[SHEET_CHECKLIST]
    base_score_col = infos[0].score_col

    # Capture the base file's OWN dropdown/list-validation template (allowed values +
    # behavior) before anything is touched - see _capture_source_validation for why this
    # must never be a hardcoded assumption about where the allowed-values list lives.
    source_validation = _capture_source_validation(ws)

    # Actually resolve that source to real values - the same resolution Excel/Calc does
    # to populate the dropdown - rather than just trusting the DataValidation object
    # exists. Handles a direct literal list, a same-sheet range, or a defined name
    # pointing anywhere (including a hidden sheet). Logged either way so this is checked
    # on every single run, not just when someone happens to open the file and click a
    # cell by hand.
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

    # A source file's own SUMMSRY sheet, if present, is never used - see module docstring.
    # (The final Summary is built fresh at the end, from the consolidated checklist only.)

    # Derive the base file's section structure from its actual header/item CONTENT -
    # never from whether a section-average formula happens to still work there (see
    # module docstring). However many sections that turns out to be is fine: a
    # 9-section checklist and a 12-section checklist are both valid on their own.
    section_headers = _find_section_headers(ws)
    if not section_headers:
        # analyze_file() already guarantees this for every file, including this one, on
        # its own earlier read - this is just a defensive re-check against the fresh copy
        # of the file loaded here.
        raise ValidationError(
            f"{infos[0].filename}: could not detect any section headers in the master "
            f"file's checklist - cannot determine its structure."
        )
    boundaries = _compute_boundaries(ws, section_headers)
    avg_rows = [b[0] for b in boundaries]
    avg_row_set = set(avg_rows)
    section_header_text = [_norm(ws.cell(row=h, column=1).value) for h in section_headers]

    # Every OTHER file must use the SAME checklist template as the master (first
    # uploaded) file - patient columns are copied by absolute row position, so if a
    # file's own section headers sit at different rows, or say something different,
    # its item rows don't line up with the master's either, and blindly copying its
    # patient values would silently attach them to the wrong checklist items. This
    # compares actual section-header row positions and text - NEVER whether a nearby
    # average formula happens to still work, since that cell is routinely overwritten
    # by manual editing in real files without the checklist itself being any
    # different (confirmed directly: P2CD and PICU have byte-identical headers and
    # item text at every one of their 12 sections, despite 1-3 average cells each
    # having been clobbered with "NA" or a hard-typed number - they are the same
    # template and are correctly treated as compatible here).
    for info in infos[1:]:
        if info.section_headers != section_headers or info.section_header_text != section_header_text:
            raise ValidationError(
                f"{info.filename}: this file's checklist sections are at rows "
                f"{info.section_headers} ({info.section_header_text}), but the master "
                f"file ({infos[0].filename}, uploaded first) has sections at rows "
                f"{section_headers} ({section_header_text}). These appear to be "
                f"different checklist templates and can't be safely combined - patient "
                f"data would misalign to the wrong checklist items. Remove this file, "
                f"or reorder so a file matching this structure is uploaded first."
            )

    # How far into the patient columns to copy: past the last detected section-average
    # row, plus a little padding for any trailing blank/signature rows below the last
    # section - not a fixed row ceiling, so a shorter or taller checklist both work.
    copy_last_row = max(avg_rows) + TRAILING_ROW_PADDING

    # --- Capture the base file's own Score-column style (font/border/fill/etc per row)
    #     and width, BEFORE that column gets overwritten by the next file's data. This is
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

    # Fallback width for any patient column whose source file never set one explicitly
    # (see _copy_patient_column docstring) - use the master file's own first patient
    # column (L) as the sane default.
    l_dim = ws.column_dimensions.get("L")
    fallback_patient_width = l_dim.width if (l_dim and l_dim.width) else 10.0

    # --- Find and dissolve every merge in the patient/Score-column region (>= column L)
    #     before any writes happen there - see _sweep_merges_in_region for why this must
    #     be a dynamic sweep rather than a fixed list of expected merge patterns. Columns
    #     A:K (and every other sheet) are never touched. ---
    wide_spans, score_col_spans, other_merge_count = _sweep_merges_in_region(
        ws, PATIENT_START_COL, base_score_col
    )

    # --- Place every file's patient columns, in upload order ---
    report_files = []
    current_col = PATIENT_START_COL
    running_index = 0
    for i, info in enumerate(infos):
        if i > 0:
            src_wb = openpyxl.load_workbook(info.path, data_only=False)
            src_ws = src_wb[SHEET_CHECKLIST]
            # This file is only ever READ from below, never saved - but if IT has its own
            # unusual merge in the patient-column area (e.g. a Unit/Ward label merged
            # across a few of its own patient columns), reading a non-anchor merged
            # cell's .value returns None, silently losing that patient's data on copy.
            # Same fix as the base file gets, applied here too: dissolve any such merge
            # and propagate the anchor's value first, so every cell we're about to copy
            # from is a real, independently-readable value.
            _, _, src_other_merge_count = _sweep_merges_in_region(src_ws, PATIENT_START_COL, info.score_col)
            warnings = list(info.warnings) + compare_to_master_text(ws, info, src_ws)
            if src_other_merge_count:
                warnings.append(
                    f"{src_other_merge_count} unexpected merged cell(s) found in this "
                    f"file's own patient columns - resolved (value copied to every cell "
                    f"the merge covered) before reading its patient data."
                )
        else:
            src_ws = ws  # first file: data's already in place, nothing to copy
            warnings = list(info.warnings)

        start_col = current_col
        for offset in info.kept_offsets:
            src_col = PATIENT_START_COL + offset
            dst_col = current_col
            running_index += 1
            if not (i == 0 and dst_col == src_col):
                _copy_patient_column(src_ws, src_col, ws, dst_col, copy_last_row)
            _set_column_width(src_ws, src_col, ws, dst_col, fallback_patient_width)
            _ensure_writable(ws, HEADER_ROW, dst_col)
            ws.cell(row=HEADER_ROW, column=dst_col).value = running_index
            current_col += 1

        report_files.append({
            "filename": info.filename, "unit": info.unit_name,
            "patient_count": info.patient_count,
            "excluded_empty_count": info.excluded_empty_count,
            "columns": f"{get_column_letter(start_col)}:{get_column_letter(current_col - 1)}",
            "warnings": warnings,
        })

    final_score_col = current_col
    final_score_letter = get_column_letter(final_score_col)
    last_patient_letter = get_column_letter(final_score_col - 1)
    total_patients = running_index

    # --- Write the new consolidated Score column ---
    for r in range(1, copy_last_row + 1):
        style = score_style_by_row[r]
        cell = _ensure_writable(ws, r, final_score_col)
        cell.font = style["font"]
        cell.border = style["border"]
        cell.fill = style["fill"]
        cell.alignment = style["alignment"]
        cell.number_format = style["number_format"]
        cell.protection = style["protection"]

    ws.cell(row=HEADER_ROW, column=final_score_col).value = "Score"

    # `boundaries` was already derived from the base file's section headers/item content
    # right after it was loaded (columns A/B, which nothing in this function ever
    # touches) - reused as-is here rather than recomputed, since recomputing against
    # `avg_rows` would be interpreting average-row positions as if they were header
    # positions, which they aren't.

    # Decide, per item row, whether a per-item formula is safe to write - based on the
    # ACTUAL consolidated patient data we just placed, not on any single source file's
    # own (possibly smaller / all-"NA") patient set. A row where every uploaded patient
    # is "NA" would divide by zero, so it's deliberately left blank, exactly matching
    # the master template's own convention (confirmed: rows that are all-"NA" even
    # across 70 patients in the reference report are left blank there too).
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

    # --- Rebuild data validation (the 0/1/2/NA dropdown) across the new patient column
    #     width, using the base file's OWN validation template captured at the very start
    #     (source_validation) - never a hardcoded reference. If the base file didn't have
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

    # --- Build a brand-new SUMMSRY sheet from the consolidated checklist only.
    #     Never reads or copies any source workbook's own SUMMSRY sheet (some don't have
    #     one at all - that's fine, this doesn't depend on it either way). ---
    _build_summary_sheet(wb, ws, boundaries, final_score_letter)

    wb.save(output_path)

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
        "section_average_rows": avg_rows,
        "blank_item_rows": blank_item_rows,
        "formula_item_rows_count": len(formula_item_rows),
        "dropdown_source": source_validation["formula1"] if source_validation else None,
        "dropdown_values": dropdown_values,
        "dropdown_verified": bool(dropdown_values) and not dropdown_resolution_note,
        "dissolved_unexpected_merges": other_merge_count,
    }

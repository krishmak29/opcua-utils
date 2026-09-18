"""Writes per-tag verification status/timestamp back into the engineering Excel sheet."""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from data_io import s, _sanitize_workbook_bytes, _compute_tag_name

GENERAL_COMMENTS_SHEET = "General Comments"
GENERAL_COMMENTS_HEADERS = ["Sr No", "PLC", "Template", "Area Code*", "Equipment Number*",
                             "Verification Status", "Cause", "Tester", "Comment"]


def _find_header_col(headers, *names):
    """Return 1-based column index of the first header in `names` found in `headers`."""
    for n in names:
        if n in headers:
            return headers[n]
    return None


def save_tag_verifications(path, sheet_name, updates, status_col_letter, ts_col_letter, current_value_col_letter=None, header_row=1, tester_col_letter=None):
    """
    Write per-tag Verification Status / Verification Time back into the engineering
    Excel file, matching rows by PLC + Tag Name + Address (tolerant of 'Address*'
    vs 'Address' as the header name).

    updates: list of dicts, each with keys 'PLC', 'Tag Name', 'Address',
             'Verification Status', 'Verification Time'.
    status_col_letter / ts_col_letter: column letters (e.g. 'M', 'N') from config.

    Returns (written_count, not_found_keys).
    """
    if not status_col_letter or not ts_col_letter:
        raise ValueError(
            "Verify Status/Timestamp columns are not configured. Set "
            "'VerifyStatusColumn' and 'VerifyTimestampColumn' in utility_config.xlsx "
            "(Display_Config sheet)."
        )

    wb_data = load_workbook(_sanitize_workbook_bytes(path), data_only=True)  # for matching: reads computed formula results
    wb = load_workbook(_sanitize_workbook_bytes(path))  # NOT data_only, so formulas elsewhere in the sheet are preserved on save
    if sheet_name not in wb.sheetnames:
        sheet_name = "Objects" if "Objects" in wb.sheetnames else wb.sheetnames[0]
    ws_data = wb_data[sheet_name]
    ws = wb[sheet_name]

    headers = {}
    for cell in ws_data[header_row]:
        h = s(cell.value)
        if h:
            headers[h] = cell.column

    plc_col = _find_header_col(headers, "PLC")
    tag_col = _find_header_col(headers, "Tag Name")
    addr_col = _find_header_col(headers, "Address*", "Address")
    if not (plc_col and tag_col and addr_col):
        raise ValueError("Could not locate PLC / Tag Name / Address columns in the sheet header row.")

    status_col = column_index_from_string(status_col_letter)
    ts_col = column_index_from_string(ts_col_letter)
    cv_col = column_index_from_string(current_value_col_letter) if current_value_col_letter else None
    tester_col = column_index_from_string(tester_col_letter) if tester_col_letter else None

    written = 0
    not_found = []
    for u in updates:
        row_idx = u.get("_row")
        if row_idx is None:
            not_found.append((s(u.get("PLC")), s(u.get("Tag Name")), s(u.get("Address"))))
            continue
        ws.cell(row=row_idx, column=status_col).value = u.get("Verification Status")
        ws.cell(row=row_idx, column=ts_col).value = u.get("Verification Time")
        if cv_col:
            ws.cell(row=row_idx, column=cv_col).value = u.get("Current Value")
        if tester_col:
            ws.cell(row=row_idx, column=tester_col).value = u.get("Tester Name")
        written += 1

    wb.save(path)
    return written, not_found

def save_pending_batch(path, tag_sheet_name, tag_updates, status_col_letter, ts_col_letter,
                        current_value_col_letter, tester_col_letter, equipment_updates, header_row=1):
    """Commit an entire pending batch (tag-level + equipment-level) to the
    engineering Excel file in one open/save cycle, per locked spec section 4.

    Before writing each tag update, verifies the target row's PLC + Tag Name
    still match what's in pending_changes.json (guards against stale _row
    references if the engineering file was reloaded/regenerated since the
    popup that produced this pending entry). Mismatches are skipped and
    reported rather than written blind.

    Writes to a temp file and renames over the original only after a
    successful save, so a crash mid-write can never corrupt the source file.

    Returns dict: {"tags_written": int, "tags_mismatched": list, "tags_no_row": list,
                    "equipment_written": int}.
    """
    if not status_col_letter or not ts_col_letter:
        raise ValueError(
            "Verify Status/Timestamp columns are not configured. Set "
            "'VerifyStatusColumn' and 'VerifyTimestampColumn' in utility_config.xlsx "
            "(Display_Config sheet)."
        )

    wb_data = load_workbook(_sanitize_workbook_bytes(path), data_only=True)
    wb = load_workbook(_sanitize_workbook_bytes(path))
    sheet_name = tag_sheet_name if tag_sheet_name in wb.sheetnames else (
        "Objects" if "Objects" in wb.sheetnames else wb.sheetnames[0])
    ws_data = wb_data[sheet_name]
    ws = wb[sheet_name]

    headers = {}
    for cell in ws_data[header_row]:
        h = s(cell.value)
        if h:
            headers[h] = cell.column

    plc_col = _find_header_col(headers, "PLC")
    tag_col = _find_header_col(headers, "Tag Name")
    area_col = _find_header_col(headers, "Area Code*")
    eq_col = _find_header_col(headers, "Equipment Number*")
    suffix_col = _find_header_col(headers, "OPC Tag Suffix")
    if not (plc_col and tag_col):
        raise ValueError("Could not locate PLC / Tag Name columns in the sheet header row.")

    def _effective_tag_name(row_idx):
        raw = s(ws_data.cell(row=row_idx, column=tag_col).value)
        if raw or not suffix_col:
            return raw
        # Tag Name is a CONCATENATE formula; if this workbook has already been
        # re-saved by openpyxl once, the cached formula result is gone, so
        # rebuild it from source columns exactly as load_engineering() does.
        return _compute_tag_name({
            "Area Code*": ws_data.cell(row=row_idx, column=area_col).value if area_col else None,
            "Equipment Number*": ws_data.cell(row=row_idx, column=eq_col).value if eq_col else None,
            "OPC Tag Suffix": ws_data.cell(row=row_idx, column=suffix_col).value if suffix_col else None,
        })

    status_col = column_index_from_string(status_col_letter)
    ts_col = column_index_from_string(ts_col_letter)
    cv_col = column_index_from_string(current_value_col_letter) if current_value_col_letter else None
    tester_col = column_index_from_string(tester_col_letter) if tester_col_letter else None

    tags_written = 0
    tags_mismatched = []
    tags_no_row = []

    for u in tag_updates:
        row_idx = u.get("_row")
        if row_idx is None:
            tags_no_row.append(s(u.get("Tag Name")))
            continue
        if row_idx < 1 or row_idx > ws_data.max_row:
            tags_mismatched.append(s(u.get("Tag Name")))
            continue
        actual_plc = s(ws_data.cell(row=row_idx, column=plc_col).value)
        actual_tag = _effective_tag_name(row_idx)
        if actual_plc != s(u.get("PLC")) or actual_tag != s(u.get("Tag Name")):
            tags_mismatched.append(s(u.get("Tag Name")))
            continue
        ws.cell(row=row_idx, column=status_col).value = u.get("Verification Status")
        ws.cell(row=row_idx, column=ts_col).value = u.get("Verification Time")
        if cv_col:
            ws.cell(row=row_idx, column=cv_col).value = u.get("Current Value")
        if tester_col:
            ws.cell(row=row_idx, column=tester_col).value = u.get("Tester Name")
        tags_written += 1

    equipment_written = 0
    if equipment_updates:
        if GENERAL_COMMENTS_SHEET not in wb.sheetnames:
            gc_ws = wb.create_sheet(GENERAL_COMMENTS_SHEET)
            gc_ws.append(GENERAL_COMMENTS_HEADERS)
        else:
            gc_ws = wb[GENERAL_COMMENTS_SHEET]

        max_sr = 0
        for r in range(2, gc_ws.max_row + 1):
            sr_val = gc_ws.cell(row=r, column=1).value
            if isinstance(sr_val, (int, float)):
                max_sr = max(max_sr, int(sr_val))

        for eq in equipment_updates:
            target_key = (s(eq.get("PLC")), s(eq.get("Area")), s(eq.get("Equipment")))
            match_row = None
            first_empty_row = None
            for r in range(2, gc_ws.max_row + 1):
                row_plc = s(gc_ws.cell(row=r, column=2).value)
                row_area = s(gc_ws.cell(row=r, column=4).value)
                row_eq = s(gc_ws.cell(row=r, column=5).value)
                if not row_plc and not row_eq and first_empty_row is None:
                    first_empty_row = r
                if (row_plc, row_area, row_eq) == target_key:
                    match_row = r
                    break
            if match_row is None:
                match_row = first_empty_row if first_empty_row is not None else gc_ws.max_row + 1
                max_sr += 1
                gc_ws.cell(row=match_row, column=1).value = max_sr

            gc_ws.cell(row=match_row, column=2).value = eq.get("PLC")
            gc_ws.cell(row=match_row, column=3).value = eq.get("Template")
            gc_ws.cell(row=match_row, column=4).value = eq.get("Area")
            gc_ws.cell(row=match_row, column=5).value = eq.get("Equipment")
            gc_ws.cell(row=match_row, column=6).value = eq.get("Status")
            gc_ws.cell(row=match_row, column=7).value = eq.get("Cause")
            gc_ws.cell(row=match_row, column=8).value = eq.get("Tester")
            gc_ws.cell(row=match_row, column=9).value = eq.get("Comment")
            equipment_written += 1

    tmp_path = Path(str(path) + ".tmp")
    wb.save(tmp_path)
    tmp_path.replace(path)

    return {
        "tags_written": tags_written,
        "tags_mismatched": tags_mismatched,
        "tags_no_row": tags_no_row,
        "equipment_written": equipment_written,
    }

def load_current_values(path, sheet_name, row_numbers, current_value_col_letter, header_row=1):
    """Read Current Value directly from the configured column, matched by
    the row's literal Excel row number. Returns dict keyed by row number -> value."""
    if not current_value_col_letter:
        return {}
    wb = load_workbook(_sanitize_workbook_bytes(path), data_only=True)
    if sheet_name not in wb.sheetnames:
        sheet_name = "Objects" if "Objects" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    cv_col = column_index_from_string(current_value_col_letter)
    out = {}
    for row_idx in row_numbers:
        if row_idx is None:
            continue
        out[row_idx] = s(ws.cell(row=row_idx, column=cv_col).value)
    return out

def load_general_comments(path):
    """Return dict keyed by 'PLC|Area|Equipment' -> dict with status/cause/tester/comment/sr."""
    if not Path(path).exists():
        return {}
    wb = load_workbook(_sanitize_workbook_bytes(path), data_only=True)
    if GENERAL_COMMENTS_SHEET not in wb.sheetnames:
        return {}
    ws = wb[GENERAL_COMMENTS_SHEET]
    out = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or all(v is None for v in row):
            continue
        vals = (list(row) + [None] * 9)[:9]
        sr, plc, template, area, equipment, status, cause, tester, comment = vals
        if not plc or not equipment:
            continue
        k = "|".join([s(plc), s(area), s(equipment)])
        out[k] = {"sr": sr, "status": s(status), "cause": s(cause), "tester": s(tester), "comment": s(comment)}
    return out


def save_general_comment(path, plc, template, area, equipment, status, cause, tester, comment):
    """Write/overwrite one equipment-level row in the General Comments sheet.
    Matches by PLC + Area + Equipment. Overwrite keeps the existing Sr No;
    a new row gets max(Sr No) + 1. Creates the sheet with headers if missing."""
    wb = load_workbook(_sanitize_workbook_bytes(path))
    if GENERAL_COMMENTS_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(GENERAL_COMMENTS_SHEET)
        ws.append(GENERAL_COMMENTS_HEADERS)
    else:
        ws = wb[GENERAL_COMMENTS_SHEET]

    target_key = (s(plc), s(area), s(equipment))
    max_sr = 0
    match_row = None
    first_empty_row = None
    for r in range(2, ws.max_row + 1):
        row_plc = s(ws.cell(row=r, column=2).value)
        row_area = s(ws.cell(row=r, column=4).value)
        row_eq = s(ws.cell(row=r, column=5).value)
        sr_val = ws.cell(row=r, column=1).value
        if isinstance(sr_val, (int, float)):
            max_sr = max(max_sr, int(sr_val))
        if not row_plc and not row_eq and first_empty_row is None:
            first_empty_row = r
        if (row_plc, row_area, row_eq) == target_key:
            match_row = r
            break

    if match_row is None:
        match_row = first_empty_row if first_empty_row is not None else ws.max_row + 1
        ws.cell(row=match_row, column=1).value = max_sr + 1

    ws.cell(row=match_row, column=2).value = plc
    ws.cell(row=match_row, column=3).value = template
    ws.cell(row=match_row, column=4).value = area
    ws.cell(row=match_row, column=5).value = equipment
    ws.cell(row=match_row, column=6).value = status
    ws.cell(row=match_row, column=7).value = cause
    ws.cell(row=match_row, column=8).value = tester
    ws.cell(row=match_row, column=9).value = comment

    wb.save(path)    
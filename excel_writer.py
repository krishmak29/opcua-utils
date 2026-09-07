"""Writes per-tag verification status/timestamp back into the engineering Excel sheet."""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from data_io import s, _sanitize_workbook_bytes

GENERAL_COMMENTS_SHEET = "General Comments"
GENERAL_COMMENTS_HEADERS = ["Sr No", "PLC", "Template", "Area Code*", "Equipment Number*",
                             "Verification Status", "Cause", "Tester", "Comment"]


def _find_header_col(headers, *names):
    """Return 1-based column index of the first header in `names` found in `headers`."""
    for n in names:
        if n in headers:
            return headers[n]
    return None


def save_tag_verifications(path, sheet_name, updates, status_col_letter, ts_col_letter, current_value_col_letter=None, header_row=1):
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
        written += 1

    wb.save(path)
    return written, not_found

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
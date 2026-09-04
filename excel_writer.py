"""Writes per-tag verification status/timestamp back into the engineering Excel sheet."""

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from data_io import s


def _find_header_col(headers, *names):
    """Return 1-based column index of the first header in `names` found in `headers`."""
    for n in names:
        if n in headers:
            return headers[n]
    return None


def save_tag_verifications(path, sheet_name, updates, status_col_letter, ts_col_letter):
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

    wb = load_workbook(path)  # NOT data_only, so formulas elsewhere in the sheet are preserved
    if sheet_name not in wb.sheetnames:
        sheet_name = "Objects" if "Objects" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]

    headers = {}
    for cell in ws[1]:
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

    pending = {(s(u.get("PLC")), s(u.get("Tag Name")), s(u.get("Address"))): u for u in updates}
    written = 0
    for row_idx in range(2, ws.max_row + 1):
        plc_v = s(ws.cell(row=row_idx, column=plc_col).value)
        tag_v = s(ws.cell(row=row_idx, column=tag_col).value)
        addr_v = s(ws.cell(row=row_idx, column=addr_col).value)
        k = (plc_v, tag_v, addr_v)
        u = pending.pop(k, None)
        if u is None:
            continue
        ws.cell(row=row_idx, column=status_col).value = u.get("Verification Status")
        ws.cell(row=row_idx, column=ts_col).value = u.get("Verification Time")
        written += 1

    wb.save(path)
    return written, list(pending.keys())
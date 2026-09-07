"""Loading engineering/config Excel data and grouping rows into objects."""

from pathlib import Path
import zipfile
import re
import io
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

def s(v):
    return "" if v is None else str(v).strip()


def _sanitize_workbook_bytes(path):
    """Some Excel files end up with an out-of-range font 'family' value in
    styles.xml (valid range 0-14), which makes openpyxl refuse to load the
    file entirely ('could not read stylesheet... invalid XML'). This clamps
    any such value to a safe default (2) in-memory before openpyxl parses it,
    so a corrupted style attribute never blocks loading."""
    with zipfile.ZipFile(path, "r") as zin:
        if "xl/styles.xml" not in zin.namelist():
            return path  # nothing to sanitize; let openpyxl handle it normally
        styles = zin.read("xl/styles.xml").decode("utf-8")

    fixed = re.sub(r'<family val="(\d+)"/>',
                    lambda m: '<family val="2"/>' if int(m.group(1)) > 14 else m.group(0),
                    styles)
    if fixed == styles:
        return path  # nothing needed fixing

    buf = io.BytesIO()
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = fixed.encode("utf-8") if item.filename == "xl/styles.xml" else zin.read(item.filename)
            zout.writestr(item, data)
    buf.seek(0)
    return buf


def rows(wb, sheet, header_row=1):
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    h = [s(c.value) for c in ws[header_row]]
    out = []
    for i, r in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True)):
        if not any(x is not None for x in r):
            continue
        d = {h[j]: r[j] if j < len(r) else None for j in range(len(h)) if h[j]}
        d["_row"] = header_row + 1 + i
        out.append(d)
    return out


def get_address(r):
    """Return the OPC UA NodeID used for live reads/writes.
    Sourced from the 'OPC Address' column (NodeID), not 'Address*'/'Address'
    which holds the symbolic address."""
    return s(r.get("OPC Address"))


ENGINEERING_HEADER_ROW = 10
ENGINEERING_SHEET = "OPC Tag Templates"


def _compute_opc_address(r, b5_val):
    """Rebuild OPC Address from its source columns (KEPWare_CHName +
    KEPWare_DVCName + Tag Name, prefixed with the fixed $B$5 namespace value),
    mirroring the sheet's own formula. Used as a fallback when the formula's
    cached value isn't available (e.g. after this app re-saves the workbook,
    which strips cached formula results)."""
    ch = s(r.get("KEPWare_CHName"))
    dvc = s(r.get("KEPWare_DVCName"))
    tag = s(r.get("Tag Name"))
    return f"ns=2;s={b5_val}{ch}.{dvc}.{tag}"


def _compute_tag_name(r):
    """Rebuild Tag Name from its source columns (Area Code* + Equipment Number* +
    OPC Tag Suffix), mirroring the sheet's own CONCATENATE formula. Used as a
    fallback when the formula's cached value isn't available (e.g. file wasn't
    re-saved in Excel, so openpyxl's data_only read returns None)."""
    area = s(r.get("Area Code*"))
    eq = s(r.get("Equipment Number*"))
    suffix = s(r.get("OPC Tag Suffix"))
    parts = []
    if area:
        parts.append(area + ".")
    if eq:
        parts.append(eq + ".")
    parts.append(suffix)
    return "".join(parts)


def load_engineering(path, status_col_letter=None, ts_col_letter=None, cv_col_letter=None, tester_col_letter=None):
    wb = load_workbook(_sanitize_workbook_bytes(path), data_only=True)
    if ENGINEERING_SHEET in wb.sheetnames:
        sheet = ENGINEERING_SHEET
    elif "Objects" in wb.sheetnames:
        sheet = "Objects"
    else:
        sheet = wb.sheetnames[0]
    ws = wb[sheet]
    data = rows(wb, sheet, header_row=ENGINEERING_HEADER_ROW)
    required = ["PLC", "Area Code*", "Equipment Number*", "Template", "Tag Name", "Data Type"]
    if not data:
        raise ValueError("No engineering data found.")
    missing = [x for x in required if x not in data[0]]
    if "OPC Address" not in data[0]:
        missing.append("OPC Address")
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))
    if "OPC Tag Suffix" in data[0]:
        for r in data:
            if not s(r.get("Tag Name")):
                r["Tag Name"] = _compute_tag_name(r)

    b5_val = s(ws.cell(row=5, column=2).value)
    for r in data:
        if not s(r.get("OPC Address")):
            r["OPC Address"] = _compute_opc_address(r, b5_val)

    # Verification Status / Verification Time / Current Value are addressed by
    # configured column letter (utility_config.xlsx), not by header text, since
    # the engineering sheet's own header row is often blank at those columns.
    status_col = column_index_from_string(status_col_letter) if status_col_letter else None
    ts_col = column_index_from_string(ts_col_letter) if ts_col_letter else None
    cv_col = column_index_from_string(cv_col_letter) if cv_col_letter else None
    tester_col = column_index_from_string(tester_col_letter) if tester_col_letter else None
    for r in data:
        row_idx = r.get("_row")
        if row_idx is None:
            continue
        if status_col:
            r["Verification Status"] = s(ws.cell(row=row_idx, column=status_col).value)
        if ts_col:
            r["Verification Time"] = s(ws.cell(row=row_idx, column=ts_col).value)
        if cv_col:
            r["Current Value"] = s(ws.cell(row=row_idx, column=cv_col).value)
        if tester_col:
            r["Tester Name"] = s(ws.cell(row=row_idx, column=tester_col).value)
    return data


def load_config(path):
    if not Path(path).exists():
        return {"plc": {}, "display": {"ObjectsPerPage": 12, "RefreshRateMs": 1000, "MainValueSuffix": "RD.PV",
                                        "VerifyStatusColumn": "", "VerifyTimestampColumn": "", "CurrentValueColumn": "", "TesterColumn": ""},
                "causes": [("C001", "Wrong PLC value"), ("C002", "Wrong SCADA value"), ("C003", "Communication failure"),
                           ("C004", "Wrong scaling"), ("C005", "Wrong engineering unit"), ("C006", "Wrong alarm status"),
                           ("C007", "Wrong tag mapping"), ("C008", "Object not available"), ("C009", "PLC not available"), ("C010", "Other")]}
    wb = load_workbook(_sanitize_workbook_bytes(path), data_only=True)
    pc = {s(r.get("PLC")): r for r in rows(wb, "PLC_Config") if s(r.get("PLC"))}
    d = {"ObjectsPerPage": 12, "RefreshRateMs": 1000, "MainValueSuffix": "RD.PV",
         "VerifyStatusColumn": "", "VerifyTimestampColumn": "", "CurrentValueColumn": "", "TesterColumn": ""}
    for r in rows(wb, "Display_Config"):
        if s(r.get("Parameter")):
            d[s(r["Parameter"])] = r.get("Value")
    c = [(s(r.get("Cause Code")), s(r.get("Cause Description"))) for r in rows(wb, "Cause_Master")
         if s(r.get("Cause Code")) and s(r.get("Cause Description"))]
    return {"plc": pc, "display": d, "causes": c or load_config.__defaults__}


def key(o):
    return "|".join([o["PLC"], o["Area"], o["Equipment"]])


def group(data):
    out = {}
    for r in data:
        plc = s(r.get("PLC"))
        area = s(r.get("Area Code*"))
        eq = s(r.get("Equipment Number*"))
        if not plc or not eq:
            continue
        k = (plc, area, eq)
        out.setdefault(k, {"PLC": plc, "Area": area, "Equipment": eq, "Template": s(r.get("Template")), "Rows": []})["Rows"].append(r)
    return list(out.values())

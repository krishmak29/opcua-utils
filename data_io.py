"""Loading engineering/config Excel data and grouping rows into objects."""

from pathlib import Path
from openpyxl import load_workbook


def s(v):
    return "" if v is None else str(v).strip()


def rows(wb, sheet):
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    h = [s(c.value) for c in ws[1]]
    return [{h[i]: r[i] if i < len(r) else None for i in range(len(h)) if h[i]}
            for r in ws.iter_rows(min_row=2, values_only=True) if any(x is not None for x in r)]


def get_address(r):
    """Return the address value for a row, tolerant of header being
    'Address*' or 'Address'."""
    return s(r.get("Address*")) or s(r.get("Address"))


def load_engineering(path):
    wb = load_workbook(path, data_only=True)
    sheet = "Objects" if "Objects" in wb.sheetnames else wb.sheetnames[0]
    data = rows(wb, sheet)
    required = ["PLC", "Area Code*", "Equipment Number*", "Template", "Tag Name", "Data Type"]
    if not data:
        raise ValueError("No engineering data found.")
    missing = [x for x in required if x not in data[0]]
    if "Address*" not in data[0] and "Address" not in data[0]:
        missing.append("Address (or Address*)")
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))
    return data


def load_config(path):
    if not Path(path).exists():
        return {"plc": {}, "display": {"ObjectsPerPage": 12, "RefreshRateMs": 1000, "MainValueSuffix": "RD.PV",
                                        "VerifyStatusColumn": "", "VerifyTimestampColumn": ""},
                "causes": [("C001", "Wrong PLC value"), ("C002", "Wrong SCADA value"), ("C003", "Communication failure"),
                           ("C004", "Wrong scaling"), ("C005", "Wrong engineering unit"), ("C006", "Wrong alarm status"),
                           ("C007", "Wrong tag mapping"), ("C008", "Object not available"), ("C009", "PLC not available"), ("C010", "Other")]}
    wb = load_workbook(path, data_only=True)
    pc = {s(r.get("PLC")): r for r in rows(wb, "PLC_Config") if s(r.get("PLC"))}
    d = {"ObjectsPerPage": 12, "RefreshRateMs": 1000, "MainValueSuffix": "RD.PV",
         "VerifyStatusColumn": "", "VerifyTimestampColumn": ""}
    for r in rows(wb, "Display_Config"):
        if s(r.get("Parameter")):
            d[s(r["Parameter"])] = r.get("Value")
    c = [(s(r.get("Cause Code")), s(r.get("Cause Description"))) for r in rows(wb, "Cause_Master")
         if s(r.get("Cause Code")) and s(r.get("Cause Description"))]
    return {"plc": pc, "display": d, "causes": c or load_config.__defaults__}


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

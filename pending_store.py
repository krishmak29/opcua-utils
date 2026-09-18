"""Pending changes JSON store.

Popup Save writes here (fast, no Excel I/O). The main screen's "Save to
Excel" action later reads everything accumulated here, writes Excel once,
and clears this file only after that write succeeds.
"""

import json
import threading
from pathlib import Path

from config import PENDING_CHANGES

_lock = threading.Lock()


def _empty():
    return {"tags": {}, "equipment": {}}


def load_pending():
    """Return the pending-changes dict. Tolerant of a missing or corrupt
    file: either case returns a fresh empty structure rather than raising,
    since a broken pending file should never block the app from opening."""
    p = Path(PENDING_CHANGES)
    if not p.exists():
        return _empty()
    try:
        with open(p, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return _empty()
    if not isinstance(data, dict):
        return _empty()
    data.setdefault("tags", {})
    data.setdefault("equipment", {})
    return data


def save_pending(data):
    """Write the pending dict to disk via temp-file + atomic rename, so a
    crash or power loss mid-write can never leave pending_changes.json
    half-written/corrupt."""
    p = Path(PENDING_CHANGES)
    tmp = p.with_suffix(p.suffix + ".tmp")
    with _lock:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        tmp.replace(p)


def _tag_key(update):
    row = update.get("_row")
    if row is not None:
        return str(row)
    # Fallback only for a tag with no source Excel row reference (shouldn't
    # normally happen) -- avoids silently dropping the update.
    return "|".join([str(update.get("PLC", "")), str(update.get("Tag Name", "")), str(update.get("Address", ""))])


def merge_tag_update(update):
    """Merge one tag-level update into pending storage. Last write for a
    given source row wins -- each call is expected to carry that row's full
    current verification field set, not a partial diff."""
    data = load_pending()
    data["tags"][_tag_key(update)] = update
    save_pending(data)


def _equipment_key(plc, area, equipment):
    return "|".join([str(plc), str(area), str(equipment)])


def merge_equipment_update(plc, template, area, equipment, status, cause, tester, comment):
    data = load_pending()
    k = _equipment_key(plc, area, equipment)
    data["equipment"][k] = {
        "PLC": plc, "Template": template, "Area": area, "Equipment": equipment,
        "Status": status, "Cause": cause, "Tester": tester, "Comment": comment,
    }
    save_pending(data)


def get_pending_tag(row_idx):
    """Pending update for a given Excel row number, or None."""
    if row_idx is None:
        return None
    return load_pending()["tags"].get(str(row_idx))


def get_pending_equipment(plc, area, equipment):
    return load_pending()["equipment"].get(_equipment_key(plc, area, equipment))


def has_pending():
    data = load_pending()
    return bool(data.get("tags")) or bool(data.get("equipment"))


def clear_pending():
    save_pending(_empty())
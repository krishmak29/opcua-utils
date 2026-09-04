"""The per-equipment verification detail window, split out of App.open()."""

import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from data_io import s, get_address
from db import key
from excel_writer import save_tag_verifications


def edit_verification_status(app, event, tr, mapping, dirty_flag):
    region = tr.identify("region", event.x, event.y)
    if region != "cell":
        return
    row_id = tr.identify_row(event.y)
    column_id = tr.identify_column(event.x)
    # Verification Status = column 7
    if not row_id or column_id != "#7":
        return
    bbox = tr.bbox(row_id, column_id)
    if not bbox:
        return
    x, y, width, height = bbox
    current_value = tr.set(row_id, column_id)
    combo = ttk.Combobox(
        tr,
        values=("Not Tested", "OK", "Recheck", "Not Available", "SCADA-Linking issue", "SCADA-Value Mismatch", "Not Ok"),
        state="readonly"
    )
    combo.set(current_value or "Not Tested")
    combo.place(x=x, y=y, width=width, height=height)
    combo.focus_set()

    def save_status(event=None):
        new_value = combo.get()
        verification_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
        tr.set(row_id, "#7", new_value)
        tr.set(row_id, "#8", verification_time)
        if row_id in mapping:
            mapping[row_id]["Verification Status"] = new_value
            mapping[row_id]["Verification Time"] = verification_time
        dirty_flag["value"] = True
        combo.destroy()

    combo.bind("<<ComboboxSelected>>", save_status)
    combo.bind("<Return>", save_status)
    combo.bind("<Escape>", lambda e: combo.destroy())


def setv(tr, iid, v):
    x = list(tr.item(iid, "values"))
    x[1] = str(v)
    tr.item(iid, values=x)


def read_all(app, o, tr, mapping):
    conn = app.conns.get(o["PLC"])
    if not conn:
        for iid in mapping:
            setv(tr, iid, "Not Connected")
        return

    def work():
        for iid, r in mapping.items():
            a = get_address(r)
            if not a:
                continue
            try:
                v = app.a.call(conn.read(a))
                app.after(0, lambda iid=iid, v=v: setv(tr, iid, v))
            except Exception as e:
                app.after(0, lambda iid=iid, e=e: setv(tr, iid, "ERROR: " + str(e)))

    threading.Thread(target=work, daemon=True).start()


def _build_tag_updates(o, mapping):
    updates = []
    for r in mapping.values():
        updates.append({
            "PLC": o["PLC"],
            "Tag Name": r.get("Tag Name"),
            "Address": get_address(r),
            "Verification Status": r.get("Verification Status") or "Not Tested",
            "Verification Time": r.get("Verification Time") or "",
        })
    return updates


def _save_tag_verifications(app, o, mapping):
    status_col = app.cfg["display"].get("VerifyStatusColumn")
    ts_col = app.cfg["display"].get("VerifyTimestampColumn")
    updates = _build_tag_updates(o, mapping)
    written, not_found = save_tag_verifications(app.eng.get(), "Objects", updates, status_col, ts_col)
    if not_found:
        messagebox.showwarning(
            "Some rows not matched",
            f"{written} row(s) saved. {len(not_found)} tag(s) could not be matched in the "
            f"Excel sheet (PLC/Tag Name/Address mismatch) and were not written."
        )


def _try_close(app, w, mapping, dirty_flag, o):
    if not dirty_flag["value"]:
        w.destroy()
        return
    resp = messagebox.askyesnocancel(
        "Unsaved Verification Changes",
        "You have unsaved Verify Status/Timestamp changes for this equipment.\n\n"
        "Save them to the engineering Excel file before closing?"
    )
    if resp is None:
        return  # Cancel: keep popup open
    if resp is True:
        try:
            _save_tag_verifications(app, o, mapping)
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return  # keep popup open so nothing is lost
    dirty_flag["value"] = False
    w.destroy()


def open_detail_window(app, o):
    t = app.colors()
    w = tk.Toplevel(app)
    w.title(o["Equipment"] + " - Verification")
    w.geometry("1050x680")
    w.configure(bg=t["bg"])

    dirty_flag = {"value": False}

    h = ttk.Frame(w, padding=10)
    h.pack(fill="x")
    ttk.Label(h, text=o["Equipment"], font=("Segoe UI", 16, "bold")).pack(anchor="w")
    ttk.Label(h, text=f"Template: {o['Template']}").pack(anchor="w")
    ttk.Label(h, text=f"PLC: {o['PLC']} | Area: {o['Area']}").pack(anchor="w")

    cols = ("tag", "value", "address", "type", "access", "unit", "verify", "verifyTS")
    tr_frame = ttk.Frame(w)
    tr_frame.pack(fill="both", expand=True, padx=10)
    tr = ttk.Treeview(tr_frame, columns=cols, show="headings")
    for c, ttext, wd in [("tag", "Tag Name", 230), ("value", "PLC / OPC Value", 130), ("address", "Address", 110),
                          ("type", "Data Type", 90), ("access", "Access", 70), ("unit", "Eng Units", 80),
                          ("verify", "Verify Sts", 80), ("verifyTS", "Verify TS", 80)]:
        tr.heading(c, text=ttext)
        tr.column(c, width=wd)
    tr.pack(side="left", fill="both", expand=True)
    tr_scroll = ttk.Scrollbar(tr_frame, orient="vertical", command=tr.yview)
    tr_scroll.pack(side="right", fill="y")
    tr.configure(yscrollcommand=tr_scroll.set)

    mapping = {}
    not_conn = o["PLC"] not in app.conns
    init_val = "Not Connected" if not_conn else "—"
    for i, r in enumerate(o["Rows"]):
        iid = str(i)
        mapping[iid] = r
        tr.insert("", "end", iid=iid, values=(
            s(r.get("Tag Name")), init_val, get_address(r), s(r.get("Data Type")),
            s(r.get("Client Access")) or "R", s(r.get("Eng Units")),
            s(r.get("Verification Status") or "Not Tested"), s(r.get("Verification Time"))))
    tr.bind("<Double-1>", lambda event: edit_verification_status(app, event, tr, mapping, dirty_flag))

    ctl = ttk.Frame(w, padding=10)
    ctl.pack(fill="x")
    result = tk.StringVar(value=(app.db.get(key(o)) or ["Correct"])[0])
    ttk.Label(ctl, text="Verification:").grid(row=0, column=0, sticky="w")
    for j, x in enumerate(("Correct", "Incorrect", "Recheck"), 1):
        ttk.Radiobutton(ctl, text=x, variable=result, value=x).grid(row=0, column=j, sticky="w")

    ttk.Label(ctl, text="Cause:").grid(row=1, column=0, sticky="w", pady=5)
    cause = tk.StringVar()
    cb = ttk.Combobox(ctl, textvariable=cause, values=[f"{a} - {b}" for a, b in app.cfg["causes"]], width=55)
    cb.grid(row=1, column=1, columnspan=3, sticky="w")

    ttk.Label(ctl, text="Tester:").grid(row=2, column=0, sticky="w")
    tester = tk.Entry(ctl, width=30, bg=t["entry_bg"], fg=t["entry_fg"], insertbackground=t["fg"], relief="flat")
    tester.grid(row=2, column=1, sticky="w")

    ttk.Label(ctl, text="Comment:").grid(row=3, column=0, sticky="nw")
    comment = tk.Text(ctl, width=80, height=4, bg=t["entry_bg"], fg=t["entry_fg"], insertbackground=t["fg"], relief="flat")
    comment.grid(row=3, column=1, columnspan=3)

    old = app.db.get(key(o))
    if old:
        if old[1]:
            cause.set(old[1])
        tester.insert(0, old[3] or "")
        comment.insert("1.0", old[2] or "")

    ttk.Button(w, text="Read All Parameters", command=lambda: read_all(app, o, tr, mapping), padding=(10, 6)).pack(side="left", padx=10, pady=8)

    def save():
        app.db.save(o, result.get(), cause.get(), comment.get("1.0", "end").strip(), tester.get().strip())
        if dirty_flag["value"]:
            try:
                _save_tag_verifications(app, o, mapping)
            except Exception as e:
                messagebox.showerror("Save failed", str(e))
                return  # keep popup open so nothing is lost
            dirty_flag["value"] = False
        w.destroy()

    ttk.Button(w, text="SAVE VERIFICATION", command=save, padding=(14, 6), style="Accent.TButton").pack(side="right", padx=10, pady=8)

    w.protocol("WM_DELETE_WINDOW", lambda: _try_close(app, w, mapping, dirty_flag, o))

    read_all(app, o, tr, mapping)
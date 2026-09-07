"""The per-equipment verification detail window, split out of App.open()."""

import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from data_io import s, get_address, key, ENGINEERING_HEADER_ROW, ENGINEERING_SHEET
from excel_writer import save_tag_verifications, save_general_comment


def edit_verification_status(app, event, tr, mapping, dirty_flag):
    region = tr.identify("region", event.x, event.y)
    if region != "cell":
        return
    row_id = tr.identify_row(event.y)
    column_id = tr.identify_column(event.x)
    # Verification Status = column 8
    if not row_id or column_id != "#8":
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
        tr.set(row_id, "#8", new_value)
        tr.set(row_id, "#9", verification_time)
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


def getv(tr, iid):
    x = tr.item(iid, "values")
    return x[1] if x else ""


def read_all(app, o, tr, mapping):
    conn = app.conns.get(o["PLC"])
    if not conn:
        for iid in mapping:
            setv(tr, iid, "Not Connected")
        return

    for iid, r in mapping.items():
        if get_address(r):
            setv(tr, iid, "Reading...")

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


def _build_tag_updates(o, mapping, tr=None):
    updates = []
    for iid, r in mapping.items():
        cv = getv(tr, iid) if tr is not None else ""
        if cv in ("Not Connected", "Reading...", "—"):
            cv = ""
        updates.append({
            "_row": r.get("_row"),
            "PLC": o["PLC"],
            "Tag Name": r.get("Tag Name"),
            "Address": get_address(r),
            "Verification Status": r.get("Verification Status") or "Not Tested",
            "Verification Time": r.get("Verification Time") or "",
            "Current Value": cv,
        })
    return updates


def _save_tag_verifications(app, o, mapping, tr=None):
    status_col = app.cfg["display"].get("VerifyStatusColumn")
    ts_col = app.cfg["display"].get("VerifyTimestampColumn")
    cv_col = app.cfg["display"].get("CurrentValueColumn")
    updates = _build_tag_updates(o, mapping, tr)
    written, not_found = save_tag_verifications(app.eng.get(), ENGINEERING_SHEET, updates, status_col, ts_col, cv_col,
                                                 header_row=ENGINEERING_HEADER_ROW)
    if not_found:
        messagebox.showwarning(
            "Some rows not matched",
            f"{written} row(s) saved. {len(not_found)} tag(s) had no source row reference "
            f"(engineering file may need reloading) and were not written."
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
            _save_tag_verifications(app, o, mapping, tr=None)
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

    cols = ("tag", "value", "saved", "address", "type", "access", "unit", "verify", "verifyTS")
    tr_frame = ttk.Frame(w)
    tr_frame.pack(fill="both", expand=True, padx=10)
    tr = ttk.Treeview(tr_frame, columns=cols, show="headings")
    for c, ttext, wd in [("tag", "Tag Name", 230), ("value", "PLC / OPC Value", 130), ("saved", "Saved Value", 110),
                          ("address", "Address", 110), ("type", "Data Type", 90), ("access", "Access", 70), ("unit", "Eng Units", 80),
                          ("verify", "Verify Sts (dbl-click)", 130), ("verifyTS", "Verify TS", 80)]:
        tr.heading(c, text=ttext)
        tr.column(c, width=wd)
    tr.pack(side="left", fill="both", expand=True)

    def _on_tr_motion(event):
        region = tr.identify("region", event.x, event.y)
        col = tr.identify_column(event.x)
        if region == "cell" and col == "#8":
            tr.configure(cursor="hand2")
        else:
            tr.configure(cursor="")

    tr.bind("<Motion>", _on_tr_motion)
    tr_scroll = ttk.Scrollbar(tr_frame, orient="vertical", command=tr.yview)
    tr_scroll.pack(side="right", fill="y")
    tr.configure(yscrollcommand=tr_scroll.set)

    mapping = {}
    not_conn = o["PLC"] not in app.conns
    for i, r in enumerate(o["Rows"]):
        iid = str(i)
        mapping[iid] = r
        saved = s(r.get("Current Value"))
        init_val = saved if saved else ("Not Connected" if not_conn else "—")
        tr.insert("", "end", iid=iid, values=(
            s(r.get("Tag Name")), init_val, saved, get_address(r), s(r.get("Data Type")),
            s(r.get("Client Access")) or "R", s(r.get("Eng Units")),
            s(r.get("Verification Status") or "Not Tested"), s(r.get("Verification Time"))))
    tr.bind("<Double-1>", lambda event: edit_verification_status(app, event, tr, mapping, dirty_flag))

    card = ttk.LabelFrame(w, text="Verification", padding=12)
    card.pack(fill="x", padx=10, pady=(8, 0))

    result = tk.StringVar(value=(app.comments.get(key(o)) or {}).get("status") or "Not Tested")
    pill_colors = {"Not Tested": t["muted"], "Correct": t["success"], "Incorrect": t["danger"], "Recheck": t["warning"]}
    pill_buttons = {}

    def refresh_pills():
        for val, btn in pill_buttons.items():
            if result.get() == val:
                btn.config(bg=pill_colors[val], fg="#FFFFFF")
            else:
                btn.config(bg=t["entry_bg"], fg=t["fg"])

    def pick_result(val):
        result.set(val)
        refresh_pills()

    top_row = ttk.Frame(card)
    top_row.pack(fill="x")

    result_frame = ttk.Frame(top_row)
    result_frame.pack(side="left")
    for val in ("Not Tested", "Correct", "Incorrect", "Recheck"):
        b = tk.Button(result_frame, text=val, width=11, relief="flat", cursor="hand2", bd=0,
                      activeforeground="#FFFFFF", command=lambda v=val: pick_result(v))
        b.pack(side="left", padx=(0, 6), ipady=4)
        pill_buttons[val] = b
    refresh_pills()

    tester_frame = ttk.Frame(top_row)
    tester_frame.pack(side="left", padx=(24, 0))
    ttk.Label(tester_frame, text="Tester:").pack(anchor="w")
    tester = tk.Entry(tester_frame, width=22, bg=t["entry_bg"], fg=t["entry_fg"], insertbackground=t["fg"], relief="flat")
    tester.pack(anchor="w")
    tester_warning = ttk.Label(tester_frame, text="", foreground=t["danger"])
    tester_warning.pack(anchor="w")

    def check_tester(*a):
        tester_warning.config(text="" if tester.get().strip() else "⚠ Tester required before saving")

    tester.bind("<KeyRelease>", check_tester)

    ttk.Label(card, text="Cause:").pack(anchor="w", pady=(12, 2))
    cause = tk.StringVar()
    cb = ttk.Combobox(card, textvariable=cause, values=[f"{a} - {b}" for a, b in app.cfg["causes"]], width=45)
    cb.pack(anchor="w")

    ttk.Label(card, text="Comment:").pack(anchor="w", pady=(12, 2))
    comment = tk.Text(card, height=5, bg=t["entry_bg"], fg=t["entry_fg"], insertbackground=t["fg"], relief="flat")
    comment.pack(fill="x")

    old = app.comments.get(key(o))
    if old:
        if old.get("cause"):
            cause.set(old["cause"])
        tester.insert(0, old.get("tester") or "")
        comment.insert("1.0", old.get("comment") or "")
    check_tester()

    def save():
        status_v = result.get()
        cause_v = cause.get()
        tester_v = tester.get().strip()
        comment_v = comment.get("1.0", "end").strip()
        try:
            save_general_comment(app.eng.get(), o["PLC"], o["Template"], o["Area"], o["Equipment"],
                                  status_v, cause_v, tester_v, comment_v)
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return
        app.comments[key(o)] = {"status": status_v, "cause": cause_v, "tester": tester_v, "comment": comment_v}
        app.populate_sidebar()
        app.render_objects()
        try:
            _save_tag_verifications(app, o, mapping, tr)
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return  # keep popup open so nothing is lost
        dirty_flag["value"] = False
        w.destroy()

    ttk.Separator(w).pack(fill="x", padx=10, pady=(10, 0))
    btn_row = ttk.Frame(w)
    btn_row.pack(fill="x")
    ttk.Button(btn_row, text="Read All Parameters", command=lambda: read_all(app, o, tr, mapping), padding=(10, 6)).pack(side="left", padx=10, pady=8)
    ttk.Button(btn_row, text="Save Verification", command=save, padding=(14, 6), style="Accent.TButton").pack(side="right", padx=10, pady=8)

    w.protocol("WM_DELETE_WINDOW", lambda: _try_close(app, w, mapping, dirty_flag, o))

    read_all(app, o, tr, mapping)
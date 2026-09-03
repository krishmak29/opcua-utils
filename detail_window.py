"""The per-equipment verification detail window, split out of App.open()."""

import threading
import tkinter as tk
from tkinter import ttk
from datetime import datetime

from data_io import s, get_address
from db import key


def edit_verification_status(app, event, tr, mapping):
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


def open_detail_window(app, o):
    t = app.colors()
    w = tk.Toplevel(app)
    w.title(o["Equipment"] + " - Verification")
    w.geometry("1050x680")
    w.configure(bg=t["bg"])

    h = ttk.Frame(w, padding=10)
    h.pack(fill="x")
    ttk.Label(h, text=o["Equipment"], font=("Segoe UI", 16, "bold")).pack(anchor="w")
    ttk.Label(h, text=f"Template: {o['Template']}").pack(anchor="w")
    ttk.Label(h, text=f"PLC: {o['PLC']} | Area: {o['Area']}").pack(anchor="w")

    cols = ("tag", "value", "address", "type", "access", "unit", "verify", "verifyTS")
    tr = ttk.Treeview(w, columns=cols, show="headings")
    for c, ttext, wd in [("tag", "Tag Name", 230), ("value", "PLC / OPC Value", 130), ("address", "Address", 110),
                          ("type", "Data Type", 90), ("access", "Access", 70), ("unit", "Eng Units", 80),
                          ("verify", "Verify Sts", 80), ("verifyTS", "Verify TS", 80)]:
        tr.heading(c, text=ttext)
        tr.column(c, width=wd)
    tr.pack(fill="both", expand=True, padx=10)

    mapping = {}
    not_conn = o["PLC"] not in app.conns
    init_val = "Not Connected" if not_conn else "—"
    for i, r in enumerate(o["Rows"]):
        iid = str(i)
        mapping[iid] = r
        tr.insert("", "end", iid=iid, values=(
            s(r.get("Tag Name")), init_val, get_address(r), s(r.get("Data Type")),
            s(r.get("Client Access")) or "R", s(r.get("Eng Units")),
            s(r.get("Verification Status") or "Not Tested"), s(r.get("Verification TS"))))
    tr.bind("<Double-1>", lambda event: edit_verification_status(app, event, tr, mapping))

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

    ttk.Button(w, text="Read All Parameters", command=lambda: read_all(app, o, tr, mapping)).pack(side="left", padx=10, pady=8)

    def save():
        app.db.save(o, result.get(), cause.get(), comment.get("1.0", "end").strip(), tester.get().strip())
        w.destroy()
        app.load()

    ttk.Button(w, text="SAVE VERIFICATION", command=save).pack(side="right", padx=10, pady=8)
    read_all(app, o, tr, mapping)

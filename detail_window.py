"""The per-equipment verification detail window, split out of App.open()."""

import threading
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime

from data_io import s, get_address, get_plc_address, key, ENGINEERING_HEADER_ROW, ENGINEERING_SHEET
from excel_writer import save_tag_verifications, save_general_comment
import pending_store as ps


def set_verification_status(tr, mapping, dirty_flag, dirty_rows, row_id, new_value):
    """Shared status-update mechanism: used by the double-click combobox editor
    and by the six Ctrl+<key> verification shortcuts alike, so the update logic
    exists in exactly one place."""
    verification_time = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    live_value = tr.set(row_id, "#2")
    tr.set(row_id, "#3", live_value)
    tr.set(row_id, "#8", new_value)
    tr.set(row_id, "#9", verification_time)
    if row_id in mapping:
        mapping[row_id]["Verification Status"] = new_value
        mapping[row_id]["Verification Time"] = verification_time
        mapping[row_id]["Current Value"] = live_value
    dirty_flag["value"] = True
    dirty_rows.add(row_id)


def edit_verification_status(app, event, tr, mapping, dirty_flag, dirty_rows):
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
        set_verification_status(tr, mapping, dirty_flag, dirty_rows, row_id, combo.get())
        combo.destroy()

    combo.bind("<<ComboboxSelected>>", save_status)
    combo.bind("<Return>", save_status)
    combo.bind("<Escape>", lambda e: combo.destroy())


VERIFICATION_SHORTCUTS = {
    "o": "OK",
    "r": "Recheck",
    "m": "SCADA-Value Mismatch",
    "i": "SCADA-Linking issue",
    "q": "Not Available",
    "k": "Not Ok",
}


def bind_verification_shortcuts(tr, mapping, dirty_flag, dirty_rows):
    """Ctrl+O/R/M/I/Q/K set the selected tag row's Verification Status via the
    same set_verification_status() mechanism as the double-click editor. No
    selected row -> no action, no popup/error (per locked spec 6.3)."""
    def handler(status):
        def _handle(event):
            sel = tr.selection()
            if not sel:
                return
            row_id = sel[0]
            set_verification_status(tr, mapping, dirty_flag, dirty_rows, row_id, status)
        return _handle

    for letter, status in VERIFICATION_SHORTCUTS.items():
        tr.bind(f"<Control-{letter}>", handler(status))
        tr.bind(f"<Control-{letter.upper()}>", handler(status))


def setv(tr, iid, v):
    x = list(tr.item(iid, "values"))
    x[1] = str(v)
    tr.item(iid, values=x)


def getv(tr, iid):
    x = tr.item(iid, "values")
    return x[2] if x else ""


def read_all(app, o, tr, mapping):
    conn = app.conns.get(o["PLC"])
    if not conn:
        for iid in mapping:
            setv(tr, iid, "-")
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


def _build_tag_updates(o, mapping, tr=None, dirty_rows=None, tester=""):
    updates = []
    for iid, r in mapping.items():
        if dirty_rows is not None and iid not in dirty_rows:
            continue
        cv = getv(tr, iid) if tr is not None else ""
        if cv in ("Not Connected", "Reading...", "—"):
            cv = "-"
        updates.append({
            "_row": r.get("_row"),
            "PLC": o["PLC"],
            "Tag Name": r.get("Tag Name"),
            "Address": get_address(r),
            "Verification Status": r.get("Verification Status") or "Not Tested",
            "Verification Time": r.get("Verification Time") or "",
            "Current Value": cv,
            "Tester Name": tester,
        })
    return updates


def _save_tag_verifications_pending(o, mapping, tr=None, dirty_rows=None, tester=""):
    """Popup Save: write dirty tag rows into pending_changes.json instead of
    Excel. Excel is only touched by the main screen's batch 'Save to Excel'
    action (Phase 3)."""
    rows_to_write = dirty_rows if dirty_rows else set(mapping.keys())
    updates = _build_tag_updates(o, mapping, tr, rows_to_write, tester)
    skipped_no_row = 0
    for u in updates:
        if u.get("_row") is None:
            skipped_no_row += 1
            continue
        ps.merge_tag_update(u)
    if skipped_no_row:
        messagebox.showwarning(
            "Some rows not matched",
            f"{skipped_no_row} tag(s) had no source row reference "
            f"(engineering file may need reloading) and were not saved."
        )


def _try_close(app, w, mapping, dirty_flag, o, dirty_rows, tester_widget):
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
        tester_v = tester_widget.get().strip()
        if not tester_v:
            messagebox.showerror("Tester required", "Please enter a Tester Name before saving.")
            return  # keep popup open so nothing is lost
        try:
            _save_tag_verifications_pending(o, mapping, tr=None, dirty_rows=dirty_rows, tester=tester_v)
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return  # keep popup open so nothing is lost
        app.last_tester_name = tester_v
    dirty_flag["value"] = False
    w.destroy()


def open_detail_window(app, o):
    t = app.colors()
    w = tk.Toplevel(app)
    w.title(o["Equipment"] + " - Verification")
    w.configure(bg=t["bg"])
    ww, wh = 1050, 680
    app.update_idletasks()
    px, py = app.winfo_rootx(), app.winfo_rooty()
    pw, ph = app.winfo_width(), app.winfo_height()
    x = px + (pw - ww) // 2
    y = py + (ph - wh) // 2
    w.geometry(f"{ww}x{wh}+{x}+{y}")

    dirty_flag = {"value": False}
    dirty_rows = set()

    h = ttk.Frame(w, padding=10)
    h.pack(fill="x")
    ttk.Label(h, text=o["Equipment"], font=("Segoe UI", 16, "bold")).pack(anchor="w")
    ttk.Label(h, text=f"Template: {o['Template']}").pack(anchor="w")
    ttk.Label(h, text=f"PLC: {o['PLC']} | Area: {o['Area']}").pack(anchor="w")

    cols = ("tag", "value", "saved", "address", "type", "access", "unit", "verify", "verifyTS")
    COLUMN_LABELS = {
        "tag": "Tag Name", "value": "PLC / OPC Value", "saved": "Saved Value",
        "address": "PLC Address", "type": "Data Type", "access": "Access",
        "unit": "Eng Units", "verify": "Verify Status", "verifyTS": "Verify TS",
    }
    tr_frame = ttk.Frame(w)
    tr_frame.pack(fill="both", expand=True, padx=10)
    tr_frame.grid_rowconfigure(0, weight=1)
    tr_frame.grid_columnconfigure(0, weight=1)
    tr = ttk.Treeview(tr_frame, columns=cols, show="headings")
    for c, ttext, wd in [("tag", "Tag Name", 230), ("value", "PLC / OPC Value", 130), ("saved", "Saved Value", 110),
                          ("address", "PLC Address", 160), ("type", "Data Type", 90), ("access", "Access", 70), ("unit", "Eng Units", 80),
                          ("verify", "Verify Status", 130), ("verifyTS", "Verify TS", 80)]:
        # anchor="w": left-align header text so it never runs into the
        # filter arrow, which is placed over the column's right edge.
        # Shortened a couple of labels ("PLC Address (click to sort)",
        # "Verify Sts (dbl-click)") that were long enough to get clipped
        # by the column width even before the arrow was added.
        tr.heading(c, text=ttext, anchor="w")
        tr.column(c, width=wd, stretch=False, anchor="w")
    tr.grid(row=0, column=0, sticky="nsew")

    def _on_tr_motion(event):
        region = tr.identify("region", event.x, event.y)
        col = tr.identify_column(event.x)
        if region == "cell" and col == "#8":
            tr.configure(cursor="hand2")
        else:
            tr.configure(cursor="")

    tr.bind("<Motion>", _on_tr_motion)
    tr_vscroll = ttk.Scrollbar(tr_frame, orient="vertical", command=tr.yview)
    tr_vscroll.grid(row=0, column=1, sticky="ns")
    tr_hscroll = ttk.Scrollbar(tr_frame, orient="horizontal", command=tr.xview)
    tr_hscroll.grid(row=1, column=0, sticky="ew")
    tr.configure(yscrollcommand=tr_vscroll.set, xscrollcommand=tr_hscroll.set)

    mapping = {}
    not_conn = o["PLC"] not in app.conns
    for i, r in enumerate(o["Rows"]):
        iid = str(i)
        pending = ps.get_pending_tag(r.get("_row"))
        if pending:
            # A pending JSON edit exists for this Excel row from an earlier
            # popup Save that hasn't been committed to Excel yet -- show
            # that instead of the (now-stale, from this popup's perspective)
            # Excel-loaded value (locked spec 3.7).
            r["Verification Status"] = pending.get("Verification Status") or r.get("Verification Status")
            r["Verification Time"] = pending.get("Verification Time") or r.get("Verification Time")
            r["Current Value"] = pending.get("Current Value") or r.get("Current Value")
        mapping[iid] = r
        saved = s(r.get("Current Value"))
        init_val = saved if saved else ("-" if not_conn else "—")
        tr.insert("", "end", iid=iid, values=(
            s(r.get("Tag Name")), init_val, saved, get_plc_address(r), s(r.get("Data Type")),
            s(r.get("Client Access")) or "R", s(r.get("Eng Units")),
            s(r.get("Verification Status") or "Not Tested"), s(r.get("Verification Time"))))
    tr.bind("<Double-1>", lambda event: edit_verification_status(app, event, tr, mapping, dirty_flag, dirty_rows))
    bind_verification_shortcuts(tr, mapping, dirty_flag, dirty_rows)

    column_filter_state = {c: None for c in cols}
    column_sort_state = {"col": None, "asc": True}
    column_filter_buttons = {}

    def _column_value(iid, col):
        if col == "address":
            return get_plc_address(mapping[iid])
        return tr.set(iid, col)

    def _column_x_offset(col):
        x = 0
        for c in cols:
            if c == col:
                break
            x += int(tr.column(c, "width"))
        return x

    def _place_column_filter_buttons():
        style = ttk.Style(tr)
        header_height = style.lookup("Treeview.Heading", "height") or 25
        try:
            header_height = int(header_height)
        except (TypeError, ValueError):
            header_height = 25
        btn_h = max(header_height - 2, 16)
        btn_w = 14
        for c, btn in column_filter_buttons.items():
            x = _column_x_offset(c)
            width = int(tr.column(c, "width"))
            # Narrow strip flush against the column's right border, not
            # overlapping the header text (which is now left-anchored),
            # so it reads as part of the header rather than a chip on top.
            btn.place(in_=tr, x=x + width - btn_w - 2, y=1, width=btn_w, height=btn_h)

    def refresh_view():
        items = list(mapping.items())
        for c, selected in column_filter_state.items():
            if selected is not None:
                items = [(iid, r) for iid, r in items if _column_value(iid, c) in selected]
        sort_col = column_sort_state["col"]
        if sort_col:
            items.sort(key=lambda kv: _column_value(kv[0], sort_col).lower(), reverse=not column_sort_state["asc"])
        else:
            items.sort(key=lambda kv: int(kv[0]))
        visible = {iid for iid, _ in items}
        for iid in mapping:
            if iid not in visible:
                tr.detach(iid)
        for idx, (iid, _r) in enumerate(items):
            tr.move(iid, "", idx)

    def open_column_filter_popup(col):
        label = COLUMN_LABELS.get(col, col)
        all_values = sorted({v for v in (_column_value(iid, col) for iid in mapping) if v})
        current = column_filter_state[col]
        checked = {v: (current is None or v in current) for v in all_values}

        pop = tk.Toplevel(w)
        pop.title(f"{label} Filter")
        pop.transient(w)
        pop.configure(bg=t["panel"])
        pop.minsize(260, 320)
        btn = column_filter_buttons[col]
        x = btn.winfo_rootx()
        y = btn.winfo_rooty() + btn.winfo_height()
        pop.geometry(f"280x480+{x}+{y}")

        sort_row = ttk.Frame(pop, padding=(8, 8, 8, 4))
        sort_row.pack(fill="x")

        def do_sort(asc):
            column_sort_state["col"] = col
            column_sort_state["asc"] = asc
            refresh_view()

        ttk.Button(sort_row, text="Sort A→Z", command=lambda: do_sort(True)).pack(side="left", padx=(0, 4))
        ttk.Button(sort_row, text="Sort Z→A", command=lambda: do_sort(False)).pack(side="left")

        ttk.Separator(pop).pack(fill="x", padx=8)

        search_row = ttk.Frame(pop, padding=(8, 6, 8, 2))
        search_row.pack(fill="x")
        ttk.Label(search_row, text="Search:").pack(side="left")
        search_var = tk.StringVar()
        search_entry = tk.Entry(search_row, textvariable=search_var, bg=t["entry_bg"], fg=t["entry_fg"],
                                 insertbackground=t["fg"], relief="flat")
        search_entry.pack(side="left", fill="x", expand=True, padx=(6, 0))

        list_container = ttk.Frame(pop)
        list_container.pack(fill="both", expand=True, padx=8, pady=4, side="top")
        canvas = tk.Canvas(list_container, bg=t["panel"], highlightthickness=0)
        vscroll = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_wheel(_e=None):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_wheel(_e=None):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)
        pop.bind("<Destroy>", lambda e: _unbind_wheel())

        check_vars = {}

        def rebuild_list(*_a):
            for child in inner.winfo_children():
                child.destroy()
            check_vars.clear()
            q = search_var.get().strip().lower()
            for v in all_values:
                if q and q not in v.lower():
                    continue
                var = tk.BooleanVar(value=checked.get(v, True))
                check_vars[v] = var
                cb = tk.Checkbutton(inner, text=v, variable=var, bg=t["panel"], fg=t["fg"],
                                     selectcolor=t["entry_bg"], anchor="w")
                cb.pack(fill="x", anchor="w")

        search_var.trace_add("write", rebuild_list)
        rebuild_list()

        def select_all():
            for var in check_vars.values():
                var.set(True)

        def deselect_all():
            for var in check_vars.values():
                var.set(False)

        def apply_and_close():
            for v in all_values:
                if v in check_vars:
                    checked[v] = check_vars[v].get()
            newly_selected = {v for v, on in checked.items() if on}
            column_filter_state[col] = None if newly_selected == set(all_values) else newly_selected
            refresh_view()
            pop.destroy()

        btn_row = ttk.Frame(pop, padding=(8, 4, 8, 8))
        btn_row.pack(fill="x", side="bottom")
        ttk.Button(btn_row, text="Apply", command=apply_and_close, style="Accent.TButton").pack(side="right")
        ttk.Button(btn_row, text="Cancel", command=pop.destroy).pack(side="right", padx=(0, 6))

        sel_row = ttk.Frame(pop, padding=(8, 2, 8, 4))
        sel_row.pack(fill="x", side="bottom")
        ttk.Button(sel_row, text="Select All", command=select_all).pack(side="left", padx=(0, 4))
        ttk.Button(sel_row, text="Deselect All", command=deselect_all).pack(side="left")

    header_bg = ttk.Style(tr).lookup("Treeview.Heading", "background") or t["panel"]
    if app.theme == "light":
        # In light mode a button that matches the (near-white) header
        # background is nearly invisible, so give it real contrast instead.
        btn_bg, btn_fg, btn_hover = t["muted"], "#FFFFFF", t["accent"]
    else:
        # In dark mode the header is already dark, so keep the subtle,
        # blended-in look.
        btn_bg, btn_fg, btn_hover = header_bg, t["muted"], t["tree_sel"]
    for c in cols:
        b = tk.Button(tr, text="▾", font=("Segoe UI", 7), relief="flat", cursor="hand2",
                      bd=0, highlightthickness=0, takefocus=0,
                      bg=btn_bg, fg=btn_fg,
                      activebackground=btn_hover, activeforeground="#FFFFFF")
        b.configure(command=lambda c=c: open_column_filter_popup(c))
        # Hover highlight -- plain tk.Button only shows activebackground on
        # press, not on mouse-over, so bind Enter/Leave to make it feel
        # responsive.
        b.bind("<Enter>", lambda e, b=b: b.configure(bg=btn_hover, fg="#FFFFFF"))
        b.bind("<Leave>", lambda e, b=b: b.configure(bg=btn_bg, fg=btn_fg))
        column_filter_buttons[c] = b

    tr.bind("<Configure>", lambda e: _place_column_filter_buttons())
    w.after(50, _place_column_filter_buttons)

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
    comment = tk.Text(card, height=5, bg=t["entry_bg"], fg=t["entry_fg"], insertbackground=t["fg"], relief="flat",
                       selectbackground=t["accent"], selectforeground="#FFFFFF")
    comment.pack(fill="x")

    pending_eq = ps.get_pending_equipment(o["PLC"], o["Area"], o["Equipment"])
    old = pending_eq or app.comments.get(key(o))
    if pending_eq:
        result.set(pending_eq.get("Status") or result.get())
        refresh_pills()
    if old:
        cause_v = pending_eq.get("Cause") if pending_eq else old.get("cause")
        tester_v = pending_eq.get("Tester") if pending_eq else old.get("tester")
        comment_v = pending_eq.get("Comment") if pending_eq else old.get("comment")
        if cause_v:
            cause.set(cause_v)
        tester.insert(0, tester_v or app.last_tester_name or "")
        comment.insert("1.0", comment_v or "")
    else:
        tester.insert(0, app.last_tester_name or "")
    check_tester()

    def save():
        status_v = result.get()
        cause_v = cause.get()
        tester_v = tester.get().strip()
        comment_v = comment.get("1.0", "end").strip()
        if not tester_v:
            check_tester()
            messagebox.showerror("Tester required", "Please enter a Tester Name before saving.")
            return
        try:
            ps.merge_equipment_update(o["PLC"], o["Template"], o["Area"], o["Equipment"],
                                       status_v, cause_v, tester_v, comment_v)
            _save_tag_verifications_pending(o, mapping, tr, dirty_rows=dirty_rows, tester=tester_v)
        except Exception as e:
            messagebox.showerror("Save failed", str(e))
            return  # keep popup open so nothing is lost
        app.last_tester_name = tester_v
        app.comments[key(o)] = {"status": status_v, "cause": cause_v, "tester": tester_v, "comment": comment_v}
        app.populate_sidebar()
        app.render_objects()
        if hasattr(app, "refresh_pending_indicator"):
            app.refresh_pending_indicator()
        dirty_flag["value"] = False
        w.destroy()

    ttk.Separator(w).pack(fill="x", padx=10, pady=(10, 0))
    btn_row = ttk.Frame(w)
    btn_row.pack(fill="x")
    ttk.Button(btn_row, text="Read All Parameters", command=lambda: read_all(app, o, tr, mapping), padding=(10, 6)).pack(side="left", padx=10, pady=8)
    ttk.Button(btn_row, text="Save Verification", command=save, padding=(14, 6), style="Accent.TButton").pack(side="right", padx=10, pady=8)

    w.protocol("WM_DELETE_WINDOW", lambda: _try_close(app, w, mapping, dirty_flag, o, dirty_rows, tester))

    read_all(app, o, tr, mapping)
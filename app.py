"""Main application window: PLC sidebar, object list, settings dialog."""

import threading
import tkinter as tk
from tkinter import ttk, messagebox

from config import ENGINEERING, CONFIG
from theme import THEMES
from data_io import load_engineering, load_config, group
from db import DB, key
from opcua_client import Async, Conn
from detail_window import open_detail_window


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PLC / SCADA Object Verification Utility")
        self.geometry("1250x780")
        self.cfg = load_config(CONFIG)
        self.db = DB()
        self.a = Async()
        self.objects = []
        self.byplc = {}
        self.conns = {}
        self.theme = "light"
        self.build()
        self.apply_theme()
        self.load()
        self.protocol("WM_DELETE_WINDOW", self.close)

    def build(self):
        self.eng = tk.StringVar(value=ENGINEERING)
        self.cf = tk.StringVar(value=CONFIG)
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="PLC / SCADA Object Verification Utility", font=("Segoe UI", 12, "bold")).pack(side="left")
        self.theme_btn = ttk.Button(top, text="🌙 Dark Mode", command=self.toggle_theme)
        self.theme_btn.pack(side="right", padx=(0, 8))
        ttk.Button(top, text="⚙ Settings", command=self.open_settings).pack(side="right")
        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, padding=(8, 0)).pack(fill="x")

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=8, pady=8)

        side = ttk.Frame(body, width=220)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self.populate_sidebar())
        ttk.Entry(side, textvariable=self.search_var).pack(fill="x", pady=(0, 5))
        ttk.Label(side, text="Search PLC...", foreground="gray").place(in_=side, x=6, y=4) if False else None
        self.plc_list = ttk.Treeview(side, columns=("count",), show="tree headings", height=25)
        self.plc_list.heading("#0", text="PLC")
        self.plc_list.heading("count", text="Tested")
        self.plc_list.column("#0", width=140)
        self.plc_list.column("count", width=70, anchor="e")
        self.plc_list.pack(fill="both", expand=True)
        self.plc_list.bind("<<TreeviewSelect>>", self.on_plc_select)

        main = ttk.Frame(body)
        main.pack(side="left", fill="both", expand=True, padx=(10, 0))
        head = ttk.Frame(main)
        head.pack(fill="x")
        self.info = ttk.Label(head, font=("Segoe UI", 10, "bold"))
        self.info.pack(side="left")
        self.filter = tk.StringVar(value="All")
        ttk.Label(head, text=" Show:").pack(side="left", padx=(20, 3))
        cb = ttk.Combobox(head, textvariable=self.filter, values=("All", "Not Tested", "Correct", "Incorrect", "Recheck"), state="readonly", width=14)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda e: self.render_objects())

        list_frame = ttk.Frame(main)
        list_frame.pack(fill="both", expand=True, pady=8)
        cols = ("template", "area", "status")
        self.obj_list = ttk.Treeview(list_frame, columns=cols, show="tree headings")
        self.obj_list.heading("#0", text="Equipment")
        self.obj_list.heading("template", text="Template")
        self.obj_list.heading("area", text="Area")
        self.obj_list.heading("status", text="Status")
        self.obj_list.column("#0", width=160)
        self.obj_list.column("template", width=220)
        self.obj_list.column("area", width=150)
        self.obj_list.column("status", width=120)
        self.obj_list.pack(side="left", fill="both", expand=True)
        obj_scroll = ttk.Scrollbar(list_frame, orient="vertical", command=self.obj_list.yview)
        obj_scroll.pack(side="right", fill="y")
        self.obj_list.configure(yscrollcommand=obj_scroll.set)
        self.obj_list.bind("<Double-1>", lambda e: self.open_selected())

        self.pvar = tk.StringVar()

        self.progress = tk.StringVar()
        ttk.Label(self, textvariable=self.progress, padding=8).pack(fill="x")

        self.current_plc = None
        self.page_no = 0
        self._objmap = {}

    def colors(self):
        return THEMES[self.theme]

    def toggle_theme(self):
        self.theme = "dark" if self.theme == "light" else "light"
        self.apply_theme()

    def apply_theme(self):
        t = self.colors()
        self.configure(bg=t["bg"])
        self.theme_btn.config(text="☀ Light Mode" if self.theme == "dark" else "🌙 Dark Mode")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background=t["bg"], foreground=t["fg"], fieldbackground=t["entry_bg"])
        style.configure("TFrame", background=t["bg"])
        style.configure("TLabel", background=t["bg"], foreground=t["fg"])
        style.configure("TButton", background=t["panel"], foreground=t["fg"], bordercolor=t["border"], focuscolor=t["accent"])
        style.map("TButton", background=[("active", t["tree_sel"])])
        style.configure("TEntry", fieldbackground=t["entry_bg"], foreground=t["entry_fg"], bordercolor=t["border"])
        style.configure("TCombobox", fieldbackground=t["entry_bg"], foreground=t["entry_fg"], background=t["panel"])
        style.map("TCombobox", fieldbackground=[("readonly", t["entry_bg"])])
        style.configure("TSeparator", background=t["border"])
        style.configure("Treeview", background=t["tree_bg"], fieldbackground=t["tree_bg"], foreground=t["tree_fg"], bordercolor=t["border"], borderwidth=0)
        style.map("Treeview", background=[("selected", t["tree_sel"])], foreground=[("selected", t["fg"])])
        style.configure("Treeview.Heading", background=t["panel"], foreground=t["fg"], bordercolor=t["border"])
        style.map("Treeview.Heading", background=[("active", t["tree_sel"])])

    def reload(self):
        self.cfg = load_config(self.cf.get())
        self.load()

    def load(self):
        try:
            self.objects = group(load_engineering(self.eng.get()))
            self.byplc = {}
            for o in self.objects:
                self.byplc.setdefault(o["PLC"], []).append(o)
            self.current_plc = None
            self.populate_sidebar()
            self.render_objects()
            self.status.set(f"Loaded {len(self.objects)} objects")
            self.update_progress()
        except Exception as e:
            messagebox.showerror("Load error", str(e))

    def open_settings(self):
        w = tk.Toplevel(self)
        w.title("Settings")
        w.geometry("520x300")
        w.transient(self)
        w.configure(bg=self.colors()["bg"])
        f = ttk.Frame(w, padding=12)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="Engineering Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.eng, width=45).grid(row=0, column=1, padx=5, pady=4)
        ttk.Button(f, text="Load", command=self.load).grid(row=0, column=2)
        ttk.Label(f, text="Config File:").grid(row=1, column=0, sticky="w")
        ttk.Entry(f, textvariable=self.cf, width=45).grid(row=1, column=1, padx=5, pady=4)
        ttk.Button(f, text="Reload Config", command=self.reload).grid(row=1, column=2)
        ttk.Separator(f).grid(row=2, column=0, columnspan=3, sticky="ew", pady=12)
        ttk.Label(f, text="Connection", font=("Segoe UI", 10, "bold")).grid(row=3, column=0, sticky="w")
        ttk.Button(f, text="Connect All", command=self.connect).grid(row=4, column=0, pady=6, sticky="w")
        ttk.Button(f, text="Disconnect", command=self.disconnect).grid(row=4, column=1, pady=6, sticky="w")
        ttk.Button(f, text="Close", command=w.destroy).grid(row=5, column=2, pady=(20, 0), sticky="e")

    def populate_sidebar(self):
        for i in self.plc_list.get_children():
            self.plc_list.delete(i)
        q = self.search_var.get().strip().lower()
        for plc in sorted(self.byplc):
            if q and q not in plc.lower():
                continue
            total = len(self.byplc[plc])
            tested = sum(bool(self.db.get(key(o))) for o in self.byplc[plc])
            self.plc_list.insert("", "end", iid=plc, text=plc, values=(f"{tested}/{total}",))
        kids = self.plc_list.get_children()
        if kids and not self.current_plc:
            self.plc_list.selection_set(kids[0])

    def on_plc_select(self, e):
        sel = self.plc_list.selection()
        if not sel:
            return
        self.current_plc = sel[0]
        self.render_objects()

    def filtered_objects(self):
        x = self.byplc.get(self.current_plc, [])
        fl = self.filter.get()
        if fl == "Not Tested":
            return [o for o in x if not self.db.get(key(o))]
        if fl in ("Correct", "Incorrect", "Recheck"):
            return [o for o in x if (self.db.get(key(o)) or [None])[0] == fl]
        return x

    def render_objects(self):
        if not self.current_plc:
            return
        for i in self.obj_list.get_children():
            self.obj_list.delete(i)
        x = self.filtered_objects()
        self._objmap = {}
        for i, o in enumerate(x):
            r = self.db.get(key(o))
            result = r[0] if r else "Not Tested"
            iid = f"row{i}"
            self._objmap[iid] = o
            self.obj_list.insert("", "end", iid=iid, text=o["Equipment"], values=(o["Template"], o["Area"], result))
        self.pvar.set(f"{len(x)} object(s)")
        total = len(self.byplc[self.current_plc])
        tested = sum(bool(self.db.get(key(o))) for o in self.byplc[self.current_plc])
        self.info.config(text=f"{self.current_plc} | Total: {total} | Tested: {tested} | Remaining: {total-tested}")

    def open_selected(self):
        sel = self.obj_list.selection()
        if sel and sel[0] in self._objmap:
            self.open(self._objmap[sel[0]])

    def open(self, o):
        open_detail_window(self, o)

    def connect(self):
        def work():
            for plc, cfg in self.cfg["plc"].items():
                try:
                    c = Conn(cfg)
                    self.a.call(c.connect())
                    self.conns[plc] = c
                except Exception as e:
                    self.status.set(f"{plc}: {e}")
            self.status.set(f"Connected {len(self.conns)} PLC(s)")
        threading.Thread(target=work, daemon=True).start()

    def disconnect(self):
        for c in self.conns.values():
            try:
                self.a.call(c.disconnect())
            except Exception:
                pass
        self.conns = {}
        self.status.set("Disconnected")

    def update_progress(self):
        n = len(self.objects)
        rec = [self.db.get(key(o)) for o in self.objects]
        t = sum(bool(x) for x in rec)
        self.progress.set(f"Overall: {n} Objects | {t} Tested | {n-t} Remaining")

    def close(self):
        self.disconnect()
        self.db.c.close()
        self.a.stop()
        self.destroy()

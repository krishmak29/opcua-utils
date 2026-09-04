"""Main application window: PLC sidebar, object list, settings dialog."""

import threading
import tkinter as tk
from tkinter import ttk, messagebox

from config import ENGINEERING, CONFIG
from theme import THEMES
from data_io import load_engineering, load_config, group, key
from excel_writer import load_general_comments
from opcua_client import Async, Conn
from detail_window import open_detail_window


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PLC / SCADA Object Verification Utility")
        self.geometry("1250x780")
        self.cfg = load_config(CONFIG)
        self.comments = {}
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
        self.theme_btn = ttk.Button(top, text="Dark Mode", command=self.toggle_theme, padding=(10, 5))
        self.theme_btn.pack(side="right", padx=(0, 8))
        ttk.Button(top, text="Settings", command=self.open_settings, padding=(10, 5)).pack(side="right", padx=(0, 8))
        ttk.Button(top, text="Disconnect", command=self.disconnect, padding=(10, 5)).pack(side="right", padx=(0, 8))
        ttk.Button(top, text="Connect All", command=self.connect, padding=(10, 5), style="Accent.TButton").pack(side="right", padx=(0, 8))
        self.status = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status, padding=(8, 0)).pack(fill="x")

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True, padx=8, pady=8)

        side = ttk.Frame(body, width=230)
        side.pack(side="left", fill="y")
        side.pack_propagate(False)
        PLACEHOLDER = "Search PLC..."
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(side, textvariable=self.search_var)
        self.search_entry.pack(fill="x", pady=(0, 5))
        self._search_placeholder_active = False

        def _clear_placeholder(e=None):
            if self._search_placeholder_active:
                self.search_entry.delete(0, "end")
                self.search_entry.config(foreground=self.colors()["fg"])
                self._search_placeholder_active = False

        def _set_placeholder(e=None):
            if not self.search_var.get():
                self._search_placeholder_active = True
                self.search_entry.insert(0, PLACEHOLDER)
                self.search_entry.config(foreground=self.colors()["muted"])

        self.search_entry.bind("<FocusIn>", _clear_placeholder)
        self.search_entry.bind("<FocusOut>", _set_placeholder)
        _set_placeholder()

        def _on_search_write(*a):
            if not self._search_placeholder_active:
                self.populate_sidebar()

        self.search_var.trace_add("write", _on_search_write)
        plc_frame = ttk.Frame(side)
        plc_frame.pack(fill="both", expand=True)
        self.plc_list = ttk.Treeview(plc_frame, columns=("count", "status"), show="tree headings", height=25)
        self.plc_list.heading("#0", text="PLC")
        self.plc_list.heading("count", text="Tested")
        self.plc_list.heading("status", text="Link")
        self.plc_list.column("#0", width=118, stretch=True)
        self.plc_list.column("count", width=52, anchor="center", stretch=False)
        self.plc_list.column("status", width=36, anchor="center", stretch=False)
        self.plc_list.pack(side="left", fill="both", expand=True)
        plc_scroll = ttk.Scrollbar(plc_frame, orient="vertical", command=self.plc_list.yview)
        plc_scroll.pack(side="right", fill="y")
        self.plc_list.configure(yscrollcommand=plc_scroll.set)
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
        self.theme_btn.config(text="Light Mode" if self.theme == "dark" else "Dark Mode")
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(".", background=t["bg"], foreground=t["fg"], fieldbackground=t["entry_bg"])
        style.configure("TFrame", background=t["bg"])
        style.configure("TLabel", background=t["bg"], foreground=t["fg"])
        style.configure("TButton", background=t["panel"], foreground=t["fg"], bordercolor=t["border"], focuscolor=t["accent"])
        style.map("TButton", background=[("active", t["tree_sel"])])
        style.configure("Accent.TButton", background=t["accent"], foreground="#FFFFFF", bordercolor=t["accent"])
        style.map("Accent.TButton", background=[("active", t["accent"])])
        style.configure("TEntry", fieldbackground=t["entry_bg"], foreground=t["entry_fg"], bordercolor=t["border"])
        style.configure("TCombobox", fieldbackground=t["entry_bg"], foreground=t["entry_fg"], background=t["panel"])
        style.map("TCombobox", fieldbackground=[("readonly", t["entry_bg"])])
        style.configure("TSeparator", background=t["border"])
        style.configure("Treeview", background=t["tree_bg"], fieldbackground=t["tree_bg"], foreground=t["tree_fg"], bordercolor=t["border"], borderwidth=0)
        style.map("Treeview", background=[("selected", t["tree_sel"])], foreground=[("selected", t["fg"])])
        style.configure("Treeview.Heading", background=t["panel"], foreground=t["fg"], bordercolor=t["border"])
        style.map("Treeview.Heading", background=[("active", t["tree_sel"])])
        self.plc_list.tag_configure("connected", foreground=t["success"])
        self.plc_list.tag_configure("disconnected", foreground=t["muted"])

    def reload(self):
        self.cfg = load_config(self.cf.get())
        self.load()

    def load(self):
        try:
            self.objects = group(load_engineering(self.eng.get()))
            self.comments = load_general_comments(self.eng.get())
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
        ttk.Button(f, text="Close", command=w.destroy).grid(row=3, column=2, pady=(20, 0), sticky="e")

    def populate_sidebar(self):
        for i in self.plc_list.get_children():
            self.plc_list.delete(i)
        q = "" if getattr(self, "_search_placeholder_active", False) else self.search_var.get().strip().lower()
        for plc in sorted(self.byplc):
            if q and q not in plc.lower():
                continue
            total = len(self.byplc[plc])
            tested = sum(bool(self.comments.get(key(o))) for o in self.byplc[plc])
            connected = plc in self.conns
            tag = "connected" if connected else "disconnected"
            self.plc_list.insert("", "end", iid=plc, text=plc, values=(f"{tested}/{total}", "●"), tags=(tag,))
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
            return [o for o in x if not self.comments.get(key(o))]
        if fl in ("Correct", "Incorrect", "Recheck"):
            return [o for o in x if (self.comments.get(key(o)) or {}).get("status") == fl]
        return x

    def render_objects(self):
        if not self.current_plc:
            return
        for i in self.obj_list.get_children():
            self.obj_list.delete(i)
        x = self.filtered_objects()
        self._objmap = {}
        for i, o in enumerate(x):
            r = self.comments.get(key(o))
            result = r["status"] if r else "Not Tested"
            iid = f"row{i}"
            self._objmap[iid] = o
            self.obj_list.insert("", "end", iid=iid, text=o["Equipment"], values=(o["Template"], o["Area"], result))
        self.pvar.set(f"{len(x)} object(s)")
        total = len(self.byplc[self.current_plc])
        tested = sum(bool(self.comments.get(key(o))) for o in self.byplc[self.current_plc])
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
                    self.after(0, self.populate_sidebar)
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
        self.populate_sidebar()

    def update_progress(self):
        n = len(self.objects)
        rec = [self.comments.get(key(o)) for o in self.objects]
        t = sum(bool(x) for x in rec)
        self.progress.set(f"Overall: {n} Objects | {t} Tested | {n-t} Remaining")

    def close(self):
        self.disconnect()
        self.a.stop()
        self.destroy()

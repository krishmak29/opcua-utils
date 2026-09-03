import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime
import sqlite3, threading, asyncio
from openpyxl import load_workbook
from asyncua import Client
from datetime import datetime

ENGINEERING = "engineering_database.xlsx"
CONFIG = "utility_config.xlsx"
DB_FILE = "verification.db"

THEMES = {
    "light": {
        "bg": "#F5F6F8", "panel": "#FFFFFF", "fg": "#1F2430", "muted": "#6B7280",
        "accent": "#2563EB", "border": "#E2E4E9",
        "tree_bg": "#FFFFFF", "tree_fg": "#1F2430", "tree_sel": "#DCE8FF",
        "entry_bg": "#FFFFFF", "entry_fg": "#1F2430",
    },
    "dark": {
        "bg": "#14161B", "panel": "#1C1F26", "fg": "#E5E7EB", "muted": "#9CA3AF",
        "accent": "#5B8DEF", "border": "#2A2E37",
        "tree_bg": "#1C1F26", "tree_fg": "#E5E7EB", "tree_sel": "#2D3B55",
        "entry_bg": "#20232B", "entry_fg": "#E5E7EB",
    },
}

def s(v): return "" if v is None else str(v).strip()
def rows(wb, sheet):
    if sheet not in wb.sheetnames: return []
    ws=wb[sheet]; h=[s(c.value) for c in ws[1]]
    return [{h[i]: r[i] if i<len(r) else None for i in range(len(h)) if h[i]}
            for r in ws.iter_rows(min_row=2, values_only=True) if any(x is not None for x in r)]

def load_engineering(path):
    wb=load_workbook(path,data_only=True)
    sheet="Objects" if "Objects" in wb.sheetnames else wb.sheetnames[0]
    data=rows(wb,sheet)
    required=["PLC","Area Code*","Equipment Number*","Template","Tag Name","Address","Data Type"]
    if not data: raise ValueError("No engineering data found.")
    missing=[x for x in required if x not in data[0]]
    if missing: raise ValueError("Missing columns: "+", ".join(missing))
    return data

def load_config(path):
    if not Path(path).exists():
        return {"plc":{}, "display":{"ObjectsPerPage":12,"RefreshRateMs":1000,"MainValueSuffix":"RD.PV"},
                "causes":[("C001","Wrong PLC value"),("C002","Wrong SCADA value"),("C003","Communication failure"),
                          ("C004","Wrong scaling"),("C005","Wrong engineering unit"),("C006","Wrong alarm status"),
                          ("C007","Wrong tag mapping"),("C008","Object not available"),("C009","PLC not available"),("C010","Other")]}
    wb=load_workbook(path,data_only=True)
    pc={s(r.get("PLC")):r for r in rows(wb,"PLC_Config") if s(r.get("PLC"))}
    d={"ObjectsPerPage":12,"RefreshRateMs":1000,"MainValueSuffix":"RD.PV"}
    for r in rows(wb,"Display_Config"):
        if s(r.get("Parameter")): d[s(r["Parameter"])]=r.get("Value")
    c=[(s(r.get("Cause Code")),s(r.get("Cause Description"))) for r in rows(wb,"Cause_Master")
       if s(r.get("Cause Code")) and s(r.get("Cause Description"))]
    return {"plc":pc,"display":d,"causes":c or load_config.__defaults__}

def group(data):
    out={}
    for r in data:
        plc=s(r.get("PLC")); area=s(r.get("Area Code*")); eq=s(r.get("Equipment Number*"))
        if not plc or not eq: continue
        k=(plc,area,eq)
        out.setdefault(k,{"PLC":plc,"Area":area,"Equipment":eq,"Template":s(r.get("Template")),"Rows":[]})["Rows"].append(r)
    return list(out.values())

class DB:
    def __init__(self):
        self.c=sqlite3.connect(DB_FILE,check_same_thread=False)
        self.c.execute("""CREATE TABLE IF NOT EXISTS verification
        (k TEXT PRIMARY KEY, plc TEXT, area TEXT, equipment TEXT, template TEXT,
         result TEXT, cause TEXT, comment TEXT, tester TEXT, timestamp TEXT)"""); self.c.commit()
        self.lock=threading.Lock()
    def get(self,k):
        with self.lock:
            r=self.c.execute("SELECT result,cause,comment,tester,timestamp FROM verification WHERE k=?",(k,)).fetchone()
        return r
    def save(self,o,result,cause,comment,tester):
        k=key(o); now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self.lock:
            self.c.execute("""INSERT INTO verification VALUES(?,?,?,?,?,?,?,?,?,?)
            ON CONFLICT(k) DO UPDATE SET result=excluded.result,cause=excluded.cause,
            comment=excluded.comment,tester=excluded.tester,timestamp=excluded.timestamp""",
            (k,o["PLC"],o["Area"],o["Equipment"],o["Template"],result,cause,comment,tester,now)); self.c.commit()

def key(o): return "|".join([o["PLC"],o["Area"],o["Equipment"]])

class Async:
    def __init__(self):
        self.loop=asyncio.new_event_loop()
        threading.Thread(target=self.run,daemon=True).start()
    def run(self):
        asyncio.set_event_loop(self.loop); self.loop.run_forever()
    def call(self,coro):
        return asyncio.run_coroutine_threadsafe(coro,self.loop).result(15)
    def stop(self): self.loop.call_soon_threadsafe(self.loop.stop)

class Conn:
    def __init__(self,cfg): self.cfg=cfg; self.client=None
    async def connect(self):
        ip=s(self.cfg.get("IP")); port=int(self.cfg.get("Port") or 4840); ep=s(self.cfg.get("EndpointPath"))
        if ep and not ep.startswith("/"): ep="/"+ep
        self.client=Client(f"opc.tcp://{ip}:{port}{ep}")
        if s(self.cfg.get("Username")): self.client.set_user(s(self.cfg["Username"])); self.client.set_password(s(self.cfg.get("Password")))
        await self.client.connect()
    async def read(self,node): return await self.client.get_node(node).read_value()
    async def write(self,node,value):
        n=self.client.get_node(node)
        dv=await n.read_data_value()
        from asyncua import ua
        await n.write_value(ua.DataValue(ua.Variant(value,dv.Value.VariantType)))
    async def disconnect(self):
        if self.client: await self.client.disconnect()

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title("PLC / SCADA Object Verification Utility"); self.geometry("1250x780")
        self.cfg=load_config(CONFIG); self.db=DB(); self.a=Async(); self.objects=[]; self.byplc={}; self.conns={}
        self.theme="light"
        self.build(); self.apply_theme(); self.load()
        self.protocol("WM_DELETE_WINDOW",self.close)
    def build(self):
        self.eng=tk.StringVar(value=ENGINEERING)
        self.cf=tk.StringVar(value=CONFIG)
        top=ttk.Frame(self,padding=8); top.pack(fill="x")
        ttk.Label(top,text="PLC / SCADA Object Verification Utility",font=("Segoe UI",12,"bold")).pack(side="left")
        self.theme_btn=ttk.Button(top,text="🌙 Dark Mode",command=self.toggle_theme); self.theme_btn.pack(side="right",padx=(0,8))
        ttk.Button(top,text="⚙ Settings",command=self.open_settings).pack(side="right")
        self.status=tk.StringVar(value="Ready"); ttk.Label(self,textvariable=self.status,padding=(8,0)).pack(fill="x")

        body=ttk.Frame(self); body.pack(fill="both",expand=True,padx=8,pady=8)

        side=ttk.Frame(body,width=220); side.pack(side="left",fill="y"); side.pack_propagate(False)
        self.search_var=tk.StringVar(); self.search_var.trace_add("write",lambda *a:self.populate_sidebar())
        ttk.Entry(side,textvariable=self.search_var).pack(fill="x",pady=(0,5))
        ttk.Label(side,text="Search PLC...",foreground="gray").place(in_=side,x=6,y=4) if False else None
        self.plc_list=ttk.Treeview(side,columns=("count",),show="tree headings",height=25)
        self.plc_list.heading("#0",text="PLC"); self.plc_list.heading("count",text="Tested")
        self.plc_list.column("#0",width=140); self.plc_list.column("count",width=70,anchor="e")
        self.plc_list.pack(fill="both",expand=True)
        self.plc_list.bind("<<TreeviewSelect>>",self.on_plc_select)

        main=ttk.Frame(body); main.pack(side="left",fill="both",expand=True,padx=(10,0))
        head=ttk.Frame(main); head.pack(fill="x")
        self.info=ttk.Label(head,font=("Segoe UI",10,"bold")); self.info.pack(side="left")
        self.filter=tk.StringVar(value="All")
        ttk.Label(head,text=" Show:").pack(side="left",padx=(20,3))
        cb=ttk.Combobox(head,textvariable=self.filter,values=("All","Not Tested","Correct","Incorrect","Recheck"),state="readonly",width=14)
        cb.pack(side="left"); cb.bind("<<ComboboxSelected>>",lambda e:self.render_objects())

        list_frame=ttk.Frame(main); list_frame.pack(fill="both",expand=True,pady=8)
        cols=("template","area","status")
        self.obj_list=ttk.Treeview(list_frame,columns=cols,show="tree headings")
        self.obj_list.heading("#0",text="Equipment"); self.obj_list.heading("template",text="Template")
        self.obj_list.heading("area",text="Area"); self.obj_list.heading("status",text="Status")
        self.obj_list.column("#0",width=160); self.obj_list.column("template",width=220)
        self.obj_list.column("area",width=150); self.obj_list.column("status",width=120)
        self.obj_list.pack(side="left",fill="both",expand=True)
        obj_scroll=ttk.Scrollbar(list_frame,orient="vertical",command=self.obj_list.yview)
        obj_scroll.pack(side="right",fill="y")
        self.obj_list.configure(yscrollcommand=obj_scroll.set)
        self.obj_list.bind("<Double-1>",lambda e:self.open_selected())

        self.pvar=tk.StringVar()

        self.progress=tk.StringVar(); ttk.Label(self,textvariable=self.progress,padding=8).pack(fill="x")

        self.current_plc=None; self.page_no=0; self._objmap={}

    def colors(self):
        return THEMES[self.theme]

    def toggle_theme(self):
        self.theme = "dark" if self.theme=="light" else "light"
        self.apply_theme()

    def apply_theme(self):
        t=self.colors()
        self.configure(bg=t["bg"])
        self.theme_btn.config(text="☀ Light Mode" if self.theme=="dark" else "🌙 Dark Mode")
        style=ttk.Style(self)
        style.theme_use("clam")
        style.configure(".",background=t["bg"],foreground=t["fg"],fieldbackground=t["entry_bg"])
        style.configure("TFrame",background=t["bg"])
        style.configure("TLabel",background=t["bg"],foreground=t["fg"])
        style.configure("TButton",background=t["panel"],foreground=t["fg"],bordercolor=t["border"],focuscolor=t["accent"])
        style.map("TButton",background=[("active",t["tree_sel"])])
        style.configure("TEntry",fieldbackground=t["entry_bg"],foreground=t["entry_fg"],bordercolor=t["border"])
        style.configure("TCombobox",fieldbackground=t["entry_bg"],foreground=t["entry_fg"],background=t["panel"])
        style.map("TCombobox",fieldbackground=[("readonly",t["entry_bg"])])
        style.configure("TSeparator",background=t["border"])
        style.configure("Treeview",background=t["tree_bg"],fieldbackground=t["tree_bg"],foreground=t["tree_fg"],bordercolor=t["border"],borderwidth=0)
        style.map("Treeview",background=[("selected",t["tree_sel"])],foreground=[("selected",t["fg"])])
        style.configure("Treeview.Heading",background=t["panel"],foreground=t["fg"],bordercolor=t["border"])
        style.map("Treeview.Heading",background=[("active",t["tree_sel"])])

    def reload(self):
        self.cfg=load_config(self.cf.get()); self.load()
    def load(self):
        try:
            self.objects=group(load_engineering(self.eng.get())); self.byplc={}
            for o in self.objects:self.byplc.setdefault(o["PLC"],[]).append(o)
            self.current_plc=None
            self.populate_sidebar(); self.render_objects()
            self.status.set(f"Loaded {len(self.objects)} objects")
            self.update_progress()
        except Exception as e: messagebox.showerror("Load error",str(e))
    def open_settings(self):
        w=tk.Toplevel(self); w.title("Settings"); w.geometry("520x300"); w.transient(self)
        w.configure(bg=self.colors()["bg"])
        f=ttk.Frame(w,padding=12); f.pack(fill="both",expand=True)
        ttk.Label(f,text="Engineering Excel:").grid(row=0,column=0,sticky="w")
        ttk.Entry(f,textvariable=self.eng,width=45).grid(row=0,column=1,padx=5,pady=4)
        ttk.Button(f,text="Load",command=self.load).grid(row=0,column=2)
        ttk.Label(f,text="Config File:").grid(row=1,column=0,sticky="w")
        ttk.Entry(f,textvariable=self.cf,width=45).grid(row=1,column=1,padx=5,pady=4)
        ttk.Button(f,text="Reload Config",command=self.reload).grid(row=1,column=2)
        ttk.Separator(f).grid(row=2,column=0,columnspan=3,sticky="ew",pady=12)
        ttk.Label(f,text="Connection",font=("Segoe UI",10,"bold")).grid(row=3,column=0,sticky="w")
        ttk.Button(f,text="Connect All",command=self.connect).grid(row=4,column=0,pady=6,sticky="w")
        ttk.Button(f,text="Disconnect",command=self.disconnect).grid(row=4,column=1,pady=6,sticky="w")
        ttk.Button(f,text="Close",command=w.destroy).grid(row=5,column=2,pady=(20,0),sticky="e")

    def populate_sidebar(self):
        for i in self.plc_list.get_children(): self.plc_list.delete(i)
        q=self.search_var.get().strip().lower()
        for plc in sorted(self.byplc):
            if q and q not in plc.lower(): continue
            total=len(self.byplc[plc]); tested=sum(bool(self.db.get(key(o))) for o in self.byplc[plc])
            self.plc_list.insert("", "end", iid=plc, text=plc, values=(f"{tested}/{total}",))
        kids=self.plc_list.get_children()
        if kids and not self.current_plc: self.plc_list.selection_set(kids[0])

    def on_plc_select(self,e):
        sel=self.plc_list.selection()
        if not sel: return
        self.current_plc=sel[0]; self.render_objects()

    def filtered_objects(self):
        x=self.byplc.get(self.current_plc,[]); fl=self.filter.get()
        if fl=="Not Tested": return [o for o in x if not self.db.get(key(o))]
        if fl in ("Correct","Incorrect","Recheck"): return [o for o in x if (self.db.get(key(o)) or [None])[0]==fl]
        return x

    def render_objects(self):
        if not self.current_plc: return
        for i in self.obj_list.get_children(): self.obj_list.delete(i)
        x=self.filtered_objects()
        self._objmap={}
        for i,o in enumerate(x):
            r=self.db.get(key(o)); result=r[0] if r else "Not Tested"
            iid=f"row{i}"; self._objmap[iid]=o
            self.obj_list.insert("", "end", iid=iid, text=o["Equipment"],
                values=(o["Template"],o["Area"],result))
        self.pvar.set(f"{len(x)} object(s)")
        total=len(self.byplc[self.current_plc]); tested=sum(bool(self.db.get(key(o))) for o in self.byplc[self.current_plc])
        self.info.config(text=f"{self.current_plc} | Total: {total} | Tested: {tested} | Remaining: {total-tested}")

    def open_selected(self):
        sel=self.obj_list.selection()
        if sel and sel[0] in self._objmap: self.open(self._objmap[sel[0]])

    def edit_verification_status(self, event, tr, mapping):
        region = tr.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = tr.identify_row(event.y)
        column_id = tr.identify_column(event.x)
        # Verification Status = column 8
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
        combo.place(
            x=x,
            y=y,
            width=width,
            height=height
        )
        combo.focus_set()
        def save_status(event=None):
            new_value = combo.get()
            # PC local date/time
            verification_time = datetime.now().strftime(
                "%d-%m-%Y %H:%M:%S"
            )
            # Update Verification Status
            tr.set(
                row_id,
                "#7",
                new_value
            )
            # Update Verification Time
            tr.set(
                row_id,
                "#8",
                verification_time
            )
            # Update underlying mapping
            if row_id in mapping:
                mapping[row_id]["Verification Status"] = new_value
                mapping[row_id]["Verification Time"] = verification_time
            combo.destroy()
        combo.bind(
            "<<ComboboxSelected>>",
            save_status
        )
        combo.bind(
            "<Return>",
            save_status
        )
        combo.bind(
            "<Escape>",
            lambda e: combo.destroy()
        )
    
    def open(self,o):
        t=self.colors()
        w=tk.Toplevel(self); w.title(o["Equipment"]+" - Verification"); w.geometry("1050x680")
        w.configure(bg=t["bg"])
        h=ttk.Frame(w,padding=10); h.pack(fill="x"); ttk.Label(h,text=o["Equipment"],font=("Segoe UI",16,"bold")).pack(anchor="w")
        ttk.Label(h,text=f"Template: {o['Template']}").pack(anchor="w"); ttk.Label(h,text=f"PLC: {o['PLC']} | Area: {o['Area']}").pack(anchor="w")
        cols=("tag","value","address","type","access","unit","verify","verifyTS"); tr=ttk.Treeview(w,columns=cols,show="headings")
        for c,ttext,wd in [("tag","Tag Name",230),("value","PLC / OPC Value",130),("address","Address",110),("type","Data Type",90),("access","Access",70),("unit","Eng Units",80),("verify","Verify Sts",80),("verifyTS","Verify TS",80)]:
            tr.heading(c,text=ttext); tr.column(c,width=wd)
        tr.pack(fill="both",expand=True,padx=10)
        mapping={}
        not_conn = o["PLC"] not in self.conns
        init_val = "Not Connected" if not_conn else "—"
        for i,r in enumerate(o["Rows"]):
            iid=str(i); mapping[iid]=r; tr.insert("", "end",iid=iid,values=(s(r.get("Tag Name")),init_val,s(r.get("Address")) or "",s(r.get("Data Type")),s(r.get("Client Access")) or "R",s(r.get("Eng Units")),s(r.get("Verification Status") or "Not Tested"),s(r.get("Verification TS")))) 
        tr.bind("<Double-1>", lambda event: self.edit_verification_status(event, tr,mapping))
        ctl=ttk.Frame(w,padding=10); ctl.pack(fill="x")
        result=tk.StringVar(value=(self.db.get(key(o)) or ["Correct"])[0])
        ttk.Label(ctl,text="Verification:").grid(row=0,column=0,sticky="w")
        for j,x in enumerate(("Correct","Incorrect","Recheck"),1): ttk.Radiobutton(ctl,text=x,variable=result,value=x).grid(row=0,column=j,sticky="w")
        ttk.Label(ctl,text="Cause:").grid(row=1,column=0,sticky="w",pady=5)
        cause=tk.StringVar(); cb=ttk.Combobox(ctl,textvariable=cause,values=[f"{a} - {b}" for a,b in self.cfg["causes"]],width=55); cb.grid(row=1,column=1,columnspan=3,sticky="w")
        ttk.Label(ctl,text="Tester:").grid(row=2,column=0,sticky="w"); tester=tk.Entry(ctl,width=30,bg=t["entry_bg"],fg=t["entry_fg"],insertbackground=t["fg"],relief="flat"); tester.grid(row=2,column=1,sticky="w")
        ttk.Label(ctl,text="Comment:").grid(row=3,column=0,sticky="nw"); comment=tk.Text(ctl,width=80,height=4,bg=t["entry_bg"],fg=t["entry_fg"],insertbackground=t["fg"],relief="flat"); comment.grid(row=3,column=1,columnspan=3)
        old=self.db.get(key(o))
        if old: 
            if old[1]: cause.set(old[1])
            tester.insert(0,old[3] or ""); comment.insert("1.0",old[2] or "")
        ttk.Button(w,text="Read All Parameters",command=lambda:self.read(o,tr,mapping)).pack(side="left",padx=10,pady=8)

        def save():
            self.db.save(o,result.get(),cause.get(),comment.get("1.0","end").strip(),tester.get().strip()); w.destroy(); self.load()
        ttk.Button(w,text="SAVE VERIFICATION",command=save).pack(side="right",padx=10,pady=8)
        self.read(o,tr,mapping)
    def read(self,o,tr,mapping):
        conn=self.conns.get(o["PLC"])
        if not conn:
            for iid in mapping: self.setv(tr,iid,"Not Connected")
            return
        def work():
            for iid,r in mapping.items():
                a=s(r.get("Address"))
                if not a: continue
                try: v=self.a.call(conn.read(a)); self.after(0,lambda iid=iid,v=v:self.setv(tr,iid,v))
                except Exception as e: self.after(0,lambda iid=iid,e=e:self.setv(tr,iid,"ERROR: "+str(e)))
        threading.Thread(target=work,daemon=True).start()
    def setv(self,tr,iid,v):
        x=list(tr.item(iid,"values")); x[1]=str(v); tr.item(iid,values=x)
    def connect(self):
        def work():
            for plc,cfg in self.cfg["plc"].items():
                try: c=Conn(cfg); self.a.call(c.connect()); self.conns[plc]=c
                except Exception as e: self.status.set(f"{plc}: {e}")
            self.status.set(f"Connected {len(self.conns)} PLC(s)")
        threading.Thread(target=work,daemon=True).start()
    def disconnect(self):
        for c in self.conns.values():
            try:self.a.call(c.disconnect())
            except:pass
        self.conns={}; self.status.set("Disconnected")
    def update_progress(self):
        n=len(self.objects); rec=[self.db.get(key(o)) for o in self.objects]; t=sum(bool(x) for x in rec)
        self.progress.set(f"Overall: {n} Objects | {t} Tested | {n-t} Remaining")
    def close(self):
        self.disconnect(); self.db.c.close(); self.a.stop(); self.destroy()

if __name__=="__main__": App().mainloop()
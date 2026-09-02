import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from datetime import datetime
import sqlite3, threading, asyncio
from openpyxl import load_workbook
from asyncua import Client

ENGINEERING = "engineering_database.xlsx"
CONFIG = "utility_config.xlsx"
DB_FILE = "verification.db"

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
    required=["PLC","Area Code*","Equipment Number*","Template","Tag Name","Address*","Data Type"]
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
        self.cfg=load_config(CONFIG); self.db=DB(); self.a=Async(); self.objects=[]; self.byplc={}; self.pages={}; self.conns={}
        self.build(); self.load()
        self.protocol("WM_DELETE_WINDOW",self.close)
    def build(self):
        top=ttk.Frame(self,padding=8); top.pack(fill="x")
        ttk.Label(top,text="Engineering Excel:").pack(side="left")
        self.eng=tk.StringVar(value=ENGINEERING); ttk.Entry(top,textvariable=self.eng,width=50).pack(side="left",padx=5)
        ttk.Button(top,text="Load",command=self.load).pack(side="left")
        ttk.Label(top,text="Config:").pack(side="left",padx=(15,3))
        self.cf=tk.StringVar(value=CONFIG); ttk.Entry(top,textvariable=self.cf,width=30).pack(side="left",padx=5)
        ttk.Button(top,text="Reload Config",command=self.reload).pack(side="left")
        ttk.Button(top,text="Connect All",command=self.connect).pack(side="left",padx=8)
        ttk.Button(top,text="Disconnect",command=self.disconnect).pack(side="left")
        self.status=tk.StringVar(value="Ready"); ttk.Label(self,textvariable=self.status,padding=8).pack(fill="x")
        self.nb=ttk.Notebook(self); self.nb.pack(fill="both",expand=True,padx=8)
        self.progress=tk.StringVar(); ttk.Label(self,textvariable=self.progress,padding=8).pack(fill="x")
    def reload(self):
        self.cfg=load_config(self.cf.get()); self.load()
    def load(self):
        try:
            self.objects=group(load_engineering(self.eng.get())); self.byplc={}
            for o in self.objects:self.byplc.setdefault(o["PLC"],[]).append(o)
            for t in self.nb.tabs(): self.nb.forget(t)
            self.pages={p:0 for p in self.byplc}
            for p in self.byplc:
                f=ttk.Frame(self.nb); self.nb.add(f,text=p); self.make_tab(f,p)
            self.status.set(f"Loaded {len(self.objects)} objects")
            self.update_progress()
        except Exception as e: messagebox.showerror("Load error",str(e))
    def make_tab(self,f,plc):
        h=ttk.Frame(f,padding=8); h.pack(fill="x")
        f.info=ttk.Label(h); f.info.pack(side="left")
        f.filter=tk.StringVar(value="All")
        ttk.Label(h,text=" Show:").pack(side="left",padx=(20,3))
        cb=ttk.Combobox(h,textvariable=f.filter,values=("All","Not Tested","Correct","Incorrect","Recheck"),state="readonly",width=14); cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>",lambda e:self.render(plc))
        f.content=ttk.Frame(f); f.content.pack(fill="both",expand=True)
        n=ttk.Frame(f,padding=8); n.pack(fill="x")
        f.prev=ttk.Button(n,text="< PREVIOUS",command=lambda:self.page(plc,-1)); f.prev.pack(side="left")
        f.pvar=tk.StringVar(); ttk.Label(n,textvariable=f.pvar).pack(side="left",expand=True)
        f.next=ttk.Button(n,text="NEXT >",command=lambda:self.page(plc,1)); f.next.pack(side="right")
        self.render(plc)
    def filtered(self,plc):
        x=self.byplc[plc]; fl=self.nb.nametowidget(self.nb.select()).filter.get()
        if fl=="Not Tested": return [o for o in x if not self.db.get(key(o))]
        if fl in ("Correct","Incorrect","Recheck"): return [o for o in x if (self.db.get(key(o)) or [None])[0]==fl]
        return x
    def render(self,plc):
        f=self.nb.nametowidget(self.nb.select()); x=self.filtered(plc); size=int(self.cfg["display"].get("ObjectsPerPage",12) or 12)
        pages=max(1,(len(x)+size-1)//size); p=min(self.pages.get(plc,0),pages-1); self.pages[plc]=p
        for w in f.content.winfo_children(): w.destroy()
        for i,o in enumerate(x[p*size:(p+1)*size]):
            card=self.card(f.content,o); card.grid(row=i//4,column=i%4,padx=8,pady=8,sticky="nsew")
        for c in range(4): f.content.columnconfigure(c,weight=1)
        f.pvar.set(f"Page {p+1} / {pages} | {len(x)} object(s)")
        f.prev.config(state="normal" if p else "disabled"); f.next.config(state="normal" if p<pages-1 else "disabled")
        total=len(self.byplc[plc]); tested=sum(bool(self.db.get(key(o))) for o in self.byplc[plc]); f.info.config(text=f"{plc} | Total: {total} | Tested: {tested} | Remaining: {total-tested}")
    def page(self,plc,d):
        self.pages[plc]=max(0,self.pages.get(plc,0)+d); self.render(plc)
    def card(self,parent,o):
        r=self.db.get(key(o)); result=r[0] if r else "Not Tested"; bg={"Correct":"#d9f2d9","Incorrect":"#f6d6d6","Recheck":"#ffe8b3"}.get(result,"#eeeeee")
        fr=tk.Frame(parent,bg=bg,bd=2,relief="groove",padx=10,pady=8)
        tk.Label(fr,text=o["Equipment"],font=("Segoe UI",12,"bold"),bg=bg).pack(anchor="w")
        tk.Label(fr,text=o["Template"],font=("Segoe UI",9),bg=bg).pack(anchor="w")
        tk.Label(fr,text=f"Area: {o['Area']}",bg=bg).pack(anchor="w")
        main=next((z for z in o["Rows"] if s(z.get("OPC Tag Suffix")).upper()==s(self.cfg["display"].get("MainValueSuffix","RD.PV")).upper()),o["Rows"][0])
        tk.Label(fr,text=s(main.get("OPC Tag Suffix")) or "PV",font=("Segoe UI",9,"bold"),bg=bg).pack(anchor="w",pady=(8,0))
        v=tk.StringVar(value="—"); tk.Label(fr,textvariable=v,font=("Segoe UI",18,"bold"),bg=bg).pack(anchor="w")
        tk.Label(fr,text=s(main.get("Eng Units")),bg=bg).pack(anchor="w")
        txt={"Correct":"✓ CORRECT","Incorrect":"✕ INCORRECT","Recheck":"↻ RECHECK"}.get(result,"NOT TESTED")
        tk.Label(fr,text=txt,font=("Segoe UI",10,"bold"),bg=bg).pack(anchor="w",pady=(8,0))
        for w in fr.winfo_children(): w.bind("<Button-1>",lambda e,o=o:self.open(o))
        fr.bind("<Button-1>",lambda e,o=o:self.open(o)); return fr
    def open(self,o):
        w=tk.Toplevel(self); w.title(o["Equipment"]+" - Verification"); w.geometry("1050x680")
        h=ttk.Frame(w,padding=10); h.pack(fill="x"); ttk.Label(h,text=o["Equipment"],font=("Segoe UI",16,"bold")).pack(anchor="w")
        ttk.Label(h,text=f"Template: {o['Template']}").pack(anchor="w"); ttk.Label(h,text=f"PLC: {o['PLC']} | Area: {o['Area']}").pack(anchor="w")
        cols=("parameter","tag","value","address","type","access","unit"); tr=ttk.Treeview(w,columns=cols,show="headings")
        for c,t,wd in [("parameter","Parameter",150),("tag","Tag Name",230),("value","PLC / OPC Value",130),("address","Address",110),("type","Data Type",90),("access","Access",70),("unit","Eng Units",80)]:
            tr.heading(c,text=t); tr.column(c,width=wd)
        tr.pack(fill="both",expand=True,padx=10)
        mapping={}
        not_conn = o["PLC"] not in self.conns
        init_val = "Not Connected" if not_conn else "—"
        for i,r in enumerate(o["Rows"]):
            iid=str(i); mapping[iid]=r; tr.insert("", "end",iid=iid,values=(s(r.get("OPC Tag Suffix")) or s(r.get("Tag Name")),s(r.get("Tag Name")),init_val,s(r.get("Address*")),s(r.get("Data Type")),s(r.get("Client Access")) or "R",s(r.get("Eng Units"))))
        ctl=ttk.Frame(w,padding=10); ctl.pack(fill="x")
        result=tk.StringVar(value=(self.db.get(key(o)) or ["Correct"])[0])
        ttk.Label(ctl,text="Verification:").grid(row=0,column=0,sticky="w")
        for j,x in enumerate(("Correct","Incorrect","Recheck"),1): ttk.Radiobutton(ctl,text=x,variable=result,value=x).grid(row=0,column=j,sticky="w")
        ttk.Label(ctl,text="Cause:").grid(row=1,column=0,sticky="w",pady=5)
        cause=tk.StringVar(); cb=ttk.Combobox(ctl,textvariable=cause,values=[f"{a} - {b}" for a,b in self.cfg["causes"]],width=55); cb.grid(row=1,column=1,columnspan=3,sticky="w")
        ttk.Label(ctl,text="Tester:").grid(row=2,column=0,sticky="w"); tester=tk.Entry(ctl,width=30); tester.grid(row=2,column=1,sticky="w")
        ttk.Label(ctl,text="Comment:").grid(row=3,column=0,sticky="nw"); comment=tk.Text(ctl,width=80,height=4); comment.grid(row=3,column=1,columnspan=3)
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
                a=s(r.get("Address*"))
                if not a: continue
                try: v=self.a.call(conn.read(a)); self.after(0,lambda iid=iid,v=v:self.setv(tr,iid,v))
                except Exception as e: self.after(0,lambda iid=iid,e=e:self.setv(tr,iid,"ERROR: "+str(e)))
        threading.Thread(target=work,daemon=True).start()
    def setv(self,tr,iid,v):
        x=list(tr.item(iid,"values")); x[2]=str(v); tr.item(iid,values=x)
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

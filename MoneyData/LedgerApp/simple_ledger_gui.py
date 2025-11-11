import csv
from datetime import date, datetime, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ----- charts -----
import matplotlib
matplotlib.use("Agg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import rcParams, font_manager

# ----- paths & defaults -----
DATA_DIR = Path(r"D:\MoneyData\LedgerApp"); DATA_DIR.mkdir(parents=True, exist_ok=True)
DEFAULT_CSV = DATA_DIR / "ledger.csv"
CATS = ["食物","娛樂","交通","服飾","日用品","醫療","房租","水電瓦斯","收入"]
ACCS = ["現金","玉山","台新","國泰","中信"]

# ----- small helpers -----
p = lambda s: datetime.strptime(s,"%Y-%m-%d").date()
def pr(base, mode):
    if mode=="daily": return base, base
    if mode=="weekly":
        st = base - timedelta(days=base.weekday()); return st, st+timedelta(days=6)
    st = base.replace(day=1); nxt = (st.replace(year=st.year+1,month=1,day=1) if st.month==12
                                     else st.replace(month=st.month+1,day=1))
    return st, nxt-timedelta(days=1)
def mr(ym):
    y,m = map(int, ym.split("-")); st = date(y,m,1)
    return pr(st,"monthly")
def setup_cjk_font():
    cand = ["Microsoft JhengHei","Microsoft YaHei","Noto Sans CJK TC","Noto Sans CJK SC",
            "PingFang TC","Arial Unicode MS","SimHei","PMingLiU"]
    avail = {f.name for f in font_manager.fontManager.ttflist}
    for n in cand:
        if n in avail:
            rcParams["font.sans-serif"]=[n]; rcParams["font.family"]="sans-serif"; rcParams["axes.unicode_minus"]=False
            break

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("簡易記帳本"); self.geometry("980x640")
        setup_cjk_font()
        # ttk styles（去外框）
        sty = ttk.Style(self)
        try: sty.theme_use("clam")
        except: pass
        self.bg = sty.lookup("TFrame","background") or self.cget("bg")
        sty.configure("Card.TFrame", background=self.bg)
        sty.configure("Plain.Treeview", borderwidth=0, relief="flat")
        sty.layout("Plain.Treeview",[("Treeview.treearea",{"sticky":"nswe"})])
        sty.configure("Plain.Treeview.Heading", borderwidth=0, relief="flat")
        sty.map("Plain.Treeview.Heading", relief=[("active","flat"),("pressed","flat")])

        self.tx = []  # transactions list of dict

        nb = ttk.Notebook(self); self.p_main=ttk.Frame(nb); self.p_dash=ttk.Frame(nb)
        nb.add(self.p_main,text="記帳"); nb.add(self.p_dash,text="Dashboard"); nb.pack(fill="both",expand=True)

        self.build_main(); self.build_dash()
        if DEFAULT_CSV.exists(): self.load_csv(DEFAULT_CSV)
        self.refresh_dash()

    # ---------- 記帳頁 ----------
    def build_main(self):
        f = ttk.Frame(self.p_main,padding=10); f.pack(fill="x")
        ttk.Label(f,text="日期YYYY-MM-DD").grid(row=0,column=0,sticky="w",padx=(0,6))
        self.v_date=tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        ttk.Entry(f,textvariable=self.v_date,width=12).grid(row=1,column=0,sticky="w",padx=(0,12))
        ttk.Label(f,text="類型").grid(row=0,column=1,sticky="w")
        self.v_type=tk.StringVar(value="expense")
        ttk.Combobox(f,textvariable=self.v_type,values=["income","expense"],state="readonly",width=8)\
            .grid(row=1,column=1,sticky="w",padx=(0,12))
        ttk.Label(f,text="金額").grid(row=0,column=2,sticky="w")
        self.v_amt=tk.StringVar(); ttk.Entry(f,textvariable=self.v_amt,width=10)\
            .grid(row=1,column=2,sticky="w",padx=(0,12))
        ttk.Label(f,text="類別").grid(row=0,column=3,sticky="w")
        self.v_cat=tk.StringVar(value=CATS[0]); self.cbo_cat=ttk.Combobox(f,textvariable=self.v_cat,values=CATS,width=10)
        self.cbo_cat.grid(row=1,column=3,sticky="w",padx=(0,12))
        ttk.Label(f,text="帳戶").grid(row=0,column=4,sticky="w")
        self.v_acc=tk.StringVar(value=ACCS[0]); self.cbo_acc=ttk.Combobox(f,textvariable=self.v_acc,values=ACCS,width=10)
        self.cbo_acc.grid(row=1,column=4,sticky="w",padx=(0,12))
        ttk.Label(f,text="說明").grid(row=0,column=5,sticky="w")
        self.v_desc=tk.StringVar(); ttk.Entry(f,textvariable=self.v_desc,width=30)\
            .grid(row=1,column=5,sticky="w",padx=(0,12))
        ttk.Button(f,text="新增",command=self.add_tx).grid(row=1,column=6,sticky="w")
        for i in range(7): f.grid_columnconfigure(i,weight=0)
        f.grid_columnconfigure(5,weight=1)

        cols=("Date","Type","Amount","Category","Account","Description")
        self.tree=ttk.Treeview(self.p_main,columns=cols,show="headings",height=18)
        for c,w in zip(cols,(110,80,100,110,110,420)):
            self.tree.heading(c,text=c); self.tree.column(c,width=w,anchor=("e" if c=="Amount" else "w"))
        self.tree.pack(fill="both",expand=True,padx=10,pady=(6,0))

        s=ttk.Frame(self.p_main,padding=10); s.pack(fill="x")
        self.lab_total=ttk.Label(s,text="收入:0 支出:0 淨額:0"); self.lab_total.pack(side="left")
        r=ttk.Frame(s); r.pack(side="right")
        ttk.Label(r,text="統計基準日").pack(side="left")
        self.v_base=tk.StringVar(value=datetime.today().strftime("%Y-%m-%d"))
        ttk.Entry(r,textvariable=self.v_base,width=12).pack(side="left",padx=6)
        self.v_mode=tk.StringVar(value="每月")
        ttk.Combobox(r,textvariable=self.v_mode,values=["每日","每週","每月"],state="readonly",width=6).pack(side="left",padx=4)
        ttk.Button(r,text="計算支出",command=self.calc_summary).pack(side="left",padx=8)

        b=ttk.Frame(self.p_main,padding=10); b.pack(fill="x"); btns=ttk.Frame(b); btns.pack(side="right")
        ttk.Button(btns,text="載入",command=self.load_csv).pack(side="left",padx=4)
        ttk.Button(btns,text="儲存",command=self.save_csv).pack(side="left",padx=4)
        ttk.Button(btns,text="另存",command=self.save_as).pack(side="left",padx=4)
        ttk.Button(btns,text="刪除選取",command=self.del_sel).pack(side="left",padx=4)
        ttk.Button(btns,text="清空列表",command=self.clear_all).pack(side="left",padx=4)

    # ---------- Dashboard ----------
    def build_dash(self):
        top=ttk.Frame(self.p_dash,padding=10,style="Card.TFrame"); top.pack(fill="x")
        ttk.Label(top,text="統計月份 YYYY-MM").pack(side="left")
        self.v_ym=tk.StringVar(value=datetime.today().strftime("%Y-%m"))
        ttk.Entry(top,textvariable=self.v_ym,width=8).pack(side="left",padx=6)
        ttk.Button(top,text="更新圖表",command=self.refresh_dash).pack(side="left",padx=8)

        body=ttk.Frame(self.p_dash,padding=10,style="Card.TFrame"); body.pack(fill="both",expand=True)
        self.chart_frame=ttk.Frame(body,style="Card.TFrame"); self.chart_frame.pack(side="left",fill="both",expand=True)
        right=ttk.Frame(body,style="Card.TFrame"); right.pack(side="right",fill="y")
        ttk.Label(right,text="帳戶餘額（累積）").pack(anchor="w")
        self.tree_acc=ttk.Treeview(right,style="Plain.Treeview",
            columns=("Account","Balance","MonthlyDelta"),show="headings",height=16)
        for c in ("Account","Balance","MonthlyDelta"):
            self.tree_acc.heading(c,text=c); self.tree_acc.column(c,width=120,anchor=("e" if c!="Account" else "w"))
        self.tree_acc.pack(fill="y",expand=False,pady=(6,0))
        ttk.Label(right,text="說明：Balance=收入−支出；MonthlyDelta=本月收入−支出").pack(anchor="w",pady=6)

    # ---------- data ops ----------
    def add_tx(self):
        try: d=p(self.v_date.get().strip())
        except: messagebox.showerror("錯誤","日期請用 YYYY-MM-DD"); return
        t=self.v_type.get().lower()
        try: a=float(self.v_amt.get().replace(",",""))
        except: messagebox.showerror("錯誤","金額需為數字"); return
        c=self.v_cat.get().strip(); acc=self.v_acc.get().strip(); desc=self.v_desc.get().strip()
        self.tx.append({"Date":d.isoformat(),"Type":t,"Amount":a,"Category":c,"Account":acc,"Description":desc})
        self.tree.insert("", "end", values=(d, t, f"{a:.2f}", c, acc, desc))
        if c and c not in self.cbo_cat["values"]: self.cbo_cat["values"]=(*self.cbo_cat["values"],c)
        if acc and acc not in self.cbo_acc["values"]: self.cbo_acc["values"]=(*self.cbo_acc["values"],acc)
        self.v_amt.set(""); self.v_desc.set(""); self.update_totals(); self.refresh_dash()

    def update_totals(self):
        inc=sum(t["Amount"] for t in self.tx if t["Type"]=="income")
        exp=sum(t["Amount"] for t in self.tx if t["Type"]=="expense")
        self.lab_total.config(text=f"收入:{inc:.2f}  支出:{exp:.2f}  淨額:{(inc-exp):.2f}")

    def save_csv(self, path:Path=None):
        path = path or DEFAULT_CSV
        with open(path,"w",newline="",encoding="utf-8-sig") as f:
            w=csv.DictWriter(f,fieldnames=["Date","Type","Amount","Category","Account","Description"])
            w.writeheader(); [w.writerow(t) for t in self.tx]
        messagebox.showinfo("已儲存",str(path))

    def save_as(self):
        fn=filedialog.asksaveasfilename(defaultextension=".csv",filetypes=[("CSV","*.csv")],
                                        initialdir=str(DATA_DIR),initialfile="ledger.csv")
        if fn: self.save_csv(Path(fn))

    def load_csv(self, path:Path=None):
        path = path or DEFAULT_CSV
        with open(path,"r",newline="",encoding="utf-8-sig") as f:
            r=csv.DictReader(f); self.tx=[{"Date":row.get("Date",""),"Type":row.get("Type",""),
                "Amount":float(row.get("Amount",0) or 0),"Category":row.get("Category",""),
                "Account":row.get("Account",""),"Description":row.get("Description","")} for row in r]
        for i in self.tree.get_children(): self.tree.delete(i)
        for t in self.tx: self.tree.insert("", "end", values=(t["Date"],t["Type"],f'{t["Amount"]:.2f}',t["Category"],t["Account"],t["Description"]))
        self.update_totals(); self.refresh_dash()
        messagebox.showinfo("已載入", f"{len(self.tx)} 筆，自 {path}")

    def del_sel(self):
        sel=self.tree.selection()
        for i in sorted([self.tree.index(x) for x in sel],reverse=True): del self.tx[i]
        for x in sel: self.tree.delete(x); self.update_totals(); self.refresh_dash()

    def clear_all(self):
        if not messagebox.askyesno("確認","清空畫面列表？(不刪硬碟CSV)"): return
        self.tx.clear(); [self.tree.delete(i) for i in self.tree.get_children()]
        self.update_totals(); self.refresh_dash()

    # ---------- summary ----------
    def calc_summary(self):
        try:
            base = p(self.v_base.get().strip())
        except Exception:
            messagebox.showerror("錯誤", "日期請用 YYYY-MM-DD")
            return

        mode = {"每日": "daily", "每週": "weekly", "每月": "monthly"}[self.v_mode.get()]
        st, ed = pr(base, mode)

        ex = [t for t in self.tx if t["Type"] == "expense" and t["Date"] and st <= p(t["Date"]) <= ed]
        tot = sum(t["Amount"] for t in ex)

        by = {}
        for t in ex:
            k = t["Category"] or "未分類"
            by[k] = by.get(k, 0.0) + t["Amount"]

        # 先把類別明細整理好；如果沒有資料，用預設訊息
        rows = [f"・{k}: {v:.2f}" for k, v in sorted(by.items(), key=lambda x: -x[1])]
        if not rows:
            rows = ["(本期無支出)"]

        lines = [
            f"區間：{st} ~ {ed}",
            f"支出合計：{tot:.2f}",
            "",
            *rows,  # 這裡才展開
        ]
        messagebox.showinfo("統計（支出）", "\n".join(lines))
    # ---------- dashboard ----------
    def refresh_dash(self):
        for w in self.chart_frame.winfo_children(): w.destroy()
        try: st,ed=mr(self.v_ym.get().strip())
        except: st,ed=pr(date.today(),"monthly")
        by={}; 
        for t in self.tx:
            if t["Type"]!="expense" or not t["Date"]: continue
            d=p(t["Date"]); 
            if st<=d<=ed: by.setdefault(t["Category"] or "未分類",0); by[t["Category"] or "未分類"]+=t["Amount"]
        if by:
            fig=Figure(figsize=(5.8,4.4),dpi=100,facecolor=self.bg); ax=fig.add_subplot(111)
            ax.set_facecolor(self.bg); ax.set_frame_on(False); [sp.set_visible(False) for sp in ax.spines.values()]
            labels,vals=list(by.keys()),list(by.values()); ax.pie(vals,labels=labels,autopct="%1.1f%%",startangle=90)
            ax.set_title(f"{st.strftime('%Y-%m')} 支出類別分佈"); fig.tight_layout()
            can=FigureCanvasTkAgg(fig,master=self.chart_frame); can.draw()
            w=can.get_tk_widget(); w.configure(background=self.bg,highlightthickness=0,bd=0); w.pack(fill="both",expand=True)
        else:
            ttk.Label(self.chart_frame,text=f"{st.strftime('%Y-%m')} 無支出資料").pack(pady=20)

        for i in self.tree_acc.get_children(): self.tree_acc.delete(i)
        bal,delta={},{}; mstart,mend=st,ed
        for t in self.tx:
            acc=(t["Account"] or "未指定").strip(); s=1 if t["Type"]=="income" else -1; a=float(t["Amount"])
            bal[acc]=bal.get(acc,0.0)+s*a
            if t["Date"] and mstart<=p(t["Date"])<=mend: delta[acc]=delta.get(acc,0.0)+s*a
        for acc,b in sorted(bal.items(),key=lambda x:-x[1]): self.tree_acc.insert("", "end", values=(acc,f"{b:.2f}",f"{delta.get(acc,0.0):.2f}"))

if __name__=="__main__": App().mainloop()

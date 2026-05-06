"""
FBC Suite — Combined Desktop App
──────────────────────────────────
  📊  Sarestock Upload Converter   (Tab 1)
  ✉   Deal Note Email Automator    (Tab 2)

Requirements:
    pip install pandas openpyxl fpdf2 pywin32 pymupdf
"""

# ════════════════════════════════════════════════════════════════════════════
#  AUTO-UPDATE
# ════════════════════════════════════════════════════════════════════════════
import sys, os, subprocess, urllib.request

VERSION       = 7
GITHUB_USER   = "Anashe-Masomeke"
GITHUB_REPO   = "fbc-suite"
GITHUB_BRANCH = "main"
EXE_NAME      = "fbc-suite.exe"

_EXE = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest/download/{EXE_NAME}"
_VER = (f"https://raw.githubusercontent.com/"
        f"{GITHUB_USER}/{GITHUB_REPO}/{GITHUB_BRANCH}/version.txt")

def _remote_ver():
    try:
        with urllib.request.urlopen(_VER, timeout=4) as r:
            return int(r.read().decode().strip())
    except Exception:
        return -1

def check_and_apply_update():
    rv = _remote_ver()
    if rv <= VERSION:
        return
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk(); root.withdraw()
    ok = messagebox.askyesno(
        "FBC Suite — Update Available",
        f"New version available  (v{rv}).\nYour version: v{VERSION}\n\nDownload and restart now?",
        icon="info")
    root.destroy()
    if not ok:
        return

    current_exe = os.path.abspath(sys.argv[0])
    exe_dir     = os.path.dirname(current_exe)
    exe_name    = os.path.basename(current_exe)
    new_exe_tmp = os.path.join(exe_dir, "_fbc_update_new.exe")
    old_exe_bak = os.path.join(exe_dir, "_fbc_update_old.exe")
    bat_path    = os.path.join(exe_dir, "_fbc_updater.bat")

    root2 = tk.Tk(); root2.withdraw()
    messagebox.showinfo("Downloading Update",
        "Downloading update — please wait.\n\nThe app will restart automatically.",
        parent=root2)
    root2.destroy()

    try:
        urllib.request.urlretrieve(_EXE, new_exe_tmp)

        bat_lines = [
            "@echo off",
            "ping 127.0.0.1 -n 7 > nul",
            f'taskkill /F /IM "{exe_name}" >nul 2>&1',
            "ping 127.0.0.1 -n 4 > nul",
            f'move /Y "{current_exe}" "{old_exe_bak}"',
            f'move /Y "{new_exe_tmp}" "{current_exe}"',
            f'start "" "{current_exe}"',
            "ping 127.0.0.1 -n 3 > nul",
            f'del "{old_exe_bak}"',
            'del "%~f0"',
        ]
        with open(bat_path, "w") as f:
            f.write("\n".join(bat_lines) + "\n")

        subprocess.Popen(
            ["cmd.exe", "/c", bat_path],
            creationflags=subprocess.CREATE_NO_WINDOW,
            close_fds=True
        )
        sys.exit(0)

    except Exception as e:
        for fp in [new_exe_tmp, bat_path]:
            try: os.remove(fp)
            except Exception: pass
        root3 = tk.Tk(); root3.withdraw()
        messagebox.showerror("Update Failed",
            f"Could not download update:\n\n{e}\n\n"
            "Please download manually from:\n"
            f"github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/latest")
        root3.destroy()


# ════════════════════════════════════════════════════════════════════════════
#  IMPORTS
# ════════════════════════════════════════════════════════════════════════════
import re, json, csv, threading
import shutil as _shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

def _require(pkg, install_name=None):
    import importlib
    try:
        return importlib.import_module(pkg)
    except ImportError:
        name = install_name or pkg
        raise ImportError(
            f"Missing package '{name}'. Open terminal and run:\n  pip install {name}")

# ════════════════════════════════════════════════════════════════════════════
#  SHARED COLOURS
# ════════════════════════════════════════════════════════════════════════════
FBC_DARK   = "#003B6F"
FBC_MID    = "#0066B3"
FBC_ACCENT = "#00A3E0"
GREEN_DARK = "#1A6B3A"
RED_DARK   = "#B71C1C"
WHITE      = "#FFFFFF"
BG         = "#F0F4F8"
CARD_BG    = "#FFFFFF"
SEP_CLR    = "#D0DAE8"
TAG_BLUE   = "#E8F1FB"
COL1_HDR   = "#003B6F"
COL2_HDR   = "#1A3A6B"
BOTTOM     = "#0D2B4E"

SIDEBAR_BG      = "#001F3F"
SIDEBAR_ACTIVE  = "#0066B3"
SIDEBAR_HOVER   = "#003B6F"
SIDEBAR_TEXT    = "#B0C8E8"
SIDEBAR_TEXT_ON = "#FFFFFF"

# ════════════════════════════════════════════════════════════════════════════
#  ── SARESTOCK LOGIC ────────────────────────────────────────────────────────
# ════════════════════════════════════════════════════════════════════════════
STATE_FILE = os.path.join(os.path.expanduser("~"), ".fbc_ticket_state.json")

EO_HEADERS = [
    "Exchange","Market","Symbol","Buy/Sell","Participant","Custodian","Client",
    "Trader","Short Sell","Price","Volume","Yield %","Accrued Interest","Order No.",
    "Ticket No.","Date/Time","Execution Date/Time","Type","Filled Volume",
    "Remaining Volume","Disc. Volume","Trigger Price","Order Initiator","Pricing Mechanism"
]
PREVIEW_COLS = ["Exchange","Market","Participant","Custodian","Client",
                "Symbol","Buy/Sell","Price","Volume","Ticket No."]

SARESTOCK_EMAIL_SUBJECT = "DEALS CONFIRMATION"
SARESTOCK_EMAIL_BODY    = (
    "Good day,\r\n\r\nKindly find attached for deals confirmation.\r\n\r\nRegards,\r\nAnashe."
)
SARESTOCK_EMAIL_TO = "Anesu.Zingundu@fbc.co.zw"
SARESTOCK_EMAIL_CC = ";".join([
    "Enock.Rukarwa@fbc.co.zw","Manatsa.Tagwireyi@fbc.co.zw",
    "Norman.Chirima@fbc.co.zw","Richard.Mashava@fbc.co.zw"
])

FIELD_MAP = [
    ("Security","Symbol"),("SCA Code","Custodian"),("Buy/Sell","Buy/Sell"),
    ("Quantity","Volume + Filled Vol."),("Price","Price"),("Trader","Trader + Order Init."),
    ("VFX → VFEX","Exchange (+E)"),("VFEX = FBCSZWVX","Participant (fixed)"),
    ("ZSE = FBCSZWHX","Participant (fixed)"),("…-02 → …-0002","Client (zero-pad)"),
    ("DD/MM/YYYY …","Date/Time (auto)"),("Auto counter","Ticket No. (unique)"),
]

def _today_prefix():
    d = datetime.now(); return f"{d.year}{d.month:02d}{d.day:02d}"

def _load_state():
    try:
        with open(STATE_FILE) as f: return json.load(f)
    except Exception: return {"date":"","seq":0}

def _save_state(s):
    with open(STATE_FILE,"w") as f: json.dump(s,f)

def peek_next():
    s=_load_state(); t=_today_prefix()
    return f"{t}{(s['seq'] if s['date']==t else 0)+1:03d}"

def allocate_tickets(count):
    s=_load_state(); t=_today_prefix()
    if s["date"]!=t: s["seq"]=0; s["date"]=t
    tickets=[]
    for _ in range(count):
        s["seq"]+=1; tickets.append(f"{t}{s['seq']:03d}")
    _save_state(s); return tickets

def reset_counter(): _save_state({"date":"","seq":0})

def get_exchange(market):
    u=(market or "").strip().upper(); return "VFEX" if u=="VFX" else (u or "ZSE")

def get_participant(exch): return "FBCSZWVX" if exch=="VFEX" else "FBCSZWHX"

def get_market(sym,exch):
    s=(sym or "").upper().strip()
    if exch=="VFEX" or s.endswith(".VX"): return "REG"
    if s.endswith(".ZW"):
        if any(r in s for r in ["FHML","ZMRE","STFL","IPFL","HAFP","REVH"]): return "REIT"
        if any(o in s for o in ["SEED","CFI","CAFCA"]): return "ODD"
        return "REG"
    return "REG"

def pad_client(c):
    s=str(c or "").strip(); d=s.rfind("-")
    return s if d==-1 else f"{s[:d]}-{s[d+1:].zfill(4)}"

def get_now():
    d=datetime.now(); h=d.hour%12 or 12; ap="PM" if d.hour>=12 else "AM"
    return f"{d.day}/{d.month}/{d.year} {h}:{d.minute:02d}:{d.second:02d} {ap}"

def stamp():
    d=datetime.now(); return f"{d.day}_{d.month}_{d.year}"

def transform_rows(raw_rows):
    now=get_now(); tickets=allocate_tickets(len(raw_rows)); out=[]
    for i,r in enumerate(raw_rows):
        exch=get_exchange(r.get("Market","")); sym=r.get("Security","")
        out.append({
            "Exchange":exch,"Market":get_market(sym,exch),"Symbol":sym,
            "Buy/Sell":r.get("Buy/Sell",""),"Participant":get_participant(exch),
            "Custodian":r.get("SCA Code",""),"Client":pad_client(r.get("CSD Account","")),
            "Trader":r.get("Trader",""),"Short Sell":"NO",
            "Price":r.get("Price",""),"Volume":r.get("Quantity",""),
            "Yield %":"0","Accrued Interest":"0","Order No.":"",
            "Ticket No.":tickets[i],"Date/Time":now,"Execution Date/Time":now,
            "Type":"Limit","Filled Volume":r.get("Quantity",""),
            "Remaining Volume":"0","Disc. Volume":"0","Trigger Price":"0",
            "Order Initiator":r.get("Trader",""),"Pricing Mechanism":""
        })
    return out,now,tickets

def generate_csv(rows, out_dir, label):
    label = label.upper()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(
        out_dir,
        f"ExportExecutedOrders_{label}_{ts}.csv"
    )
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=EO_HEADERS)
        w.writeheader()
        for r in rows:
            w.writerow({h: r.get(h, "") for h in EO_HEADERS})
    return path

def generate_matched_excel(source_path,raw_rows,out_dir):
    exch=get_exchange(raw_rows[0].get("Market","")) if raw_rows else ""
    label="VFEX" if exch=="VFEX" else "ZSE"
    ext=os.path.splitext(source_path)[1] if source_path else ".xlsx"
    dest=os.path.join(out_dir,f"MATCHED TRADES, {label}{ext}")
    _shutil.copy2(source_path,dest); return dest

def generate_pdf(raw_rows,raw_headers,out_dir):
    _require("fpdf","fpdf2")
    from fpdf import FPDF
    exch=get_exchange(raw_rows[0].get("Market","")) if raw_rows else ""
    label="VFEX" if exch=="VFEX" else "ZSE"
    path=os.path.join(out_dir,f"MATCHED TRADES, {label}.pdf")
    col_chars=[]
    for h in raw_headers:
        mx=len(str(h))
        for row in raw_rows:
            v=str(row.get(h,"")); mx=max(mx,len(v))
        col_chars.append(min(mx,30))
    FS=6.5; CW=FS*0.52; RH=FS*0.9
    total_w=sum(c*CW for c in col_chars)+len(col_chars)*2
    if total_w<277: fmt,orient="A4","L"
    else: fmt,orient,FS,CW,RH="A3","L",5.5,5.5*0.52,5.5*0.85
    pdf=FPDF(orientation=orient,unit="mm",format=fmt)
    pdf.set_margins(6,6,6); pdf.set_auto_page_break(auto=True,margin=10); pdf.add_page()
    uw=pdf.w-12; cws=[max(c*CW,6) for c in col_chars]
    sc=uw/sum(cws); cws=[w*sc for w in cws]
    def _hdr():
        pdf.set_font("Courier",style="B",size=FS); pdf.set_text_color(0,0,0)
        for i,h in enumerate(raw_headers): pdf.cell(cws[i],RH,str(h)[:col_chars[i]],border=0,align="L")
        pdf.ln(); y=pdf.get_y(); pdf.set_draw_color(180,180,180)
        pdf.line(6,y,pdf.w-6,y); pdf.set_draw_color(0,0,0); pdf.ln(0.5)
    _hdr(); pdf.set_font("Courier",size=FS)
    for row in raw_rows:
        if pdf.get_y()>pdf.h-14: pdf.add_page(); _hdr(); pdf.set_font("Courier",size=FS)
        for i,h in enumerate(raw_headers):
            pdf.cell(cws[i],RH,str(row.get(h,""))[:col_chars[i]],border=0,align="L")
        pdf.ln()
    pdf.output(path); return path

def open_sarestock_outlook(file_paths):
    _require("win32com.client","pywin32")
    import win32com.client as win32
    outlook=win32.Dispatch("outlook.application"); mail=outlook.CreateItem(0)
    mail.Subject=SARESTOCK_EMAIL_SUBJECT; mail.Body=SARESTOCK_EMAIL_BODY
    mail.To=SARESTOCK_EMAIL_TO; mail.CC=SARESTOCK_EMAIL_CC
    for fp in file_paths:
        if fp and os.path.exists(fp): mail.Attachments.Add(fp)
    mail.Display(True)

# ════════════════════════════════════════════════════════════════════════════
#  ── EMAILER LOGIC ──────────────────────────────────────────────────────────
# ════════════════════════════════════════════════════════════════════════════
CONTACTS_FILE    = os.path.join(os.path.expanduser("~"),".fbc_dealnote_contacts.json")
KNOWN_CUSTODIANS = ["FBCZSEZW","CBZCZWHX","STINZWVX","CBCZSEZW","FBCSZWVX"]

CUSTODIAN_PREFIX_MAP = [
    ("FBC","FBCZSEZW"),("CBC","CBCZSEZW"),("CBZ","CBZCZWHX"),("STIN","STINZWVX"),("STIZ", "STINZWVX"),
]

_FBC_CC = [
    "Manatsa Tagwireyi <Manatsa.Tagwireyi@fbc.co.zw>",
    "Norman Chirima <Norman.Chirima@fbc.co.zw>",
    "Enock Rukarwa <Enock.Rukarwa@fbc.co.zw>",
    "Richard Mashava <Richard.Mashava@fbc.co.zw>",
    "Anesu Zingundu <Anesu.Zingundu@fbc.co.zw>",
]

CUSTODIAN_ROUTING = {
    "FBCZSEZW":{"label":"FBC Securities (ZSE)",
        "to":["Faith Chikati <Faith.Chikati@fbc.co.zw>"],
        "cc":["Custodial Services <CustodialServices@fbc.co.zw>"]+_FBC_CC},
    "CBZCZWHX":{"label":"CBZ (ZSE)",
        "to":["Sharleen Kapininga <skapininga@cbz.co.zw>","Phillipa Gurure <pgurure@cbz.co.zw>"],
        "cc":["Custodial Services <custodialservices@cbz.co.zw>"]+_FBC_CC},
    "STINZWVX":{"label":"Stanbic",
        "to":["Maigurira, Debra D <maigurirad@stanbic.com>","Chibvongodze, Kudakwashe K <chibvongodzek@stanbic.com>"],
        "cc":["custodyzim <custodyzim@standardbank.co.za>"]+_FBC_CC},
    "CBCZSEZW":{"label":"CABS / Old Mutual",
        "to":["Darlington Tatenda Maenda <darlingtonm@oldmutual.co.zw>"],
        "cc":["Custodial Services Division <custodialservicesdivision@cabs.co.zw>"]+_FBC_CC},
    "FBCSZWVX":{"label":"FBC Securities (VFEX)",
        "to":["Faith Chikati <Faith.Chikati@fbc.co.zw>"],
        "cc":["Custodial Services <CustodialServices@fbc.co.zw>"]+_FBC_CC},
}

CUSTODIAN_BODY_SINGLE = "Good day,\r\n\r\nKindly find attached today's deal note.\r\n\r\nRegards,\r\nAnashe."
CUSTODIAN_BODY_MULTI  = "Good day,\r\n\r\nKindly find attached today's deal notes.\r\n\r\nRegards,\r\nAnashe."

CLIENT_BODY_SINGLE = "Dear {client},\r\n\r\nPlease find attached your deal note for today's transaction.\r\n\r\nRegards,\r\nAnashe."
CLIENT_BODY_MULTI  = "Dear {client},\r\n\r\nPlease find attached your deal notes for today's transactions.\r\n\r\nRegards,\r\nAnashe."

CLIENT_CC = _FBC_CC

def _name_tokens(name):
    """Return a frozenset of uppercase words — order-independent name matching."""
    return frozenset(w.strip() for w in name.upper().split() if w.strip())

def _names_match(a, b):
    """True if two names share all tokens regardless of order.
    e.g. 'MAKWASHA TANYARADZWA' matches 'TANYARADZWA MAKWASHA'."""
    ta, tb = _name_tokens(a), _name_tokens(b)
    if not ta or not tb:
        return False
    # Exact token-set match  OR  one is a subset of the other (handles short saved names)
    return ta == tb or ta.issubset(tb) or tb.issubset(ta)

def find_contact(contacts, client_name):
    # 1. Exact match
    if client_name in contacts:
        return contacts[client_name]
    # 2. Order-independent token match (handles surname-first filenames)
    for saved, data in contacts.items():
        if _names_match(client_name, saved):
            return data
    # 3. Prefix fallback (legacy)
    for saved, data in contacts.items():
        if client_name.startswith(saved) or saved.startswith(client_name):
            return data
    return {}

def load_contacts():
    try:
        with open(CONTACTS_FILE) as f: return json.load(f)
    except Exception: return {}

def save_contacts(data):
    with open(CONTACTS_FILE,"w") as f: json.dump(data,f,indent=2)

def parse_client_name_from_filename(fname):
    base=os.path.splitext(fname)[0]
    base=re.sub(r'_+\d+_*$','',base)
    base=base.replace("_"," ").strip()
    base=re.sub(r'\s*\(\d+\)\s*$','',base)
    base=re.sub(r'[,\.]\s*.*$','',base)
    base=re.sub(r'\s+\d+.*$','',base)
    return base.strip().upper()

def parse_custodian_from_pdf(pdf_path):
    try:
        import fitz
        doc=fitz.open(pdf_path); text="".join(p.get_text() for p in doc); doc.close()
        for code in KNOWN_CUSTODIANS:
            if code in text: return code
        candidates=re.findall(r'\b([A-Z]{4,10})\b',text)
        for c in candidates:
            for prefix,canonical in CUSTODIAN_PREFIX_MAP:
                if c.startswith(prefix): return canonical
    except Exception: pass
    return None

def parse_deal_info_from_pdf(pdf_path):
    info={"deal_number":"","counter":"","deal_date":""}
    try:
        import fitz
        doc=fitz.open(pdf_path); text="".join(p.get_text() for p in doc); doc.close()
        m=re.search(r'Deal Number\s+(\d+)',text)
        if m: info["deal_number"]=m.group(1)
        m=re.search(r'Deal Date\s+([\d/]+)',text)
        if m: info["deal_date"]=m.group(1)
        m=re.search(r'\b([A-Z]{2,6}\.ZW|[A-Z]{2,6}\.VX)\b',text)
        if m: info["counter"]=m.group(1)
    except Exception: pass
    return info

def open_outlook(to_list,cc_list,subject,body,attachments):
    try:
        import win32com.client as win32
        outlook=win32.Dispatch("outlook.application"); mail=outlook.CreateItem(0)
        mail.To="; ".join(to_list); mail.CC="; ".join(cc_list)
        mail.Subject=subject; mail.Body=body
        for path in attachments:
            if os.path.exists(path): mail.Attachments.Add(path)
        mail.Display(True)
    except ImportError: raise ImportError("pywin32 not installed.\n\nRun:  pip install pywin32")
    except Exception as e: raise RuntimeError(f"Outlook error: {e}")

# ════════════════════════════════════════════════════════════════════════════
#  CONTACTS DIALOG
# ════════════════════════════════════════════════════════════════════════════
class ContactsDialog(tk.Toplevel):
    def __init__(self,parent,contacts,on_save):
        super().__init__(parent)
        self.title("Manage Client Contacts"); self.geometry("740x560")
        self.configure(bg=BG); self.contacts=dict(contacts)
        self.on_save=on_save; self._current_name=None
        self.grab_set(); self._build()

    def _build(self):
        hdr=tk.Frame(self,bg=FBC_DARK,pady=10,padx=14); hdr.pack(fill="x")
        tk.Label(hdr,text="👥  Manage Client Contacts",bg=FBC_DARK,fg=WHITE,
                 font=("Segoe UI",11,"bold")).pack(side="left")
        tk.Label(hdr,text=f"  {len(self.contacts)} clients saved",bg=FBC_DARK,fg="#90CAF9",
                 font=("Segoe UI",9)).pack(side="left",padx=8)

        body=tk.Frame(self,bg=BG); body.pack(fill="both",expand=True,padx=12,pady=10)

        # ── LEFT: search + list ──────────────────────────────────────────
        left=tk.Frame(body,bg=WHITE,relief="flat",bd=1); left.pack(side="left",fill="y",padx=(0,8))

        tk.Label(left,text="Clients",bg=FBC_MID,fg=WHITE,
                 font=("Segoe UI",9,"bold"),pady=6,padx=8).pack(fill="x")

        # search bar
        s_frame=tk.Frame(left,bg=WHITE,padx=4,pady=4); s_frame.pack(fill="x")
        tk.Label(s_frame,text="🔍",bg=WHITE,font=("Segoe UI",10)).pack(side="left")
        self.search_var=tk.StringVar()
        self.search_var.trace_add("write",lambda *_:self._filter_list())
        search_entry=tk.Entry(s_frame,textvariable=self.search_var,
                              font=("Segoe UI",9),relief="flat",bd=0,
                              bg="#F0F4F8",width=20)
        search_entry.pack(side="left",fill="x",expand=True,padx=4)
        tk.Button(s_frame,text="✕",command=lambda:(self.search_var.set(""),search_entry.focus()),
                  bg=WHITE,fg="#8096B0",relief="flat",font=("Segoe UI",8),
                  cursor="hand2",padx=2).pack(side="left")

        # listbox
        lb_frame=tk.Frame(left,bg=WHITE); lb_frame.pack(fill="both",expand=True,padx=4,pady=(0,4))
        self.listbox=tk.Listbox(lb_frame,width=26,font=("Segoe UI",9),
                                selectbackground=FBC_MID,activestyle="none",
                                relief="flat",bd=0)
        lb_sb=ttk.Scrollbar(lb_frame,orient="vertical",command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=lb_sb.set)
        self.listbox.pack(side="left",fill="both",expand=True)
        lb_sb.pack(side="right",fill="y")
        self.listbox.bind("<<ListboxSelect>>",self._on_select)

        # add / delete
        br=tk.Frame(left,bg=WHITE); br.pack(fill="x",padx=4,pady=(0,6))
        tk.Button(br,text="+ Add",command=self._add,bg=GREEN_DARK,fg=WHITE,
                  relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2").pack(side="left",padx=(0,4))
        tk.Button(br,text="✕ Delete",command=self._delete,bg=RED_DARK,fg=WHITE,
                  relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2").pack(side="left")

        # ── RIGHT: detail panel ──────────────────────────────────────────
        right=tk.Frame(body,bg=WHITE,relief="flat",bd=1); right.pack(side="left",fill="both",expand=True)
        tk.Label(right,text="Contact Details",bg=FBC_MID,fg=WHITE,
                 font=("Segoe UI",9,"bold"),pady=6,padx=8).pack(fill="x")
        self.detail=tk.Frame(right,bg=WHITE); self.detail.pack(fill="both",expand=True,padx=14,pady=12)
        self._show_detail(None)

        # ── bottom bar ───────────────────────────────────────────────────
        bot=tk.Frame(self,bg=BG); bot.pack(fill="x",padx=12,pady=(0,10))
        tk.Button(bot,text="💾  Save & Close",command=self._save,bg=FBC_MID,fg=WHITE,
                  font=("Segoe UI",10,"bold"),relief="flat",padx=16,pady=8,cursor="hand2").pack(side="right")

        self._filter_list()

    # ── list helpers ─────────────────────────────────────────────────────────
    def _filter_list(self):
        term=self.search_var.get().strip().upper()
        self.listbox.delete(0,tk.END)
        for n in sorted(self.contacts):
            if term in n.upper():
                self.listbox.insert(tk.END,n)
        # reselect current if still visible
        if self._current_name:
            for i in range(self.listbox.size()):
                if self.listbox.get(i)==self._current_name:
                    self.listbox.selection_set(i); break

    def _refresh_list(self):
        self._filter_list()

    def _on_select(self,_=None):
        sel=self.listbox.curselection()
        if sel:
            self._current_name=self.listbox.get(sel[0])
            self._show_detail(self._current_name)

    def _show_detail(self,name):
        for w in self.detail.winfo_children(): w.destroy()
        if not name:
            tk.Label(self.detail,text="Select a client on the left\nto view or edit their details.",
                     bg=WHITE,fg="#8096B0",font=("Segoe UI",9),justify="center").pack(pady=30); return

        data=self.contacts.get(name,{"email":""})

        # ── client name (editable) ──
        tk.Label(self.detail,text="Client Name:",bg=WHITE,fg="#607080",
                 font=("Segoe UI",8,"bold")).pack(anchor="w")
        name_row=tk.Frame(self.detail,bg=WHITE); name_row.pack(fill="x",pady=(2,10))
        self.entry_name=tk.Entry(name_row,font=("Segoe UI",10),width=34)
        self.entry_name.insert(0,name)
        self.entry_name.pack(side="left")
        tk.Button(name_row,text="✏ Rename",
                  command=lambda n=name:self._rename(n),
                  bg=FBC_MID,fg=WHITE,relief="flat",
                  font=("Segoe UI",8,"bold"),cursor="hand2",
                  padx=8,pady=4).pack(side="left",padx=6)

        # ── email ──
        tk.Label(self.detail,text="Client Email:",bg=WHITE,fg="#607080",
                 font=("Segoe UI",8,"bold")).pack(anchor="w")
        self.entry_email=tk.Entry(self.detail,font=("Segoe UI",10),width=42)
        self.entry_email.insert(0,data.get("email",""))
        self.entry_email.pack(anchor="w",pady=(2,4))

        # show current saved email as hint
        saved=data.get("email","")
        hint_txt="No email saved yet" if not saved else f"Saved: {saved}"
        hint_col=RED_DARK if not saved else GREEN_DARK
        tk.Label(self.detail,text=hint_txt,bg=WHITE,fg=hint_col,
                 font=("Segoe UI",8)).pack(anchor="w",pady=(0,12))

        tk.Button(self.detail,text="✔  Apply Email",
                  command=lambda n=name:self._apply(n),
                  bg=GREEN_DARK,fg=WHITE,relief="flat",
                  font=("Segoe UI",9,"bold"),cursor="hand2",
                  padx=12,pady=6).pack(anchor="w")

    def _rename(self,old_name):
        new_name=self.entry_name.get().strip().upper()
        if not new_name:
            messagebox.showwarning("Empty","Name cannot be empty.",parent=self); return
        if new_name==old_name:
            messagebox.showinfo("No Change","Name is the same.",parent=self); return
        if new_name in self.contacts:
            messagebox.showwarning("Duplicate",f"'{new_name}' already exists.",parent=self); return
        # rename in dict
        self.contacts[new_name]=self.contacts.pop(old_name)
        self._current_name=new_name
        self._filter_list()
        self._show_detail(new_name)
        messagebox.showinfo("Renamed",f"'{old_name}' → '{new_name}'\n\nClick 'Save & Close' to keep this change.",parent=self)

    def _apply(self,name):
        self.contacts[name]={"email":self.entry_email.get().strip()}
        self._show_detail(name)
        messagebox.showinfo("Saved",f"Email saved for {name}.\nClick 'Save & Close' to write to disk.",parent=self)

    def _add(self):
        dlg=tk.Toplevel(self); dlg.title("Add Client")
        dlg.geometry("340x130"); dlg.configure(bg=BG); dlg.grab_set()
        tk.Label(dlg,text="Client name (as it appears in the filename):",
                 bg=BG,font=("Segoe UI",9)).pack(pady=(14,4),padx=12)
        e=tk.Entry(dlg,font=("Segoe UI",10),width=34); e.pack(padx=12); e.focus()
        def ok():
            n=e.get().strip().upper()
            if not n: return
            if n in self.contacts:
                messagebox.showwarning("Duplicate",f"'{n}' already exists.",parent=dlg); return
            self.contacts[n]={"email":""}
            self._current_name=n
            self._filter_list(); dlg.destroy(); self._show_detail(n)
        e.bind("<Return>",lambda _:ok())
        tk.Button(dlg,text="Add",command=ok,bg=FBC_MID,fg=WHITE,
                  relief="flat",font=("Segoe UI",9,"bold"),cursor="hand2",
                  padx=12,pady=6).pack(pady=10)

    def _delete(self):
        sel=self.listbox.curselection()
        if not sel: return
        name=self.listbox.get(sel[0])
        if messagebox.askyesno("Delete",f"Delete '{name}' from contacts?",parent=self):
            self.contacts.pop(name,None); self._current_name=None
            self._filter_list(); self._show_detail(None)

    def _save(self):
        save_contacts(self.contacts); self.on_save(self.contacts); self.destroy()

# ════════════════════════════════════════════════════════════════════════════
#  SARESTOCK PAGE
# ════════════════════════════════════════════════════════════════════════════
class SarestockPage(tk.Frame):
    def __init__(self,parent):
        super().__init__(parent,bg=BG)
        self.raw_rows=[]; self.raw_headers=[]; self.conv_rows=[]
        self.source_path=None; self.gen_csv=self.gen_pdf=self.gen_mt_xlsx=None
        self.raw_rows2=[]; self.raw_headers2=[]; self.conv_rows2=[]
        self.source_path2=None; self.gen_csv2=self.gen_pdf2=self.gen_mt_xlsx2=None
        self.out_dir=os.path.join(os.path.expanduser("~"),"Downloads")
        self._build()

    def _build(self):
        # top info bar
        info=tk.Frame(self,bg=FBC_MID,padx=16,pady=8); info.pack(fill="x")
        tk.Label(info,text="📊  Sarestock Upload Converter",bg=FBC_MID,fg=WHITE,
                 font=("Segoe UI",11,"bold")).pack(side="left")
        right=tk.Frame(info,bg=FBC_MID); right.pack(side="right")
        self.lbl_ticket=tk.Label(right,text="Next ticket: …",bg=FBC_MID,fg="#D0EAFF",
                                  font=("Consolas",10)); self.lbl_ticket.pack(side="right",padx=(10,0))
        tk.Button(right,text="⟳  Reset Counter",command=self._reset,
                  bg=FBC_DARK,fg="#90CAF9",relief="flat",font=("Segoe UI",9),
                  cursor="hand2",padx=8,pady=3).pack(side="right")

        # two-column paned area
        self.paned=tk.PanedWindow(self,orient="horizontal",bg=SEP_CLR,sashwidth=4,sashrelief="flat")
        self.paned.pack(fill="both",expand=True)
        self.left_frame,self.left_canvas,self.left_body=self._scroll_pane(self.paned)
        self.paned.add(self.left_frame,stretch="always")
        self.right_frame,self.right_canvas,self.right_body=self._scroll_pane(self.paned)
        self.paned.add(self.right_frame,stretch="always")

        self._build_bottom_bar()
        self._build_left_column()
        self._build_right_column()
        self._refresh_ticket()

    def _scroll_pane(self,parent):
        frame=tk.Frame(parent,bg=BG)
        canvas=tk.Canvas(frame,bg=BG,highlightthickness=0)
        vsb=ttk.Scrollbar(frame,orient="vertical",command=canvas.yview)
        inner=tk.Frame(canvas,bg=BG)
        inner.bind("<Configure>",lambda e:canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0),window=inner,anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left",fill="both",expand=True); vsb.pack(side="right",fill="y")
        canvas.bind("<Enter>",lambda e,c=canvas:c.bind_all("<MouseWheel>",
            lambda ev:c.yview_scroll(-1*(ev.delta//120),"units")))
        canvas.bind("<Leave>",lambda e,c=canvas:c.unbind_all("<MouseWheel>"))
        return frame,canvas,inner

    def _build_bottom_bar(self):
        bar=tk.Frame(self,bg=BOTTOM,pady=10); bar.pack(fill="x",side="bottom")
        path_row=tk.Frame(bar,bg=BOTTOM); path_row.pack(fill="x",padx=16,pady=(0,6))
        tk.Label(path_row,text="📁  Files saved to:",bg=BOTTOM,fg="#8BAAC8",
                 font=("Segoe UI",8)).pack(side="left")
        self.lbl_outdir=tk.Label(path_row,text=self.out_dir,bg=BOTTOM,fg="#90CAF9",
                                  font=("Segoe UI",8)); self.lbl_outdir.pack(side="left",padx=6)
        tk.Button(path_row,text="Change…",command=self._pick_outdir,bg="#1A3A6B",fg="#90CAF9",
                  relief="flat",font=("Segoe UI",8),cursor="hand2",padx=6,pady=2).pack(side="left")
        btn_row=tk.Frame(bar,bg=BOTTOM); btn_row.pack(fill="x",padx=16)
        btn_row.columnconfigure(0,weight=1); btn_row.columnconfigure(1,weight=1); btn_row.columnconfigure(2,weight=2)
        self.btn_email=tk.Button(btn_row,text="✉  Send — ZSE Only",command=self._send_email,
            bg=GREEN_DARK,fg=WHITE,font=("Segoe UI",10,"bold"),relief="flat",pady=9,
            cursor="hand2",state="disabled")
        self.btn_email.grid(row=0,column=0,sticky="ew",padx=(0,6))
        self.btn_email2=tk.Button(btn_row,text="✉  Send — VFEX Only",command=self._send_email2,
            bg="#1A3A6B",fg=WHITE,font=("Segoe UI",10,"bold"),relief="flat",pady=9,
            cursor="hand2",state="disabled")
        self.btn_email2.grid(row=0,column=1,sticky="ew",padx=(0,6))
        self.btn_email_both=tk.Button(btn_row,text="✉  Send BOTH ZSE + VFEX in One Email",
            command=self._send_email_both,bg=FBC_MID,fg=WHITE,font=("Segoe UI",11,"bold"),
            relief="flat",pady=9,cursor="hand2",state="disabled")
        self.btn_email_both.grid(row=0,column=2,sticky="ew")
        tk.Label(bar,text="Each email attaches: Matched Trades PDF + original file  ·  Pre-fills To, CC and Subject",
                 bg=BOTTOM,fg="#5D7A99",font=("Segoe UI",8)).pack(pady=(4,0))

    def _build_left_column(self):
        p=self.left_body
        hdr=tk.Frame(p,bg=COL1_HDR); hdr.pack(fill="x",padx=12,pady=(12,0))
        tk.Label(hdr,text=" 1 ",bg=WHITE,fg=COL1_HDR,font=("Segoe UI",9,"bold"),
                 padx=4,pady=4).pack(side="left",padx=(8,0),pady=6)
        tk.Label(hdr,text="  FIRST EXCHANGE  (ZSE or VFEX)",bg=COL1_HDR,fg=WHITE,
                 font=("Segoe UI",10,"bold")).pack(side="left",pady=6)
        ucard=self._card(p,COL1_HDR)
        dz=tk.Frame(ucard,bg="#F4F8FE",relief="groove",bd=2); dz.pack(fill="x",pady=(0,10))
        inner=tk.Frame(dz,bg="#F4F8FE"); inner.pack(pady=16)
        tk.Label(inner,text="🗂",bg="#F4F8FE",font=("Segoe UI",22)).pack()
        tk.Label(inner,text="Upload Matched Trades File",bg="#F4F8FE",fg=FBC_MID,
                 font=("Segoe UI",10,"bold")).pack(pady=(4,0))
        tk.Label(inner,text=".csv or .xlsx",bg="#F4F8FE",fg="#8096B0",font=("Segoe UI",8)).pack()
        tk.Button(inner,text="  Browse…  ",command=self._pick_file,bg=FBC_MID,fg=WHITE,
                  font=("Segoe UI",10,"bold"),relief="flat",padx=14,pady=6,cursor="hand2").pack(pady=(8,0))
        self.info_bar1=tk.Frame(ucard,bg=TAG_BLUE,highlightbackground=FBC_ACCENT,highlightthickness=1)
        self.lbl_file1=tk.Label(self.info_bar1,text="",bg=TAG_BLUE,fg=FBC_DARK,font=("Segoe UI",9,"bold"))
        self.lbl_rows1=tk.Label(self.info_bar1,text="",bg=TAG_BLUE,fg=FBC_MID,font=("Consolas",8))
        self.btn_reupload1=tk.Button(self.info_bar1,text="↩ Change",command=self._pick_file,
                                     bg=TAG_BLUE,fg=FBC_MID,relief="flat",font=("Segoe UI",8),cursor="hand2")
        dcard=self._card(p,COL1_HDR)
        tk.Label(dcard,text="DOWNLOAD",bg=CARD_BG,fg="#8096B0",font=("Segoe UI",8,"bold")).pack(anchor="w")
        btn_row=tk.Frame(dcard,bg=CARD_BG); btn_row.pack(fill="x",pady=(6,2))
        self.btn_csv=self._col_btn(btn_row,"⬇  Download CSV",self._dl_csv,FBC_MID)
        self.btn_pdf=self._col_btn(btn_row,"⬇  Download PDF",self._dl_pdf,RED_DARK)
        for b in (self.btn_csv,self.btn_pdf): b.config(state="disabled")
        self.lbl_csv_done=tk.Label(dcard,text="",bg=CARD_BG,fg=GREEN_DARK,font=("Segoe UI",8)); self.lbl_csv_done.pack(anchor="w")
        self.lbl_pdf_done=tk.Label(dcard,text="",bg=CARD_BG,fg=GREEN_DARK,font=("Segoe UI",8)); self.lbl_pdf_done.pack(anchor="w")
        self.prev_outer1=tk.Frame(p,bg=BG)
        self._build_preview_shell(self.prev_outer1,COL1_HDR,"PREVIEW — FIRST EXCHANGE","prev_body1","lbl_showing1")
        self.prev_outer1.pack_forget()

    def _build_right_column(self):
        p=self.right_body
        hdr=tk.Frame(p,bg=COL2_HDR); hdr.pack(fill="x",padx=12,pady=(12,0))
        tk.Label(hdr,text=" 2 ",bg=WHITE,fg=COL2_HDR,font=("Segoe UI",9,"bold"),
                 padx=4,pady=4).pack(side="left",padx=(8,0),pady=6)
        tk.Label(hdr,text="  SECOND EXCHANGE  (ZSE or VFEX)",bg=COL2_HDR,fg=WHITE,
                 font=("Segoe UI",10,"bold")).pack(side="left",pady=6)
        ucard=self._card(p,COL2_HDR)
        dz=tk.Frame(ucard,bg="#F4F8FE",relief="groove",bd=2); dz.pack(fill="x",pady=(0,10))
        inner=tk.Frame(dz,bg="#F4F8FE"); inner.pack(pady=16)
        tk.Label(inner,text="🗂",bg="#F4F8FE",font=("Segoe UI",22)).pack()
        tk.Label(inner,text="Upload Matched Trades File",bg="#F4F8FE",fg=FBC_MID,
                 font=("Segoe UI",10,"bold")).pack(pady=(4,0))
        tk.Label(inner,text=".csv or .xlsx",bg="#F4F8FE",fg="#8096B0",font=("Segoe UI",8)).pack()
        tk.Button(inner,text="  Browse…  ",command=self._pick_file2,bg=COL2_HDR,fg=WHITE,
                  font=("Segoe UI",10,"bold"),relief="flat",padx=14,pady=6,cursor="hand2").pack(pady=(8,0))
        self.info_bar2=tk.Frame(ucard,bg=TAG_BLUE,highlightbackground=FBC_ACCENT,highlightthickness=1)
        self.lbl_file2=tk.Label(self.info_bar2,text="",bg=TAG_BLUE,fg=FBC_DARK,font=("Segoe UI",9,"bold"))
        self.lbl_rows2=tk.Label(self.info_bar2,text="",bg=TAG_BLUE,fg=FBC_MID,font=("Consolas",8))
        self.btn_reupload2=tk.Button(self.info_bar2,text="↩ Change",command=self._pick_file2,
                                     bg=TAG_BLUE,fg=FBC_MID,relief="flat",font=("Segoe UI",8),cursor="hand2")
        dcard=self._card(p,COL2_HDR)
        tk.Label(dcard,text="DOWNLOAD",bg=CARD_BG,fg="#8096B0",font=("Segoe UI",8,"bold")).pack(anchor="w")
        btn_row=tk.Frame(dcard,bg=CARD_BG); btn_row.pack(fill="x",pady=(6,2))
        self.btn_csv2=self._col_btn(btn_row,"⬇  Download CSV",self._dl_csv2,FBC_MID)
        self.btn_pdf2=self._col_btn(btn_row,"⬇  Download PDF",self._dl_pdf2,RED_DARK)
        for b in (self.btn_csv2,self.btn_pdf2): b.config(state="disabled")
        self.lbl_csv2_done=tk.Label(dcard,text="",bg=CARD_BG,fg=GREEN_DARK,font=("Segoe UI",8)); self.lbl_csv2_done.pack(anchor="w")
        self.lbl_pdf2_done=tk.Label(dcard,text="",bg=CARD_BG,fg=GREEN_DARK,font=("Segoe UI",8)); self.lbl_pdf2_done.pack(anchor="w")
        self.prev_outer2=tk.Frame(p,bg=BG)
        self._build_preview_shell(self.prev_outer2,COL2_HDR,"PREVIEW — SECOND EXCHANGE","prev_body2","lbl_showing2")
        self.prev_outer2.pack_forget()

    def _card(self,parent,accent):
        wrapper=tk.Frame(parent,bg=BG); wrapper.pack(fill="x",padx=12,pady=(4,0))
        tk.Frame(wrapper,bg=accent,height=2).pack(fill="x")
        body=tk.Frame(wrapper,bg=CARD_BG,padx=14,pady=12,highlightbackground=SEP_CLR,highlightthickness=1)
        body.pack(fill="x"); return body

    def _col_btn(self,parent,text,cmd,bg):
        b=tk.Button(parent,text=text,command=cmd,bg=bg,fg=WHITE,
                    font=("Segoe UI",9,"bold"),relief="flat",padx=10,pady=7,cursor="hand2")
        b.pack(side="left",padx=(0,8),pady=2); return b

    def _build_preview_shell(self,outer,color,title,body_attr,label_attr):
        outer.pack(fill="x",padx=12,pady=(4,0))
        hdr=tk.Frame(outer,bg=color); hdr.pack(fill="x")
        tk.Label(hdr,text=f"  {title}",bg=color,fg=WHITE,
                 font=("Segoe UI",9,"bold")).pack(side="left",pady=5,padx=6)
        lbl=tk.Label(hdr,text="",bg=color,fg="#90CAF9",font=("Segoe UI",8))
        lbl.pack(side="right",padx=8,pady=5); setattr(self,label_attr,lbl)
        body=tk.Frame(outer,bg=CARD_BG,padx=14,pady=10,highlightbackground=SEP_CLR,highlightthickness=1)
        body.pack(fill="x"); setattr(self,body_attr,body)

    def _refresh_ticket(self): self.lbl_ticket.config(text=f"Next ticket:  {peek_next()}")

    def _build_preview(self,body_attr,label_attr,rows,tickets,now):
        body=getattr(self,body_attr); lbl=getattr(self,label_attr)
        for w in body.winfo_children(): w.destroy()
        summ=tk.Frame(body,bg=CARD_BG); summ.pack(fill="x",pady=(0,8))
        def ibox(parent,ltext,val,col):
            f=tk.Frame(parent,bg="#F0F7FF",padx=10,pady=6,highlightbackground=SEP_CLR,highlightthickness=1)
            f.grid(row=0,column=col,sticky="nsew",padx=(0,6)); parent.columnconfigure(col,weight=1)
            tk.Label(f,text=ltext,bg="#F0F7FF",fg="#8096B0",font=("Segoe UI",7,"bold")).pack(anchor="w")
            tk.Label(f,text=val,bg="#F0F7FF",fg=FBC_DARK,font=("Segoe UI",11,"bold")).pack(anchor="w")
        ibox(summ,"ROWS",str(len(rows)),0); ibox(summ,"DATE/TIME",now,1)
        ibox(summ,"TICKET RANGE",f"{tickets[0]}→{tickets[-1]}",2)
        style=ttk.Style()
        try: style.theme_use("default")
        except Exception: pass
        style.configure("Treeview.Heading",background=FBC_DARK,foreground=WHITE,relief="flat",font=("Segoe UI",8,"bold"))
        style.map("Treeview.Heading",background=[("active",FBC_MID)],foreground=[("active",WHITE)])
        style.configure("Treeview",font=("Segoe UI",8),rowheight=22)
        frm=tk.Frame(body,bg=CARD_BG); frm.pack(fill="x")
        xsb=ttk.Scrollbar(frm,orient="horizontal"); ysb=ttk.Scrollbar(frm,orient="vertical")
        tv=ttk.Treeview(frm,columns=PREVIEW_COLS,show="headings",height=min(len(rows),7),
                        xscrollcommand=xsb.set,yscrollcommand=ysb.set)
        xsb.config(command=tv.xview); ysb.config(command=tv.yview)
        for col in PREVIEW_COLS: tv.heading(col,text=col); tv.column(col,width=95,minwidth=70,anchor="w")
        for i,row in enumerate(rows):
            vals=[row.get(c,"") for c in PREVIEW_COLS]
            bs=row.get("Buy/Sell","").strip().lower()
            tag="buy" if bs=="buy" else("sell" if bs=="sell" else("even" if i%2==0 else "odd"))
            tv.insert("","end",values=vals,tags=(tag,))
        tv.tag_configure("even",background="#F7FAFF"); tv.tag_configure("odd",background=CARD_BG)
        tv.tag_configure("buy",foreground=GREEN_DARK,background="#F2FBF5")
        tv.tag_configure("sell",foreground=RED_DARK,background="#FFF5F5")
        tv.grid(row=0,column=0,sticky="nsew"); ysb.grid(row=0,column=1,sticky="ns")
        xsb.grid(row=1,column=0,sticky="ew"); frm.columnconfigure(0,weight=1)
        lbl.config(text=f"showing {len(rows)} rows")

    def _pick_file(self):
        path=filedialog.askopenfilename(title="Select First Exchange Matched Trades File",
            filetypes=[("CSV / Excel","*.csv *.xlsx *.xls"),("All files","*.*")])
        if path: self._load_file(path)

    def _load_file(self,path):
        try:
            pd=_require("pandas")
            df=pd.read_csv(path,dtype=str).fillna("") if path.lower().endswith(".csv") else pd.read_excel(path,dtype=str).fillna("")
            self.raw_headers=list(df.columns); self.raw_rows=df.to_dict("records")
            if not self.raw_rows: raise ValueError("File is empty.")
            self.source_path=path; self.conv_rows,now,tickets=transform_rows(self.raw_rows)
            self.gen_csv=self.gen_pdf=self.gen_mt_xlsx=None
            self.lbl_csv_done.config(text=""); self.lbl_pdf_done.config(text="")
            self.btn_csv.config(text="⬇  Download CSV",bg=FBC_MID)
            self.btn_pdf.config(text="⬇  Download PDF",bg=RED_DARK)
            fname=os.path.basename(path); exch=get_exchange(self.raw_rows[0].get("Market",""))
            for w in self.info_bar1.winfo_children(): w.pack_forget()
            self.info_bar1.pack(fill="x",pady=(0,6))
            tk.Label(self.info_bar1,text="✅",bg=TAG_BLUE,font=("Segoe UI",10)).pack(side="left",padx=(6,2),pady=4)
            self.lbl_file1.config(text=fname); self.lbl_file1.pack(side="left",pady=4)
            self.lbl_rows1.config(text=f"  {len(self.conv_rows)} rows  ·  {tickets[0]}→{tickets[-1]}")
            self.lbl_rows1.pack(side="left",pady=4); self.btn_reupload1.pack(side="right",padx=6,pady=4)
            for b in (self.btn_csv,self.btn_pdf): b.config(state="normal")
            self.btn_email.config(text=f"✉  Send — {exch} Only",state="normal")
            self.prev_outer1.pack(fill="x",padx=12,pady=(4,0))
            self._build_preview("prev_body1","lbl_showing1",self.conv_rows,tickets,now)
            if self.source_path2: self.btn_email_both.config(state="normal")
            self._refresh_ticket()
        except Exception as e: messagebox.showerror("Error loading file",str(e))

    def _dl_csv(self):
        try:
            exch=get_exchange(self.raw_rows[0].get("Market",""))
            self.gen_csv=generate_csv(self.conv_rows,self.out_dir,exch)
            self.lbl_csv_done.config(text=f"✅  {os.path.basename(self.gen_csv)} saved")
            self.btn_csv.config(text="✅  CSV Downloaded",bg="#1B5E20")
        except Exception as e: messagebox.showerror("CSV Error",str(e))

    def _dl_pdf(self):
        try:
            self.gen_pdf=generate_pdf(self.raw_rows,self.raw_headers,self.out_dir)
            self.lbl_pdf_done.config(text=f"✅  {os.path.basename(self.gen_pdf)} saved")
            self.btn_pdf.config(text="✅  PDF Downloaded",bg="#7B1010")
        except Exception as e: messagebox.showerror("PDF Error",str(e))

    def _ensure_email_files(self):
        if not self.gen_pdf:
            self.gen_pdf=generate_pdf(self.raw_rows,self.raw_headers,self.out_dir)
            self.lbl_pdf_done.config(text=f"✅  {os.path.basename(self.gen_pdf)} saved")
        if not self.gen_mt_xlsx:
            self.gen_mt_xlsx=generate_matched_excel(self.source_path,self.raw_rows,self.out_dir)

    def _send_email(self):
        try: self._ensure_email_files(); open_sarestock_outlook([self.gen_pdf,self.gen_mt_xlsx])
        except ImportError: messagebox.showerror("pywin32 not installed","Run:  pip install pywin32")
        except Exception as e: messagebox.showerror("Outlook Error",str(e))

    def _pick_file2(self):
        path=filedialog.askopenfilename(title="Select Second Exchange Matched Trades File",
            filetypes=[("CSV / Excel","*.csv *.xlsx *.xls"),("All files","*.*")])
        if path: self._load_file2(path)

    def _load_file2(self,path):
        try:
            pd=_require("pandas")
            df=pd.read_csv(path,dtype=str).fillna("") if path.lower().endswith(".csv") else pd.read_excel(path,dtype=str).fillna("")
            self.raw_headers2=list(df.columns); self.raw_rows2=df.to_dict("records")
            if not self.raw_rows2: raise ValueError("File is empty.")
            self.source_path2=path; self.conv_rows2,now,tickets=transform_rows(self.raw_rows2)
            self.gen_csv2=self.gen_pdf2=self.gen_mt_xlsx2=None
            self.lbl_csv2_done.config(text=""); self.lbl_pdf2_done.config(text="")
            self.btn_csv2.config(text="⬇  Download CSV",bg=FBC_MID)
            self.btn_pdf2.config(text="⬇  Download PDF",bg=RED_DARK)
            fname=os.path.basename(path); exch2=get_exchange(self.raw_rows2[0].get("Market",""))
            for w in self.info_bar2.winfo_children(): w.pack_forget()
            self.info_bar2.pack(fill="x",pady=(0,6))
            tk.Label(self.info_bar2,text="✅",bg=TAG_BLUE,font=("Segoe UI",10)).pack(side="left",padx=(6,2),pady=4)
            self.lbl_file2.config(text=fname); self.lbl_file2.pack(side="left",pady=4)
            self.lbl_rows2.config(text=f"  {len(self.conv_rows2)} rows  ·  {tickets[0]}→{tickets[-1]}")
            self.lbl_rows2.pack(side="left",pady=4); self.btn_reupload2.pack(side="right",padx=6,pady=4)
            for b in (self.btn_csv2,self.btn_pdf2): b.config(state="normal")
            self.btn_email2.config(text=f"✉  Send — {exch2} Only",state="normal")
            self.prev_outer2.pack(fill="x",padx=12,pady=(4,0))
            self._build_preview("prev_body2","lbl_showing2",self.conv_rows2,tickets,now)
            if self.source_path: self.btn_email_both.config(state="normal")
            self._refresh_ticket()
        except Exception as e: messagebox.showerror("Error loading 2nd file",str(e))

    def _dl_csv2(self):
        try:
            exch2=get_exchange(self.raw_rows2[0].get("Market",""))
            self.gen_csv2=generate_csv(self.conv_rows2,self.out_dir,exch2)
            self.lbl_csv2_done.config(text=f"✅  {os.path.basename(self.gen_csv2)} saved")
            self.btn_csv2.config(text="✅  CSV Downloaded",bg="#1B5E20")
        except Exception as e: messagebox.showerror("CSV Error (2nd)",str(e))

    def _dl_pdf2(self):
        try:
            self.gen_pdf2=generate_pdf(self.raw_rows2,self.raw_headers2,self.out_dir)
            self.lbl_pdf2_done.config(text=f"✅  {os.path.basename(self.gen_pdf2)} saved")
            self.btn_pdf2.config(text="✅  PDF Downloaded",bg="#7B1010")
        except Exception as e: messagebox.showerror("PDF Error (2nd)",str(e))

    def _ensure_email_files2(self):
        if not self.gen_pdf2:
            self.gen_pdf2=generate_pdf(self.raw_rows2,self.raw_headers2,self.out_dir)
            self.lbl_pdf2_done.config(text=f"✅  {os.path.basename(self.gen_pdf2)} saved")
        if not self.gen_mt_xlsx2:
            self.gen_mt_xlsx2=generate_matched_excel(self.source_path2,self.raw_rows2,self.out_dir)

    def _send_email2(self):
        try: self._ensure_email_files2(); open_sarestock_outlook([self.gen_pdf2,self.gen_mt_xlsx2])
        except ImportError: messagebox.showerror("pywin32 not installed","Run:  pip install pywin32")
        except Exception as e: messagebox.showerror("Outlook Error",str(e))

    def _send_email_both(self):
        try:
            self._ensure_email_files(); self._ensure_email_files2()
            open_sarestock_outlook([self.gen_pdf,self.gen_mt_xlsx,self.gen_pdf2,self.gen_mt_xlsx2])
        except ImportError: messagebox.showerror("pywin32 not installed","Run:  pip install pywin32")
        except Exception as e: messagebox.showerror("Outlook Error",str(e))

    def _pick_outdir(self):
        d=filedialog.askdirectory(title="Choose output folder",initialdir=self.out_dir)
        if d: self.out_dir=d; self.lbl_outdir.config(text=d)

    def _reset(self):
        if messagebox.askyesno("Reset Counter","Reset ticket counter?\n\nOnly do this if Sarestock has also been reset."):
            reset_counter()
            self.raw_rows=[]; self.raw_headers=[]; self.conv_rows=[]
            self.gen_csv=self.gen_pdf=self.gen_mt_xlsx=None; self.source_path=None
            self.raw_rows2=[]; self.raw_headers2=[]; self.conv_rows2=[]
            self.gen_csv2=self.gen_pdf2=self.gen_mt_xlsx2=None; self.source_path2=None
            self.prev_outer1.pack_forget(); self.prev_outer2.pack_forget()
            for b in (self.btn_csv,self.btn_pdf,self.btn_email,
                      self.btn_csv2,self.btn_pdf2,self.btn_email2,self.btn_email_both):
                b.config(state="disabled")
            self._refresh_ticket()

# ════════════════════════════════════════════════════════════════════════════
#  EMAILER PAGE
# ════════════════════════════════════════════════════════════════════════════
class EmailerPage(tk.Frame):
    def __init__(self,parent):
        super().__init__(parent,bg=BG)
        self.contacts=load_contacts(); self.deal_items=[]; self.pdf_folder=""
        self._build()

    def _build(self):
        # top bar
        bar=tk.Frame(self,bg=FBC_MID,padx=16,pady=8); bar.pack(fill="x")
        tk.Label(bar,text="✉  Deal Note Email Automator",bg=FBC_MID,fg=WHITE,
                 font=("Segoe UI",11,"bold")).pack(side="left")
        tk.Button(bar,text="👤  Manage Client Contacts",command=self._open_contacts,
                  bg=FBC_DARK,fg=WHITE,relief="flat",font=("Segoe UI",9,"bold"),
                  cursor="hand2",padx=10,pady=4).pack(side="right",padx=4)

        # ── upload panel ─────────────────────────────────────────────────────
        fp=tk.Frame(self,bg=WHITE,padx=16,pady=12); fp.pack(fill="x",padx=16,pady=(12,0))

        # row 1: upload buttons
        btn_row=tk.Frame(fp,bg=WHITE); btn_row.pack(fill="x")
        tk.Button(btn_row,text="📂  Choose Deal Notes Folder",command=self._pick_folder,
                  bg=FBC_MID,fg=WHITE,relief="flat",font=("Segoe UI",10,"bold"),
                  cursor="hand2",padx=14,pady=8).pack(side="left")

        tk.Label(btn_row,text="  or  ",bg=WHITE,fg="#8096B0",
                 font=("Segoe UI",9)).pack(side="left")

        tk.Button(btn_row,text="📄  Select Individual Deal Note(s)",command=self._pick_individual_files,
                  bg="#4051B5",fg=WHITE,relief="flat",font=("Segoe UI",10,"bold"),
                  cursor="hand2",padx=14,pady=8).pack(side="left")

        # clear all button
        self.btn_clear=tk.Button(btn_row,text="🗑  Clear All Uploads",command=self._clear_uploads,
                  bg=RED_DARK,fg=WHITE,relief="flat",font=("Segoe UI",9,"bold"),
                  cursor="hand2",padx=10,pady=8,state="disabled")
        self.btn_clear.pack(side="right")

        # row 2: status labels
        info_row=tk.Frame(fp,bg=WHITE); info_row.pack(fill="x",pady=(6,0))
        self.lbl_folder=tk.Label(info_row,text="No files loaded",bg=WHITE,fg="#8096B0",font=("Segoe UI",9))
        self.lbl_folder.pack(side="left")
        self.lbl_found=tk.Label(info_row,text="",bg=WHITE,fg=FBC_MID,font=("Consolas",9))
        self.lbl_found.pack(side="left",padx=10)
        self.lbl_file_list=tk.Label(info_row,text="",bg=WHITE,fg="#607080",font=("Segoe UI",8))
        self.lbl_file_list.pack(side="left")

        # tabs
        style=ttk.Style()
        style.configure("TNotebook.Tab",font=("Segoe UI",10,"bold"),padding=[14,6])
        nb=ttk.Notebook(self); nb.pack(fill="both",expand=True,padx=16,pady=10)
        self.tab_cust=tk.Frame(nb,bg=BG)
        self.tab_client=tk.Frame(nb,bg=BG)
        nb.add(self.tab_cust,text="  ✉  Custodian Emails  ")
        nb.add(self.tab_client,text="  ✉  Client Emails  ")
        self._build_custodian_tab()
        self._build_client_tab()

    # ── CUSTODIAN TAB ──
    def _build_custodian_tab(self):
        p=self.tab_cust
        tk.Label(p,text="Groups all PDFs by custodian → one email per custodian with all their deal notes attached.",
                 bg=BG,fg="#607080",font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=(10,0))
        self.btn_send_all_cust=tk.Button(p,text="✉  Send ALL Custodian Emails",
            command=self._cust_send_all,bg=GREEN_DARK,fg=WHITE,font=("Segoe UI",11,"bold"),
            relief="flat",padx=16,pady=10,cursor="hand2",state="disabled")
        self.btn_send_all_cust.pack(fill="x",padx=16,pady=(8,4))
        self.lbl_cust_hint=tk.Label(p,text="Load a folder above to begin.",bg=BG,fg="#607080",font=("Segoe UI",9))
        self.lbl_cust_hint.pack(anchor="w",padx=16,pady=(0,8))
        canvas=tk.Canvas(p,bg=BG,highlightthickness=0)
        sb=ttk.Scrollbar(p,orient="vertical",command=canvas.yview)
        self.cust_body=tk.Frame(canvas,bg=BG)
        self.cust_body.bind("<Configure>",lambda e:canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0),window=self.cust_body,anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left",fill="both",expand=True,padx=(16,0))
        sb.pack(side="right",fill="y",pady=4)

    def _render_custodian_tab(self):
        for w in self.cust_body.winfo_children(): w.destroy()
        self.cust_btn_map={}
        groups={}
        for item in self.deal_items: groups.setdefault(item["custodian"],[]).append(item)
        for code,items in sorted(groups.items()):
            routing=CUSTODIAN_ROUTING.get(code)
            head_color=FBC_MID if routing else RED_DARK
            card=tk.Frame(self.cust_body,bg=WHITE,pady=0,padx=0)
            card.pack(fill="x",padx=4,pady=(0,10))
            head=tk.Frame(card,bg=head_color,pady=7,padx=12); head.pack(fill="x")
            label=routing["label"] if routing else "UNKNOWN CUSTODIAN"
            count=len(items)
            tk.Label(head,text=f"{code}  —  {label}",bg=head_color,fg=WHITE,
                     font=("Segoe UI",10,"bold")).pack(side="left")
            tk.Label(head,text=f"{count} deal note{'s' if count>1 else ''}",
                     bg=head_color,fg=WHITE,font=("Segoe UI",9)).pack(side="right")
            inner=tk.Frame(card,bg=WHITE,padx=12,pady=8); inner.pack(fill="x")
            for it in items:
                tk.Label(inner,text=f"  📄 {it['fname']}",bg=WHITE,fg="#2D3748",font=("Segoe UI",9)).pack(anchor="w")
            if routing:
                subj=f"DEAL NOTE{'S' if count>1 else ''} — {datetime.now().strftime('%d %b %Y')}"
                tk.Label(inner,text=f"Subject: {subj}",bg=WHITE,fg="#607080",font=("Segoe UI",8,"italic")).pack(anchor="w",pady=(6,0))
                tk.Label(inner,text=f"To: {'; '.join(routing['to'])}",bg=WHITE,fg="#607080",font=("Segoe UI",8)).pack(anchor="w")
                btn=tk.Button(inner,text=f"✉  Open in Outlook  ({count} file{'s' if count>1 else ''} attached)",
                    command=lambda c=code:self._cust_send_one(c),bg=FBC_MID,fg=WHITE,relief="flat",
                    font=("Segoe UI",9,"bold"),cursor="hand2",padx=10,pady=6)
                btn.pack(anchor="w",pady=(8,0)); self.cust_btn_map[code]=btn
            else:
                tk.Label(inner,text="⚠  No routing configured for this custodian code.",
                         bg=WHITE,fg=RED_DARK,font=("Segoe UI",9)).pack(anchor="w")
        known=sum(1 for c in groups if c in CUSTODIAN_ROUTING)
        self.btn_send_all_cust.config(
            state="normal" if known else "disabled",
            text=f"✉  Send ALL {known} Custodian Email{'s' if known!=1 else ''} in Outlook")
        self.lbl_cust_hint.config(
            text=f"✅  {len(self.deal_items)} deal note(s) across {len(groups)} custodian(s).",fg=GREEN_DARK)

    def _cust_send_one(self,code):
        routing=CUSTODIAN_ROUTING.get(code)
        if not routing: messagebox.showwarning("Unknown",f"No routing for {code}."); return
        items=[it for it in self.deal_items if it["custodian"]==code]
        count=len(items)
        subj=f"DEAL NOTE{'S' if count>1 else ''} — {datetime.now().strftime('%d %b %Y')}"
        body=CUSTODIAN_BODY_MULTI if count>1 else CUSTODIAN_BODY_SINGLE
        try:
            open_outlook(routing["to"],routing["cc"],subj,body,[it["path"] for it in items])
            if code in self.cust_btn_map: self.cust_btn_map[code].config(text="✅  Sent",bg="#2E7D32")
        except ImportError as e: messagebox.showerror("pywin32 not installed",str(e))
        except Exception as e: messagebox.showerror("Outlook Error",str(e))

    def _cust_send_all(self):
        codes=sorted(set(it["custodian"] for it in self.deal_items if it["custodian"] in CUSTODIAN_ROUTING))
        if not codes: messagebox.showinfo("Nothing","No known custodians found."); return
        if not messagebox.askyesno("Send All",f"Open {len(codes)} Outlook window(s), one per custodian?\n\nContinue?"): return
        for code in codes: self._cust_send_one(code)
        messagebox.showinfo("Done",f"✅  {len(codes)} Outlook window(s) opened.")

    # ── CLIENT TAB ──
    def _build_client_tab(self):
        p=self.tab_client
        tk.Label(p,text="Sends one email per client with all their deal notes attached. Add emails via 'Manage Client Contacts'.",
                 bg=BG,fg="#607080",font=("Segoe UI",9)).pack(anchor="w",padx=16,pady=(10,0))
        self.btn_send_everything=tk.Button(p,text="✉  Send ALL Emails (Custodian + Client)",
            command=self._send_everything,bg=FBC_DARK,fg=WHITE,font=("Segoe UI",11,"bold"),
            relief="flat",padx=16,pady=10,cursor="hand2",state="disabled")
        self.btn_send_everything.pack(fill="x",padx=16,pady=(8,2))
        self.btn_send_all_client=tk.Button(p,text="✉  Send ALL Client Emails",
            command=self._client_send_all,bg=GREEN_DARK,fg=WHITE,font=("Segoe UI",11,"bold"),
            relief="flat",padx=16,pady=10,cursor="hand2",state="disabled")
        self.btn_send_all_client.pack(fill="x",padx=16,pady=(2,4))
        self.lbl_client_hint=tk.Label(p,text="Load a folder above to begin.",bg=BG,fg="#607080",font=("Segoe UI",9))
        self.lbl_client_hint.pack(anchor="w",padx=16,pady=(0,4))
        wrap=tk.Frame(p,bg=BG); wrap.pack(fill="both",expand=True,padx=16,pady=4)
        canvas=tk.Canvas(wrap,bg=BG,highlightthickness=0)
        sb=ttk.Scrollbar(wrap,orient="vertical",command=canvas.yview)
        self.client_body=tk.Frame(canvas,bg=BG)
        self.client_body.bind("<Configure>",lambda e:canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0),window=self.client_body,anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left",fill="both",expand=True); sb.pack(side="right",fill="y")

    def _client_groups(self):
        groups={}
        for item in self.deal_items:
            contact = find_contact(self.contacts, item["client"])
            email   = contact.get("email", "")
            # Resolve the canonical key: prefer the saved contact name so that
            # "MAKWASHA TANYARADZWA" (from filename) maps to "TANYARADZWA MAKWASHA"
            # (as saved in contacts), keeping the UI consistent.
            key = item["client"]
            for saved in self.contacts:
                if _names_match(item["client"], saved):
                    key = saved; break
            # Legacy prefix fallback
            if key == item["client"]:
                for saved in self.contacts:
                    if item["client"].startswith(saved) or saved.startswith(item["client"]):
                        key = saved; break
            if key not in groups:
                groups[key] = {"email": email, "items": [], "sent": False}
            groups[key]["items"].append(item)
            groups[key]["email"] = groups[key]["email"] or email
        for g in groups.values():
            g["status"] = "ready" if g["email"] else "missing"
        return groups

    def _render_client_tab(self):
        for w in self.client_body.winfo_children(): w.destroy()
        self.client_group_btns={}
        if not self.deal_items: return
        groups=self._client_groups()
        headers=["#","Client Name","Files","Custodian","Client Email","Status","Action"]
        widths=[3,22,18,10,26,9,10]
        hdr=tk.Frame(self.client_body,bg=FBC_DARK); hdr.pack(fill="x")
        for h,w in zip(headers,widths):
            tk.Label(hdr,text=h,bg=FBC_DARK,fg=WHITE,font=("Segoe UI",8,"bold"),
                     width=w,anchor="w",padx=4,pady=6).pack(side="left")
        for i,(client_name,grp) in enumerate(groups.items()):
            email=grp["email"]; status=grp["status"]; count=len(grp["items"])
            custodians=", ".join(sorted(set(it["custodian"] for it in grp["items"])))
            bg="#F8FBFF" if i%2==0 else WHITE
            row=tk.Frame(self.client_body,bg=bg); row.pack(fill="x")
            sc=GREEN_DARK if status=="ready" else RED_DARK
            st="✅ Ready" if status=="ready" else "⚠ Missing"
            file_label=f"{count} file{'s' if count>1 else ''}"
            for v,w in zip([str(i+1),client_name[:20],file_label,custodians[:10],email[:24] or "—",st],widths):
                tk.Label(row,text=v,bg=bg,fg=(sc if v==st else "#2D3748"),
                         font=("Segoe UI",8),width=w,anchor="w",padx=4,pady=5).pack(side="left")
            sent=grp.get("sent",False)
            btn=tk.Button(row,text="✅ Sent" if sent else "✉ Send",
                bg="#2E7D32" if sent else FBC_MID,fg=WHITE,relief="flat",
                font=("Segoe UI",8,"bold"),cursor="hand2",padx=6,pady=3,
                command=lambda cn=client_name:self._client_send_group(cn))
            btn.pack(side="left",padx=4); self.client_group_btns[client_name]=btn
        ready=sum(1 for g in groups.values() if g["status"]=="ready")
        missing=len(groups)-ready
        self.btn_send_all_client.config(
            state="normal" if ready else "disabled",
            text=f"✉  Send ALL {ready} Client Email{'s' if ready!=1 else ''} in Outlook")
        self.btn_send_everything.config(
            state="normal" if ready or any(it["custodian"] in CUSTODIAN_ROUTING for it in self.deal_items) else "disabled")
        self.lbl_client_hint.config(
            text=(f"✅  All {ready} clients matched." if not missing
                  else f"⚠  {missing} client(s) missing — click 'Manage Client Contacts'."),
            fg=GREEN_DARK if not missing else RED_DARK)

    def _client_send_group(self,client_name):
        groups=self._client_groups(); grp=groups.get(client_name)
        if not grp: return
        if grp["status"]=="missing":
            messagebox.showwarning("Missing Email",f"No email for '{client_name}'.\n\nUse 'Manage Client Contacts'."); return
        count=len(grp["items"])
        subj=f"DEAL{'S' if count>1 else ''} CONFIRMATION"
        body=(CLIENT_BODY_MULTI if count>1 else CLIENT_BODY_SINGLE).format(client=client_name.title())
        paths=[it["path"] for it in grp["items"]]
        try:
            open_outlook([grp["email"]],CLIENT_CC,subj,body,paths)
            grp["sent"]=True
            if client_name in self.client_group_btns:
                self.client_group_btns[client_name].config(text="✅ Sent",bg="#2E7D32")
        except ImportError as e: messagebox.showerror("pywin32 not installed",str(e))
        except Exception as e: messagebox.showerror("Outlook Error",str(e))

    def _client_send_all(self):
        groups=self._client_groups(); ready=[cn for cn,g in groups.items() if g["status"]=="ready"]
        if not ready: messagebox.showinfo("Nothing","No clients with emails found."); return
        if not messagebox.askyesno("Send All Client Emails",f"Open {len(ready)} Outlook window(s), one per client?\n\nContinue?"): return
        for cn in ready: self._client_send_group(cn)
        messagebox.showinfo("Done",f"✅  {len(ready)} Outlook window(s) opened.")

    def _send_everything(self):
        cust_codes=sorted(set(it["custodian"] for it in self.deal_items if it["custodian"] in CUSTODIAN_ROUTING))
        groups=self._client_groups(); ready_clients=[cn for cn,g in groups.items() if g["status"]=="ready"]
        total=len(cust_codes)+len(ready_clients)
        if total==0: messagebox.showinfo("Nothing","No emails to send."); return
        if not messagebox.askyesno("Send ALL Emails",
            f"This will open {total} Outlook window(s):\n"
            f"  • {len(cust_codes)} custodian email(s)\n"
            f"  • {len(ready_clients)} client email(s)\n\nContinue?"): return
        for code in cust_codes: self._cust_send_one(code)
        for cn in ready_clients: self._client_send_group(cn)
        messagebox.showinfo("Done",f"✅  {total} Outlook window(s) opened.")

    def _open_contacts(self):
        ContactsDialog(self,self.contacts,self._on_contacts_saved)

    def _on_contacts_saved(self,new_contacts):
        self.contacts=new_contacts
        if self.deal_items: self._render_client_tab()

    def _pick_folder(self):
        folder=filedialog.askdirectory(title="Select folder containing deal note PDFs")
        if not folder: return
        pdfs=sorted(f for f in os.listdir(folder) if f.lower().endswith(".pdf"))
        if not pdfs: messagebox.showwarning("No PDFs","No PDF files found in that folder."); return
        self.pdf_folder=folder
        self.lbl_folder.config(text=f"📂  {os.path.basename(folder)}",fg=FBC_DARK)
        self.lbl_found.config(text=f"Scanning {len(pdfs)} PDF(s)…")
        self.lbl_file_list.config(text="")
        self._disable_send_buttons()
        self.btn_clear.config(state="normal")
        threading.Thread(target=self._scan,args=(folder,pdfs),daemon=True).start()

    def _pick_individual_files(self):
        paths=filedialog.askopenfilenames(
            title="Select Deal Note PDF(s)",
            filetypes=[("PDF files","*.pdf"),("All files","*.*")])
        if not paths: return
        pdf_paths=list(paths)
        # add to existing items (don't wipe, let user accumulate)
        already={it["path"] for it in self.deal_items}
        new_paths=[p for p in pdf_paths if p not in already]
        if not new_paths:
            messagebox.showinfo("No New Files","All selected files are already loaded."); return
        self.lbl_found.config(text=f"Scanning {len(new_paths)} new PDF(s)…")
        self.btn_clear.config(state="normal")
        # show short names
        names=", ".join(os.path.basename(p) for p in new_paths[:3])
        if len(new_paths)>3: names+=f" +{len(new_paths)-3} more"
        self.lbl_folder.config(text=f"📄  {names}",fg=FBC_DARK)
        threading.Thread(target=self._scan_files,args=(new_paths,),daemon=True).start()

    def _clear_uploads(self):
        if not self.deal_items: return
        if not messagebox.askyesno("Clear All Uploads",
            f"Remove all {len(self.deal_items)} loaded deal note(s) and start fresh?\n\nThis does NOT delete the files from disk."):
            return
        self.deal_items=[]
        self.pdf_folder=""
        self.lbl_folder.config(text="No files loaded",fg="#8096B0")
        self.lbl_found.config(text="")
        self.lbl_file_list.config(text="")
        self.btn_clear.config(state="disabled")
        self._disable_send_buttons()
        # clear rendered tabs
        for w in self.cust_body.winfo_children(): w.destroy()
        for w in self.client_body.winfo_children(): w.destroy()
        self.lbl_cust_hint.config(text="Load files above to begin.",fg="#607080")
        self.lbl_client_hint.config(text="Load files above to begin.",fg="#607080")
        self.btn_send_all_cust.config(state="disabled",text="✉  Send ALL Custodian Emails")
        self.btn_send_all_client.config(state="disabled",text="✉  Send ALL Client Emails")
        self.btn_send_everything.config(state="disabled")

    def _disable_send_buttons(self):
        self.btn_send_all_cust.config(state="disabled")
        self.btn_send_all_client.config(state="disabled")
        self.btn_send_everything.config(state="disabled")

    def _scan(self,folder,pdfs):
        items=[]
        for fname in pdfs:
            path=os.path.join(folder,fname)
            items.append({
                "fname":fname,"path":path,
                "client":parse_client_name_from_filename(fname),
                "custodian":parse_custodian_from_pdf(path) or "UNKNOWN",
                "deal_info":parse_deal_info_from_pdf(path),"sent":False,
            })
        self.deal_items=items
        self._after_scan(len(items))

    def _scan_files(self,paths):
        """Scan individually selected files and append to existing deal_items."""
        new_items=[]
        for path in paths:
            fname=os.path.basename(path)
            new_items.append({
                "fname":fname,"path":path,
                "client":parse_client_name_from_filename(fname),
                "custodian":parse_custodian_from_pdf(path) or "UNKNOWN",
                "deal_info":parse_deal_info_from_pdf(path),"sent":False,
            })
        self.deal_items.extend(new_items)
        self._after_scan(len(self.deal_items))

    def _after_scan(self,total):
        self.after(0,self._render_custodian_tab)
        self.after(0,self._render_client_tab)
        self.after(0,lambda:self.lbl_found.config(text=f"✅  {total} PDF(s) loaded"))

# ════════════════════════════════════════════════════════════════════════════
#  MAIN APP SHELL  (sidebar + page switcher)
# ════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"FBC Suite  v{VERSION}")
        self.state("zoomed")
        self.configure(bg=SIDEBAR_BG)
        self._active_page = None
        self._build()

    def _build(self):
        # ── sidebar ──
        sidebar = tk.Frame(self, bg=SIDEBAR_BG, width=200)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)

        # logo area
        logo = tk.Frame(sidebar, bg=SIDEBAR_BG, pady=20)
        logo.pack(fill="x")
        tk.Label(logo, text="FBC", bg=FBC_ACCENT, fg=WHITE,
                 font=("Segoe UI",16,"bold"), padx=10, pady=6).pack()
        tk.Label(logo, text="Suite", bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                 font=("Segoe UI",10)).pack(pady=(4,0))

        # divider
        tk.Frame(sidebar, bg=FBC_MID, height=1).pack(fill="x", padx=16, pady=(0,10))

        # nav buttons
        self.nav_buttons = {}
        nav_items = [
            ("📊", "Converter", "converter"),
            ("✉",  "Deal Note\nEmailer", "emailer"),
        ]
        for icon, label, key in nav_items:
            btn = tk.Button(sidebar,
                text=f"{icon}\n{label}",
                command=lambda k=key: self._switch(k),
                bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                activebackground=SIDEBAR_ACTIVE, activeforeground=WHITE,
                font=("Segoe UI",10,"bold"), relief="flat",
                cursor="hand2", pady=16, width=16,
                justify="center")
            btn.pack(fill="x", padx=8, pady=2)
            btn.bind("<Enter>", lambda e, b=btn, k=key: b.config(
                bg=SIDEBAR_ACTIVE if self._active_page==k else SIDEBAR_HOVER))
            btn.bind("<Leave>", lambda e, b=btn, k=key: b.config(
                bg=SIDEBAR_ACTIVE if self._active_page==k else SIDEBAR_BG))
            self.nav_buttons[key] = btn

        # bottom version label
        tk.Label(sidebar, text=f"v{VERSION}", bg=SIDEBAR_BG, fg="#2A4A6A",
                 font=("Segoe UI",8)).pack(side="bottom", pady=10)

        # ── content area ──
        self.content = tk.Frame(self, bg=BG)
        self.content.pack(side="left", fill="both", expand=True)

        # create pages (hidden initially)
        self.pages = {
            "converter": SarestockPage(self.content),
            "emailer":   EmailerPage(self.content),
        }

        # start on converter
        self._switch("converter")

    def _switch(self, key):
        # hide all pages
        for page in self.pages.values():
            page.pack_forget()
        # show selected
        self.pages[key].pack(fill="both", expand=True)
        self._active_page = key
        # update sidebar button highlights
        for k, btn in self.nav_buttons.items():
            if k == key:
                btn.config(bg=SIDEBAR_ACTIVE, fg=WHITE)
            else:
                btn.config(bg=SIDEBAR_BG, fg=SIDEBAR_TEXT)

# ════════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    check_and_apply_update()
    app = App()
    app.mainloop()

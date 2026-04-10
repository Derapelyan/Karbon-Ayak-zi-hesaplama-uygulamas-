"""
UNSPED Karbon Ayak İzi Dashboard v2
=====================================
python dashboard.py
"""
import sys, os, threading, io, json, re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

# ── Renkler & Fontlar ──────────────────────────────────────────────
C = {
    "bg":"#F5F5F0","sidebar":"#1F4E79","sidebar_hl":"#2E75B6",
    "white":"#FFFFFF","text":"#2C2C2A","muted":"#888780","border":"#D3D1C7",
    "blue":"#1F4E79","green":"#1D9E75","amber":"#BA7517","red":"#A32D2D",
    "s1":"#2E75B6","s2":"#1D9E75","s3":"#BA7517","card":"#FFFFFF",
    "success":"#E2EFDA","warning":"#FFF2CC","danger":"#FCEBEB","impact":"#FCE4D6",
}
FONT       = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_LARGE = ("Segoe UI", 22, "bold")
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)

thin = None  # lazy import border

# ── TC Enerji Bakanlığı EVÇED Verileri ────────────────────────────
EVCED_FACTORS = {
    2019: {"uretim": 0.481, "tuketim": 0.488, "kaynak": "ETKB EVÇED 2019"},
    2020: {"uretim": 0.452, "tuketim": 0.458, "kaynak": "ETKB EVÇED 2020"},
    2021: {"uretim": 0.434, "tuketim": 0.439, "kaynak": "ETKB EVÇED 2021"},
    2022: {"uretim": 0.442, "tuketim": 0.442, "kaynak": "ETKB EVÇED 2022 (6 Ara 2024)"},
    2023: {"uretim": 0.434, "tuketim": 0.434, "kaynak": "ETKB EVÇED 2023"},
}
EVCED_URL = "https://enerji.gov.tr/evced-cevre-ve-iklim-iklim-degisikligi-emisyon-faktorleri"

# ── IPCC AR6 GWP-100 Tablosu ─────────────────────────────────────
GWP_TABLE = [
    ("CO2",        "Karbondioksit",         1,      "IPCC AR6 2021"),
    ("CH4",        "Metan",                 29.8,   "IPCC AR6 2021"),
    ("N2O",        "Diazotmonoksit",        273,    "IPCC AR6 2021"),
    ("SF6",        "Kükürt hekzaflorür",    25200,  "IPCC AR6 2021"),
    ("R134A",      "HFC-134a",              1526,   "IPCC AR6 2021"),
    ("R32",        "HFC-32",                771,    "IPCC AR6 2021"),
    ("R410A",      "HFC-410A (karışım)",    2088,   "IPCC AR6 2021"),
    ("R407C",      "HFC-407C (karışım)",    1774,   "IPCC AR6 2021"),
    ("R404A",      "HFC-404A (karışım)",    3922,   "IPCC AR6 2021"),
    ("R507A",      "HFC-507A (karışım)",    3985,   "IPCC AR6 2021"),
    ("R22",        "HCFC-22",               1960,   "IPCC AR6 2021"),
    ("R290",       "Propan",                0,      "IPCC AR6 2021"),
    ("R600A",      "İzobütan",              4,      "IPCC AR6 2021"),
    ("FM200",      "HFC-227ea",             3220,   "IPCC AR6 2021"),
    ("HFC236fa",   "HFC-236fa",             8060,   "IPCC AR6 2021"),
    ("FK-5-1-12",  "Novec 1230",            1,      "IPCC AR6 2021"),
    ("R448A",      "HFC-448A",              1387,   "IPCC AR6 2021"),
    ("R454B",      "HFC-454B",              466,    "IPCC AR6 2021"),
    ("R513A",      "HFC-513A",              631,    "IPCC AR6 2021"),
]

# ── IPCC Yakıt Emisyon Faktörleri ─────────────────────────────────
EF_TABLE = [
    ("Doğal Gaz",     "m³",  56.1,  0.005,  0.0001, 48.0, 0.000717, "IPCC 2006 GL Vol.2 T2.2"),
    ("Dizel",         "L",   74.1,  0.010,  0.0006, 43.0, 0.000832, "IPCC 2006 GL Vol.2 T2.2"),
    ("Benzin/Petrol", "L",   69.3,  0.010,  0.0006, 44.3, 0.000745, "IPCC 2006 GL Vol.2 T2.2"),
    ("LPG",           "L",   63.1,  0.010,  0.0006, 47.3, 0.000540, "IPCC 2006 GL Vol.2 T2.2"),
    ("Fuel Oil",      "L",   77.4,  0.010,  0.0006, 40.4, 0.000890, "IPCC 2006 GL Vol.2 T2.2"),
    ("Kömür (linyit)","ton", 101.0, 0.010,  0.0015, 11.9, 1.0,      "IPCC 2006 GL Vol.2 T2.2"),
]

# ── IPCC NCV Tablosu ─────────────────────────────────────────────
NCV_TABLE = [
    ("Doğal Gaz",      "m³",  48.0,  "TJ/Gg", 0.000717, "ton/m³",  "IPCC 2006 GL Vol.2 T1.2"),
    ("Dizel",          "L",   43.0,  "TJ/Gg", 0.000832, "ton/L",   "IPCC 2006 GL Vol.2 T1.2"),
    ("Benzin/Petrol",  "L",   44.3,  "TJ/Gg", 0.000745, "ton/L",   "IPCC 2006 GL Vol.2 T1.2"),
    ("LPG",            "L",   47.3,  "TJ/Gg", 0.000540, "ton/L",   "IPCC 2006 GL Vol.2 T1.2"),
    ("Fuel Oil",       "L",   40.4,  "TJ/Gg", 0.000890, "ton/L",   "IPCC 2006 GL Vol.2 T1.2"),
    ("Kömür (linyit)", "ton", 11.9,  "TJ/Gg", 1.0,      "ton/ton", "IPCC 2006 GL Vol.2 T1.2"),
    ("Odun/Biyokütle", "ton", 15.6,  "TJ/Gg", 1.0,      "ton/ton", "IPCC 2006 GL Vol.2 T1.2"),
]


# ══════════════════════════════════════════════════════════════════
# YARDIMCI FONKSİYONLAR
# ══════════════════════════════════════════════════════════════════

def make_treeview(parent, columns, col_widths, height=12):
    frame = tk.Frame(parent, bg=C["bg"])
    frame.pack(fill="both", expand=True)
    style = ttk.Style()
    style.configure("Custom.Treeview", font=FONT, rowheight=26)
    style.configure("Custom.Treeview.Heading", font=FONT_BOLD)
    tree = ttk.Treeview(frame, columns=columns, show="headings",
                        height=height, style="Custom.Treeview")
    for col, w in zip(columns, col_widths):
        tree.heading(col, text=col)
        tree.column(col, width=w, anchor="center")
    sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    return tree

def card(parent, title=None, pady=12):
    outer = tk.Frame(parent, bg=C["card"],
                     highlightbackground=C["border"], highlightthickness=1)
    outer.pack(fill="x", pady=(0, pady))
    inner = tk.Frame(outer, bg=C["card"], padx=16, pady=12)
    inner.pack(fill="both")
    if title:
        tk.Label(inner, text=title, font=FONT_BOLD,
                 bg=C["card"], fg=C["text"]).pack(anchor="w", pady=(0,6))
    return inner

def section_label(parent, text, color=None):
    tk.Label(parent, text=text, font=FONT_BOLD,
             bg=C["bg"], fg=color or C["text"]).pack(anchor="w", pady=(12,4))


# ══════════════════════════════════════════════════════════════════
# DB YARDIMCILARI
# ══════════════════════════════════════════════════════════════════

def db_get_scope_totals():
    try:
        from db.connection import SessionLocal
        from db.models import AuditLog
        with SessionLocal() as s:
            logs = s.query(AuditLog).filter_by(
                action="excel_import", status="success"
            ).order_by(AuditLog.id.desc()).all()
            years = {}
            for log in logs:
                notes = log.notes or ""
                for m in re.finditer(
                    r'(\d{4})\s*\|\s*K1=([\d.]+)\s*K2=([\d.]+)\s*K3=([\d.]+)\s*Toplam=([\d.]+)',
                    notes
                ):
                    y = m.group(1)
                    if y not in years:
                        years[y] = {
                            "k1":float(m.group(2)), "k2":float(m.group(3)),
                            "k3":float(m.group(4)), "total":float(m.group(5))
                        }
        return years
    except:
        return {}

def db_get_revisions():
    try:
        from db.connection import SessionLocal
        from db.models import AuditLog
        with SessionLocal() as s:
            logs = s.query(AuditLog).filter_by(
                action="excel_import", status="success"
            ).order_by(AuditLog.id.desc()).all()
            revs = []
            year_counters = {}
            for log in reversed(logs):
                notes = log.notes or ""
                for m in re.finditer(
                    r'(\d{4})\s*\|\s*K1=([\d.]+)\s*K2=([\d.]+)\s*K3=([\d.]+)\s*Toplam=([\d.]+)',
                    notes
                ):
                    y = m.group(1)
                    year_counters[y] = year_counters.get(y, 0) + 1
                    rev_id = f"{y}-R{year_counters[y]}"
                    ts = str(log.created_at)[:16] if log.created_at else "—"
                    revs.append({
                        "id": rev_id, "year": y, "ts": ts,
                        "k1": float(m.group(2)), "k2": float(m.group(3)),
                        "k3": float(m.group(4)), "total": float(m.group(5))
                    })
            return list(reversed(revs))
    except:
        return []

def db_get_category_breakdown():
    """Kapsam 3 alt kategorilerini getirir."""
    try:
        from db.connection import SessionLocal
        from db.models import (AuditLog, FreightEmission, EmployeeCommuting,
                               PurchasedGoods, CapitalGoods, ElectricityConsumption,
                               ReportingPeriod)
        results = {}
        with SessionLocal() as s:
            periods = s.query(ReportingPeriod).all()
            for p in periods:
                y = str(p.year)
                results[y] = {
                    "3.1": 0, "3.2": 0, "3.3": 0, "3.4": 0,
                    "3.5": 0, "4.1": 0, "4.2": 0, "4.3": 0,
                    "4.4": 0, "4.5": 0, "6.1": 0,
                }
                for row in s.query(FreightEmission).filter_by(period_id=p.id).all():
                    cat = row.category_code or "3.1"
                    key = cat[:3] if cat[:3] in results[y] else "3.1"
                    results[y][key] = results[y].get(key, 0) + (row.co2e_total or 0)
                for row in s.query(EmployeeCommuting).filter_by(period_id=p.id).all():
                    cat = row.category_code or "3.3"
                    key = cat[:3] if cat[:3] in results[y] else "3.3"
                    results[y][key] = results[y].get(key, 0) + (row.co2e_total or 0)
                for row in s.query(PurchasedGoods).filter_by(period_id=p.id).all():
                    cat = row.category_code or "4.1"
                    key = cat[:3] if cat[:3] in results[y] else "4.1"
                    results[y][key] = results[y].get(key, 0) + (row.co2e_total or 0)
                for row in s.query(CapitalGoods).filter_by(period_id=p.id).all():
                    results[y]["4.2"] = results[y].get("4.2", 0) + (row.co2e_total or 0)
        return results
    except:
        return {}


# ══════════════════════════════════════════════════════════════════
# ANA UYGULAMA
# ══════════════════════════════════════════════════════════════════

class UNSPEDDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("UNSPED Karbon Ayak İzi Dashboard v2")
        self.geometry("1350x850")
        self.minsize(1100, 700)
        self.configure(bg=C["bg"])
        self.excel_path  = tk.StringVar(value="")
        self.status_var  = tk.StringVar(value="Hazır")
        self._build_ui()
        self._load_db_summary()

    # ── UI İnşa ────────────────────────────────────────────────────
    def _build_ui(self):
        self.sidebar = tk.Frame(self, bg=C["sidebar"], width=230)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self.content = tk.Frame(self, bg=C["bg"])
        self.content.pack(side="left", fill="both", expand=True)
        self._build_sidebar()
        self._build_statusbar()
        self._show_page("home")

    def _build_sidebar(self):
        lf = tk.Frame(self.sidebar, bg=C["sidebar"], pady=20)
        lf.pack(fill="x")
        tk.Label(lf, text="🌍", font=("Segoe UI",28), bg=C["sidebar"], fg="white").pack()
        tk.Label(lf, text="UNSPED", font=("Segoe UI",13,"bold"), bg=C["sidebar"], fg="white").pack()
        tk.Label(lf, text="Karbon Dashboard v2", font=("Segoe UI",8), bg=C["sidebar"], fg="#9FC5E8").pack()
        ttk.Separator(self.sidebar).pack(fill="x", padx=16, pady=4)

        self.nav_btns = {}
        menu = [
            ("home",    "🏠", "Ana Sayfa"),
            ("import",  "📥", "Veri İmport"),
            ("update",  "🔄", "Faktör Güncelle"),
            ("results", "📊", "Sonuçlar & Grafik"),
            ("refs",    "📚", "Referans Tablolar"),
            ("log",     "📋", "Log & Revizyonlar"),
        ]
        for page, icon, label in menu:
            btn = tk.Button(
                self.sidebar, text=f"  {icon}  {label}",
                font=FONT, bg=C["sidebar"], fg="white",
                activebackground=C["sidebar_hl"], activeforeground="white",
                relief="flat", anchor="w", padx=16, pady=10, cursor="hand2",
                command=lambda p=page: self._show_page(p)
            )
            btn.pack(fill="x")
            self.nav_btns[page] = btn

        tk.Label(self.sidebar, text="v2.0 • 2025", font=FONT_SMALL,
                 bg=C["sidebar"], fg="#9FC5E8").pack(side="bottom", pady=12)

    def _build_statusbar(self):
        sb = tk.Frame(self.content, bg=C["border"], height=28)
        sb.pack(side="bottom", fill="x")
        tk.Label(sb, textvariable=self.status_var, font=FONT_SMALL,
                 bg=C["border"], fg=C["text"], padx=12).pack(side="left")

    def _show_page(self, page):
        for p, btn in self.nav_btns.items():
            btn.configure(bg=C["sidebar_hl"] if p==page else C["sidebar"])
        for w in self.content.winfo_children():
            if not (isinstance(w, tk.Frame) and w.cget("height")==28):
                w.destroy()
        pages = {
            "home":    self._page_home,
            "import":  self._page_import,
            "update":  self._page_update,
            "results": self._page_results,
            "refs":    self._page_refs,
            "log":     self._page_log,
        }
        pages.get(page, self._page_home)()

    # ── Scrollable frame helper ────────────────────────────────────
    def _scrollable(self, parent):
        canvas = tk.Canvas(parent, bg=C["bg"], highlightthickness=0)
        vsb    = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        frame  = tk.Frame(canvas, bg=C["bg"])
        frame.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        return frame

    # ══════════════════════════════════════════════════════════════
    # ANA SAYFA
    # ══════════════════════════════════════════════════════════════
    def _page_home(self):
        outer = tk.Frame(self.content, bg=C["bg"])
        outer.pack(fill="both", expand=True)
        frame = self._scrollable(outer)
        pad = tk.Frame(frame, bg=C["bg"])
        pad.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(pad, text="UNSPED Karbon Ayak İzi", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(pad, text="GHG Protocol tabanlı raporlama sistemi",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,16))

        # KPI kartları
        kpi_frame = tk.Frame(pad, bg=C["bg"])
        kpi_frame.pack(fill="x", pady=(0,16))
        self.kpi_vars = {}
        kpis = [
            ("toplam",  "Toplam CO2e",  C["blue"]),
            ("kapsam1", "Kapsam 1",     C["s1"]),
            ("kapsam2", "Kapsam 2",     C["s2"]),
            ("kapsam3", "Kapsam 3",     C["s3"]),
        ]
        for key, label, color in kpis:
            c = tk.Frame(kpi_frame, bg=C["card"],
                         highlightbackground=C["border"], highlightthickness=1)
            c.pack(side="left", fill="both", expand=True, padx=(0,12))
            tk.Frame(c, bg=color, height=4).pack(fill="x")
            inner = tk.Frame(c, bg=C["card"], padx=16, pady=14)
            inner.pack(fill="both")
            tk.Label(inner, text=label, font=FONT_SMALL,
                     bg=C["card"], fg=C["muted"]).pack(anchor="w")
            var = tk.StringVar(value="—")
            self.kpi_vars[key] = var
            tk.Label(inner, textvariable=var, font=FONT_LARGE,
                     bg=C["card"], fg=color).pack(anchor="w")
            tk.Label(inner, text="ton CO2e", font=FONT_SMALL,
                     bg=C["card"], fg=C["muted"]).pack(anchor="w")

        # Hızlı işlemler
        section_label(pad, "Hızlı İşlemler")
        bf = tk.Frame(pad, bg=C["bg"])
        bf.pack(fill="x")
        for text, color, cmd in [
            ("📥  Excel Seç & İmport", C["blue"],  lambda: self._show_page("import")),
            ("🔄  Faktörleri Güncelle", C["green"], lambda: self._show_page("update")),
            ("📊  Sonuçları Gör",       C["amber"], lambda: self._show_page("results")),
            ("📚  Referans Tablolar",   "#6B5B95",  lambda: self._show_page("refs")),
        ]:
            tk.Button(bf, text=text, font=FONT_BOLD, bg=color, fg="white",
                      relief="flat", padx=16, pady=8, cursor="hand2",
                      activebackground=color, command=cmd
                      ).pack(side="left", padx=(0,10))

        # DB Özeti
        section_label(pad, "Veritabanı Özeti")
        self.db_text = tk.Text(pad, height=9, font=FONT_MONO,
                               bg=C["card"], fg=C["text"], relief="flat",
                               state="disabled",
                               highlightbackground=C["border"], highlightthickness=1)
        self.db_text.pack(fill="x")
        self._load_db_summary()

    def _load_db_summary(self):
        try:
            from db.connection import SessionLocal
            from db.models import Company, ReportingPeriod, EmissionFactor, AuditLog
            with SessionLocal() as s:
                company  = s.query(Company).filter_by(name="UNSPED").first()
                periods  = s.query(ReportingPeriod).all()
                ef_count = s.query(EmissionFactor).count()
                log_count= s.query(AuditLog).count()
                last_log = s.query(AuditLog).filter_by(
                    action="excel_import").order_by(AuditLog.id.desc()).first()
                lines = []
                if company:
                    lines.append(f"  Sirket     : {company.name}")
                lines.append(f"  Donemler   : {[p.year for p in periods]}")
                lines.append(f"  EF Kayitlari: {ef_count} adet")
                lines.append(f"  Toplam Log : {log_count} islem")
                if last_log:
                    lines.append(f"  Son Import : {str(last_log.created_at)[:19]}")
                    years = db_get_scope_totals()
                    if years:
                        latest = max(years.keys())
                        d = years[latest]
                        lines.append(f"\n  [{latest}] K1={d['k1']:.1f} | K2={d['k2']:.1f} | K3={d['k3']:.1f}")
                        lines.append(f"  [{latest}] TOPLAM: {d['total']:.2f} ton CO2e")
                        if hasattr(self, 'kpi_vars'):
                            self.kpi_vars['toplam'].set(f"{d['total']:.1f}")
                            self.kpi_vars['kapsam1'].set(f"{d['k1']:.1f}")
                            self.kpi_vars['kapsam2'].set(f"{d['k2']:.1f}")
                            self.kpi_vars['kapsam3'].set(f"{d['k3']:.1f}")
                if hasattr(self, 'db_text'):
                    self.db_text.configure(state="normal")
                    self.db_text.delete("1.0","end")
                    self.db_text.insert("end", "\n".join(lines))
                    self.db_text.configure(state="disabled")
        except Exception as e:
            if hasattr(self, 'db_text'):
                self.db_text.configure(state="normal")
                self.db_text.delete("1.0","end")
                self.db_text.insert("end", f"  DB baglantisi kurulamadi: {e}\n  Once 'python main.py' calistirin.")
                self.db_text.configure(state="disabled")

    # ══════════════════════════════════════════════════════════════
    # IMPORT SAYFASI
    # ══════════════════════════════════════════════════════════════
    def _page_import(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)
        tk.Label(frame, text="Veri İmport", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="Excel dosyanızı seçin ve sisteme aktarın",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,16))

        fi = card(frame, "Excel Dosyası")
        row = tk.Frame(fi, bg=C["card"])
        row.pack(fill="x")
        tk.Entry(row, textvariable=self.excel_path, font=FONT,
                 width=55, relief="solid", bd=1).pack(side="left", padx=(0,8))
        tk.Button(row, text="Gözat...", font=FONT, bg=C["blue"], fg="white",
                  relief="flat", padx=12, pady=4, cursor="hand2",
                  command=self._browse_excel).pack(side="left")

        oi = card(frame, "Seçenekler")
        self.opt_update = tk.BooleanVar(value=False)
        tk.Checkbutton(oi, text="İmport öncesi emisyon faktörlerini güncelle",
                       variable=self.opt_update, font=FONT,
                       bg=C["card"], fg=C["text"], activebackground=C["card"]
                       ).pack(anchor="w")

        bf = tk.Frame(frame, bg=C["bg"])
        bf.pack(fill="x", pady=8)
        self.import_btn = tk.Button(bf, text="▶  İmport Et",
                                    font=FONT_BOLD, bg=C["blue"], fg="white",
                                    relief="flat", padx=24, pady=10, cursor="hand2",
                                    command=self._run_import)
        self.import_btn.pack(side="left")
        tk.Button(bf, text="📊 Sonuçlara Git", font=FONT, bg=C["bg"],
                  fg=C["blue"], relief="flat", padx=12, pady=10, cursor="hand2",
                  command=lambda: self._show_page("results")
                  ).pack(side="left", padx=8)

        section_label(frame, "İmport Logu")
        self.import_log = scrolledtext.ScrolledText(
            frame, height=18, font=FONT_MONO,
            bg="#1E1E1E", fg="#D4D4D4",
            relief="flat", highlightbackground=C["border"], highlightthickness=1)
        self.import_log.pack(fill="both", expand=True)
        self.import_log.insert("end", "Import logu burada gorunecek...\n")

    def _browse_excel(self):
        p = filedialog.askopenfilename(
            title="Excel Dosyası Seç",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tüm", "*.*")])
        if p: self.excel_path.set(p)

    def _run_import(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("Uyarı", "Lütfen Excel dosyası seçin."); return
        if not os.path.exists(path):
            messagebox.showerror("Hata", f"Dosya bulunamadı:\n{path}"); return
        self.import_btn.configure(state="disabled", text="⏳  Çalışıyor...")
        self.import_log.delete("1.0","end")
        self.status_var.set("İmport çalışıyor...")

        class LogW:
            def __init__(self, w): self.w = w
            def write(self, t):
                self.w.after(0, lambda x=t: (
                    self.w.configure(state="normal"),
                    self.w.insert("end", x),
                    self.w.see("end")
                ))
            def flush(self): pass

        def run():
            try:
                old = sys.stdout; sys.stdout = LogW(self.import_log)
                if self.opt_update.get():
                    from run_update import run_update
                    run_update(excel_path=path)
                from pipeline.Importer import run_import
                run_import(path)
                sys.stdout = old
                self.after(0, lambda: self._import_done(True))
            except Exception as e:
                import traceback
                self.after(0, lambda: self.import_log.insert("end", f"\nHATA:\n{traceback.format_exc()}"))
                self.after(0, lambda: self._import_done(False))

        threading.Thread(target=run, daemon=True).start()

    def _import_done(self, ok):
        self.import_btn.configure(state="normal", text="▶  İmport Et")
        self.status_var.set("✅ Tamamlandı" if ok else "❌ Hata")
        self._load_db_summary()
        if ok and messagebox.askyesno("Tamamlandı", "İmport başarılı!\nSonuçları görüntüle?"):
            self._show_page("results")

    # ══════════════════════════════════════════════════════════════
    # FAKTÖR GÜNCELLEME
    # ══════════════════════════════════════════════════════════════
    def _page_update(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)
        tk.Label(frame, text="Emisyon Faktörü Güncelleme", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="IPCC, DEFRA ve TC Enerji Bakanlığı EVÇED kaynaklarından günceller",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,16))

        si = card(frame, "Güncellenecek Kaynaklar")
        self.src_vars = {}
        for key, label in [
            ("gwp",   "🧪  GWP Değerleri (IPCC AR6 2021)"),
            ("ipcc",  "🔥  Yakıt Faktörleri (IPCC 2006 GL)"),
            ("teias", "⚡  Elektrik Grid Faktörü (TC ETKB EVÇED)"),
            ("defra", "🚛  Taşımacılık Faktörleri (DEFRA 2025)"),
        ]:
            v = tk.BooleanVar(value=True); self.src_vars[key] = v
            tk.Checkbutton(si, text=label, variable=v, font=FONT,
                           bg=C["card"], fg=C["text"], activebackground=C["card"]
                           ).pack(anchor="w", pady=2)

        xi = card(frame, "Excel Güncelleme (opsiyonel)")
        tk.Label(xi, text="Boş yeşil hücreler otomatik doldurulur",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(anchor="w", pady=(0,6))
        xr = tk.Frame(xi, bg=C["card"]); xr.pack(fill="x")
        tk.Entry(xr, textvariable=self.excel_path, font=FONT,
                 width=50, relief="solid", bd=1).pack(side="left", padx=(0,8))
        tk.Button(xr, text="Gözat...", font=FONT, bg=C["blue"], fg="white",
                  relief="flat", padx=10, pady=3, cursor="hand2",
                  command=self._browse_excel).pack(side="left")

        # EVÇED bilgi kutusu
        ei = card(frame, "TC Enerji Bakanlığı EVÇED — Bilinen Değerler")
        tk.Label(ei, text=f"Kaynak: {EVCED_URL}",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(anchor="w", pady=(0,6))
        cols = ["Yıl","Üretim EF (tCO2/MWh)","Tüketim EF (tCO2/MWh)","Kaynak"]
        tree = make_treeview(ei, cols, [80,180,180,250], height=5)
        for y, d in sorted(EVCED_FACTORS.items()):
            tree.insert("","end", values=(y, d["uretim"], d["tuketim"], d["kaynak"]))

        self.update_btn = tk.Button(frame, text="🔄  Güncellemeyi Başlat",
                                    font=FONT_BOLD, bg=C["green"], fg="white",
                                    relief="flat", padx=24, pady=10, cursor="hand2",
                                    command=self._run_update)
        self.update_btn.pack(anchor="w", pady=8)

        section_label(frame, "Güncelleme Logu")
        self.update_log = scrolledtext.ScrolledText(
            frame, height=10, font=FONT_MONO, bg="#1E1E1E", fg="#D4D4D4",
            relief="flat", highlightbackground=C["border"], highlightthickness=1)
        self.update_log.pack(fill="both", expand=True)
        self.update_log.insert("end", "Guncelleme logu burada gorunecek...\n")

    def _run_update(self):
        sources = [k for k,v in self.src_vars.items() if v.get()]
        if not sources:
            messagebox.showwarning("Uyarı","En az bir kaynak seçin."); return
        excel = self.excel_path.get().strip() or None
        self.update_btn.configure(state="disabled", text="⏳  Güncelleniyor...")
        self.update_log.delete("1.0","end")
        self.status_var.set("Güncelleniyor...")

        class LogW:
            def __init__(self, w): self.w = w
            def write(self, t):
                self.w.after(0, lambda x=t: (
                    self.w.configure(state="normal"),
                    self.w.insert("end", x), self.w.see("end")))
            def flush(self): pass

        def run():
            try:
                old = sys.stdout; sys.stdout = LogW(self.update_log)
                from run_update import run_update
                run_update(sources=sources, excel_path=excel)
                sys.stdout = old
                self.after(0, lambda: self._update_done(True))
            except Exception as e:
                import traceback
                self.after(0, lambda: self.update_log.insert("end", f"\nHATA:\n{traceback.format_exc()}"))
                self.after(0, lambda: self._update_done(False))

        threading.Thread(target=run, daemon=True).start()

    def _update_done(self, ok):
        self.update_btn.configure(state="normal", text="🔄  Güncellemeyi Başlat")
        self.status_var.set("✅ Güncelleme tamamlandı" if ok else "❌ Hata")

    # ══════════════════════════════════════════════════════════════
    # SONUÇLAR & GRAFİK
    # ══════════════════════════════════════════════════════════════
    def _page_results(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)
        tk.Label(frame, text="Sonuçlar & Grafikler", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="Emisyon hesaplama sonuçları", font=FONT,
                 bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,8))

        # Sekme çubuğu
        tabf = tk.Frame(frame, bg=C["bg"]); tabf.pack(fill="x", pady=(0,8))
        self.res_tab = tk.StringVar(value="chart")
        tabs = [("📊 Yıllık Grafik","chart"),("📈 Değişim Grafiği","trend"),
                ("📋 Kapsam Tablosu","table"),("🎯 Etki Analizi","impact"),
                ("🏢 Kapsam 3 Detay","s3detail"),("👤 Kişi Başına","percap")]
        for text, val in tabs:
            tk.Radiobutton(tabf, text=text, variable=self.res_tab, value=val,
                           font=FONT, bg=C["bg"], fg=C["text"],
                           activebackground=C["bg"], cursor="hand2",
                           command=self._refresh_results
                           ).pack(side="left", padx=(0,12))

        self.res_content = tk.Frame(frame, bg=C["bg"])
        self.res_content.pack(fill="both", expand=True)
        self._refresh_results()

    def _refresh_results(self):
        for w in self.res_content.winfo_children(): w.destroy()
        tab = self.res_tab.get()
        {
            "chart":    self._res_chart,
            "trend":    self._res_trend,
            "table":    self._res_table,
            "impact":   self._res_impact,
            "s3detail": self._res_s3detail,
            "percap":   self._res_percap,
        }.get(tab, self._res_chart)()

    def _res_chart(self):
        try:
            import matplotlib
            matplotlib.use("TkAgg")
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            years_data = db_get_scope_totals()
            if not years_data:
                tk.Label(self.res_content,
                         text="Henüz import edilmiş veri yok.\nÖnce Veri İmport sayfasından import yapın.",
                         font=FONT, bg=C["bg"], fg=C["muted"]).pack(expand=True); return

            fig = Figure(figsize=(12, 5.5), facecolor=C["bg"])
            years  = sorted(years_data.keys())
            k1v = [years_data[y]["k1"]    for y in years]
            k2v = [years_data[y]["k2"]    for y in years]
            k3v = [years_data[y]["k3"]    for y in years]
            x   = list(range(len(years)))
            w   = 0.25

            ax1 = fig.add_subplot(121)
            ax1.bar([i-w for i in x], k1v, w, label="Kapsam 1", color=C["s1"], alpha=0.85)
            ax1.bar([i   for i in x], k2v, w, label="Kapsam 2", color=C["s2"], alpha=0.85)
            ax1.bar([i+w for i in x], k3v, w, label="Kapsam 3", color=C["s3"], alpha=0.85)
            ax1.set_xticks(x); ax1.set_xticklabels(years, fontsize=9)
            ax1.set_ylabel("ton CO2e", fontsize=9)
            ax1.set_title("Yıllık Kapsam Karşılaştırması", fontsize=10, fontweight="bold")
            ax1.legend(fontsize=8); ax1.set_facecolor(C["bg"])
            ax1.spines["top"].set_visible(False); ax1.spines["right"].set_visible(False)

            ax2 = fig.add_subplot(122)
            latest = max(years_data.keys()); d = years_data[latest]
            vals   = [d["k1"], d["k2"], d["k3"]]
            labels = [f"Kapsam 1\n{d['k1']:.0f}t", f"Kapsam 2\n{d['k2']:.0f}t", f"Kapsam 3\n{d['k3']:.0f}t"]
            nz = [(v,l,c) for v,l,c in zip(vals,labels,[C["s1"],C["s2"],C["s3"]]) if v>0]
            if nz:
                v2,l2,c2 = zip(*nz)
                ax2.pie(v2, labels=l2, colors=c2, autopct="%1.1f%%",
                        startangle=90, textprops={"fontsize":8})
            ax2.set_title(f"{latest} Dağılımı\nToplam: {d['total']:.1f} ton CO2e",
                          fontsize=10, fontweight="bold")
            ax2.set_facecolor(C["bg"]); fig.tight_layout(pad=2)

            canvas = FigureCanvasTkAgg(fig, master=self.res_content)
            canvas.draw(); canvas.get_tk_widget().pack(fill="both", expand=True)
        except ImportError:
            tk.Label(self.res_content,
                     text="Grafik için: pip install matplotlib",
                     font=FONT, bg=C["bg"], fg=C["red"]).pack(expand=True)

    def _res_trend(self):
        """Yıllara göre değişim grafiği — kapsamlar ayrı."""
        try:
            import matplotlib
            matplotlib.use("TkAgg")
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            years_data = db_get_scope_totals()
            if len(years_data) < 1:
                tk.Label(self.res_content,
                         text="En az 1 yıl verisi gerekli.",
                         font=FONT, bg=C["bg"], fg=C["muted"]).pack(expand=True); return

            fig = Figure(figsize=(12, 5.5), facecolor=C["bg"])
            years  = sorted(years_data.keys())
            k1v = [years_data[y]["k1"]    for y in years]
            k2v = [years_data[y]["k2"]    for y in years]
            k3v = [years_data[y]["k3"]    for y in years]
            tot = [years_data[y]["total"] for y in years]

            # Sol: Çizgi grafik — kapsamlar ayrı
            ax1 = fig.add_subplot(121)
            ax1.plot(years, k1v, "o-", color=C["s1"], linewidth=2,
                     markersize=7, label="Kapsam 1")
            ax1.plot(years, k2v, "s-", color=C["s2"], linewidth=2,
                     markersize=7, label="Kapsam 2")
            ax1.plot(years, k3v, "^-", color=C["s3"], linewidth=2,
                     markersize=7, label="Kapsam 3")
            ax1.plot(years, tot, "D--", color=C["blue"], linewidth=2,
                     markersize=7, label="Toplam", alpha=0.7)

            # Değerleri noktaların üstüne yaz
            for i, y in enumerate(years):
                ax1.annotate(f"{k1v[i]:.0f}", (y, k1v[i]),
                             textcoords="offset points", xytext=(0,6), fontsize=7, color=C["s1"])
                ax1.annotate(f"{k2v[i]:.0f}", (y, k2v[i]),
                             textcoords="offset points", xytext=(0,6), fontsize=7, color=C["s2"])
                ax1.annotate(f"{k3v[i]:.0f}", (y, k3v[i]),
                             textcoords="offset points", xytext=(0,6), fontsize=7, color=C["s3"])

            ax1.set_title("Yıllara Göre Değişim Trendi", fontsize=10, fontweight="bold")
            ax1.set_ylabel("ton CO2e", fontsize=9)
            ax1.legend(fontsize=8); ax1.set_facecolor(C["bg"])
            ax1.spines["top"].set_visible(False); ax1.spines["right"].set_visible(False)
            ax1.grid(axis="y", alpha=0.3)

            # Sağ: Stacked bar — yıl bazlı toplam dağılımı
            ax2 = fig.add_subplot(122)
            xi = list(range(len(years)))
            ax2.bar(xi, k1v, label="Kapsam 1", color=C["s1"], alpha=0.85)
            ax2.bar(xi, k2v, bottom=k1v, label="Kapsam 2", color=C["s2"], alpha=0.85)
            k12 = [a+b for a,b in zip(k1v,k2v)]
            ax2.bar(xi, k3v, bottom=k12, label="Kapsam 3", color=C["s3"], alpha=0.85)

            # Yüzde değişimini yıl başlıklarına ekle
            xlabels = []
            for i, y in enumerate(years):
                if i == 0:
                    xlabels.append(y)
                else:
                    prev = tot[i-1]; curr = tot[i]
                    if prev > 0:
                        pct = (curr - prev) / prev * 100
                        arrow = "▲" if pct > 0 else "▼"
                        xlabels.append(f"{y}\n{arrow}{abs(pct):.1f}%")
                    else:
                        xlabels.append(y)

            ax2.set_xticks(xi); ax2.set_xticklabels(xlabels, fontsize=8)
            ax2.set_title("Yıllık Toplam (Yığılmış)", fontsize=10, fontweight="bold")
            ax2.set_ylabel("ton CO2e", fontsize=9)
            ax2.legend(fontsize=8); ax2.set_facecolor(C["bg"])
            ax2.spines["top"].set_visible(False); ax2.spines["right"].set_visible(False)

            fig.tight_layout(pad=2)
            canvas = FigureCanvasTkAgg(fig, master=self.res_content)
            canvas.draw(); canvas.get_tk_widget().pack(fill="both", expand=True)

        except ImportError:
            tk.Label(self.res_content, text="pip install matplotlib",
                     font=FONT, bg=C["bg"], fg=C["red"]).pack(expand=True)

    def _res_table(self):
        years_data = db_get_scope_totals()
        if not years_data:
            tk.Label(self.res_content, text="Veri yok.",
                     font=FONT, bg=C["bg"], fg=C["muted"]).pack(expand=True); return
        cols = ["Yıl","Kapsam 1 (ton)","Kapsam 2 (ton)","Kapsam 3 (ton)","Toplam (ton)"]
        tree = make_treeview(self.res_content, cols, [80,180,180,180,180], height=15)
        for y in sorted(years_data.keys()):
            d = years_data[y]
            tree.insert("","end", values=(
                y, f"{d['k1']:.2f}", f"{d['k2']:.2f}", f"{d['k3']:.2f}", f"{d['total']:.2f}"
            ))

    def _res_impact(self):
        """Etki analizi — en büyük emisyon kaynakları."""
        years_data = db_get_scope_totals()
        if not years_data:
            tk.Label(self.res_content, text="Veri yok.",
                     font=FONT, bg=C["bg"], fg=C["muted"]).pack(expand=True); return

        latest = max(years_data.keys()); d = years_data[latest]

        header = tk.Frame(self.res_content, bg=C["bg"])
        header.pack(fill="x", pady=(0,8))
        tk.Label(header,
                 text=f"Etki Analizi — {latest} Yılı En Büyük Emisyon Kaynakları (%95 Eşiği)",
                 font=FONT_BOLD, bg=C["bg"], fg=C["text"]).pack(side="left")
        tk.Label(header, text=f"  Toplam: {d['total']:.2f} ton CO2e",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(side="left")

        sources = [
            ("Kapsam 1","1.1","Sabit Yakma",           d["k1"]*0.022),
            ("Kapsam 1","1.2","Hareketli Yakma",       d["k1"]*0.429),
            ("Kapsam 1","1.4","Soğutucu Gazlar",       d["k1"]*0.549),
            ("Kapsam 2","2.1","Elektrik Tüketimi",     d["k2"]),
            ("Kapsam 3","3.x","Kapsam 3 Toplam",       d["k3"]),
        ]
        sources_s = sorted(sources, key=lambda x: x[3], reverse=True)
        total = sum(v for _,_,_,v in sources_s)

        cols = ["Sıra","Kapsam","Kategori","Emisyon Kaynağı","CO2e (ton)","Pay (%)","Kümülatif (%)"]
        tree = make_treeview(self.res_content, cols, [50,90,80,200,120,80,100], height=12)

        cumulative = 0
        for i, (scope, cat, label, val) in enumerate(sources_s, 1):
            if val <= 0: continue
            pct = (val/total*100) if total else 0
            cumulative += pct
            tag = "above" if cumulative <= 95 else "below"
            tree.insert("","end", values=(
                i, scope, cat, label,
                f"{val:.2f}", f"{pct:.1f}%", f"{cumulative:.1f}%"
            ), tags=(tag,))

        tree.tag_configure("above", background=C["impact"])
        tree.tag_configure("below", background=C["bg"])

        tk.Label(self.res_content,
                 text="🟠 Turuncu = emisyonların %95'ini oluşturan kritik kaynaklar",
                 font=FONT_SMALL, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=4)

    def _res_s3detail(self):
        """Kapsam 3 içinde 3.1-6 alt başlıklar — açılır/kapanır."""
        outer = tk.Frame(self.res_content, bg=C["bg"])
        outer.pack(fill="both", expand=True)
        sf = self._scrollable(outer)
        pad = tk.Frame(sf, bg=C["bg"])
        pad.pack(fill="both", expand=True, padx=8, pady=8)

        tk.Label(pad, text="Kapsam 3 Alt Kategori Detayı",
                 font=FONT_BOLD, bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(0,8))
        tk.Label(pad, text="Her kategoriye tıklayarak satırları açın/kapatın.",
                 font=FONT_SMALL, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,12))

        years_data  = db_get_scope_totals()
        cat_data    = db_get_category_breakdown()

        if not years_data:
            tk.Label(pad, text="Veri yok.", font=FONT, bg=C["bg"], fg=C["muted"]).pack()
            return

        latest = max(years_data.keys())
        cats = {
            "3 — Kapsam 3 (Dolaylı)": {
                "color": C["s3"], "children": {
                    "3.1 — Hammadde Sevkiyatı":    cat_data.get(latest,{}).get("3.1",0),
                    "3.2 — Ürün Sevkiyatı":        cat_data.get(latest,{}).get("3.2",0),
                    "3.3 — Personel İşe Gidiş":   cat_data.get(latest,{}).get("3.3",0),
                    "3.4 — İş Seyahati (Kara)":   cat_data.get(latest,{}).get("3.4",0),
                    "3.5 — İş Seyahati (Hava)":   cat_data.get(latest,{}).get("3.5",0),
                }
            },
            "4 — Tedarik Zinciri": {
                "color": "#8B4513", "children": {
                    "4.1 — Satın Alınan Mallar":   cat_data.get(latest,{}).get("4.1",0),
                    "4.2 — Sermaye Varlıkları":    cat_data.get(latest,{}).get("4.2",0),
                    "4.3 — Atık Bertarafı":        cat_data.get(latest,{}).get("4.3",0),
                    "4.4 — Kiralanan Ekipmanlar":  cat_data.get(latest,{}).get("4.4",0),
                    "4.5 — Hizmet Alımları":       cat_data.get(latest,{}).get("4.5",0),
                }
            },
            "6 — Elektrik T&D": {
                "color": "#4B0082", "children": {
                    "6.1 — T&D Kayıpları":         cat_data.get(latest,{}).get("6.1",0),
                }
            },
        }

        self._s3_states = {}
        for group_name, group_data in cats.items():
            color = group_data["color"]
            children = group_data["children"]
            group_total = sum(children.values())

            # Grup başlığı (tıklanabilir)
            state_var = tk.BooleanVar(value=True)
            self._s3_states[group_name] = state_var

            header_frame = tk.Frame(pad, bg=color, cursor="hand2")
            header_frame.pack(fill="x", pady=(4,0))

            child_frame = tk.Frame(pad, bg=C["card"],
                                   highlightbackground=color, highlightthickness=1)

            def toggle(cf=child_frame, sv=state_var, hf=header_frame):
                if sv.get():
                    cf.pack(fill="x")
                else:
                    cf.pack_forget()

            header_btn = tk.Button(
                header_frame,
                text=f"  ▼  {group_name}    {group_total:.2f} ton CO2e  →  tıkla aç/kapat",
                font=FONT_BOLD, bg=color, fg="white",
                relief="flat", anchor="w", padx=12, pady=8, cursor="hand2",
                activebackground=color,
                command=lambda sv=state_var, cf=child_frame, hf=header_frame: (
                    sv.set(not sv.get()),
                    cf.pack(fill="x") if sv.get() else cf.pack_forget()
                )
            )
            header_btn.pack(fill="x")

            # Çocuk satırları
            for i, (cat_name, val) in enumerate(children.items()):
                row_bg = C["card"] if i%2==0 else "#F8F8F5"
                rf = tk.Frame(child_frame, bg=row_bg)
                rf.pack(fill="x")
                tk.Label(rf, text=f"    {cat_name}", font=FONT,
                         bg=row_bg, fg=C["text"], padx=16, pady=6,
                         anchor="w").pack(side="left", fill="x", expand=True)
                pct = (val/group_total*100) if group_total else 0
                tk.Label(rf, text=f"{val:.2f} ton CO2e",
                         font=FONT_MONO, bg=row_bg, fg=C["text"],
                         padx=16).pack(side="right")
                tk.Label(rf, text=f"{pct:.1f}%",
                         font=FONT_MONO, bg=row_bg, fg=C["muted"],
                         padx=8).pack(side="right")

            child_frame.pack(fill="x")

    def _res_percap(self):
        tk.Label(self.res_content,
                 text="Kişi Başına Emisyon — Lokasyon Bazında",
                 font=FONT_BOLD, bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(0,8))
        excel = self.excel_path.get().strip()
        if not excel or not os.path.exists(excel):
            tk.Label(self.res_content,
                     text="Excel dosyasını 'Veri İmport' sayfasından seçin.",
                     font=FONT, bg=C["bg"], fg=C["muted"]).pack(); return
        try:
            from openpyxl import load_workbook
            wb = load_workbook(excel, read_only=True, data_only=True)
            ws = wb["PER_CAPITA"]
            cols = ["Yıl","Lokasyon","Personel","Toplam CO2e",
                    "Kapsam 1","Kapsam 2","Kapsam 3","ton CO2e/Kişi"]
            tree = make_treeview(self.res_content, cols,
                                 [60,200,80,120,100,100,100,120], height=14)
            for row in ws.iter_rows(min_row=4, values_only=True):
                if row[0] and row[1]:
                    vals = tuple(
                        f"{v:.4f}" if isinstance(v,float) else str(v) if v else "—"
                        for v in row[:8]
                    )
                    tag = "total" if "TOPLAM" in str(row[1]) else "loc"
                    tree.insert("","end", values=vals, tags=(tag,))
            tree.tag_configure("total", background=C["warning"], font=FONT_BOLD)
        except Exception as e:
            tk.Label(self.res_content, text=f"Hata: {e}",
                     font=FONT, bg=C["bg"], fg=C["red"]).pack()

    # ══════════════════════════════════════════════════════════════
    # REFERANS TABLOLAR
    # ══════════════════════════════════════════════════════════════
    def _page_refs(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True)
        outer = tk.Frame(frame, bg=C["bg"])
        outer.pack(fill="both", expand=True)
        sf = self._scrollable(outer)
        pad = tk.Frame(sf, bg=C["bg"])
        pad.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(pad, text="Referans Tablolar", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(pad, text="GWP, Emisyon Faktörleri ve NCV değerleri ile kaynakları",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,16))

        # ── 1. GWP Tablosu ─────────────────────────────────────────
        section_label(pad, "1. GWP Değerleri (GWP-100) — IPCC AR6 2021")
        gi = card(pad, pady=16)
        tk.Label(gi, text="Kaynak: IPCC AR6 WGI Chapter 7 Supplementary Material, Tablo 7.SM.7",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(anchor="w", pady=(0,6))
        cols_gwp = ["Formül","Gaz Adı","GWP-100","Kaynak"]
        tree_gwp = make_treeview(gi, cols_gwp, [100,200,100,250], height=10)
        for row in GWP_TABLE:
            tree_gwp.insert("","end", values=row)

        # ── 2. Emisyon Faktörleri Tablosu ──────────────────────────
        section_label(pad, "2. Yakıt Emisyon Faktörleri — IPCC 2006 GL")
        efi = card(pad, pady=16)
        tk.Label(efi, text="Kaynak: IPCC 2006 Guidelines for National GHG Inventories, Volume 2, Table 2.2",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(anchor="w", pady=(0,6))
        cols_ef = ["Yakıt","Birim","EF CO2\n(ton/TJ)","EF CH4\n(ton/TJ)","EF N2O\n(ton/TJ)","NCV\n(TJ/Gg)","Yoğunluk","Kaynak"]
        tree_ef = make_treeview(efi, cols_ef, [120,60,90,90,90,80,100,200], height=7)
        for row in EF_TABLE:
            tree_ef.insert("","end", values=row)

        # ── 3. NCV Tablosu ─────────────────────────────────────────
        section_label(pad, "3. Net Kalorifik Değer (NCV) — IPCC 2006 GL")
        ni = card(pad, pady=16)
        tk.Label(ni, text="Kaynak: IPCC 2006 Guidelines, Volume 2, Table 1.2",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(anchor="w", pady=(0,6))
        cols_ncv = ["Yakıt","Birim","NCV","NCV Birimi","Yoğunluk","Yoğunluk Birimi","Kaynak"]
        tree_ncv = make_treeview(ni, cols_ncv, [130,60,80,80,90,120,200], height=8)
        for row in NCV_TABLE:
            tree_ncv.insert("","end", values=row)

        # ── 4. TC Enerji Bakanlığı EVÇED ───────────────────────────
        section_label(pad, "4. TC Enerji Bakanlığı EVÇED Elektrik Emisyon Faktörleri")
        evi = card(pad, pady=16)
        tk.Label(evi,
                 text="Kaynak: TC Enerji ve Tabii Kaynaklar Bakanlığı, Enerji Verimliliği ve Çevre Dairesi (EVÇED)",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(anchor="w")
        tk.Label(evi, text=EVCED_URL,
                 font=FONT_SMALL, bg=C["card"], fg=C["blue"]).pack(anchor="w", pady=(0,8))

        cols_ev = ["Yıl","Üretim EF\n(tCO2/MWh)","Tüketim EF\n(tCO2/MWh)","Yayın Tarihi","Kaynak"]
        tree_ev = make_treeview(evi, cols_ev, [60,130,130,180,280], height=7)
        ev_dates = {
            2019:"Mart 2021", 2020:"Mart 2021",
            2021:"Eylül 2022", 2022:"Aralık 2024", 2023:"Aralık 2025"
        }
        for y, d in sorted(EVCED_FACTORS.items()):
            tree_ev.insert("","end", values=(
                y, d["uretim"], d["tuketim"],
                ev_dates.get(y,"—"), d["kaynak"]
            ))

        tk.Label(evi,
                 text="Not: Üretim EF = elektrik üretim noktasındaki emisyon. "
                      "Tüketim EF = şebekeden alınan elektrikteki emisyon (T&D kayıpları dahil).",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"],
                 wraplength=900, justify="left").pack(anchor="w", pady=(8,0))

    # ══════════════════════════════════════════════════════════════
    # LOG & REVİZYONLAR
    # ══════════════════════════════════════════════════════════════
    def _page_log(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)
        tk.Label(frame, text="Log & Revizyonlar", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="Tüm import ve güncelleme revizyonları",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,8))

        # Sekme
        tf = tk.Frame(frame, bg=C["bg"]); tf.pack(fill="x", pady=(0,8))
        self.log_tab = tk.StringVar(value="revisions")
        for text, val in [("📋 Revizyonlar","revisions"),("🔧 Tüm İşlemler","all")]:
            tk.Radiobutton(tf, text=text, variable=self.log_tab, value=val,
                           font=FONT, bg=C["bg"], fg=C["text"],
                           activebackground=C["bg"], cursor="hand2",
                           command=self._refresh_log).pack(side="left", padx=(0,16))
        tk.Button(tf, text="🔄 Yenile", font=FONT, bg=C["blue"], fg="white",
                  relief="flat", padx=12, pady=4, cursor="hand2",
                  command=self._refresh_log).pack(side="right")

        self.log_content = tk.Frame(frame, bg=C["bg"])
        self.log_content.pack(fill="both", expand=True)
        self._refresh_log()

    def _refresh_log(self):
        for w in self.log_content.winfo_children(): w.destroy()
        if self.log_tab.get() == "revisions":
            self._log_revisions()
        else:
            self._log_all()

    def _log_revisions(self):
        """Tüm revizyonlar — 2025-R1, 2025-R2 formatında."""
        revs = db_get_revisions()

        info = tk.Frame(self.log_content, bg=C["warning"],
                        highlightbackground=C["border"], highlightthickness=1)
        info.pack(fill="x", pady=(0,8))
        inf = tk.Frame(info, bg=C["warning"], padx=16, pady=8); inf.pack(fill="x")
        tk.Label(inf,
                 text="Revizyon Sistemi: Her import aynı yıl için yeni bir revizyon oluşturur.\n"
                      "Örnek: 2025-R1 = 2025 yılının ilk importu, 2025-R2 = ikinci importu.",
                 font=FONT_SMALL, bg=C["warning"], fg=C["text"],
                 justify="left").pack(anchor="w")

        if not revs:
            tk.Label(self.log_content, text="Henüz import yapılmadı.",
                     font=FONT, bg=C["bg"], fg=C["muted"]).pack(pady=20); return

        cols = ["Revizyon ID","Yıl","Tarih","Kapsam 1","Kapsam 2","Kapsam 3","Toplam","Durum"]
        tree = make_treeview(self.log_content, cols,
                             [100,60,150,110,110,110,120,80], height=16)

        # En son revizyon hangisi?
        latest_by_year = {}
        for r in revs:
            y = r["year"]
            if y not in latest_by_year:
                latest_by_year[y] = r["id"]

        for r in revs:
            is_latest = latest_by_year.get(r["year"]) == r["id"]
            tag = "latest" if is_latest else "old"
            tree.insert("","end", values=(
                r["id"], r["year"], r["ts"],
                f"{r['k1']:.2f}", f"{r['k2']:.2f}",
                f"{r['k3']:.2f}", f"{r['total']:.2f}",
                "✅ Güncel" if is_latest else "📦 Arşiv"
            ), tags=(tag,))

        tree.tag_configure("latest", background=C["success"], font=FONT_BOLD)
        tree.tag_configure("old",    background=C["bg"])

        tk.Label(self.log_content,
                 text="✅ Yeşil = her yılın en son (geçerli) revizyonu   📦 Gri = eski revizyon (arşiv)",
                 font=FONT_SMALL, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=4)

    def _log_all(self):
        """Tüm audit log kayıtları."""
        try:
            from db.connection import SessionLocal
            from db.models import AuditLog
            cols = ["Tarih","İşlem","Kapsam","Durum","Notlar"]
            tree = make_treeview(self.log_content, cols,
                                 [150,140,100,80,450], height=18)
            with SessionLocal() as s:
                logs = s.query(AuditLog).order_by(AuditLog.id.desc()).limit(200).all()
                for log in logs:
                    tag = "success" if log.status == "success" else "fail"
                    tree.insert("","end", values=(
                        str(log.created_at)[:19] if log.created_at else "—",
                        log.action or "—", log.scope or "—",
                        log.status or "—", (log.notes or "")[:80]
                    ), tags=(tag,))
            tree.tag_configure("success", background=C["success"])
            tree.tag_configure("fail",    background=C["danger"])
        except Exception as e:
            tk.Label(self.log_content, text=f"DB Hatası: {e}",
                     font=FONT, bg=C["bg"], fg=C["red"]).pack()


# ══════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = UNSPEDDashboard()
    app.mainloop()
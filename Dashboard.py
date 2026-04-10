"""
UNSPED Karbon Ayak İzi Dashboard
==================================
Kullanım:
    python dashboard.py

Gereksinimler:
    pip install matplotlib
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import io

# Encoding fix for Windows
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

# Project root
ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, ROOT)

# ── Renkler ────────────────────────────────────────────────────────
C = {
    "bg":        "#F5F5F0",
    "sidebar":   "#1F4E79",
    "sidebar_hl":"#2E75B6",
    "white":     "#FFFFFF",
    "text":      "#2C2C2A",
    "muted":     "#888780",
    "border":    "#D3D1C7",
    "blue":      "#1F4E79",
    "green":     "#1D9E75",
    "amber":     "#BA7517",
    "red":       "#A32D2D",
    "s1":        "#2E75B6",
    "s2":        "#1D9E75",
    "s3":        "#BA7517",
    "card":      "#FFFFFF",
    "success":   "#E2EFDA",
    "warning":   "#FFF2CC",
    "danger":    "#FCEBEB",
}

FONT       = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_TITLE = ("Segoe UI", 14, "bold")
FONT_LARGE = ("Segoe UI", 22, "bold")
FONT_SMALL = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)


# ══════════════════════════════════════════════════════════════════
# ANA UYGULAMA
# ══════════════════════════════════════════════════════════════════

class UNSPEDDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("UNSPED Karbon Ayak İzi Dashboard")
        self.geometry("1280x800")
        self.minsize(1000, 650)
        self.configure(bg=C["bg"])

        # State
        self.excel_path   = tk.StringVar(value="")
        self.status_var   = tk.StringVar(value="Hazır")
        self.current_page = tk.StringVar(value="home")

        self._build_ui()
        self._load_db_summary()

    # ── UI İnşa ────────────────────────────────────────────────────

    def _build_ui(self):
        # Ana layout: sidebar + content
        self.sidebar  = tk.Frame(self, bg=C["sidebar"], width=220)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.content  = tk.Frame(self, bg=C["bg"])
        self.content.pack(side="left", fill="both", expand=True)

        self._build_sidebar()
        self._build_statusbar()
        self._show_page("home")

    def _build_sidebar(self):
        # Logo alanı
        logo_frame = tk.Frame(self.sidebar, bg=C["sidebar"], pady=20)
        logo_frame.pack(fill="x")
        tk.Label(logo_frame, text="🌍", font=("Segoe UI", 28),
                 bg=C["sidebar"], fg="white").pack()
        tk.Label(logo_frame, text="UNSPED", font=("Segoe UI", 13, "bold"),
                 bg=C["sidebar"], fg="white").pack()
        tk.Label(logo_frame, text="Karbon Dashboard", font=("Segoe UI", 9),
                 bg=C["sidebar"], fg="#9FC5E8").pack()

        ttk.Separator(self.sidebar, orient="horizontal").pack(fill="x", padx=16, pady=4)

        # Menü
        self.nav_buttons = {}
        menu_items = [
            ("home",    "🏠", "Ana Sayfa"),
            ("import",  "📥", "Veri İmport"),
            ("update",  "🔄", "Faktör Güncelle"),
            ("results", "📊", "Sonuçlar & Grafik"),
            ("log",     "📋", "Log Geçmişi"),
        ]
        for page, icon, label in menu_items:
            btn = tk.Button(
                self.sidebar, text=f"  {icon}  {label}",
                font=FONT, bg=C["sidebar"], fg="white",
                activebackground=C["sidebar_hl"], activeforeground="white",
                relief="flat", anchor="w", padx=16, pady=10, cursor="hand2",
                command=lambda p=page: self._show_page(p)
            )
            btn.pack(fill="x")
            self.nav_buttons[page] = btn

        # Alt bilgi
        info_frame = tk.Frame(self.sidebar, bg=C["sidebar"])
        info_frame.pack(side="bottom", fill="x", pady=12, padx=16)
        tk.Label(info_frame, text="v3.0 • 2025", font=FONT_SMALL,
                 bg=C["sidebar"], fg="#9FC5E8").pack()

    def _build_statusbar(self):
        status = tk.Frame(self.content, bg=C["border"], height=28)
        status.pack(side="bottom", fill="x")
        tk.Label(status, textvariable=self.status_var, font=FONT_SMALL,
                 bg=C["border"], fg=C["text"], padx=12).pack(side="left")

    def _show_page(self, page):
        # Sidebar highlight
        for p, btn in self.nav_buttons.items():
            btn.configure(bg=C["sidebar_hl"] if p == page else C["sidebar"])

        # Content temizle
        for w in self.content.winfo_children():
            if not isinstance(w, tk.Frame) or w.cget("height") != 28:
                w.destroy()

        self.current_page.set(page)
        if   page == "home":    self._page_home()
        elif page == "import":  self._page_import()
        elif page == "update":  self._page_update()
        elif page == "results": self._page_results()
        elif page == "log":     self._page_log()

    # ── Ana Sayfa ──────────────────────────────────────────────────

    def _page_home(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(frame, text="UNSPED Karbon Ayak İzi", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="2025 Raporlama Yılı", font=FONT,
                 bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,16))

        # KPI kartları
        kpi_frame = tk.Frame(frame, bg=C["bg"])
        kpi_frame.pack(fill="x", pady=(0,16))

        self.kpi_vars = {}
        kpis = [
            ("toplam",  "Toplam CO2e",  "ton", C["blue"]),
            ("kapsam1", "Kapsam 1",     "ton", C["s1"]),
            ("kapsam2", "Kapsam 2",     "ton", C["s2"]),
            ("kapsam3", "Kapsam 3",     "ton", C["s3"]),
        ]
        for key, label, unit, color in kpis:
            card = tk.Frame(kpi_frame, bg=C["card"], relief="flat",
                            highlightbackground=C["border"],
                            highlightthickness=1)
            card.pack(side="left", fill="both", expand=True, padx=(0,12))

            tk.Frame(card, bg=color, height=4).pack(fill="x")
            inner = tk.Frame(card, bg=C["card"], padx=16, pady=14)
            inner.pack(fill="both")
            tk.Label(inner, text=label, font=FONT_SMALL,
                     bg=C["card"], fg=C["muted"]).pack(anchor="w")
            var = tk.StringVar(value="—")
            self.kpi_vars[key] = var
            tk.Label(inner, textvariable=var, font=FONT_LARGE,
                     bg=C["card"], fg=color).pack(anchor="w")
            tk.Label(inner, text=unit, font=FONT_SMALL,
                     bg=C["card"], fg=C["muted"]).pack(anchor="w")

        # Hızlı işlemler
        tk.Label(frame, text="Hızlı İşlemler", font=FONT_BOLD,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(8,8))

        btn_frame = tk.Frame(frame, bg=C["bg"])
        btn_frame.pack(fill="x")

        actions = [
            ("📥  Excel Seç & İmport Et", C["blue"],   lambda: self._show_page("import")),
            ("🔄  Faktörleri Güncelle",   C["green"],  lambda: self._show_page("update")),
            ("📊  Sonuçları Gör",         C["amber"],  lambda: self._show_page("results")),
        ]
        for text, color, cmd in actions:
            tk.Button(btn_frame, text=text, font=FONT_BOLD,
                      bg=color, fg="white", relief="flat",
                      padx=20, pady=10, cursor="hand2",
                      activebackground=color, command=cmd
                      ).pack(side="left", padx=(0,12))

        # Son çalıştırma
        tk.Label(frame, text="Veritabanı Özeti", font=FONT_BOLD,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(20,8))
        self.db_text = tk.Text(frame, height=8, font=FONT_MONO,
                               bg=C["card"], fg=C["text"],
                               relief="flat", state="disabled",
                               highlightbackground=C["border"],
                               highlightthickness=1)
        self.db_text.pack(fill="x")
        self._load_db_summary()

    def _load_db_summary(self):
        try:
            from db.connection import SessionLocal
            from db.models import (Company, ReportingPeriod, EmissionFactor,
                                   AuditLog, StationaryCombustion,
                                   MobileCombustion, ElectricityConsumption)
            with SessionLocal() as s:
                company   = s.query(Company).filter_by(name="UNSPED").first()
                periods   = s.query(ReportingPeriod).all()
                ef_count  = s.query(EmissionFactor).count()
                log_count = s.query(AuditLog).count()
                last_log  = s.query(AuditLog).filter_by(
                    action="excel_import").order_by(AuditLog.id.desc()).first()

                lines = []
                if company:
                    lines.append(f"  Şirket     : {company.name} — {company.site}")
                lines.append(f"  Dönemler   : {[p.year for p in periods]}")
                lines.append(f"  EF Kayıtları: {ef_count} adet")
                lines.append(f"  Toplam Log : {log_count} işlem")
                if last_log:
                    lines.append(f"  Son Import : {last_log.created_at}")
                    # Parse KPIs from notes
                    import json
                    try:
                        notes = last_log.notes or ""
                        # Extract totals
                        import re
                        years = {}
                        for m in re.finditer(r'(\d{4})\s*\|\s*K1=([\d.]+)\s*K2=([\d.]+)\s*K3=([\d.]+)\s*Toplam=([\d.]+)', notes):
                            years[m.group(1)] = {
                                'k1': float(m.group(2)),
                                'k2': float(m.group(3)),
                                'k3': float(m.group(4)),
                                'total': float(m.group(5))
                            }
                        if years:
                            latest = max(years.keys())
                            d = years[latest]
                            lines.append(f"\n  [{latest}] Kapsam 1: {d['k1']:.1f} | Kapsam 2: {d['k2']:.1f} | Kapsam 3: {d['k3']:.1f}")
                            lines.append(f"  [{latest}] TOPLAM  : {d['total']:.2f} ton CO2e")
                            # Update KPIs
                            if hasattr(self, 'kpi_vars'):
                                self.kpi_vars['toplam'].set(f"{d['total']:.1f}")
                                self.kpi_vars['kapsam1'].set(f"{d['k1']:.1f}")
                                self.kpi_vars['kapsam2'].set(f"{d['k2']:.1f}")
                                self.kpi_vars['kapsam3'].set(f"{d['k3']:.1f}")
                    except: pass

                if hasattr(self, 'db_text'):
                    self.db_text.configure(state="normal")
                    self.db_text.delete("1.0", "end")
                    self.db_text.insert("end", "\n".join(lines))
                    self.db_text.configure(state="disabled")
        except Exception as e:
            if hasattr(self, 'db_text'):
                self.db_text.configure(state="normal")
                self.db_text.delete("1.0", "end")
                self.db_text.insert("end", f"  DB bağlantısı kurulamadı: {e}\n  Önce 'python main.py' çalıştırın.")
                self.db_text.configure(state="disabled")

    # ── Import Sayfası ─────────────────────────────────────────────

    def _page_import(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(frame, text="Veri İmport", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="Excel dosyanızı seçin ve sisteme aktarın",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,16))

        # Dosya seçici
        file_card = tk.Frame(frame, bg=C["card"],
                             highlightbackground=C["border"], highlightthickness=1)
        file_card.pack(fill="x", pady=(0,12))
        fc_inner = tk.Frame(file_card, bg=C["card"], padx=16, pady=14)
        fc_inner.pack(fill="x")

        tk.Label(fc_inner, text="Excel Dosyası", font=FONT_BOLD,
                 bg=C["card"], fg=C["text"]).grid(row=0, column=0, sticky="w")
        tk.Entry(fc_inner, textvariable=self.excel_path, font=FONT,
                 width=55, relief="solid", bd=1
                 ).grid(row=1, column=0, sticky="ew", pady=6, padx=(0,8))
        tk.Button(fc_inner, text="Gözat...", font=FONT,
                  bg=C["blue"], fg="white", relief="flat",
                  padx=12, pady=4, cursor="hand2",
                  command=self._browse_excel
                  ).grid(row=1, column=1)
        fc_inner.columnconfigure(0, weight=1)

        # Seçenekler
        opt_card = tk.Frame(frame, bg=C["card"],
                            highlightbackground=C["border"], highlightthickness=1)
        opt_card.pack(fill="x", pady=(0,12))
        opt_inner = tk.Frame(opt_card, bg=C["card"], padx=16, pady=14)
        opt_inner.pack(fill="x")
        tk.Label(opt_inner, text="Seçenekler", font=FONT_BOLD,
                 bg=C["card"], fg=C["text"]).pack(anchor="w", pady=(0,6))

        self.opt_update_ef = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_inner,
                       text="İmport öncesi emisyon faktörlerini online'dan güncelle",
                       variable=self.opt_update_ef, font=FONT,
                       bg=C["card"], fg=C["text"], activebackground=C["card"]
                       ).pack(anchor="w")

        # Çalıştır butonu
        btn_frame = tk.Frame(frame, bg=C["bg"])
        btn_frame.pack(fill="x", pady=8)
        self.import_btn = tk.Button(
            btn_frame, text="▶  İmport Et",
            font=FONT_BOLD, bg=C["blue"], fg="white",
            relief="flat", padx=24, pady=10, cursor="hand2",
            command=self._run_import
        )
        self.import_btn.pack(side="left")

        tk.Button(btn_frame, text="📂  Sonuçlara Git",
                  font=FONT, bg=C["bg"], fg=C["blue"],
                  relief="flat", padx=12, pady=10, cursor="hand2",
                  command=lambda: self._show_page("results")
                  ).pack(side="left", padx=8)

        # Log alanı
        tk.Label(frame, text="İmport Logu", font=FONT_BOLD,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(12,4))
        self.import_log = scrolledtext.ScrolledText(
            frame, height=16, font=FONT_MONO,
            bg="#1E1E1E", fg="#D4D4D4",
            insertbackground="white", relief="flat",
            highlightbackground=C["border"], highlightthickness=1
        )
        self.import_log.pack(fill="both", expand=True)
        self.import_log.insert("end", "İmport logu burada görünecek...\n")

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="Excel Dosyası Seç",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls"), ("Tüm Dosyalar", "*.*")]
        )
        if path:
            self.excel_path.set(path)

    def _run_import(self):
        path = self.excel_path.get().strip()
        if not path:
            messagebox.showwarning("Uyarı", "Lütfen önce bir Excel dosyası seçin.")
            return
        if not os.path.exists(path):
            messagebox.showerror("Hata", f"Dosya bulunamadı:\n{path}")
            return

        self.import_btn.configure(state="disabled", text="⏳  Çalışıyor...")
        self.import_log.delete("1.0", "end")
        self.status_var.set("İmport çalışıyor...")

        def run():
            try:
                # Redirect stdout to log widget
                class LogWriter:
                    def __init__(self, widget):
                        self.widget = widget
                    def write(self, text):
                        self.widget.after(0, lambda t=text: self._append(t))
                    def _append(self, text):
                        self.widget.configure(state="normal")
                        self.widget.insert("end", text)
                        self.widget.see("end")
                    def flush(self): pass

                import sys
                old_stdout = sys.stdout
                sys.stdout = LogWriter(self.import_log)

                if self.opt_update_ef.get():
                    from run_update import run_update
                    run_update(excel_path=path)

                from pipeline.Importer import run_import
                run_import(path)

                sys.stdout = old_stdout

                self.after(0, lambda: self._import_done(success=True))
            except Exception as e:
                import traceback
                err = traceback.format_exc()
                self.after(0, lambda: self.import_log.insert("end", f"\n❌ HATA:\n{err}"))
                self.after(0, lambda: self._import_done(success=False))

        threading.Thread(target=run, daemon=True).start()

    def _import_done(self, success):
        self.import_btn.configure(state="normal", text="▶  İmport Et")
        if success:
            self.status_var.set("✅ İmport tamamlandı")
            self._load_db_summary()
            if messagebox.askyesno("Tamamlandı",
                "İmport başarıyla tamamlandı!\nSonuçları görüntülemek ister misiniz?"):
                self._show_page("results")
        else:
            self.status_var.set("❌ İmport başarısız")

    # ── Faktör Güncelleme Sayfası ──────────────────────────────────

    def _page_update(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(frame, text="Emisyon Faktörü Güncelleme", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="IPCC, DEFRA, TEİAŞ kaynaklarından güncel faktörleri çeker",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,16))

        # Kaynak seçimi
        src_card = tk.Frame(frame, bg=C["card"],
                            highlightbackground=C["border"], highlightthickness=1)
        src_card.pack(fill="x", pady=(0,12))
        src_inner = tk.Frame(src_card, bg=C["card"], padx=16, pady=14)
        src_inner.pack(fill="x")
        tk.Label(src_inner, text="Güncellenecek Kaynaklar", font=FONT_BOLD,
                 bg=C["card"], fg=C["text"]).pack(anchor="w", pady=(0,8))

        self.src_vars = {}
        sources = [
            ("gwp",   "🧪  GWP Değerleri (IPCC AR6)"),
            ("ipcc",  "🔥  Yakıt Faktörleri (IPCC 2006)"),
            ("teias", "⚡  Elektrik Grid Faktörü (TEİAŞ)"),
            ("defra", "🚛  Taşımacılık Faktörleri (DEFRA 2025)"),
        ]
        for key, label in sources:
            var = tk.BooleanVar(value=True)
            self.src_vars[key] = var
            tk.Checkbutton(src_inner, text=label, variable=var,
                           font=FONT, bg=C["card"], fg=C["text"],
                           activebackground=C["card"]
                           ).pack(anchor="w", pady=2)

        # Excel seçimi
        ex_card = tk.Frame(frame, bg=C["card"],
                           highlightbackground=C["border"], highlightthickness=1)
        ex_card.pack(fill="x", pady=(0,12))
        ex_inner = tk.Frame(ex_card, bg=C["card"], padx=16, pady=14)
        ex_inner.pack(fill="x")
        tk.Label(ex_inner, text="Excel Güncelleme (opsiyonel)",
                 font=FONT_BOLD, bg=C["card"], fg=C["text"]).pack(anchor="w")
        tk.Label(ex_inner,
                 text="Seçilirse boş yeşil hücreler otomatik doldurulur",
                 font=FONT_SMALL, bg=C["card"], fg=C["muted"]).pack(anchor="w", pady=(2,6))

        ex_row = tk.Frame(ex_inner, bg=C["card"])
        ex_row.pack(fill="x")
        tk.Entry(ex_row, textvariable=self.excel_path, font=FONT,
                 width=50, relief="solid", bd=1
                 ).pack(side="left", padx=(0,8))
        tk.Button(ex_row, text="Gözat...", font=FONT,
                  bg=C["blue"], fg="white", relief="flat",
                  padx=10, pady=3, cursor="hand2",
                  command=self._browse_excel
                  ).pack(side="left")

        # Çalıştır
        self.update_btn = tk.Button(
            frame, text="🔄  Güncellemeyi Başlat",
            font=FONT_BOLD, bg=C["green"], fg="white",
            relief="flat", padx=24, pady=10, cursor="hand2",
            command=self._run_update
        )
        self.update_btn.pack(anchor="w", pady=8)

        # Log
        tk.Label(frame, text="Güncelleme Logu", font=FONT_BOLD,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(8,4))
        self.update_log = scrolledtext.ScrolledText(
            frame, height=14, font=FONT_MONO,
            bg="#1E1E1E", fg="#D4D4D4",
            relief="flat", highlightbackground=C["border"], highlightthickness=1
        )
        self.update_log.pack(fill="both", expand=True)
        self.update_log.insert("end", "Güncelleme logu burada görünecek...\n")

    def _run_update(self):
        sources = [k for k, v in self.src_vars.items() if v.get()]
        if not sources:
            messagebox.showwarning("Uyarı", "En az bir kaynak seçin.")
            return

        excel = self.excel_path.get().strip() or None
        self.update_btn.configure(state="disabled", text="⏳  Güncelleniyor...")
        self.update_log.delete("1.0", "end")
        self.status_var.set("Faktörler güncelleniyor...")

        def run():
            try:
                class LogWriter:
                    def __init__(self, widget):
                        self.widget = widget
                    def write(self, text):
                        self.widget.after(0, lambda t=text: self._append(t))
                    def _append(self, text):
                        self.widget.configure(state="normal")
                        self.widget.insert("end", text)
                        self.widget.see("end")
                    def flush(self): pass

                import sys
                old_stdout = sys.stdout
                sys.stdout  = LogWriter(self.update_log)
                from run_update import run_update
                run_update(sources=sources, excel_path=excel)
                sys.stdout = old_stdout
                self.after(0, lambda: self._update_done(True))
            except Exception as e:
                import traceback
                err = traceback.format_exc()
                self.after(0, lambda: self.update_log.insert("end", f"\n❌ HATA:\n{err}"))
                self.after(0, lambda: self._update_done(False))

        threading.Thread(target=run, daemon=True).start()

    def _update_done(self, success):
        self.update_btn.configure(state="normal", text="🔄  Güncellemeyi Başlat")
        self.status_var.set("✅ Güncelleme tamamlandı" if success else "❌ Hata oluştu")

    # ── Sonuçlar & Grafik Sayfası ──────────────────────────────────

    def _page_results(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(frame, text="Sonuçlar & Grafikler", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="Emisyon hesaplama sonuçları ve görselleştirme",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,12))

        # Sekme çubuğu
        tab_frame = tk.Frame(frame, bg=C["bg"])
        tab_frame.pack(fill="x", pady=(0,12))

        self.result_tab = tk.StringVar(value="chart")
        for text, val in [("📊 Grafik", "chart"), ("📋 Tablo", "table"),
                          ("🎯 Etki Analizi", "impact"), ("👤 Kişi Başına", "percap")]:
            tk.Radiobutton(tab_frame, text=text, variable=self.result_tab,
                           value=val, font=FONT, bg=C["bg"], fg=C["text"],
                           activebackground=C["bg"], cursor="hand2",
                           command=self._refresh_results
                           ).pack(side="left", padx=(0,16))

        self.results_content = tk.Frame(frame, bg=C["bg"])
        self.results_content.pack(fill="both", expand=True)
        self._refresh_results()

    def _refresh_results(self):
        for w in self.results_content.winfo_children():
            w.destroy()

        tab = self.result_tab.get()
        if   tab == "chart":  self._results_chart()
        elif tab == "table":  self._results_table()
        elif tab == "impact": self._results_impact()
        elif tab == "percap": self._results_percap()

    def _get_scope_totals(self):
        """DB'den yıl bazlı kapsam toplamlarını çeker."""
        try:
            from db.connection import SessionLocal
            from db.models import AuditLog
            import re, json
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
                                'k1': float(m.group(2)),
                                'k2': float(m.group(3)),
                                'k3': float(m.group(4)),
                                'total': float(m.group(5))
                            }
            return years
        except:
            return {}

    def _results_chart(self):
        try:
            import matplotlib
            matplotlib.use("TkAgg")
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            years_data = self._get_scope_totals()
            if not years_data:
                tk.Label(self.results_content,
                         text="Henüz import edilmiş veri yok.\nÖnce 'Veri İmport' sayfasından import yapın.",
                         font=FONT, bg=C["bg"], fg=C["muted"]).pack(expand=True)
                return

            fig = Figure(figsize=(11, 6), facecolor=C["bg"])

            # Sol: Yıllık karşılaştırma (grouped bar)
            ax1 = fig.add_subplot(121)
            years  = sorted(years_data.keys())
            k1_vals = [years_data[y]['k1']    for y in years]
            k2_vals = [years_data[y]['k2']    for y in years]
            k3_vals = [years_data[y]['k3']    for y in years]
            x      = range(len(years))
            w      = 0.25

            ax1.bar([i-w for i in x], k1_vals, w, label="Kapsam 1",
                    color=C["s1"], alpha=0.85)
            ax1.bar([i   for i in x], k2_vals, w, label="Kapsam 2",
                    color=C["s2"], alpha=0.85)
            ax1.bar([i+w for i in x], k3_vals, w, label="Kapsam 3",
                    color=C["s3"], alpha=0.85)
            ax1.set_xticks(list(x))
            ax1.set_xticklabels(years, fontsize=9)
            ax1.set_ylabel("ton CO2e", fontsize=9)
            ax1.set_title("Yıllık Kapsam Karşılaştırması", fontsize=10, fontweight="bold")
            ax1.legend(fontsize=8)
            ax1.set_facecolor(C["bg"])
            fig.patch.set_facecolor(C["bg"])
            ax1.spines['top'].set_visible(False)
            ax1.spines['right'].set_visible(False)

            # Sağ: En son yıl pie chart
            ax2 = fig.add_subplot(122)
            latest = max(years_data.keys())
            d = years_data[latest]
            vals   = [d['k1'], d['k2'], d['k3']]
            labels = [f"Kapsam 1\n{d['k1']:.0f}t",
                      f"Kapsam 2\n{d['k2']:.0f}t",
                      f"Kapsam 3\n{d['k3']:.0f}t"]
            colors = [C["s1"], C["s2"], C["s3"]]
            non_zero = [(v, l, c) for v, l, c in zip(vals, labels, colors) if v > 0]
            if non_zero:
                v2, l2, c2 = zip(*non_zero)
                ax2.pie(v2, labels=l2, colors=c2, autopct='%1.1f%%',
                        startangle=90, textprops={'fontsize': 8})
            ax2.set_title(f"{latest} Kapsam Dağılımı\nToplam: {d['total']:.1f} ton CO2e",
                          fontsize=10, fontweight="bold")
            ax2.set_facecolor(C["bg"])

            fig.tight_layout(pad=2)

            canvas = FigureCanvasTkAgg(fig, master=self.results_content)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except ImportError:
            tk.Label(self.results_content,
                     text="Grafik için matplotlib gerekli:\n pip install matplotlib",
                     font=FONT, bg=C["bg"], fg=C["red"]).pack(expand=True)

    def _results_table(self):
        years_data = self._get_scope_totals()
        if not years_data:
            tk.Label(self.results_content,
                     text="Veri bulunamadı. Önce import yapın.",
                     font=FONT, bg=C["bg"], fg=C["muted"]).pack(expand=True)
            return

        # Treeview tablosu
        cols = ["Yıl", "Kapsam 1", "Kapsam 2", "Kapsam 3", "Toplam"]
        tree = ttk.Treeview(self.results_content, columns=cols,
                            show="headings", height=12)
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=160, anchor="center")

        style = ttk.Style()
        style.configure("Treeview", font=FONT, rowheight=28)
        style.configure("Treeview.Heading", font=FONT_BOLD)

        for y in sorted(years_data.keys()):
            d = years_data[y]
            tree.insert("", "end", values=(
                y,
                f"{d['k1']:.2f} ton",
                f"{d['k2']:.2f} ton",
                f"{d['k3']:.2f} ton",
                f"{d['total']:.2f} ton",
            ))

        tree.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(self.results_content, orient="vertical",
                           command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

    def _results_impact(self):
        """IMPACT_ANALYSIS verilerini gösterir."""
        try:
            from db.connection import SessionLocal
            from db.models import AuditLog
            import re

            # En son impact analysis log'undan çek
            # (veya doğrudan hesapla)
            tk.Label(self.results_content,
                     text="En Büyük Emisyon Kaynakları (%95 Eşiği)",
                     font=FONT_BOLD, bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(0,8))

            years_data = self._get_scope_totals()
            if not years_data:
                tk.Label(self.results_content, text="Veri yok.",
                         font=FONT, bg=C["bg"], fg=C["muted"]).pack()
                return

            latest = max(years_data.keys())
            d = years_data[latest]

            sources = [
                ("Kapsam 1", "1.1", "Sabit Yakma",            d['k1'] * 0.02),
                ("Kapsam 1", "1.2", "Hareketli Yakma",        d['k1'] * 0.43),
                ("Kapsam 1", "1.4", "Soğutucu Gazlar",        d['k1'] * 0.55),
                ("Kapsam 2", "2.1", "Elektrik Tüketimi",      d['k2']),
                ("Kapsam 3", "6.1", "T&D Kayıpları",          d['k3']),
            ]
            total = sum(v for _, _, _, v in sources)

            cols = ["Sıra", "Kapsam", "Kategori", "Kaynak", "CO2e (ton)", "Pay (%)"]
            tree = ttk.Treeview(self.results_content, columns=cols,
                                show="headings", height=10)
            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=130, anchor="center")
            tree.column("Kaynak", width=200)

            sources_sorted = sorted(sources, key=lambda x: x[3], reverse=True)
            cumulative = 0
            for i, (scope, cat, label, val) in enumerate(sources_sorted, 1):
                pct       = (val / total * 100) if total else 0
                cumulative += pct
                tag = "above" if cumulative <= 95 else "below"
                tree.insert("", "end", values=(
                    i, scope, cat, label,
                    f"{val:.2f}", f"{pct:.1f}%"
                ), tags=(tag,))

            tree.tag_configure("above", background="#FCE4D6")
            tree.tag_configure("below", background=C["bg"])
            tree.pack(fill="both", expand=True)

            tk.Label(self.results_content,
                     text="🟠 Turuncu = %95 eşiğine kadar olan önemli kaynaklar",
                     font=FONT_SMALL, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=4)

        except Exception as e:
            tk.Label(self.results_content, text=f"Hata: {e}",
                     font=FONT, bg=C["bg"], fg=C["red"]).pack()

    def _results_percap(self):
        """Kişi başına emisyon tablosunu gösterir."""
        try:
            from db.connection import SessionLocal
            from db.models import ReportingPeriod

            tk.Label(self.results_content,
                     text="Kişi Başına Emisyon — Lokasyon Bazında",
                     font=FONT_BOLD, bg=C["bg"], fg=C["text"]).pack(anchor="w", pady=(0,8))

            # PER_CAPITA Excel sayfasından oku
            excel = self.excel_path.get().strip()
            if not excel or not os.path.exists(excel):
                tk.Label(self.results_content,
                         text="Excel dosyasını 'Veri İmport' sayfasından seçin.",
                         font=FONT, bg=C["bg"], fg=C["muted"]).pack()
                return

            from openpyxl import load_workbook
            wb = load_workbook(excel, read_only=True, data_only=True)
            ws = wb['PER_CAPITA']

            cols = ["Yıl", "Lokasyon", "Personel", "Toplam CO2e",
                    "Kapsam 1", "Kapsam 2", "Kapsam 3", "ton CO2e/Kişi"]
            tree = ttk.Treeview(self.results_content, columns=cols,
                                show="headings", height=12)
            for col in cols:
                tree.heading(col, text=col)
                tree.column(col, width=120, anchor="center")
            tree.column("Lokasyon", width=200)

            for row in ws.iter_rows(min_row=4, values_only=True):
                if row[0] and row[1]:
                    vals = tuple(
                        f"{v:.4f}" if isinstance(v, float) else str(v) if v else "—"
                        for v in row[:8]
                    )
                    tag = "total" if row[1] and "TOPLAM" in str(row[1]) else "loc"
                    tree.insert("", "end", values=vals, tags=(tag,))

            tree.tag_configure("total", background=C["warning"], font=FONT_BOLD)
            tree.pack(fill="both", expand=True)

        except Exception as e:
            tk.Label(self.results_content, text=f"Hata: {e}",
                     font=FONT, bg=C["bg"], fg=C["red"]).pack()

    # ── Log Sayfası ────────────────────────────────────────────────

    def _page_log(self):
        frame = tk.Frame(self.content, bg=C["bg"])
        frame.pack(fill="both", expand=True, padx=24, pady=20)

        tk.Label(frame, text="Log Geçmişi", font=FONT_TITLE,
                 bg=C["bg"], fg=C["text"]).pack(anchor="w")
        tk.Label(frame, text="Tüm import ve güncelleme işlemlerinin kaydı",
                 font=FONT, bg=C["bg"], fg=C["muted"]).pack(anchor="w", pady=(0,12))

        # Treeview
        cols = ["Tarih", "İşlem", "Kapsam", "Durum", "Notlar"]
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=20)
        for col, w in zip(cols, [160, 140, 100, 80, 400]):
            tree.heading(col, text=col)
            tree.column(col, width=w)

        style = ttk.Style()
        style.configure("Treeview", font=FONT, rowheight=26)

        try:
            from db.connection import SessionLocal
            from db.models import AuditLog
            with SessionLocal() as s:
                logs = s.query(AuditLog).order_by(AuditLog.id.desc()).limit(100).all()
                for log in logs:
                    tag = "success" if log.status == "success" else "fail"
                    notes = (log.notes or "")[:80]
                    tree.insert("", "end", values=(
                        str(log.created_at)[:19] if log.created_at else "—",
                        log.action or "—",
                        log.scope or "—",
                        log.status or "—",
                        notes,
                    ), tags=(tag,))
        except Exception as e:
            tree.insert("", "end", values=("—", "DB Hatası", "—", "error", str(e)))

        tree.tag_configure("success", background=C["success"])
        tree.tag_configure("fail",    background=C["danger"])

        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        tk.Button(frame, text="🔄 Yenile", font=FONT,
                  bg=C["blue"], fg="white", relief="flat",
                  padx=12, pady=6, cursor="hand2",
                  command=lambda: self._show_page("log")
                  ).pack(side="bottom", anchor="w", pady=8)


# ══════════════════════════════════════════════════════════════════
# BAŞLAT
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app = UNSPEDDashboard()
    app.mainloop()
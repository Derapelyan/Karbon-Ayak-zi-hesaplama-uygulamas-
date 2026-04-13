"""
UNSPED Carbon Footprint System — Setup Script
Run this ONCE on a new machine to initialize the database.

    python main.py

What it does:
    1. Creates all DB tables
    2. Creates UNSPED company + reporting periods (dynamic — last 2 years + current)
    3. Prints confirmation

Emission factors come from the Excel green cells — NOT hardcoded here.
To fill empty green cells automatically, run:
    python run_update.py --excel data/UNSPED_Karbon_Veri_Girisi_v3.xlsx

To import data after filling Excel:
    python pipeline/Importer.py data/UNSPED_Karbon_Veri_Girisi_v3.xlsx
"""

import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from db.connection import engine, Base, SessionLocal
from db.models import Company, ReportingPeriod, AuditLog


# ══════════════════════════════════════════════════════════════════
# STEP 1 — CREATE TABLES
# ══════════════════════════════════════════════════════════════════

def init_db():
    """Create all tables in carbon.db. Safe to run multiple times."""
    Base.metadata.create_all(engine)
    print("OK  Veritabani ve tablolar olusturuldu.")


# ══════════════════════════════════════════════════════════════════
# STEP 2 — COMPANY & REPORTING PERIODS
# ══════════════════════════════════════════════════════════════════

# Raporlamanın başladığı yıl — değiştirmeyin
REPORTING_START_YEAR = 2023

# Grafik/tablolarda gösterilecek maksimum yıl sayısı
DISPLAY_YEARS = 3

def setup_company():
    """
    Create UNSPED company and annual reporting periods.

    Mantık:
        - DB'de REPORTING_START_YEAR'dan (current_year - 1)'e kadar tüm yıllar tutulur
        - current_year (2026 gibi) henüz bitmediği için oluşturulmaz
        - Grafik/tablolarda sadece son DISPLAY_YEARS yıl gösterilir

    Örnek (2026'da çalıştırılınca):
        DB dönemleri → [2023, 2024, 2025]
        Grafik        → [2023, 2024, 2025]  (son 3)

    Örnek (2027'de çalıştırılınca):
        DB dönemleri → [2023, 2024, 2025, 2026]
        Grafik        → [2024, 2025, 2026]  (son 3)

    Safe to run multiple times — skips already existing periods.
    """
    current_year = datetime.now().year
    last_complete_year = current_year - 1  # bu yıl henüz bitmedi
    years = list(range(REPORTING_START_YEAR, last_complete_year + 1))

    with SessionLocal() as session:

        # Get or create company
        company = session.query(Company).filter_by(name="UNSPED").first()
        if not company:
            company = Company(
                name="UNSPED",
                site="Ana Yerleske - Makine Uretim Tesisi"
            )
            session.add(company)
            session.flush()
            print("OK  UNSPED sirketi olusturuldu.")
        else:
            print("--  UNSPED zaten mevcut.")

        # Create any missing periods
        created = []
        for year in years:
            exists = session.query(ReportingPeriod).filter_by(
                company_id=company.id,
                year=year,
                period="annual"
            ).first()
            if not exists:
                session.add(ReportingPeriod(
                    company_id=company.id,
                    year=year,
                    period="annual"
                ))
                created.append(year)

        if created:
            session.add(AuditLog(
                action="setup_company",
                scope="all",
                status="success",
                notes=f"Yeni donemler olusturuldu: {created}"
            ))
            session.commit()
            print(f"OK  Raporlama donemleri olusturuldu: {created}")
        else:
            print(f"--  Tum donemler zaten mevcut: {years}")


# ══════════════════════════════════════════════════════════════════
# STEP 3 — SUMMARY
# ══════════════════════════════════════════════════════════════════

def show_summary():
    with SessionLocal() as session:
        companies = session.query(Company).all()
        periods   = session.query(ReportingPeriod).all()

        print("\n" + "="*50)
        print("   UNSPED KARBON SISTEMI - HAZIR")
        print("="*50)

        for c in companies:
            print(f"  Sirket  : {c.name}")
            print(f"  Tesis   : {c.site}")

        print(f"\n  Raporlama Donemleri:")
        for p in sorted(periods, key=lambda x: x.year):
            status = "Kapali" if p.is_closed else "Acik"
            print(f"    {p.year}  {p.period}  {status}")

        print(f"\n  Emisyon Faktorleri: Excel'den okunur")
        print(f"  Bos yesil hucreler icin:")
        print(f"    python run_update.py --excel data/UNSPED_Karbon_Veri_Girisi_v3.xlsx")
        print(f"\n  Import icin:")
        print(f"    python pipeline/Importer.py data/UNSPED_Karbon_Veri_Girisi_v3.xlsx")
        print("="*50 + "\n")


# ══════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print(f"\nUNSPED Karbon Sistemi Kurulumu - {datetime.now().year}\n")
    init_db()
    setup_company()
    show_summary()
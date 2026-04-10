"""
UNSPED Carbon Footprint System — Setup Script
Run this ONCE on a new machine to initialize the database.

    python main.py

What it does:
    1. Creates all tables in carbon.db
    2. Creates UNSPED company + reporting periods
    3. Seeds emission factors (IPCC, TEİAŞ)
    4. Prints a summary to confirm everything is ready

After this, use the importer for annual data:
    python pipeline/importer.py data/UNSPED_Karbon_Veri_Girisi_v3.xlsx 2024
"""

from db.connection import engine, Base, SessionLocal
from db.models import (
    Company, ReportingPeriod, EmissionFactor, AuditLog
)


# ══════════════════════════════════════════════════════════════════
# STEP 1 — CREATE TABLES
# ══════════════════════════════════════════════════════════════════

def init_db():
    """Create all tables in carbon.db. Safe to run multiple times."""
    Base.metadata.create_all(engine)
    print("✅ Veritabanı ve tablolar oluşturuldu.")


# ══════════════════════════════════════════════════════════════════
# STEP 2 — COMPANY & REPORTING PERIODS
# ══════════════════════════════════════════════════════════════════

def setup_company():
    """Create UNSPED company and annual reporting periods."""
    with SessionLocal() as session:
        existing = session.query(Company).filter_by(name="UNSPED").first()
        if existing:
            print("ℹ️  UNSPED zaten mevcut, atlanıyor.")
            return

        company = Company(
            name="UNSPED",
            site="Ana Yerleşke - Makine Üretim Tesisi"
        )
        session.add(company)
        session.flush()

        for year in [2023, 2024]:
            session.add(ReportingPeriod(
                company_id=company.id,
                year=year,
                period="annual"
            ))

        session.commit()
        print("✅ UNSPED şirketi ve raporlama dönemleri (2023, 2024) oluşturuldu.")


# ══════════════════════════════════════════════════════════════════
# STEP 3 — SEED EMISSION FACTORS
# ══════════════════════════════════════════════════════════════════

def seed_emission_factors():
    """
    Load base emission factors into the database.

    These are the starting values — the fetchers will update them
    automatically from IPCC / DEFRA / TEİAŞ when run_update.py is used.

    Unit conversions stored here:
        Natural gas : m³  → ton  (density 0.000717 ton/m³)
        Diesel      : L   → ton  (density 0.000832 ton/L)
        Petrol      : L   → ton  (density 0.000745 ton/L)
    """
    with SessionLocal() as session:
        if session.query(EmissionFactor).first():
            print("ℹ️  Emisyon faktörleri zaten yüklü, atlanıyor.")
            return

        factors = [

            # ── Scope 1 | 1.1 Sabit Yakma ──────────────────────────
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope1", category="1.1",
                fuel_type="natural_gas", region="TR", unit="Ton/TJ",
                co2_factor=56.1, ch4_factor=0.005, n2o_factor=0.0001,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=48, ncv_unit="TJ/Gg",
                density=0.000717, density_unit="ton/m³",
                input_unit="m³", valid_year=2024
            ),
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope1", category="1.1",
                fuel_type="diesel", region="TR", unit="Ton/TJ",
                co2_factor=74.1, ch4_factor=0.01, n2o_factor=0.0006,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=43, ncv_unit="TJ/Gg",
                density=0.000832, density_unit="ton/L",
                input_unit="L", valid_year=2024
            ),
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope1", category="1.1",
                fuel_type="petrol", region="TR", unit="Ton/TJ",
                co2_factor=69.3, ch4_factor=0.01, n2o_factor=0.0006,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=44.3, ncv_unit="TJ/Gg",
                density=0.000745, density_unit="ton/L",
                input_unit="L", valid_year=2024
            ),

            # ── Scope 1 | 1.2 Hareketli Yakma ──────────────────────
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope1", category="1.2",
                fuel_type="petrol", region="TR", unit="Ton/TJ",
                co2_factor=69.3, ch4_factor=0.01, n2o_factor=6e-05,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=44.3, ncv_unit="TJ/Gg",
                density=0.000745, density_unit="ton/L",
                input_unit="L", valid_year=2024
            ),
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope1", category="1.2",
                fuel_type="diesel", region="TR", unit="Ton/TJ",
                co2_factor=74.1, ch4_factor=0.01, n2o_factor=6e-05,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=43, ncv_unit="TJ/Gg",
                density=0.000832, density_unit="ton/L",
                input_unit="L", valid_year=2024
            ),

            # ── Scope 2 | 2.1 Elektrik ──────────────────────────────
            EmissionFactor(
                source="TEİAŞ/EPDK REF#10", scope="scope2", category="2.1",
                fuel_type="electricity", region="TR", unit="ton CO2e/MWh",
                co2_factor=0.442, input_unit="MWh", valid_year=2023
            ),
            EmissionFactor(
                source="TEİAŞ/EPDK REF#10", scope="scope2", category="2.1",
                fuel_type="electricity", region="TR", unit="ton CO2e/MWh",
                co2_factor=0.434, input_unit="MWh", valid_year=2024
            ),

            # ── Scope 3 | 3.3 Personel Ulaşım ───────────────────────
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope3", category="3.3",
                fuel_type="petrol", region="TR", unit="Ton/TJ",
                co2_factor=69.3, ch4_factor=0.01, n2o_factor=6e-05,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=44.3, ncv_unit="TJ/Gg",
                density=0.000745, density_unit="ton/L",
                input_unit="L", valid_year=2024
            ),
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope3", category="3.3",
                fuel_type="diesel", region="TR", unit="Ton/TJ",
                co2_factor=74.1, ch4_factor=0.01, n2o_factor=6e-05,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=43, ncv_unit="TJ/Gg",
                density=0.000832, density_unit="ton/L",
                input_unit="L", valid_year=2024
            ),
            EmissionFactor(
                source="IPCC REF#3/REF#5", scope="scope3", category="3.3",
                fuel_type="diesel_shuttle", region="TR", unit="Ton/TJ",
                co2_factor=74.1, ch4_factor=0.01, n2o_factor=6e-05,
                gwp_ch4=29.8, gwp_n2o=273,
                ncv=43, ncv_unit="TJ/Gg",
                density=0.000832, density_unit="ton/L",
                input_unit="L", valid_year=2024
            ),
        ]

        for f in factors:
            session.add(f)

        session.add(AuditLog(
            action="seed_emission_factors",
            scope="all",
            status="success",
            notes=f"{len(factors)} emisyon faktörü yüklendi."
        ))
        session.commit()
        print(f"✅ {len(factors)} emisyon faktörü yüklendi.")


# ══════════════════════════════════════════════════════════════════
# STEP 4 — CONFIRM SETUP
# ══════════════════════════════════════════════════════════════════

def show_summary():
    """Print a summary confirming the DB is ready."""
    with SessionLocal() as session:
        companies = session.query(Company).all()
        periods   = session.query(ReportingPeriod).all()
        factors   = session.query(EmissionFactor).all()

        print("\n─────────────────────────────────────────")
        print("       UNSPED KARBON SİSTEMİ — HAZIR     ")
        print("─────────────────────────────────────────")
        for c in companies:
            print(f"  🏢 Şirket  : {c.name}")
            print(f"  📍 Tesis   : {c.site}")

        print(f"\n  📅 Raporlama Dönemleri:")
        for p in periods:
            status = "🔒 Kapalı" if p.is_closed else "🟢 Açık"
            print(f"     {p.year} — {p.period} — {status}")

        print(f"\n  📊 Emisyon Faktörleri ({len(factors)} adet):")
        for f in factors:
            print(f"     [{f.category}] {f.fuel_type:<18} "
                  f"input: {str(f.input_unit):<5}  "
                  f"co2_ef: {f.co2_factor}")

        print("\n─────────────────────────────────────────")
        print("  ✅ Kurulum tamamlandı. İmporter'ı çalıştırabilirsiniz:")
        print("  python pipeline/importer.py data/UNSPED_Karbon_Veri_Girisi_v3.xlsx 2024")
        print("─────────────────────────────────────────\n")


# ══════════════════════════════════════════════════════════════════
# ENTRY POINT
# ══════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("\n🚀 UNSPED Karbon Sistemi Kurulumu Başlıyor...\n")
    init_db()
    setup_company()
    seed_emission_factors()
    show_summary()
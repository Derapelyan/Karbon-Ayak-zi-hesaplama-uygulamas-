from sqlalchemy import Column, Integer, String, Float, Boolean, DateTime, Text, ForeignKey
from sqlalchemy.sql import func
from .connection import Base

# ══════════════════════════════════════════════════════════════════
# COMPANY & REPORTING PERIOD
# ══════════════════════════════════════════════════════════════════

class Company(Base):
    __tablename__ = "companies"

    id         = Column(Integer, primary_key=True)
    name       = Column(String(100), nullable=False)  # UNSPED
    site       = Column(String(200))                  # tesis / şube adı
    created_at = Column(DateTime, default=func.now())


class ReportingPeriod(Base):
    __tablename__ = "reporting_periods"

    id         = Column(Integer, primary_key=True)
    company_id = Column(Integer, ForeignKey("companies.id"), nullable=False)
    year       = Column(Integer, nullable=False)           # 2024, 2025
    period     = Column(String(20), default="annual")      # "annual" | "monthly"
    month      = Column(Integer, nullable=True)            # 1-12 if monthly, NULL if annual
    is_closed  = Column(Boolean, default=False)            # True → period kilitli
    created_at = Column(DateTime, default=func.now())


# ══════════════════════════════════════════════════════════════════
# EMISSION FACTORS — Referans Tablo
# Sadece DEFRA online çekilerek buraya kaydedilir.
# IPCC/EVCED değerleri Excel'den okunur, buraya kaydedilmez.
# ══════════════════════════════════════════════════════════════════

class EmissionFactor(Base):
    __tablename__ = "emission_factors"

    id           = Column(Integer, primary_key=True)
    source       = Column(String(100))          # "DEFRA 2025", "TC ETKB EVCED"
    scope        = Column(String(10))           # "scope1", "scope2", "scope3"
    category     = Column(String(50))           # "1.1", "2.1", "3.x"
    fuel_type    = Column(String(100))          # "All HGVs", "electricity"
    region       = Column(String(50))           # "TR", "global"
    unit         = Column(String(50))           # "kg CO2e/tonne.km", "ton CO2e/MWh"
    co2_factor   = Column(Float)               # ana emisyon faktörü
    ch4_factor   = Column(Float, nullable=True)
    n2o_factor   = Column(Float, nullable=True)
    gwp_ch4      = Column(Float, nullable=True) # GWP değeri (soğutucu için)
    gwp_n2o      = Column(Float, nullable=True)
    ncv          = Column(Float, nullable=True) # Net Kalorifik Değer
    ncv_unit     = Column(String(50), nullable=True)
    density      = Column(Float, nullable=True)
    density_unit = Column(String(50), nullable=True)
    input_unit   = Column(String(20), nullable=True)  # "km", "MWh", "L"
    is_active    = Column(Boolean, default=True)
    version      = Column(Integer, default=1)
    valid_year   = Column(Integer, nullable=True)
    created_at   = Column(DateTime, default=func.now())


# ══════════════════════════════════════════════════════════════════
# SCOPE 1 — STATIONARY COMBUSTION (1.1)
# ══════════════════════════════════════════════════════════════════

class StationaryCombustion(Base):
    __tablename__ = "scope1_stationary"

    id             = Column(Integer, primary_key=True)
    period_id      = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    emission_source= Column(String(100))   # "Kombi", "Jeneratör"  ← typo düzeltildi
    fuel_type      = Column(String(50))    # "natural_gas", "diesel"
    activity_value = Column(Float)         # tüketim miktarı
    activity_unit  = Column(String(20))    # "m³", "L"
    co2e_total     = Column(Float)         # hesaplanan ton CO2e
    created_at     = Column(DateTime, default=func.now())

    # NOT: emission_factor_id kaldırıldı
    # Faktörler Excel yeşil hücrelerden okunur, DB'ye kaydedilmez


# ══════════════════════════════════════════════════════════════════
# SCOPE 1 — MOBILE COMBUSTION (1.2)
# ══════════════════════════════════════════════════════════════════

class MobileCombustion(Base):
    __tablename__ = "scope1_mobile"

    id             = Column(Integer, primary_key=True)
    period_id      = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    emission_source= Column(String(100))   # "Binek Araçlar"
    fuel_type      = Column(String(50))    # "petrol", "diesel"
    activity_value = Column(Float)
    activity_unit  = Column(String(20))    # "L"
    co2e_total     = Column(Float)
    created_at     = Column(DateTime, default=func.now())

    # NOT: emission_factor_id kaldırıldı


# ══════════════════════════════════════════════════════════════════
# SCOPE 1 — REFRIGERANTS (1.4)
# ══════════════════════════════════════════════════════════════════

class Refrigerant(Base):
    __tablename__ = "scope1_refrigerants"

    id             = Column(Integer, primary_key=True)
    period_id      = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    gas_type       = Column(String(50))    # "R134A", "FM200", "R410A"
    activity_value = Column(Float)         # kaçak miktarı (kg)
    activity_unit  = Column(String(20))    # "kg"
    gwp            = Column(Float)         # Excel'den okunan GWP değeri
    co2e_total     = Column(Float)         # activity_value * gwp / 1000
    created_at     = Column(DateTime, default=func.now())


# ══════════════════════════════════════════════════════════════════
# SCOPE 2 — ELECTRICITY (2.1)
# ══════════════════════════════════════════════════════════════════

class ElectricityConsumption(Base):
    __tablename__ = "scope2_electricity"

    id              = Column(Integer, primary_key=True)
    period_id       = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    consumption_mwh = Column(Float)        # MWh cinsinden tüketim
    emission_factor = Column(Float)        # ton CO2e/MWh (EVCED/TEİAŞ)
    ef_source       = Column(String(100))  # "TC ETKB EVCED 2023"
    co2e_total      = Column(Float)        # consumption_mwh × emission_factor
    created_at      = Column(DateTime, default=func.now())


# ══════════════════════════════════════════════════════════════════
# SCOPE 3 — FREIGHT / SHIPMENT (3.1, 3.2, 3.4)
# ══════════════════════════════════════════════════════════════════

class FreightEmission(Base):
    __tablename__ = "scope3_freight"

    id              = Column(Integer, primary_key=True)
    period_id       = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    category_code   = Column(String(10))   # "3.1.1", "3.2.1", "3.4.1"
    emission_source = Column(String(100))  # "Hammadde Sevkiyatı"
    transport_type  = Column(String(100))  # "Karayolu-%100 Laden"
    vehicle_type    = Column(String(100))  # "All HGVs", "Small car"
    activity_value  = Column(Float)        # mesafe (km)
    activity_unit   = Column(String(20))   # "km"
    emission_factor = Column(Float)        # kg CO2e/tonne.km
    co2e_total      = Column(Float)
    created_at      = Column(DateTime, default=func.now())


# ══════════════════════════════════════════════════════════════════
# SCOPE 3 — EMPLOYEE COMMUTING (3.3)
# ══════════════════════════════════════════════════════════════════

class EmployeeCommuting(Base):
    __tablename__ = "scope3_commuting"

    id              = Column(Integer, primary_key=True)
    period_id       = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    category_code   = Column(String(10))   # "3.3.1", "3.3.2", "3.3.3"
    emission_source = Column(String(100))  # "Şahsi Araç", "Servis"
    fuel_type       = Column(String(50))   # "petrol", "diesel", "diesel_shuttle"
    activity_value  = Column(Float)        # tüketim (L)
    activity_unit   = Column(String(20))   # "L"
    co2e_total      = Column(Float)
    created_at      = Column(DateTime, default=func.now())

    # NOT: emission_factor_id kaldırıldı


# ══════════════════════════════════════════════════════════════════
# SCOPE 3 — PURCHASED GOODS & SERVICES (4.1, 4.3, 4.4, 4.5, 3.5)
# ══════════════════════════════════════════════════════════════════

class PurchasedGoods(Base):
    __tablename__ = "scope3_purchased_goods"

    id              = Column(Integer, primary_key=True)
    period_id       = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    category_code   = Column(String(10))   # "4.1.1", "3.5.1"
    item_name       = Column(String(200))  # "Printer Kağıt", "Uçak Bileti"
    activity_value  = Column(Float)        # miktar veya USD değeri
    activity_unit   = Column(String(50))   # "ton", "2021 USD", "TL"
    emission_factor = Column(Float)        # kg CO2e/birim
    ef_source       = Column(String(200))  # "DEFRA", "SupplyChainGHG"
    co2e_total      = Column(Float)
    created_at      = Column(DateTime, default=func.now())


# ══════════════════════════════════════════════════════════════════
# SCOPE 3 — CAPITAL GOODS (4.2)
# ══════════════════════════════════════════════════════════════════

class CapitalGoods(Base):
    __tablename__ = "scope3_capital_goods"

    id                 = Column(Integer, primary_key=True)
    period_id          = Column(Integer, ForeignKey("reporting_periods.id"), nullable=False)
    category_code      = Column(String(10))   # "4.2.1"
    asset_name         = Column(String(200))  # "Bilgi İşlem Cihazları"
    activity_value     = Column(Float)        # USD değeri
    activity_unit      = Column(String(50))   # "2021 USD"
    depreciation_years = Column(Integer)      # amortisman süresi
    emission_factor    = Column(Float)        # kg CO2e/2021 USD
    ef_source          = Column(String(200))
    co2e_total         = Column(Float)
    created_at         = Column(DateTime, default=func.now())


# ══════════════════════════════════════════════════════════════════
# AUDIT LOG
# ══════════════════════════════════════════════════════════════════

class AuditLog(Base):
    __tablename__ = "audit_log"

    id         = Column(Integer, primary_key=True)
    action     = Column(String(100))                               # "excel_import", "run_update"
    scope      = Column(String(20))                                # "all", "scope1"
    period_id  = Column(Integer, ForeignKey("reporting_periods.id"), nullable=True)  # ← nullable
    status     = Column(String(20))                                # "success", "failed"
    notes      = Column(Text)
    created_at = Column(DateTime, default=func.now())
from sqlalchemy import Column, Integer, String, Float, Boolean, DateTime, Text, ForeignKey
from sqlalchemy.orm import relationship
from sqlalchemy.sql import func
from .connection import Base

# ----------------------------------------
# ÜNSPED & YILLIK RAPORLAMA
# ----------------------------------------

class Company(Base): 
    __tablename__ = "companies"

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False)  #ünsped
    site = Column(String(200))                  #Şube ve ismi
    created_at = Column(DateTime, default=func.now())

class ReportingPeriod(Base):
    __tablename__ = "reporting_periods"

    id = Column(Integer, primary_key=True)
    company_id = Column(Integer,ForeignKey("companies.id"))
    year = Column(Integer, nullable=False)
    period = Column(String(20), default="annual") # "annual" or " monthly"
    month = Column(Integer, nullable=True)  #1-12 if monthly, null if annual
    is_closed = Column(Boolean,default=False)   #lock period after submission
    created_at = Column(DateTime, default=func.now())
#---------------------------------------------------------------
#Emission Factors (referance table)
#---------------------------------------------------------------
class EmissionFactor(Base):
    __tablename__ = "emission_factors"

    id = Column(Integer, primary_key=True)
    source = Column(String(100)) #"ipcc", "Defra", "epa"
    scope = Column(String(10)) # scope1, scope2
    category = Column(String(50)) # 1.1, 1.2
    fuel_type= Column(String(100)) #natural_gas, diesel petrol
    region = Column(String(50)) #mtr , gloabl
    unit = Column(String(50))
    co2_factor = Column(Float)
    ch4_factor = Column(Float)
    n2o_factor = Column(Float)
    gwp_ch4 = Column(Float)
    gwp_n2o = Column(Float)
    ncv = Column(Float) #net calorific value
    ncv_unit = Column(String(50))
    density = Column(Float, nullable=True)
    density_unit = Column(String(50), nullable=True)
    input_unit = Column(String(20), nullable=True)   # 'm³', 'L', 'MWh', 'ton'
    is_active = Column(Boolean,default=True)
    version = Column(Integer,default=1)
    valid_year = Column(Integer)
    created_at = Column(DateTime, default=func.now())

#--------------------------------------------------
# scope 1 - statıonary combustion (1.1)
#--------------------------------------------------

class StationaryCombustion(Base):
    __tablename__ = "scope1_stationary"

    id = Column(Integer, primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    emisssion_source = Column(String(100)) #kombi, jeneratörler vs
    fuel_type = Column(String(50)) #doğalgaz, motorin vs
    activity_value = Column(Float) #FV
    activity_unit = Column(String(20))
    emission_factor_id = Column(Integer, ForeignKey("emission_factors.id"))
    co2e_total = Column(Float)
    created_at = Column(DateTime,default=func.now())

#--------------------------------------------------
#scope 1 - MOBİLE COMBUSTİON (1.2)
#--------------------------------------------------

class MobileCombustion(Base):
    __tablename__ = "scope1_mobile"

    id = Column(Integer,primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    emission_source = Column(String(100))   #binek araçlar on road
    fuel_type = Column(String(50))
    activity_value = Column(Float)
    activity_unit = Column(String(20))
    emission_factor_id = Column(Integer,ForeignKey("emission_factors.id"))
    co2e_total = Column(Float)
    created_at = Column(DateTime,default=func.now())

# ─────────────────────────────────────────
# SCOPE 1 — REFRIGERANTS (1.4)
# ─────────────────────────────────────────
class Refrigerant(Base):
    __tablename__ = "scope1_refrigerants"

    id = Column(Integer, primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    gas_type = Column(String(50))           # 'R407c', 'FM200', 'HFC236FA', 'R134A', 'R32'
    activity_value = Column(Float)          # amount leaked in ton
    activity_unit = Column(String(20))      # 'ton'
    gwp = Column(Float)                     # GWP value for that gas
    co2e_total = Column(Float)              # activity_value * gwp
    created_at = Column(DateTime, default=func.now())

# ─────────────────────────────────────────
# SCOPE 2 — ELECTRICITY (2.1)
# ─────────────────────────────────────────
class ElectricityConsumption(Base):
    __tablename__ = "scope2_electricity"

    id = Column(Integer, primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    consumption_mwh = Column(Float)         # electricity consumed in MWh
    emission_factor = Column(Float)         # ton CO2e/MWh (Turkey grid factor)
    ef_source = Column(String(100))         # 'REF#10'
    co2e_total = Column(Float)              # consumption_mwh * emission_factor
    created_at = Column(DateTime, default=func.now())

# ─────────────────────────────────────────
# SCOPE 3 — FREIGHT / SHIPMENT (3.1)
# ─────────────────────────────────────────
class FreightEmission(Base):
    __tablename__ = "scope3_freight"

    id = Column(Integer, primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    category_code = Column(String(10))      # '3.1.1', '3.1.2'
    emission_source = Column(String(100))   # 'Hammadde sevkiyatı'
    transport_type = Column(String(100))    # 'karayolu-%100 LADEN'
    vehicle_type = Column(String(100))      # 'All HGVs', 'Small car'
    activity_value = Column(Float)          # distance in km
    activity_unit = Column(String(20))      # 'km'
    emission_factor = Column(Float)         # kg CO2e/tonne.km
    co2e_total = Column(Float)
    created_at = Column(DateTime, default=func.now())

# ─────────────────────────────────────────
# SCOPE 3 — EMPLOYEE COMMUTING (3.3)
# ─────────────────────────────────────────
class EmployeeCommuting(Base):
    __tablename__ = "scope3_commuting"

    id = Column(Integer, primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    category_code = Column(String(10))      # '3.3.1', '3.3.2', '3.3.3'
    emission_source = Column(String(100))   # 'Şahsi Araç', 'Servisler'
    fuel_type = Column(String(50))          # 'Benzin', 'Dizel'
    activity_value = Column(Float)          # fuel consumed in ton
    activity_unit = Column(String(20))      # 'ton'
    emission_factor_id = Column(Integer, ForeignKey("emission_factors.id"))
    co2e_total = Column(Float)
    created_at = Column(DateTime, default=func.now())

# ─────────────────────────────────────────
# SCOPE 3 — PURCHASED GOODS (4.1)
# ─────────────────────────────────────────
class PurchasedGoods(Base):
    __tablename__ = "scope3_purchased_goods"

    id = Column(Integer, primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    category_code = Column(String(10))      # '4.1.1', '4.1.2' etc.
    item_name = Column(String(200))         # 'Printer-Kağıt', 'Toner'
    activity_value = Column(Float)          # amount or USD value
    activity_unit = Column(String(50))      # 'Ton Kağıt' or '2021 USD'
    emission_factor = Column(Float)         # kg CO2e/unit
    ef_source = Column(String(200))         # 'DEFRA', 'SupplyChainGHG...'
    co2e_total = Column(Float)
    created_at = Column(DateTime, default=func.now())

# ─────────────────────────────────────────
# SCOPE 3 — CAPITAL GOODS (4.2)
# ─────────────────────────────────────────
class CapitalGoods(Base):
    __tablename__ = "scope3_capital_goods"

    id = Column(Integer, primary_key=True)
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    category_code = Column(String(10))      # '4.2.1', '4.2.2' etc.
    asset_name = Column(String(200))        # 'BİLGİ İŞLEM CİHAZLARI'
    activity_value = Column(Float)          # USD value
    activity_unit = Column(String(50))      # '2021 USD'
    depreciation_years = Column(Integer)    # amortization period
    emission_factor = Column(Float)         # kg CO2e/2021 USD
    ef_source = Column(String(200))
    co2e_total = Column(Float)
    created_at = Column(DateTime, default=func.now())

# ─────────────────────────────────────────
# AUDIT LOG
# ─────────────────────────────────────────
class AuditLog(Base):
    __tablename__ = "audit_log"

    id = Column(Integer, primary_key=True)
    action = Column(String(100))            # 'data_entry', 'calculation', 'import'
    scope = Column(String(20))              # 'scope1', 'scope2', 'scope3'
    period_id = Column(Integer, ForeignKey("reporting_periods.id"))
    status = Column(String(20))             # 'success', 'failed'
    notes = Column(Text)
    created_at = Column(DateTime, default=func.now())
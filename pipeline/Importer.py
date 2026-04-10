"""
UNSPED Carbon Footprint Importer v3
- Excel'deki Yıl sütununu okur, her satırı kendi yılına kaydeder
- GWP: DB'de yoksa ekler, varsa günceller
- Aynı yıl tekrar import edilirse o yılın verisi silinir ve yeniden yazılır
- Legend/not satırlarını atlar (ilk sütun geçerli yıl değilse)

Kullanım:
    python pipeline/Importer.py data/UNSPED_Karbon_Veri_Girisi_v3.xlsx
"""

import sys
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from datetime import datetime
from collections import defaultdict

sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from db.connection import SessionLocal, engine, Base
from db.models import (
    Company, ReportingPeriod, AuditLog,
    StationaryCombustion, MobileCombustion, Refrigerant,
    ElectricityConsumption, FreightEmission, EmployeeCommuting,
    PurchasedGoods, CapitalGoods
)


# ══════════════════════════════════════════════════════════════════
# BİRİM DÖNÜŞÜM SİSTEMİ
# ══════════════════════════════════════════════════════════════════

# Tüm birimleri standart birime çeviren katsayılar
# Standart birimler: ton (kütle), m³ (hacim), MWh (enerji), USD (para)
UNIT_CONVERSIONS = {
    # ── Kütle → ton ───────────────────────────────────────────────
    "ton":   1.0,
    "tonne": 1.0,
    "t":     1.0,
    "kg":    0.001,
    "g":     0.000001,
    "mg":    0.000000001,
    "lb":    0.000453592,
    "lbs":   0.000453592,
    "pound": 0.000453592,
    "gg":    1000.0,        # Gigagram

    # ── Hacim → L ─────────────────────────────────────────────────
    "l":     1.0,
    "lt":    1.0,
    "litre": 1.0,
    "liter": 1.0,
    "m³":    1000.0,
    "m3":    1000.0,
    "ml":    0.001,

    # ── Enerji → MWh ──────────────────────────────────────────────
    "mwh":   1.0,
    "kwh":   0.001,
    "gwh":   1000.0,
    "twh":   1000000.0,
    "gj":    0.27778,
    "tj":    277.78,
    "mj":    0.000277778,
    "kj":    0.000000277778,

    # ── Mesafe → km ───────────────────────────────────────────────
    "km":    1.0,
    "m":     0.001,
    "mil":   1.60934,
    "mile":  1.60934,
    "miles": 1.60934,
    "nm":    1.852,         # Deniz mili

    # ── Para → USD (Anlık kur DB'den gelir) ───────────────────────
    # Para birimi dönüşümleri özel fonksiyon kullanır
}

# Her hesaplama kategorisi için standart birim
STANDARD_UNITS = {
    "combustion":    {"kütle": "ton",  "hacim_gaz": "m³", "hacim_sivi": "L"},
    "electricity":   {"enerji": "MWh"},
    "refrigerant":   {"kütle": "kg"},   # kg olarak girilir, hesapta kg kullanılır
    "transport":     {"mesafe": "km"},
    "spend":         {"para": "USD"},
}

# Birim grupları — hangi grup olduğunu anlamak için
UNIT_GROUPS = {
    "kütle":     {"ton","tonne","t","kg","g","mg","lb","lbs","pound","gg"},
    "hacim_gaz": {"m³","m3"},
    "hacim_sivi":{"l","lt","litre","liter","ml"},
    "enerji":    {"mwh","kwh","gwh","twh","gj","tj","mj","kj"},
    "mesafe":    {"km","m","mil","mile","miles","nm"},
    "para":      {"tl","try","usd","eur","gbp","₺","$","€","£"},
}

# Dönüşüm uyarıları için log
conversion_warnings = []

def get_unit_group(unit: str) -> str:
    """Birimin hangi gruba ait olduğunu döndürür."""
    ul = unit.lower().strip()
    for group, units in UNIT_GROUPS.items():
        if ul in units:
            return group
    return "unknown"

def convert_to_standard(value: float, unit: str, target: str,
                         source_label: str = "", usd_rate: float = None) -> tuple:
    """
    Değeri standart birime çevirir.
    Returns: (converted_value, conversion_note)
    conversion_note: None ise dönüşüm yapılmadı, string ise yapıldı ve uyarı var.
    """
    if value is None or value == 0:
        return 0.0, None

    ul     = unit.lower().strip()
    tl     = target.lower().strip()
    note   = None

    # Zaten standart birimde
    if ul == tl:
        return float(value), None

    # Para birimi dönüşümü
    if get_unit_group(ul) == "para":
        if usd_rate and usd_rate > 0 and ul in ("tl","try","₺"):
            converted = float(value) / usd_rate
            note = f"{source_label}: {value:,.2f} TL ÷ {usd_rate} = {converted:,.4f} USD"
            return converted, note
        elif ul in ("eur","€"):
            # EUR → USD yaklaşık (güncel kur için DB kullanılabilir)
            converted = float(value) * 1.08
            note = f"{source_label}: {value:,.2f} EUR × 1.08 ≈ {converted:,.4f} USD (yaklaşık)"
            return converted, note
        elif ul in ("gbp","£"):
            converted = float(value) * 1.27
            note = f"{source_label}: {value:,.2f} GBP × 1.27 ≈ {converted:,.4f} USD (yaklaşık)"
            return converted, note
        else:
            return float(value), None

    # Standart dönüşüm tablosundan
    factor_from = UNIT_CONVERSIONS.get(ul)
    factor_to   = UNIT_CONVERSIONS.get(tl)

    if factor_from is None:
        return float(value), None  # Bilinmeyen birim — olduğu gibi kullan

    if factor_to is None:
        return float(value), None

    # Aynı grup mu kontrol et
    group_from = get_unit_group(ul)
    group_to   = get_unit_group(tl)
    if group_from != group_to and group_from != "unknown" and group_to != "unknown":
        return float(value), None  # Farklı gruplar — dönüşüm yapma

    converted = float(value) * (factor_from / factor_to)
    note = f"{source_label}: {value} {unit} → {converted:.6f} {target}"
    return converted, note

def log_conversion(note: str, sheet: str):
    """Dönüşüm uyarısını loglar."""
    if note:
        msg = f"[{sheet}] {note}"
        conversion_warnings.append(msg)
        print(f"    🔄 {msg}")

def print_conversion_summary():
    """İmport sonunda tüm dönüşümleri özetler."""
    if not conversion_warnings:
        return
    print(f"\n{'─'*55}")
    print(f"  🔄 BİRİM DÖNÜŞÜM ÖZETI ({len(conversion_warnings)} dönüşüm yapıldı)")
    print(f"{'─'*55}")
    for w in conversion_warnings:
        print(f"  • {w}")
    print(f"{'─'*55}")

# ── Style helpers ──────────────────────────────────────────────────
LOCKED_BG = "F2F2F2"
TOTAL_BG  = "FFF2CC"
HEADER_BG = "1F4E79"
IMPACT_BG = "FCE4D6"
INPUT_BG  = "D9E1F2"

thin = Side(style="thin", color="BFBFBF")
B    = Border(left=thin, right=thin, top=thin, bottom=thin)

def lck(cell, value=None):
    if value is not None: cell.value = value
    cell.font      = Font(name="Arial", size=10, color="595959")
    cell.fill      = PatternFill("solid", start_color=LOCKED_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = B

def tot(cell, value=None):
    if value is not None: cell.value = value
    cell.font      = Font(name="Arial", size=10, bold=True, color="7B3F00")
    cell.fill      = PatternFill("solid", start_color=TOTAL_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = B

def imp_c(cell, value=None):
    if value is not None: cell.value = value
    cell.font      = Font(name="Arial", size=10, color="843C0C")
    cell.fill      = PatternFill("solid", start_color=IMPACT_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = B

def inp(cell, value=None):
    if value is not None: cell.value = value
    cell.font      = Font(name="Arial", size=10, color="00008B")
    cell.fill      = PatternFill("solid", start_color=INPUT_BG)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = B

def hdr(cell, value=None, bg=HEADER_BG):
    if value is not None: cell.value = value
    cell.font      = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    cell.fill      = PatternFill("solid", start_color=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = B


# ══════════════════════════════════════════════════════════════════
# EXCEL OKUMA — AKILLI FİLTRE
# ══════════════════════════════════════════════════════════════════

def is_valid_year(val):
    """İlk sütunun geçerli bir yıl olup olmadığını kontrol eder."""
    try:
        y = int(float(str(val)))
        return 2000 <= y <= 2100
    except (ValueError, TypeError):
        return False

def read_sheet(wb, sheet_name, header_row=3):
    """
    Sayfayı DataFrame olarak okur.
    - Sadece ilk sütunu geçerli yıl olan satırları alır
    - Legend, not, boş satırları otomatik atlar
    """
    ws = wb[sheet_name]
    headers = [c.value for c in ws[header_row] if c.value]
    headers = [str(h).split('\n')[0].strip() for h in headers]

    data = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        vals = list(row[:len(headers)])
        if not any(v is not None for v in vals):
            continue
        if not is_valid_year(vals[0]):
            continue  # legend, not, başlık satırı → atla
        data.append(vals)

    if not data:
        return pd.DataFrame(columns=headers)

    df = pd.DataFrame(data, columns=headers)
    df['_year'] = df.iloc[:, 0].apply(lambda v: int(float(str(v))))
    return df


# ══════════════════════════════════════════════════════════════════
# DB HELPERS
# ══════════════════════════════════════════════════════════════════

def get_or_create_period(session, year):
    """Yıl için raporlama dönemi getirir veya oluşturur."""
    period = session.query(ReportingPeriod).join(Company).filter(
        Company.name == "UNSPED",
        ReportingPeriod.year == year,
        ReportingPeriod.period == "annual"
    ).first()

    if not period:
        company = session.query(Company).filter_by(name="UNSPED").first()
        if not company:
            raise ValueError(
                "❌ UNSPED şirketi DB'de bulunamadı. Önce 'python main.py' çalıştırın."
            )
        period = ReportingPeriod(
            company_id=company.id,
            year=year,
            period="annual"
        )
        session.add(period)
        session.flush()
        print(f"    ℹ️  {year} dönemi otomatik oluşturuldu.")
    return period

def clear_year_data(session, period_id, year):
    """Belirtilen yılın tüm hesaplanan verisini siler."""
    deleted = 0
    for Model in [StationaryCombustion, MobileCombustion, Refrigerant,
                  ElectricityConsumption, FreightEmission, EmployeeCommuting,
                  PurchasedGoods, CapitalGoods]:
        n = session.query(Model).filter_by(period_id=period_id).delete()
        deleted += n
    print(f"    🗑️  {year} yılı temizlendi ({deleted} kayıt silindi, period_id={period_id})")

def get_periods_from_df(session, df):
    """
    DataFrame'deki tüm yılları bulur, her biri için period_id döndürür.
    Returns: {year: period_id}
    """
    years = df['_year'].unique()
    year_to_period = {}
    for year in years:
        period = get_or_create_period(session, int(year))
        year_to_period[int(year)] = period.id
    return year_to_period


# ══════════════════════════════════════════════════════════════════
# KALKULATÖRLER
# ══════════════════════════════════════════════════════════════════

def calc_combustion(amount, unit, ncv, ef_co2, ef_ch4, ef_n2o,
                    gwp_ch4, gwp_n2o, density, density_unit,
                    sheet="", source=""):
    """
    Yanma hesabı — birim dönüşümü otomatik.
    density_unit: ton/m³ ise amount m³, ton/L ise amount L cinsinden.
    Kullanıcı farklı birim girerse otomatik çevirir ve uyarır.
    """
    try:
        amount      = float(amount      or 0)
        ncv         = float(ncv         or 0)
        ef_co2      = float(ef_co2      or 0)
        ef_ch4      = float(ef_ch4      or 0)
        ef_n2o      = float(ef_n2o      or 0)
        gwp_ch4     = float(gwp_ch4     or 0)
        gwp_n2o     = float(gwp_n2o     or 0)
        density     = float(density     or 0)

        if amount == 0 or density == 0 or ncv == 0:
            return 0.0

        # Standart birim density_unit'ten çıkar: "ton/m³" → "m³", "ton/L" → "L"
        if density_unit and "/" in str(density_unit):
            std_unit = str(density_unit).split("/")[-1].strip()
        else:
            std_unit = str(unit or "").strip()

        # Birim dönüşümü
        unit_str = str(unit or std_unit).strip()
        if unit_str.lower() != std_unit.lower():
            converted, note = convert_to_standard(
                amount, unit_str, std_unit,
                source_label=f"{source}"
            )
            if note:
                log_conversion(note, sheet)
                amount = converted

        activity_tj = (amount * density / 1000) * ncv
        return round(
            activity_tj * ef_co2 +
            activity_tj * ef_ch4 * gwp_ch4 +
            activity_tj * ef_n2o * gwp_n2o, 6
        )
    except Exception as e:
        print(f"    ⚠ Yanma hesap hatası: {e}")
        return 0.0

def calc_electricity(kwh, unit, ef, sheet="", source=""):
    """Elektrik hesabı — kWh, MWh, GWh otomatik çevrilir."""
    try:
        val     = float(kwh or 0)
        unit_str = str(unit or "kWh").strip()
        # Standart: MWh
        if unit_str.lower() != "mwh":
            converted, note = convert_to_standard(val, unit_str, "MWh",
                                                   source_label=source)
            if note:
                log_conversion(note, sheet)
                val = converted
        else:
            val = val / 1000  # kWh → MWh (varsayılan)
        return round(val * float(ef or 0), 6)
    except: return 0.0

def calc_refrigerant(amount, unit, gwp, sheet="", source=""):
    """Soğutucu gaz — kg, ton, g otomatik çevrilir. Sonuç ton CO2e."""
    try:
        val      = float(amount or 0)
        unit_str = str(unit or "kg").strip()
        # Standart: kg
        if unit_str.lower() != "kg":
            converted, note = convert_to_standard(val, unit_str, "kg",
                                                   source_label=source)
            if note:
                log_conversion(note, sheet)
                val = converted
        return round(val * float(gwp or 0) / 1000, 6)
    except: return 0.0

def calc_spend(amount, unit, ef, sheet="", source="", usd_rate=None):
    """Harcama — TL, EUR, GBP otomatik USD'ye çevrilir."""
    try:
        val      = float(amount or 0)
        unit_str = str(unit or "USD").strip()
        if unit_str.lower() not in ("usd","$","2021 usd","2022 usd","2023 usd","2024 usd","2025 usd"):
            converted, note = convert_to_standard(val, unit_str, "USD",
                                                   source_label=source,
                                                   usd_rate=usd_rate)
            if note:
                log_conversion(note, sheet)
                val = converted
        return round(val * float(ef or 0) / 1000, 6)
    except: return 0.0

def calc_transport(distance, unit, ef, sheet="", source=""):
    """Taşımacılık — km, mil otomatik çevrilir."""
    try:
        val      = float(distance or 0)
        unit_str = str(unit or "km").strip()
        if unit_str.lower() != "km":
            converted, note = convert_to_standard(val, unit_str, "km",
                                                   source_label=source)
            if note:
                log_conversion(note, sheet)
                val = converted
        return round(val * float(ef or 0) / 1000, 6)
    except: return 0.0

def calc_capital(amount, unit, ef, depreciation, sheet="", source="", usd_rate=None):
    """Sermaye varlığı — TL/EUR otomatik USD'ye çevrilir."""
    try:
        val      = float(amount or 0)
        unit_str = str(unit or "USD").strip()
        if unit_str.lower() not in ("usd","$","2021 usd","2022 usd","2023 usd","2024 usd","2025 usd"):
            converted, note = convert_to_standard(val, unit_str, "USD",
                                                   source_label=source,
                                                   usd_rate=usd_rate)
            if note:
                log_conversion(note, sheet)
                val = converted
        dep = float(depreciation or 1) or 1
        return round(val * float(ef or 0) / (dep * 1000), 6)
    except: return 0.0

def calc_td_loss(kwh, loc_ef, mkt_ef):
    try: return round((float(kwh or 0) / 1000) * (float(mkt_ef or 0) - float(loc_ef or 0)), 6)
    except: return 0.0


# ══════════════════════════════════════════════════════════════════
# EXCEL'E GERİ YAZ
# ══════════════════════════════════════════════════════════════════

def write_co2e_column(ws, col_index, results, start_row=4, notes=None):
    """
    Hesaplanan CO2e değerlerini Excel'deki gri sütuna yazar.
    notes: dönüşüm yapılan satırlar için hücre yorumu eklenir.
    """
    from openpyxl.comments import Comment
    for i, val in enumerate(results):
        cell = ws.cell(row=start_row + i, column=col_index)
        lck(cell, round(val, 4) if val else 0)
        if notes and i < len(notes) and notes[i]:
            try:
                cell.comment = Comment(
                    f"🔄 Birim dönüşümü yapıldı:\n{notes[i]}",
                    "UNSPED Importer"
                )
            except: pass


# ══════════════════════════════════════════════════════════════════
# HER SAYFA İÇİN İŞLEM FONKSİYONLARI
# Her satır kendi yılına kaydedilir
# ══════════════════════════════════════════════════════════════════

def process_11(wb, session, year_to_period):
    print("  [1.1] Sabit Yakma...")
    df = read_sheet(wb, '1.1_Stationary')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year     = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        birim  = str(r.get('Birim') or 'L').strip()
        kaynak = f"{r.get('Emisyon Kaynağı')} ({r.get('Yakıt Türü')})"
        density_unit = 'ton/m³' if birim.lower() in ('m³','m3') else 'ton/L'
        _prev = len(conversion_warnings)
        co2e = calc_combustion(
            r.get('Miktar'), birim,
            r.get('NCV'), r.get('EF CO2'),
            r.get('EF CH4'), r.get('EF N2O'),
            r.get('GWP CH4'), r.get('GWP N2O'),
            r.get('Yoğunluk'), density_unit,
            sheet='1.1_Stationary', source=kaynak
        )
        conv_notes.append(conversion_warnings[-1] if len(conversion_warnings) > _prev else None)
        results.append(co2e)
        totals_by_year[year] += co2e
        session.add(StationaryCombustion(
            period_id=period_id,
            emisssion_source=str(r.get('Emisyon Kaynağı') or ''),
            fuel_type=str(r.get('Yakıt Türü') or ''),
            activity_value=float(r.get('Miktar') or 0),
            activity_unit=birim,
            co2e_total=co2e
        ))

    write_co2e_column(wb['1.1_Stationary'], 13, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_12(wb, session, year_to_period):
    print("  [1.2] Hareketli Yakma...")
    df = read_sheet(wb, '1.2_Mobile')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        birim  = str(r.get('Birim') or 'L').strip()
        kaynak = f"{r.get('Emisyon Kaynağı')} ({r.get('Yakıt Türü')})"
        density_unit = 'ton/m³' if birim.lower() in ('m³','m3') else 'ton/L'
        _prev = len(conversion_warnings)
        co2e = calc_combustion(
            r.get('Miktar'), birim,
            r.get('NCV'), r.get('EF CO2'),
            r.get('EF CH4'), r.get('EF N2O'),
            r.get('GWP CH4'), r.get('GWP N2O'),
            r.get('Yoğunluk'), density_unit,
            sheet='1.2_Mobile', source=kaynak
        )
        conv_notes.append(conversion_warnings[-1] if len(conversion_warnings) > _prev else None)
        results.append(co2e)
        totals_by_year[year] += co2e
        session.add(MobileCombustion(
            period_id=period_id,
            emission_source=str(r.get('Emisyon Kaynağı') or ''),
            fuel_type=str(r.get('Yakıt Türü') or ''),
            activity_value=float(r.get('Miktar') or 0),
            activity_unit=str(r.get('Birim') or ''),
            co2e_total=co2e
        ))

    write_co2e_column(wb['1.2_Mobile'], 13, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_14(wb, session, year_to_period):
    """
    Soğutucu gazlar — GWP akıllı yönetim:
    Excel'deki GWP değeri alınır.
    DB'de o gaz yoksa yeni kayıt, varsa GWP güncellenir.
    """
    print("  [1.4] Soğutucu Gazlar (GWP akıllı)...")
    df = read_sheet(wb, '1.4_Refrigerants')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        gas_type  = str(r.get('Gaz Türü') or '')
        amount_kg = float(r.get('Kaçak Miktar (kg)') or 0)
        gwp_excel = float(r.get('GWP') or 0)

        # ── GWP akıllı yönetim ──────────────────────────────────
        if gwp_excel == 0:
            print(f"    ⚠ {gas_type} için GWP girilmemiş, hesaplama yapılamıyor.")
            co2e = 0.0
            conv_notes.append(None)
        else:
            birim  = str(r.get('Birim') or 'kg').strip()
            kaynak = gas_type
            _prev  = len(conversion_warnings)
            co2e = calc_refrigerant(
                amount_kg, birim, gwp_excel,
                sheet='1.4_Refrigerants', source=kaynak
            )
            conv_notes.append(conversion_warnings[-1] if len(conversion_warnings) > _prev else None)

        results.append(co2e)
        totals_by_year[year] += co2e

        session.add(Refrigerant(
            period_id=period_id,
            gas_type=gas_type,
            activity_value=amount_kg,
            activity_unit=str(r.get('Birim') or 'kg'),
            gwp=gwp_excel,
            co2e_total=co2e
        ))

    write_co2e_column(wb['1.4_Refrigerants'], 6, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_21(wb, session, year_to_period):
    print("  [2.1] Elektrik...")
    df = read_sheet(wb, '2.1_Electricity')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        kwh       = float(r.get('Tüketim (kWh)') or 0)
        ef        = float(r.get('EF') or 0)
        birim     = str(r.get('Birim') or 'kWh').strip() if 'Birim' in r.index else 'kWh'
        _prev = len(conversion_warnings)
        co2e      = calc_electricity(kwh, birim, ef,
                        sheet='2.1_Electricity', source='Elektrik')
        conv_notes.append(conversion_warnings[-1] if len(conversion_warnings) > _prev else None)

        results.append(co2e)
        totals_by_year[year] += co2e
        session.add(ElectricityConsumption(
            period_id=period_id,
            consumption_mwh=kwh / 1000,
            emission_factor=ef,
            ef_source=str(r.get('EF Kaynak') or ''),
            co2e_total=co2e
        ))

    write_co2e_column(wb['2.1_Electricity'], 5, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_freight(wb, session, year_to_period, sheet, label):
    print(f"  {label}...")
    df = read_sheet(wb, sheet)
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        birim  = str(r.get('Birim') or 'km').strip() if 'Birim' in r.index else 'km'
        kaynak = str(r.get('Emisyon Kaynağı') or '')
        co2e = calc_transport(r.get('Mesafe (km)'), birim, r.get('EF'),
                              sheet=sheet, source=kaynak)
        results.append(co2e)
        conv_notes.append(None)
        totals_by_year[year] += co2e
        session.add(FreightEmission(
            period_id=period_id,
            category_code=str(r.get('Kategori') or ''),
            emission_source=str(r.get('Emisyon Kaynağı') or ''),
            transport_type=str(r.get('Taşıma Tipi') or ''),
            vehicle_type=str(r.get('Araç Tipi') or ''),
            activity_value=float(r.get('Mesafe (km)') or 0),
            activity_unit='km',
            emission_factor=float(r.get('EF') or 0),
            co2e_total=co2e
        ))

    write_co2e_column(wb[sheet], 8, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_33(wb, session, year_to_period):
    print("  [3.3] Personel Ulaşım...")
    df = read_sheet(wb, '3.3_Commuting')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        birim  = str(r.get('Birim') or 'L').strip()
        kaynak = f"{r.get('Emisyon Kaynağı')} ({r.get('Yakıt Türü')})"
        density_unit = 'ton/m³' if birim.lower() in ('m³','m3') else 'ton/L'
        co2e = calc_combustion(
            r.get('Miktar'), birim,
            r.get('NCV'), r.get('EF CO2'),
            r.get('EF CH4'), r.get('EF N2O'),
            r.get('GWP CH4'), r.get('GWP N2O'),
            r.get('Yoğunluk'), density_unit,
            sheet='3.3_Commuting', source=kaynak
        )
        results.append(co2e)
        conv_notes.append(None)
        totals_by_year[year] += co2e
        session.add(EmployeeCommuting(
            period_id=period_id,
            category_code=str(r.get('Kategori') or ''),
            emission_source=str(r.get('Emisyon Kaynağı') or ''),
            fuel_type=str(r.get('Yakıt Türü') or ''),
            activity_value=float(r.get('Miktar') or 0),
            activity_unit=str(r.get('Birim') or ''),
            co2e_total=co2e
        ))

    write_co2e_column(wb['3.3_Commuting'], 14, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_35(wb, session, year_to_period):
    print("  [3.5] Uçak & Konaklama...")
    df = read_sheet(wb, '3.5_FlightsHotels')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        try:
            tl   = float(r.get('Miktar (TL)') or 0)
            kur  = float(r.get('Dolar Kuru')  or 1) or 1
            ef   = float(r.get('EF')          or 0)
            co2e = round((tl / kur) * ef / 1000, 6)
        except:
            co2e = 0.0
            tl, ef = 0, 0

        results.append(co2e)
        conv_notes.append(None)
        totals_by_year[year] += co2e
        session.add(PurchasedGoods(
            period_id=period_id,
            category_code=str(r.get('Kategori') or ''),
            item_name=str(r.get('Kalem') or ''),
            activity_value=tl,
            activity_unit='TL',
            emission_factor=ef,
            ef_source='SupplyChainGHG / DEFRA',
            co2e_total=co2e
        ))

    write_co2e_column(wb['3.5_FlightsHotels'], 8, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_purchased(wb, session, year_to_period, sheet, amount_col, co2e_col, label):
    print(f"  {label}...")
    df = read_sheet(wb, sheet)
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        birim  = str(r.get('Birim') or 'USD').strip() if 'Birim' in r.index else 'USD'
        kaynak = str(r.get(list(df.columns)[2]) or '')
        co2e = calc_spend(r.get(amount_col), birim, r.get('EF'),
                          sheet=sheet, source=kaynak)
        results.append(co2e)
        conv_notes.append(None)
        totals_by_year[year] += co2e
        item_col = [c for c in df.columns if c not in ('_year', 'Kategori', amount_col, 'Birim', 'EF', 'EF Kaynak', 'Not')][0]
        session.add(PurchasedGoods(
            period_id=period_id,
            category_code=str(r.get('Kategori') or ''),
            item_name=str(r.get(item_col) or ''),
            activity_value=float(r.get(amount_col) or 0),
            activity_unit=str(r.get('Birim') or ''),
            emission_factor=float(r.get('EF') or 0),
            ef_source=str(r.get('EF Kaynak') or ''),
            co2e_total=co2e
        ))

    write_co2e_column(wb[sheet], co2e_col, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_42(wb, session, year_to_period):
    print("  [4.2] Sermaye Varlıkları...")
    df = read_sheet(wb, '4.2_Capital')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        birim  = str(r.get('Birim') or '2021 USD').strip()
        kaynak = str(r.get('Varlık Adı') or '')
        co2e = calc_capital(
            r.get('Tutar (2021 USD)'), birim,
            r.get('EF'), r.get('Amortisman (Yıl)'),
            sheet='4.2_Capital', source=kaynak
        )
        results.append(co2e)
        conv_notes.append(None)
        totals_by_year[year] += co2e
        session.add(CapitalGoods(
            period_id=period_id,
            category_code=str(r.get('Kategori') or ''),
            asset_name=str(r.get('Varlık Adı') or ''),
            activity_value=float(r.get('Tutar (2021 USD)') or 0),
            activity_unit='2021 USD',
            depreciation_years=int(float(r.get('Amortisman (Yıl)') or 1)),
            emission_factor=float(r.get('EF') or 0),
            ef_source=str(r.get('EF Kaynak') or ''),
            co2e_total=co2e
        ))

    write_co2e_column(wb['4.2_Capital'], 9, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year

def process_61(wb, session, year_to_period):
    print("  [6.1] T&D Kayıpları...")
    df = read_sheet(wb, '6.1_TDLosses')
    results, totals_by_year, conv_notes = [], defaultdict(float), []

    for _, r in df.iterrows():
        year      = r['_year']
        if year not in year_to_period:
            p = get_or_create_period(session, int(year))
            year_to_period[year] = p.id
        period_id = year_to_period[year]
        kwh    = float(r.get('Tüketim (kWh)') or 0)
        loc_ef = float(r.get('Konum EF') or 0)
        mkt_ef = float(r.get('Pazar EF') or 0)
        co2e   = calc_td_loss(kwh, loc_ef, mkt_ef)

        results.append(co2e)
        conv_notes.append(None)
        totals_by_year[year] += co2e
        session.add(ElectricityConsumption(
            period_id=period_id,
            consumption_mwh=kwh / 1000,
            emission_factor=round(mkt_ef - loc_ef, 6),
            ef_source='TEİAŞ T&D kayıp faktörü',
            co2e_total=co2e
        ))

    write_co2e_column(wb['6.1_TDLosses'], 6, results, notes=conv_notes)
    for y, t in totals_by_year.items():
        print(f"    → {y}: {t:.4f} ton CO2e | DB'ye kaydedildi ✅")
    return totals_by_year


# ══════════════════════════════════════════════════════════════════
# ÖZET SAYFALAR
# ══════════════════════════════════════════════════════════════════

def merge_totals(all_results_by_year):
    """
    Her kategorinin yıl bazlı toplamlarını tek dict'te birleştirir.
    all_results_by_year = {'1.1': {2023: x, 2024: y}, '1.2': {...}, ...}
    Returns: {year: {'1.1': x, '1.2': y, ..., 'GENEL TOPLAM': z}}
    """
    all_years = set()
    for d in all_results_by_year.values():
        all_years.update(d.keys())

    scope_totals = {}
    for year in sorted(all_years):
        t11  = all_results_by_year.get('1.1',  {}).get(year, 0)
        t12  = all_results_by_year.get('1.2',  {}).get(year, 0)
        t14  = all_results_by_year.get('1.4',  {}).get(year, 0)
        t21  = all_results_by_year.get('2.1',  {}).get(year, 0)
        t31  = all_results_by_year.get('3.1',  {}).get(year, 0)
        t32  = all_results_by_year.get('3.2',  {}).get(year, 0)
        t33  = all_results_by_year.get('3.3',  {}).get(year, 0)
        t34  = all_results_by_year.get('3.4',  {}).get(year, 0)
        t35  = all_results_by_year.get('3.5',  {}).get(year, 0)
        t41  = all_results_by_year.get('4.1',  {}).get(year, 0)
        t42  = all_results_by_year.get('4.2',  {}).get(year, 0)
        t43  = all_results_by_year.get('4.3',  {}).get(year, 0)
        t44  = all_results_by_year.get('4.4',  {}).get(year, 0)
        t45  = all_results_by_year.get('4.5',  {}).get(year, 0)
        t61  = all_results_by_year.get('6.1',  {}).get(year, 0)
        s1   = t11 + t12 + t14
        s2   = t21
        s3   = t31+t32+t33+t34+t35+t41+t42+t43+t44+t45+t61
        grand = s1 + s2 + s3

        scope_totals[year] = {
            '1.1': t11, '1.2': t12, '1.4': t14, 'Kapsam 1 Top.': s1,
            '2.1': t21, 'Kapsam 2 Top.': s2,
            '3.1': t31, '3.2': t32, '3.3': t33, '3.4': t34, '3.5': t35,
            '4.1': t41, '4.2': t42, '4.3': t43, '4.4': t44, '4.5': t45,
            '6.1': t61, 'Kapsam 3 Top.': s3, 'GENEL TOPLAM': grand,
        }
    return scope_totals

def write_totals(wb, scope_totals):
    ws = wb['TOTAL_EMISSIONS']
    category_rows = {
        'KAPSAM 1':4,'1.1':5,'1.2':6,'1.4':7,'Kapsam 1 Top.':8,
        'KAPSAM 2':10,'2.1':11,'Kapsam 2 Top.':12,
        'KAPSAM 3':14,'3.1':15,'3.2':16,'3.3':17,'3.4':18,
        '3.5':19,'4.1':20,'4.2':21,'4.3':22,'4.4':23,
        '4.5':24,'6.1':25,'Kapsam 3 Top.':26,'GENEL TOPLAM':28,
    }
    # Yılları sırala — ilk yıl C sütunu, ikinci yıl D sütunu
    sorted_years = sorted(scope_totals.keys())
    year_cols    = {y: 3 + i for i, y in enumerate(sorted_years)}

    for year, totals in scope_totals.items():
        col = year_cols[year]
        # Yıl başlığını sütun başına yaz
        ws.cell(row=3, column=col).value = f"{year} (ton CO2e)"

        for cat, val in totals.items():
            row = category_rows.get(cat)
            if not row: continue
            cell = ws.cell(row=row, column=col)
            if cat in ('KAPSAM 1','KAPSAM 2','KAPSAM 3'):
                hdr(cell, round(val,2) if val else '-', bg="2E75B6")
            elif 'Top.' in cat or cat == 'GENEL TOPLAM':
                tot(cell, round(val,2) if val else 0)
            else:
                lck(cell, round(val,4) if val else 0)

    print("    ✅ TOTAL_EMISSIONS güncellendi.")

def write_impact(wb, scope_totals):
    """En son yılın verisiyle impact analizi yazar."""
    latest_year = max(scope_totals.keys())
    totals      = scope_totals[latest_year]
    ws          = wb['IMPACT_ANALYSIS']

    all_sources = [
        {'scope':'Kapsam 1','category':'1.1','label':'Sabit Yakma',             'co2e': totals.get('1.1',0)},
        {'scope':'Kapsam 1','category':'1.2','label':'Hareketli Yakma',         'co2e': totals.get('1.2',0)},
        {'scope':'Kapsam 1','category':'1.4','label':'Soğutucu Gazlar',         'co2e': totals.get('1.4',0)},
        {'scope':'Kapsam 2','category':'2.1','label':'Elektrik Tüketimi',       'co2e': totals.get('2.1',0)},
        {'scope':'Kapsam 3','category':'3.1','label':'Hammadde Sevkiyatı',      'co2e': totals.get('3.1',0)},
        {'scope':'Kapsam 3','category':'3.2','label':'Ürün Sevkiyatı',          'co2e': totals.get('3.2',0)},
        {'scope':'Kapsam 3','category':'3.3','label':'Personel Ulaşım',         'co2e': totals.get('3.3',0)},
        {'scope':'Kapsam 3','category':'3.4','label':'İş Seyahati',             'co2e': totals.get('3.4',0)},
        {'scope':'Kapsam 3','category':'3.5','label':'Uçak & Konaklama',        'co2e': totals.get('3.5',0)},
        {'scope':'Kapsam 3','category':'4.1','label':'Satın Alınan Malzemeler', 'co2e': totals.get('4.1',0)},
        {'scope':'Kapsam 3','category':'4.2','label':'Sermaye Varlıkları',      'co2e': totals.get('4.2',0)},
        {'scope':'Kapsam 3','category':'4.3','label':'Atık Bertarafı',          'co2e': totals.get('4.3',0)},
        {'scope':'Kapsam 3','category':'4.4','label':'Kiralanan Ekipmanlar',    'co2e': totals.get('4.4',0)},
        {'scope':'Kapsam 3','category':'4.5','label':'Hizmet Alımları',         'co2e': totals.get('4.5',0)},
        {'scope':'Kapsam 3','category':'6.1','label':'T&D Kayıpları',           'co2e': totals.get('6.1',0)},
    ]

    sorted_sources = sorted(all_sources, key=lambda x: x['co2e'], reverse=True)
    grand_total    = sum(s['co2e'] for s in sorted_sources)

    for row in ws.iter_rows(min_row=4, max_row=35):
        for cell in row:
            try: cell.value = None; cell.fill = PatternFill("solid", start_color="FFFFFF")
            except: pass

    cumulative = 0
    active_count = 0
    for i, source in enumerate(sorted_sources):
        if source['co2e'] == 0: continue
        row        = 4 + active_count
        pct        = (source['co2e'] / grand_total * 100) if grand_total else 0
        cumulative += pct
        above_95   = cumulative <= 95.0
        for ci, val in enumerate([
            active_count + 1, source['scope'], source['category'],
            source['label'], round(source['co2e'],4),
            round(pct,2), round(cumulative,2)
        ], 1):
            cell = ws.cell(row=row, column=ci, value=val)
            imp_c(cell) if above_95 else lck(cell)
        active_count += 1

    # %95 eşik çizgisi
    mr = 4 + active_count
    ws.merge_cells(f"A{mr}:G{mr}")
    m = ws.cell(row=mr, column=1,
        value=f"▲ %95 EŞIĞI ({latest_year}) — Aşağıdaki kaynaklar toplam emisyonun %5'inden az")
    m.font      = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    m.fill      = PatternFill("solid", start_color="843C0C")
    m.alignment = Alignment(horizontal="center", vertical="center")
    print(f"    ✅ IMPACT_ANALYSIS güncellendi ({latest_year} yılı bazında).")

def write_per_capita(wb, scope_totals):
    """Her yıl için kişi başına emisyon yazar."""
    ws_pc  = wb['PER_CAPITA']
    ws_per = wb['PERSONNEL']

    # Tüm satırları temizle
    for row in ws_pc.iter_rows(min_row=4, max_row=40):
        for cell in row:
            try: cell.value = None; cell.fill = PatternFill("solid", start_color="FFFFFF")
            except: pass

    current_row = 4
    for year in sorted(scope_totals.keys()):
        # O yılın personelini oku
        personnel = []
        for row in ws_per.iter_rows(min_row=4, values_only=True):
            if row[0] and row[1] and row[2] and is_valid_year(row[0]):
                if int(float(str(row[0]))) == year:
                    personnel.append({
                        'location': row[1],
                        'headcount': int(float(str(row[2])))
                    })

        if not personnel:
            print(f"    ⚠ {year} için PERSONNEL verisi bulunamadı, atlanıyor.")
            continue

        totals     = scope_totals[year]
        s1         = totals.get('Kapsam 1 Top.', 0)
        s2         = totals.get('Kapsam 2 Top.', 0)
        s3         = totals.get('Kapsam 3 Top.', 0)
        grand      = s1 + s2 + s3
        total_head = sum(p['headcount'] for p in personnel)

        for p in personnel:
            ratio     = p['headcount'] / total_head if total_head else 0
            loc_s1    = round(s1 * ratio, 4)
            loc_s2    = round(s2 * ratio, 4)
            loc_s3    = round(s3 * ratio, 4)
            loc_total = round(loc_s1 + loc_s2 + loc_s3, 4)
            per_cap   = round(loc_total / p['headcount'], 4) if p['headcount'] else 0
            ws_pc.row_dimensions[current_row].height = 20
            for ci, val in enumerate([year, p['location'], p['headcount'],
                                       loc_total, loc_s1, loc_s2, loc_s3, per_cap], 1):
                try:
                    inp(ws_pc.cell(row=current_row, column=ci), val)
                except AttributeError:
                    pass  # skip merged cells
            current_row += 1

        # Yıl toplamı
        grand_cap = round(grand / total_head, 4) if total_head else 0
        ws_pc.row_dimensions[current_row].height = 20
        for ci, val in enumerate([year, f'ŞİRKET TOPLAMI ({year})', total_head,
                                   round(grand,4), round(s1,4),
                                   round(s2,4), round(s3,4), grand_cap], 1):
            try:
                tot(ws_pc.cell(row=tot_row, column=ci), val)
            except AttributeError:
                pass
        current_row += 2  # boş satır bırak

        print(f"    ✅ {year}: {len(personnel)} lokasyon | {grand:.2f} ton CO2e | {grand_cap} ton/kişi")


# ══════════════════════════════════════════════════════════════════
# ANA FONKSİYON
# ══════════════════════════════════════════════════════════════════

def run_import(excel_path):
    print(f"\n{'═'*58}")
    print(f"  UNSPED KARBON İMPORTER v3")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Dosya: {excel_path}")
    print(f"{'═'*58}\n")

    wb = load_workbook(excel_path)

    with SessionLocal() as session:

        # ── 1. Excel'deki tüm yılları bul ────────────────────────
        # Her sayfayı tarayarak hangi yılların olduğunu öğren
        all_years = set()
        for sheet_name in ['1.1_Stationary','1.2_Mobile','2.1_Electricity']:
            try:
                df = read_sheet(wb, sheet_name)
                if '_year' in df.columns:
                    all_years.update(df['_year'].unique())
            except: pass

        if not all_years:
            raise ValueError("❌ Excel'de geçerli yıl verisi bulunamadı.")

        print(f"  📅 Bulunan yıllar: {sorted(all_years)}\n")

        # ── 2. Her yıl için period_id al + mevcut veriyi temizle ─
        year_to_period = {}
        for year in sorted(all_years):
            year = int(year)
            period = get_or_create_period(session, year)
            year_to_period[year] = period.id
            clear_year_data(session, period.id, year)
        session.flush()

        # ── 3. Her sayfayı işle ───────────────────────────────────
        print("📌 KAPSAM 1")
        r11 = process_11(wb, session, year_to_period)
        r12 = process_12(wb, session, year_to_period)
        r14 = process_14(wb, session, year_to_period)

        print("\n📌 KAPSAM 2")
        r21 = process_21(wb, session, year_to_period)

        print("\n📌 KAPSAM 3")
        r31 = process_freight(wb, session, year_to_period, '3.1_Freight',         '[3.1] Hammadde Sevkiyatı')
        r32 = process_freight(wb, session, year_to_period, '3.2_ProductShipment', '[3.2] Ürün Sevkiyatı')
        r33 = process_33(wb, session, year_to_period)
        r34 = process_freight(wb, session, year_to_period, '3.4_BusinessTravel',  '[3.4] İş Seyahati')
        r35 = process_35(wb, session, year_to_period)
        r41 = process_purchased(wb, session, year_to_period, '4.1_Purchased',       'Miktar/USD',       8, '[4.1] Satın Alınan Malzemeler')
        r42 = process_42(wb, session, year_to_period)
        r43 = process_purchased(wb, session, year_to_period, '4.3_Waste',           'Miktar (kg)',       8, '[4.3] Atık')
        r44 = process_purchased(wb, session, year_to_period, '4.4_LeasedEquipment', 'Tutar (2021 USD)',  8, '[4.4] Kiralanan Ekipmanlar')
        r45 = process_purchased(wb, session, year_to_period, '4.5_Services',        'Tutar (2021 USD)',  8, '[4.5] Hizmet Alımları')
        r61 = process_61(wb, session, year_to_period)

        # ── 4. Yıl bazlı toplamları birleştir ────────────────────
        all_results = {
            '1.1':r11,'1.2':r12,'1.4':r14,
            '2.1':r21,
            '3.1':r31,'3.2':r32,'3.3':r33,'3.4':r34,'3.5':r35,
            '4.1':r41,'4.2':r42,'4.3':r43,'4.4':r44,'4.5':r45,
            '6.1':r61,
        }
        scope_totals = merge_totals(all_results)

        # ── 5. Audit log ──────────────────────────────────────────
        for year, totals in scope_totals.items():
            session.add(AuditLog(
                action="excel_import",
                scope="all",
                period_id=year_to_period[year],
                status="success",
                notes=(f"{year} | "
                       f"K1={totals['Kapsam 1 Top.']:.2f} "
                       f"K2={totals['Kapsam 2 Top.']:.2f} "
                       f"K3={totals['Kapsam 3 Top.']:.2f} "
                       f"Toplam={totals['GENEL TOPLAM']:.2f} ton CO2e")
            ))
        session.commit()
        print(f"\n  💾 Tüm veriler carbon.db'ye kaydedildi.")

    # ── 6. Özet sayfaları Excel'e yaz ────────────────────────────
    print("\n📌 ÖZET SAYFALAR")
    write_totals(wb, scope_totals)
    write_impact(wb, scope_totals)
    write_per_capita(wb, scope_totals)
    wb.save(excel_path)

    # ── 7. Final özet ────────────────────────────────────────────
    print_conversion_summary()

    print(f"\n{'═'*58}")
    print(f"  ✅ İMPORT TAMAMLANDI")
    for year in sorted(scope_totals.keys()):
        t = scope_totals[year]
        print(f"\n  📅 {year}")
        print(f"     Kapsam 1 : {t['Kapsam 1 Top.']:.4f} ton CO2e")
        print(f"     Kapsam 2 : {t['Kapsam 2 Top.']:.4f} ton CO2e")
        print(f"     Kapsam 3 : {t['Kapsam 3 Top.']:.4f} ton CO2e")
        print(f"     TOPLAM   : {t['GENEL TOPLAM']:.4f} ton CO2e")
    print(f"\n  Excel : {excel_path}")
    print(f"  DB    : carbon.db")
    print(f"{'═'*58}\n")


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "data/UNSPED_Karbon_Veri_Girisi_v3.xlsx"
    run_import(path)
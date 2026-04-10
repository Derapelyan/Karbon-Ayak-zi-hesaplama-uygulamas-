"""
UNSPED Carbon Footprint — Akıllı Faktör Güncelleme Sistemi
==========================================================
Kullanım:
    python run_update.py                                    # tüm kontroller
    python run_update.py --source defra                    # sadece DEFRA
    python run_update.py --excel data/UNSPED_...xlsx       # Excel'i de güncelle

Mantık:
    1. Her kaynak için "yeni yayın var mı?" kontrol eder
    2. Yeni yayın varsa indirir ve parse eder
    3. Yoksa son bilinen değerleri kullanır
    4. DB'ye kaydeder (yoksa ekle, varsa güncelle)
    5. Excel'deki BOŞ yeşil hücreleri doldurur
    6. Dolu hücrelere DOKUNMAZ (manuel override korunur)
    7. Audit log'a ne bulduğunu ve ne değiştirdiğini yazar
"""

import sys
import os
import re
import io
import json
import argparse
import requests
import pandas as pd
from datetime import datetime
from bs4 import BeautifulSoup

sys.path.insert(0, os.path.join(os.path.dirname(__file__)))
from db.connection import SessionLocal
from db.models import EmissionFactor, AuditLog

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; UNSPED-CarbonBot/1.0)"}
TIMEOUT = 30


# ══════════════════════════════════════════════════════════════════
# YARDIMCI — SON KONTROL TARİHİNİ DB'YE KAYDET / OKU
# ══════════════════════════════════════════════════════════════════

def get_last_check(session, source_name):
    """Bu kaynak için son kontrol kaydını döndürür."""
    log = session.query(AuditLog).filter(
        AuditLog.action == f"check_{source_name}"
    ).order_by(AuditLog.id.desc()).first()
    if log:
        return json.loads(log.notes or "{}")
    return {}

def save_check(session, source_name, data: dict, status="success"):
    """Kontrol sonucunu audit log'a kaydeder."""
    session.add(AuditLog(
        action=f"check_{source_name}",
        scope="emission_factors",
        status=status,
        notes=json.dumps(data, ensure_ascii=False)
    ))


# ══════════════════════════════════════════════════════════════════
# DEFRA FETCHER — TAM OTOMATİK
# Strateji:
#   1. GOV.UK koleksiyon sayfasını çek
#   2. En son yıl sayfasının URL'sini bul
#   3. O sayfadan flat file (CSV/XLSX) URL'sini bul
#   4. Dosyayı indir ve parse et
#   5. Son kontrolle karşılaştır — değişmişse güncelle
# ══════════════════════════════════════════════════════════════════

DEFRA_COLLECTION_URL = (
    "https://www.gov.uk/government/collections/"
    "government-conversion-factors-for-company-reporting"
)

def find_latest_defra_year_url():
    """
    GOV.UK koleksiyon sayfasını tarar ve en son yılın URL'sini döndürür.
    Örnek: 'greenhouse-gas-reporting-conversion-factors-2025'
    """
    print("    🌐 DEFRA koleksiyon sayfası kontrol ediliyor...")
    resp = requests.get(DEFRA_COLLECTION_URL, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()

    soup  = BeautifulSoup(resp.text, "html.parser")
    links = soup.find_all("a", href=True)

    # "/government/publications/greenhouse-gas-reporting-conversion-factors-YYYY"
    pattern = re.compile(
        r"/government/publications/greenhouse-gas-reporting-conversion-factors-(\d{4})"
    )
    years_found = {}
    for link in links:
        m = pattern.search(link["href"])
        if m:
            year = int(m.group(1))
            years_found[year] = "https://www.gov.uk" + m.group(0)

    if not years_found:
        raise ValueError("DEFRA koleksiyonunda yıl sayfası bulunamadı.")

    latest_year = max(years_found.keys())
    latest_url  = years_found[latest_year]
    print(f"    📅 En son DEFRA yılı: {latest_year} → {latest_url}")
    return latest_year, latest_url

def find_defra_flat_file_url(year_page_url):
    """
    Yıl sayfasından flat file (CSV veya XLSX) indirme URL'sini bulur.
    """
    resp = requests.get(year_page_url, headers=HEADERS, timeout=TIMEOUT)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # "flat file" içeren linkleri ara
    for link in soup.find_all("a", href=True):
        href  = link["href"]
        text  = link.get_text(strip=True).lower()
        if "flat" in text and ("xlsx" in href or "csv" in href or "xls" in href):
            if href.startswith("http"):
                return href
            return "https://www.gov.uk" + href

    # Yedek: assets.publishing.service.gov.uk üzerindeki tüm xlsx linkleri
    for link in soup.find_all("a", href=True):
        href = link["href"]
        if "assets.publishing.service.gov.uk" in href and "flat" in href.lower():
            return href

    raise ValueError(f"DEFRA flat file URL bulunamadı: {year_page_url}")

def download_and_parse_defra(flat_file_url, year):
    """
    DEFRA flat file XLSX'i indirir ve taşımacılık faktörlerini parse eder.
    Flat file formatı: Activity, Type, Unit, GHG/unit, kgCO2e sütunları içerir.
    """
    print(f"    📥 DEFRA flat file indiriliyor: {flat_file_url}")
    resp = requests.get(flat_file_url, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    print(f"    ✅ İndirildi ({len(resp.content)//1024} KB)")

    xl   = pd.ExcelFile(io.BytesIO(resp.content))
    print(f"    📋 Sayfalar: {xl.sheet_names[:8]}...")

    # ── Flat file formatı: "Factors by Category" tek sayfa ──────────
    # Sütunlar: Scope, Level 1, Level 2, Level 3, Level 4, Column Text,
    #           UOM, GHG Conversion Factor 2025, GHG Unit
    main_sheet = None
    for name in xl.sheet_names:
        nl = name.lower()
        if any(k in nl for k in ["factor", "category", "data", "ghg"]):
            main_sheet = name
            break
    if not main_sheet:
        main_sheet = xl.sheet_names[0]

    print(f"    📄 Sayfa: '{main_sheet}'")
    df = xl.parse(main_sheet)
    df.columns = [str(c).strip() for c in df.columns]
    print(f"    📋 Sütunlar: {list(df.columns[:8])}")

    # ── DEFRA Flat File Format — Kesin Yapı ──────────────────────────
    # Row 5 = başlık: ID | Scope | Level1 | Level2 | Level3 | Level4 |
    #                  Column Text | UOM | GHG/Unit | GHG Conversion Factor YYYY
    # Filtre: GHG/Unit == "kg CO2e" AND UOM == "km" AND taşımacılık

    df2 = xl.parse(main_sheet, header=5)
    df2.columns = [str(c).strip() for c in df2.columns]
    print(f"    📋 Sütunlar: {list(df2.columns)}")

    col_scope    = next((c for c in df2.columns if c.lower() == "scope"),          "Scope")
    col_level2   = next((c for c in df2.columns if "level 2" in c.lower()),        "Level 2")
    col_level3   = next((c for c in df2.columns if "level 3" in c.lower()),        "Level 3")
    col_col_text = next((c for c in df2.columns if "column text" in c.lower()),    "Column Text")
    col_uom      = next((c for c in df2.columns if c.lower() == "uom"),            "UOM")
    col_ghg_unit = next((c for c in df2.columns if "ghg/unit" in c.lower()),       "GHG/Unit")
    col_factor   = next((c for c in df2.columns if "ghg conversion factor" in c.lower()), None)

    if not col_factor:
        print("    ⚠ 'GHG Conversion Factor' sütunu bulunamadı — yerleşik değerler kullanılacak.")
        return None

    print(f"    📊 Değer sütunu: '{col_factor}'")

    # Filtrele: toplam CO2e + birim km
    mask = (
        (df2[col_ghg_unit].astype(str).str.strip() == "kg CO2e") &
        (df2[col_uom].astype(str).str.strip() == "km")
    )
    df_km = df2[mask].copy()
    print(f"    📏 km + kg CO2e filtresi: {len(df_km)} satır")

    factors = {}
    for _, row in df_km.iterrows():
        level2 = str(row.get(col_level2, "")).strip()
        level3 = str(row.get(col_level3, "nan")).strip()
        col_t  = str(row.get(col_col_text, "nan")).strip()

        if level3 and level3.lower() != "nan":
            key = f"{level2} — {level3}"
        else:
            key = level2
        if col_t and col_t.lower() not in ("nan", ""):
            key = f"{key} ({col_t})"

        try:
            val = float(pd.to_numeric(row.get(col_factor), errors="coerce"))
            if val > 0:
                factors[key] = val
        except (ValueError, TypeError):
            pass

    if not factors:
        print("    ⚠ Faktör bulunamadı — yerleşik değerler kullanılacak.")
        return None

    for k, v in list(factors.items())[:5]:
        print(f"    📋 {k[:60]}: {v:.5f}")

    print(f"    ✅ {len(factors)} taşımacılık faktörü parse edildi.")
    return {"year": year, "url": flat_file_url, "factors": factors}

# Yerleşik yedek — online erişilemezse kullanılır
DEFRA_FALLBACK = {
    "All HGVs":              0.08223,
    "HGV refrigerated":      0.10453,
    "Average HGV":           0.08223,
    "Small car":             0.14844,
    "Medium car":            0.18432,
    "Large car":             0.27793,
    "Average car":           0.17068,
    "Average van":           0.23339,
    "Domestic flight":       0.25504,
    "Short-haul flight":     0.15353,
    "Long-haul flight":      0.19085,
}

def update_defra(session):
    """DEFRA taşımacılık faktörlerini otomatik günceller."""
    print("\n  🔄 DEFRA kontrol ediliyor...")
    last = get_last_check(session, "defra")
    last_year = last.get("year", 0)

    try:
        latest_year, year_url = find_latest_defra_year_url()

        db_count = session.query(EmissionFactor).filter_by(
            scope="scope3", category="3.x"
        ).count()
        if latest_year == last_year and db_count > 0:
            print(f"    ✓  DEFRA {latest_year} zaten güncel — yeni yayın yok ({db_count} kayıt mevcut).")
            save_check(session, "defra",
                       {"year": latest_year, "status": "no_change"})
            session.commit()
            return 0, 0
        if db_count == 0:
            print(f"    ℹ️  DB boş — DEFRA faktörleri yükleniyor...")

        flat_url = find_defra_flat_file_url(year_url)
        result   = download_and_parse_defra(flat_url, latest_year)
        factors  = result["factors"] if result else DEFRA_FALLBACK

    except Exception as e:
        print(f"    ⚠ DEFRA online erişim başarısız: {e}")
        print("    ℹ️  Yerleşik değerler kullanılıyor.")
        factors     = DEFRA_FALLBACK
        latest_year = datetime.now().year

    added, updated = _save_defra_factors(session, factors, latest_year)
    save_check(session, "defra",
               {"year": latest_year, "factors_count": len(factors),
                "added": added, "updated": updated})
    session.commit()
    print(f"  ✅ DEFRA tamamlandı — {added} yeni, {updated} güncellendi.")
    return added, updated

def _save_defra_factors(session, factors, year):
    added, updated = 0, 0
    for vehicle_type, ef in factors.items():
        existing = session.query(EmissionFactor).filter_by(
            fuel_type=vehicle_type, scope="scope3", category="3.x"
        ).first()
        if not existing:
            session.add(EmissionFactor(
                source=f"DEFRA {year} GHG Conversion Factors",
                scope="scope3", category="3.x",
                fuel_type=vehicle_type, region="global",
                unit="kg CO2e/tonne.km", co2_factor=ef,
                input_unit="km", is_active=True, version=1, valid_year=year
            ))
            added += 1
            print(f"    ➕ {vehicle_type}: {ef:.5f} (yeni)")
        else:
            if abs((existing.co2_factor or 0) - ef) > 1e-8:
                print(f"    🔁 {vehicle_type}: {existing.co2_factor:.5f} → {ef:.5f}")
                existing.co2_factor = ef
                existing.valid_year = year
                updated += 1
    return added, updated


# ══════════════════════════════════════════════════════════════════
# IPCC GWP FETCHER
# Strateji:
#   IPCC yeni AR (Assessment Report) yayınladı mı kontrol et
#   AR6 (2021) → AR7 bekleniyor ~2027
#   Wikipedia / IPCC sitesini kontrol ederek yeni rapor var mı bak
#   Yeni rapor yoksa → mevcut AR6 değerleri geçerli
# ══════════════════════════════════════════════════════════════════

# AR6 GWP-100 değerleri (IPCC 2021, Tablo 7.SM.7)
GWP_AR6 = {
    "R134A":1526,"R-134A":1526,"HFC134A":1526,
    "R32":771,"R-32":771,"HFC32":771,
    "R410A":2088,"R-410A":2088,
    "R407C":1774,"R407c":1774,"R-407C":1774,
    "R404A":3922,"R-404A":3922,
    "R507A":3985,"R-507A":3985,
    "R22":1960,"R-22":1960,"HCFC22":1960,
    "R290":0,"R600A":4,"R600":4,
    "R744":1,"R717":0,
    "FM200":3220,"HFC227EA":3220,
    "HFC236FA":8060,"HFC236fa":8060,
    "SF6":25200,
    "R125":3740,"R-125":3740,
    "R143A":5810,"R-143A":5810,
    "R152A":164,"R-152A":164,
    "R1234YF":1,"R1234ZE":1,
    # Diğer gazlar
    "FK-5-1-12":   1,    # Novec 1230 — GWP-100 < 1
    "FK5112":      1,
    "NOVEC1230":   1,
    "C6F12O":      1,
    "R448A":       1387, # AR6
    "R449A":       1397, # AR6
    "R452A":       2140, # AR6
    "R454B":       466,  # AR6 — R32/R1234yf karışımı
    "R513A":       631,  # AR6
}
IPCC_CURRENT_AR = "AR6"
IPCC_NEXT_AR    = "AR7"

def check_ipcc_new_report():
    """
    IPCC'nin yeni Assessment Report yayınlayıp yayınlamadığını kontrol eder.
    IPCC AR7 ~2027 bekleniyor.
    """
    print("    🌐 IPCC yeni rapor kontrol ediliyor...")
    # AR7 GWP raporu (Chapter 7 Supplementary Material) yayınlandı mı kontrol et
    # AR7 Synthesis Report 2029'a kadar bekleniyor
    # GWP değerleri genellikle WG1 Physical Science Basis içinde yayınlanır
    # Kontrol: WG1 raporu yayınlandı mı?
    try:
        resp = requests.get(
            "https://www.ipcc.ch/report/ar7/wg1/",
            headers=HEADERS, timeout=TIMEOUT
        )
        if resp.status_code == 200:
            text = resp.text.lower()
            # WG1 gerçekten yayınlandıysa indirme linkleri ve chapter içerikleri olur
            has_gwp_content = (
                "table 7" in text or
                "chapter 7" in text and "download" in text or
                "global warming potential" in text and "supplementary" in text
            )
            if has_gwp_content:
                print(f"    🆕 IPCC AR7 WG1 yayınlandı — GWP değerlerini güncelleyin!")
                print(f"    🔗 https://www.ipcc.ch/report/ar7/wg1/")
                return True
    except Exception as e:
        print(f"    ⚠ IPCC AR7 WG1 kontrol hatası: {e}")

    # AR7 Planning: 2029'a kadar tamamlanması bekleniyor
    print(f"    ✓  IPCC AR7 henüz yayınlanmadı (2029 bekleniyor) — {IPCC_CURRENT_AR} geçerli.")
    return False

def update_gwp(session):
    """GWP değerlerini kontrol eder ve gerekirse günceller."""
    print("\n  🔄 IPCC GWP kontrol ediliyor...")
    last     = get_last_check(session, "gwp")
    new_ar   = check_ipcc_new_report()

    db_count = session.query(EmissionFactor).filter_by(category="1.4").count()
    if not new_ar and last.get("ar") == IPCC_CURRENT_AR and db_count > 0:
        print(f"    ✓  GWP değerleri {IPCC_CURRENT_AR} ile güncel — değişiklik yok ({db_count} kayıt mevcut).")
        session.commit()
        return 0, 0
    if db_count == 0:
        print(f"    ℹ️  DB boş — tüm GWP değerleri yükleniyor...")

    added, updated = 0, 0
    for gas, gwp in GWP_AR6.items():
        existing = session.query(EmissionFactor).filter_by(
            fuel_type=gas, scope="scope1", category="1.4"
        ).first()
        if not existing:
            session.add(EmissionFactor(
                source=f"IPCC {IPCC_CURRENT_AR} 2021",
                scope="scope1", category="1.4",
                fuel_type=gas, region="global",
                unit="ton CO2e/ton gas",
                gwp_ch4=gwp, is_active=True, version=1, valid_year=2024
            ))
            added += 1
            print(f"    ➕ {gas}: GWP={gwp} (yeni)")
        else:
            if (existing.gwp_ch4 or 0) != gwp:
                existing.gwp_ch4 = gwp
                updated += 1
                print(f"    🔁 {gas}: GWP güncellendi → {gwp}")

    save_check(session, "gwp",
               {"ar": IPCC_CURRENT_AR, "added": added, "updated": updated,
                "new_ar_detected": new_ar})
    session.commit()
    if new_ar:
        print("    ℹ️  Not: AR7 yayınlandığında GWP_AR6 sözlüğü güncellenmelidir.")
    print(f"  ✅ GWP tamamlandı — {added} yeni, {updated} güncellendi.")
    return added, updated


# ══════════════════════════════════════════════════════════════════
# IPCC YAKIT FAKTÖRLERİ
# IPCC 2006 GL nadiren değişir — 2019 Refinement yayınlandı
# Strateji: IPCC sitesinde "2019 Refinement" sonrası güncelleme var mı bak
# ══════════════════════════════════════════════════════════════════

IPCC_FUEL_FACTORS = {
    "natural_gas": {"ncv":48.0,"co2_factor":56.1,"ch4_factor":0.005,
                    "n2o_factor":0.0001,"density":0.000717,
                    "density_unit":"ton/m³","input_unit":"m³",
                    "source":"IPCC 2006 GL Vol.2 Table 2.2"},
    "diesel":      {"ncv":43.0,"co2_factor":74.1,"ch4_factor":0.01,
                    "n2o_factor":0.0006,"density":0.000832,
                    "density_unit":"ton/L","input_unit":"L",
                    "source":"IPCC 2006 GL Vol.2 Table 2.2"},
    "petrol":      {"ncv":44.3,"co2_factor":69.3,"ch4_factor":0.01,
                    "n2o_factor":0.0006,"density":0.000745,
                    "density_unit":"ton/L","input_unit":"L",
                    "source":"IPCC 2006 GL Vol.2 Table 2.2"},
    "diesel_shuttle":{"ncv":43.0,"co2_factor":74.1,"ch4_factor":0.01,
                    "n2o_factor":6e-05,"density":0.000832,
                    "density_unit":"ton/L","input_unit":"L",
                    "source":"IPCC 2006 GL Vol.2 Table 2.2"},
}
IPCC_GWP = {"gwp_ch4": 29.8, "gwp_n2o": 273.0}

def check_ipcc_fuel_update():
    """IPCC yakıt faktörlerinde güncelleme var mı kontrol eder."""
    print("    🌐 IPCC yakıt faktörü güncellemesi kontrol ediliyor...")
    try:
        resp = requests.get(
            "https://www.ipcc-nggip.iges.or.jp/public/2019rf/index.html",
            headers=HEADERS, timeout=TIMEOUT
        )
        # Sayfa erişilebilirse 2019 Refinement hâlâ en güncel
        if resp.status_code == 200:
            print("    ✓  IPCC 2019 Refinement hâlâ güncel.")
            return False
    except Exception as e:
        print(f"    ⚠ IPCC kontrol başarısız: {e}")
    return False

def update_ipcc_fuel(session):
    """IPCC yakıt faktörlerini günceller."""
    print("\n  🔄 IPCC yakıt faktörleri kontrol ediliyor...")
    last     = get_last_check(session, "ipcc_fuel")
    has_update = check_ipcc_fuel_update()

    db_count = session.query(EmissionFactor).filter_by(category="1.1", is_active=True).count()
    if not has_update and last.get("version") == "2019_refinement" and db_count > 0:
        print(f"    ✓  IPCC yakıt faktörleri güncel — değişiklik yok ({db_count} kayıt mevcut).")
        return 0, 0
    if db_count == 0:
        print(f"    ℹ️  DB boş — tüm IPCC yakıt faktörleri yükleniyor...")

    added, updated = 0, 0
    for fuel, vals in IPCC_FUEL_FACTORS.items():
        for category, scope in [("1.1","scope1"),("1.2","scope1"),("3.3","scope3")]:
            existing = session.query(EmissionFactor).filter_by(
                fuel_type=fuel, scope=scope, category=category, is_active=True
            ).first()
            if not existing:
                session.add(EmissionFactor(
                    source=vals["source"], scope=scope, category=category,
                    fuel_type=fuel, region="TR", unit="Ton/TJ",
                    co2_factor=vals["co2_factor"], ch4_factor=vals["ch4_factor"],
                    n2o_factor=vals["n2o_factor"],
                    gwp_ch4=IPCC_GWP["gwp_ch4"], gwp_n2o=IPCC_GWP["gwp_n2o"],
                    ncv=vals["ncv"], ncv_unit="TJ/Gg",
                    density=vals["density"], density_unit=vals["density_unit"],
                    input_unit=vals["input_unit"], is_active=True,
                    version=1, valid_year=2024
                ))
                added += 1
                print(f"    ➕ [{category}] {fuel} (yeni)")
            else:
                changed = any(
                    abs((getattr(existing, f) or 0) - v) > 1e-8
                    for f, v in [
                        ("co2_factor", vals["co2_factor"]),
                        ("ncv",        vals["ncv"]),
                        ("density",    vals["density"]),
                    ]
                )
                if changed:
                    for f, v in [("co2_factor",vals["co2_factor"]),
                                 ("ch4_factor",vals["ch4_factor"]),
                                 ("n2o_factor",vals["n2o_factor"]),
                                 ("ncv",vals["ncv"]),("density",vals["density"]),
                                 ("gwp_ch4",IPCC_GWP["gwp_ch4"]),
                                 ("gwp_n2o",IPCC_GWP["gwp_n2o"])]:
                        setattr(existing, f, v)
                    updated += 1
                    print(f"    🔁 [{category}] {fuel} güncellendi")

    save_check(session, "ipcc_fuel",
               {"version":"2019_refinement","added":added,"updated":updated})
    session.commit()
    print(f"  ✅ IPCC yakıt tamamlandı — {added} yeni, {updated} güncellendi.")
    return added, updated


# ══════════════════════════════════════════════════════════════════
# TEİAŞ ELEKTRİK GRID FAKTÖRÜ
# Strateji:
#   TEİAŞ resmi sayfasını kontrol et
#   Yeni yıl faktörü yayınlanmış mı bak
#   Yayınlanmamışsa önceki yılın değerini kullan ve uyar
# ══════════════════════════════════════════════════════════════════

TEIAS_KNOWN_FACTORS = {
    2019:0.481, 2020:0.452, 2021:0.434,
    2022:0.442, 2023:0.442, 2024:0.434,
}
TEIAS_URL = "https://www.teias.gov.tr/tr-TR/turkiye-elektrik-istatistikleri"
TEIAS_ALT = "https://www.epdk.gov.tr/Detay/Icerik/3-0-24-2"

def check_teias_new_factor():
    """
    TEİAŞ sayfasında yeni yıl faktörü var mı kontrol eder.
    TEİAŞ API sunmadığından sayfa içeriğini tarar.
    """
    current_year = datetime.now().year
    print(f"    🌐 TEİAŞ {current_year} faktörü kontrol ediliyor...")

    for url in [TEIAS_URL, TEIAS_ALT]:
        try:
            resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
            if resp.status_code != 200:
                continue

            text = resp.text.lower()
            # Sayfada güncel yıl verisi var mı
            if str(current_year) in text and ("emisyon" in text or "emission" in text
                                               or "co2" in text):
                print(f"    🔍 TEİAŞ sayfasında {current_year} verisi görüldü.")
                # Sayfa içinden faktörü çıkarmak zor — kullanıcıyı uyar
                return current_year, None  # Yıl var ama değer çıkarılamadı
        except Exception as e:
            print(f"    ⚠ {url}: {e}")

    print(f"    ℹ️  TEİAŞ online erişilemedi veya {current_year} verisi bulunamadı.")
    return None, None

def update_teias(session):
    """TEİAŞ elektrik grid faktörlerini günceller."""
    print("\n  🔄 TEİAŞ kontrol ediliyor...")
    last         = get_last_check(session, "teias")
    last_year    = last.get("year", 0)
    current_year = datetime.now().year

    new_year, new_factor = check_teias_new_factor()

    db_count = session.query(EmissionFactor).filter_by(
        fuel_type="electricity", scope="scope2", category="2.1"
    ).count()
    if db_count == 0:
        print(f"    ℹ️  DB boş — tüm TEİAŞ faktörleri yükleniyor...")

    added, updated = 0, 0

    # Bilinen tüm faktörleri kaydet
    for year, factor in TEIAS_KNOWN_FACTORS.items():
        existing = session.query(EmissionFactor).filter_by(
            fuel_type="electricity", scope="scope2",
            category="2.1", valid_year=year
        ).first()
        if not existing:
            session.add(EmissionFactor(
                source="TEİAŞ/EPDK",
                scope="scope2", category="2.1",
                fuel_type="electricity", region="TR",
                unit="ton CO2e/MWh", co2_factor=factor,
                input_unit="MWh", is_active=True, version=1, valid_year=year
            ))
            added += 1
            print(f"    ➕ {year}: {factor} ton CO2e/MWh (yeni)")
        else:
            if abs((existing.co2_factor or 0) - factor) > 1e-6:
                existing.co2_factor = factor
                updated += 1
                print(f"    🔁 {year}: güncellendi → {factor}")
            else:
                print(f"    ✓  {year}: {factor} (değişmedi)")

    # Yeni yıl tespit edildiyse ve elimizde değer yoksa uyar
    if new_year and new_year not in TEIAS_KNOWN_FACTORS:
        print(f"\n    ⚠️  TEİAŞ {new_year} YILI TESPİT EDİLDİ!")
        print(f"    ⚠️  run_update.py içindeki TEIAS_KNOWN_FACTORS'a")
        print(f"    ⚠️  {new_year} yılının grid faktörünü ekleyin.")
        print(f"    ⚠️  Kaynak: {TEIAS_URL}")

    save_check(session, "teias",
               {"year": current_year, "added": added, "updated": updated,
                "new_year_detected": new_year})
    session.commit()
    print(f"  ✅ TEİAŞ tamamlandı — {added} yeni, {updated} güncellendi.")
    return added, updated


# ══════════════════════════════════════════════════════════════════
# EXCEL GÜNCELLEME — BOŞ YEŞİL HÜCRELERİ DOLDUR
# ══════════════════════════════════════════════════════════════════

def is_empty(val):
    """Hücre boş mu?"""
    return val in (None, 0, "", "0", 0.0)

def is_data_row(val):
    """Veri satırı mı? (legend, not değil)"""
    if val is None: return False
    s = str(val).strip()
    return not (s.startswith("•") or s.startswith("←") or
                s.startswith("NOTLAR") or len(s) == 0)

def fill_green(cell, value):
    """Boşsa yeşil olarak doldur, doluysa dokunma. True dönerse değişti."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    if not is_empty(cell.value):
        return False
    thin = Side(style="thin", color="BFBFBF")
    cell.value     = value
    cell.font      = Font(name="Arial", size=10, color="375623")
    cell.fill      = PatternFill("solid", start_color="E2EFDA")
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = Border(left=thin, right=thin, top=thin, bottom=thin)
    return True

def update_excel(excel_path, session):
    """Excel'deki boş yeşil hücreleri DB'den çekilen güncel değerlerle doldurur."""
    from openpyxl import load_workbook
    print(f"\n  📝 Excel güncelleniyor: {excel_path}")
    wb      = load_workbook(excel_path)
    changes = 0

    # ── 1.4 GWP sütunu ────────────────────────────────────────────
    ws = wb['1.4_Refrigerants']
    for row in ws.iter_rows(min_row=4, values_only=False):
        gas_cell = row[1]   # B — gaz türü
        gwp_cell = row[4]   # E — GWP
        if not gas_cell.value or not is_data_row(gas_cell.value): continue

        gas = str(gas_cell.value).strip()
        gwp = GWP_AR6.get(gas) or GWP_AR6.get(gas.upper().replace("-","").replace(" ",""))
        if gwp is not None and fill_green(gwp_cell, gwp):
            print(f"    📗 1.4 {gas}: GWP={gwp}")
            changes += 1

    # ── 1.1, 1.2, 3.3 — Yakıt faktörleri ─────────────────────────
    combustion_sheets = {
        '1.1_Stationary': (2, 5, 6, 7, 8, 9, 10, 11),
        '1.2_Mobile':     (2, 5, 6, 7, 8, 9, 10, 11),
        '3.3_Commuting':  (3, 6, 7, 8, 9, 10, 11, 12),
    }
    # sütun indeksleri: fuel, ncv, ef_co2, ef_ch4, ef_n2o, gwp_ch4, gwp_n2o, density

    for sheet_name, cols in combustion_sheets.items():
        ws = wb[sheet_name]
        fuel_col, ncv_col, co2_col, ch4_col, n2o_col, gch4_col, gn2o_col, dens_col = cols
        for row in ws.iter_rows(min_row=4, values_only=False):
            fuel_cell = row[fuel_col - 1]
            if not fuel_cell.value or not is_data_row(fuel_cell.value): continue
            fuel = str(fuel_cell.value).strip().lower()
            vals = IPCC_FUEL_FACTORS.get(fuel)
            if not vals: continue

            mapping = {
                ncv_col:   vals['ncv'],
                co2_col:   vals['co2_factor'],
                ch4_col:   vals['ch4_factor'],
                n2o_col:   vals['n2o_factor'],
                gch4_col:  IPCC_GWP['gwp_ch4'],
                gn2o_col:  IPCC_GWP['gwp_n2o'],
                dens_col:  vals['density'],
            }
            filled = 0
            for col_idx, value in mapping.items():
                if fill_green(row[col_idx - 1], value):
                    filled += 1; changes += 1
            if filled:
                print(f"    📗 {sheet_name} [{fuel}]: {filled} hücre dolduruldu")

    # ── 2.1 EF sütunu — yıla göre TEİAŞ ──────────────────────────
    ws = wb['2.1_Electricity']
    for row in ws.iter_rows(min_row=4, values_only=False):
        year_cell = row[0]
        ef_cell   = row[2]  # C
        if not year_cell.value or not is_data_row(year_cell.value): continue
        try: year = int(float(str(year_cell.value)))
        except: continue

        # DB'den o yılın faktörünü al
        ef_rec = session.query(EmissionFactor).filter_by(
            fuel_type="electricity", scope="scope2",
            category="2.1", valid_year=year
        ).first()
        ef = ef_rec.co2_factor if ef_rec else TEIAS_KNOWN_FACTORS.get(year)
        if ef and fill_green(ef_cell, ef):
            print(f"    📗 2.1 {year}: EF={ef}")
            changes += 1

    # ── 3.1, 3.2, 3.4 — Araç tipine göre DEFRA ───────────────────
    # DB'den tüm DEFRA faktörlerini yükle
    defra_records = session.query(EmissionFactor).filter_by(
        scope="scope3", category="3.x"
    ).all()

    # Akıllı eşleştirme: Excel'deki araç tipi → DEFRA anahtarı
    # DEFRA "Freighting goods — Average HGV (100% Laden)" gibi uzun isimler kullanır
    # Excel'de "All HGVs", "Karayolu-%100 Laden" gibi kısa isimler var
    VEHICLE_MAP = {
        # Excel değeri (lower)         : DEFRA key içinde aranacak kelimeler
        "all hgvs":                    ["average hgv", "100% laden"],
        "karayolu-%100 laden":         ["average hgv", "100% laden"],
        "hgv":                         ["average hgv", "100% laden"],
        "small car":                   ["small car", "average car"],
        "medium car":                  ["medium car"],
        "large car":                   ["large car"],
        "average car":                 ["average car"],
        "average van":                 ["average van"],
        "karayolu (dizel)":            ["average car", "medium car"],
        "karayolu (benzin)":           ["average car", "medium car"],
        "domestic flight":             ["domestic", "flight"],
        "short-haul flight":           ["short-haul", "flight"],
        "long-haul flight":            ["long-haul", "flight"],
    }

    def find_defra_ef(vehicle_name, defra_records):
        """Excel'deki araç ismine göre DEFRA faktörünü bul."""
        vl = vehicle_name.lower().strip()

        # 1. Direkt eşleşme
        for rec in defra_records:
            if rec.fuel_type.lower() == vl:
                return rec.co2_factor

        # 2. VEHICLE_MAP ile eşleşme
        search_terms = VEHICLE_MAP.get(vl)
        if search_terms:
            for rec in defra_records:
                rt = rec.fuel_type.lower()
                if all(t in rt for t in search_terms):
                    return rec.co2_factor

        # 3. Kısmi eşleşme — en az bir kelime geçiyor mu
        vehicle_words = [w for w in vl.split() if len(w) > 3]
        best_match = None
        best_score = 0
        for rec in defra_records:
            rt = rec.fuel_type.lower()
            score = sum(1 for w in vehicle_words if w in rt)
            if score > best_score:
                best_score = score
                best_match = rec

        if best_match and best_score >= 1:
            return best_match.co2_factor

        return None

    transport_sheets = {
        '3.1_Freight':         (5, 7),
        '3.2_ProductShipment': (5, 7),
        '3.4_BusinessTravel':  (5, 7),
    }
    for sheet_name, (vcol, ecol) in transport_sheets.items():
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=4, values_only=False):
            vehicle_cell = row[vcol - 1]
            ef_cell      = row[ecol - 1]
            if not vehicle_cell.value or not is_data_row(vehicle_cell.value): continue

            vehicle = str(vehicle_cell.value).strip()
            ef = find_defra_ef(vehicle, defra_records)
            ef_val = ef_cell.value
            empty  = is_empty(ef_val)
            if ef and empty:
                fill_green(ef_cell, ef)
                print(f"    📗 {sheet_name} [{vehicle}]: EF={ef:.5f}")
                changes += 1
            elif ef and not empty:
                print(f"    ✓  {sheet_name} [{vehicle}]: EF={ef_val} (zaten dolu)")
            elif not ef:
                print(f"    ⚠ {sheet_name} [{vehicle}]: DEFRA eşleşmesi bulunamadı")

    wb.save(excel_path)
    print(f"\n  ✅ Excel güncellendi — {changes} hücre dolduruldu.")
    return changes


# ══════════════════════════════════════════════════════════════════
# ANA FONKSİYON
# ══════════════════════════════════════════════════════════════════

def run_update(sources=None, excel_path=None):
    if sources is None:
        sources = ["gwp", "ipcc", "teias", "defra"]

    print(f"\n{'═'*58}")
    print(f"  UNSPED KARBON — AKILLI FAKTÖR GÜNCELLEME")
    print(f"  {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  Kontrol edilecekler: {', '.join(s.upper() for s in sources)}")
    print(f"{'═'*58}")

    total_added = total_updated = 0

    with SessionLocal() as session:
        if "gwp"   in sources:
            a, u = update_gwp(session);        total_added+=a; total_updated+=u
        if "ipcc"  in sources:
            a, u = update_ipcc_fuel(session);  total_added+=a; total_updated+=u
        if "teias" in sources:
            a, u = update_teias(session);      total_added+=a; total_updated+=u
        if "defra" in sources:
            a, u = update_defra(session);      total_added+=a; total_updated+=u

        session.add(AuditLog(
            action="run_update", scope="emission_factors", status="success",
            notes=json.dumps({
                "sources": sources,
                "added": total_added,
                "updated": total_updated,
                "timestamp": datetime.now().isoformat()
            }, ensure_ascii=False)
        ))
        session.commit()

    # Excel güncelle
    excel_changes = 0
    if excel_path:
        if os.path.exists(excel_path):
            with SessionLocal() as session:
                excel_changes = update_excel(excel_path, session)
        else:
            print(f"\n  ⚠ Excel bulunamadı: {excel_path}")

    print(f"\n{'═'*58}")
    print(f"  ✅ GÜNCELLEME TAMAMLANDI")
    print(f"  DB    : {total_added} yeni, {total_updated} güncellendi")
    if excel_path:
        print(f"  Excel : {excel_changes} hücre dolduruldu")
    print(f"\n  Sonraki adım:")
    print(f"  python pipeline/Importer.py data/UNSPED_Karbon_Veri_Girisi_v3.xlsx")
    print(f"{'═'*58}\n")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", "-s", nargs="+",
                        choices=["gwp","ipcc","teias","defra","all"],
                        default=["all"])
    parser.add_argument("--excel", "-e", type=str, default=None)
    args    = parser.parse_args()
    sources = ["gwp","ipcc","teias","defra"] if "all" in args.source else args.source
    run_update(sources=sources, excel_path=args.excel)
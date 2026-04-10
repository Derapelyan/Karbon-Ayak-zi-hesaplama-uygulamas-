"""
DEFRA flat file yapısını inceler.
carbon_project klasöründe çalıştır:
    python inspect_defra.py
"""
import sys, os, io
import pandas as pd

# DEFRA dosyasını bul
possible_paths = [
    "defra_temp.xlsx",
    os.path.join(os.path.expanduser("~"), "Downloads", "ghg-conversion-factors-2025-flat-format.xlsx"),
]

# DB'den URL'yi al ve indir
try:
    sys.path.insert(0, os.path.dirname(__file__))
    from db.connection import SessionLocal
    from db.models import AuditLog
    import json, requests

    with SessionLocal() as s:
        log = s.query(AuditLog).filter(
            AuditLog.action == "check_defra"
        ).order_by(AuditLog.id.desc()).first()
        if log:
            data = json.loads(log.notes or "{}")
            url  = data.get("url") or data.get("factors", {})

    # Direkt URL dene
    url = "https://assets.publishing.service.gov.uk/media/6846b6ea57f3515d9611f0dd/ghg-conversion-factors-2025-flat-format.xlsx"
    headers = {"User-Agent": "Mozilla/5.0 (compatible; UNSPED-CarbonBot/1.0)"}
    print(f"İndiriliyor: {url}")
    resp = requests.get(url, headers=headers, timeout=60)
    resp.raise_for_status()
    content = io.BytesIO(resp.content)
    print(f"İndirildi: {len(resp.content)//1024} KB")

except Exception as e:
    print(f"İndirme hatası: {e}")
    print("Lütfen DEFRA flat file'ı manuel indirin ve yolu aşağıya yazın.")
    path = input("Dosya yolu: ").strip().strip('"')
    content = open(path, 'rb')

xl = pd.ExcelFile(content)
print(f"\nSayfalar: {xl.sheet_names}")

df = xl.parse('Factors by Category', header=None)
print(f"\nSatır sayısı: {len(df)}, Sütun sayısı: {len(df.columns)}")

print("\n=== İlk 10 satır (ham) ===")
pd.set_option('display.max_columns', 20)
pd.set_option('display.width', 200)
pd.set_option('display.max_colwidth', 30)
for i in range(min(10, len(df))):
    print(f"Row {i:2d}: {[str(v)[:25] for v in df.iloc[i]]}")

print("\n=== HGV içeren satırlar ===")
count = 0
for i, row in df.iterrows():
    row_str = " ".join(str(v) for v in row.values).lower()
    if "hgv" in row_str or "heavy" in row_str:
        print(f"Row {i:3d}: {[str(v)[:20] for v in row]}")
        count += 1
        if count >= 5:
            break

print("\n=== Sütun başlıkları olan satırı bul ===")
for i in range(min(15, len(df))):
    row_str = " ".join(str(v) for v in df.iloc[i]).lower()
    if any(k in row_str for k in ["scope", "level", "ghg", "unit", "kg"]):
        print(f"Row {i}: {[str(v)[:30] for v in df.iloc[i]]}")
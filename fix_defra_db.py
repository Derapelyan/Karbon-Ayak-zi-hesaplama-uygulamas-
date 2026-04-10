"""
DEFRA kayıtlarını DB'den temizler ve kontrol geçmişini sıfırlar.
carbon_project klasöründe çalıştır:
    python fix_defra_db.py
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from db.connection import SessionLocal
from db.models import EmissionFactor, AuditLog

with SessionLocal() as session:
    # Tüm scope3 / 3.x kayıtlarını göster
    all_defra = session.query(EmissionFactor).filter_by(
        scope="scope3", category="3.x"
    ).all()

    print(f"Toplam DEFRA kaydı: {len(all_defra)}")
    for r in all_defra[:10]:
        print(f"  id={r.id} | {r.fuel_type[:50]:<50} | ef={r.co2_factor}")

    if all_defra:
        confirm = input(f"\nTüm {len(all_defra)} DEFRA kaydı silinsin mi? (e/h): ").strip().lower()
        if confirm == "e":
            for r in all_defra:
                session.delete(r)

            # Kontrol geçmişini de temizle
            old_checks = session.query(AuditLog).filter(
                AuditLog.action == "check_defra"
            ).all()
            for c in old_checks:
                session.delete(c)

            session.commit()
            print(f"\n✅ {len(all_defra)} DEFRA kaydı ve kontrol geçmişi silindi.")
            print("Şimdi şunu çalıştırın:")
            print("  python run_update.py --source defra --excel data/UNSPED_Karbon_Veri_Girisi_v3.xlsx")
        else:
            print("İptal edildi.")
    else:
        print("Silinecek kayıt yok.")
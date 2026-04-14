@echo off
chcp 65001 >nul 2>&1
title UNSPED Karbon Dashboard

echo.
echo  ================================================
echo   UNSPED Karbon Ayak Izi Dashboard
echo  ================================================
echo.

:: ?? Python kontrolu ??????????????????????????????????????????????
python --version >nul 2>&1
if errorlevel 1 (
    echo  [HATA] Python bulunamadi!
    echo.
    echo  Lutfen once Python yukleyin:
    echo  https://www.python.org/downloads/
    echo.
    echo  Yuklerken "Add python.exe to PATH" secenegini isaretleyin!
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b 1
)
echo  [OK] Python bulundu.

:: ?? Sanal ortam olustur (yoksa) ??????????????????????????????????
if not exist "venv\Scripts\activate.bat" (
    echo  [KURULUM] Sanal ortam olusturuluyor...
    python -m venv venv
    if errorlevel 1 (
        echo  [HATA] Sanal ortam olusturulamadi!
        pause
        exit /b 1
    )
    echo  [OK] Sanal ortam olusturuldu.
)

:: ?? Sanal ortami aktive et ????????????????????????????????????????
call venv\Scripts\activate.bat
echo  [OK] Sanal ortam aktif.

:: ?? Paketleri kur (requirements.txt ile tek seferde) ?????????????
echo  [KURULUM] Paketler kontrol ediliyor...
python -m pip install --upgrade pip --quiet --disable-pip-version-check

if exist "requirements.txt" (
    python -m pip install -r requirements.txt --quiet
    if errorlevel 1 (
        echo  [HATA] Paket kurulumu basarisiz!
        echo  Internet baglantinizi kontrol edin.
        pause
        exit /b 1
    )
    echo  [OK] Tum paketler hazir.
) else (
    echo  [UYARI] requirements.txt bulunamadi, tek tek kurulmaya calisiliyor...
    python -m pip install sqlalchemy pandas openpyxl requests beautifulsoup4 matplotlib --quiet
    if errorlevel 1 (
        echo  [HATA] Paket kurulumu basarisiz!
        pause
        exit /b 1
    )
    echo  [OK] Paketler kuruldu.
)

:: Word export python-docx ile yapilir (Node.js gerekmez)
echo  [OK] Word export icin python-docx kullaniliyor (requirements.txt ile kuruldu).

:: DB ilk kurulum (yoksa)????????????
if not exist "carbon.db" (
    echo.
    echo  [KURULUM] Veritabani ilk kez olusturuluyor...
    python main.py
    if errorlevel 1 (
        echo  [HATA] Veritabani olusturulamadi!
        pause
        exit /b 1
    )
    echo  [OK] Veritabani olusturuldu.

    :: ?? ?lk kurulumda Excel ye?il h?crelerini doldur ????????????
    set EXCEL=data\UNSPED_Karbon_Veri_Girisi_v3.xlsx
    if exist "%EXCEL%" (
        echo.
        echo  [KURULUM] Excel bos yesil hucreler dolduruluyor...
        python run_update.py --excel "%EXCEL%"
        if errorlevel 1 (
            echo  [UYARI] Faktor guncelleme basarisiz - internet baglantisini kontrol edin.
            echo  Daha sonra GUNCELLE.bat ile tekrar deneyebilirsiniz.
        ) else (
            echo  [OK] Excel guncellendi.
        )
    ) else (
        echo  [UYARI] Excel dosyasi bulunamadi: %EXCEL%
        echo  Fakt?r guncelleme icin GUNCELLE.bat calistirin.
    )
)

:: ?? Dashboard baslat ?????????????????????????????????????????????
echo.
echo  [OK] Dashboard aciliyor...
echo  ================================================
echo.

python dashboard.py

if errorlevel 1 (
    echo.
    echo  [HATA] Dashboard baslatilirken sorun olustu!
    echo  Yukaridaki hata mesajini not alin.
    pause
)
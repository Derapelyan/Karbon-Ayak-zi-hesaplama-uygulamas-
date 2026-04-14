@echo off
chcp 65001 >nul
echo.
echo  ================================================
echo   UNSPED Karbon Sistemi - Ilk Kurulum
echo  ================================================
echo.

:: Python kontrol
python --version >nul 2>&1
if errorlevel 1 (
    echo  [HATA] Python bulunamadi.
    echo  Indirin: https://python.org  ^(Add to PATH secin^)
    start https://python.org/downloads
    pause
    exit /b 1
)
echo  [OK] Python bulundu.

:: Sanal ortam
if not exist "venv\Scripts\activate.bat" (
    echo  [KURULUM] Sanal ortam olusturuluyor...
    python -m venv venv
    echo  [OK] Sanal ortam olusturuldu.
)
call venv\Scripts\activate.bat
echo  [OK] Sanal ortam aktif.

:: pip paketleri (python-docx dahil, Node.js gerekmez)
echo.
echo  [KURULUM] Python paketleri yukleniyor...
python -m pip install --upgrade pip --quiet --disable-pip-version-check
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo  [HATA] Paket kurulumu basarisiz. Internet baglantisini kontrol edin.
    pause
    exit /b 1
)
echo  [OK] Tum paketler yuklendi (sqlalchemy, pandas, openpyxl, python-docx, matplotlib...).

:: Word export kontrolu
if exist "word_export\generate_report.py" (
    echo  [OK] Word export hazir ^(python-docx ile, Node.js gerekmez^).
) else (
    echo  [UYARI] word_export\generate_report.py bulunamadi.
)

:: DB kurulum
echo.
if not exist "carbon.db" (
    echo  [KURULUM] Veritabani olusturuluyor...
    python main.py
    echo  [OK] Veritabani hazir.
) else (
    echo  [OK] Veritabani zaten mevcut.
)

echo.
echo  ================================================
echo   Kurulum tamamlandi!
echo   Artik BASLAT.bat ile sistemi acabilirsiniz.
echo  ================================================
echo.
pause
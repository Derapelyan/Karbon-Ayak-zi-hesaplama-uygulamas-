@echo off
chcp 65001 >nul 2>&1
title UNSPED - Emisyon Faktoru Guncelleme

echo.
echo  ================================================
echo   UNSPED Emisyon Faktoru Guncelleme
echo  ================================================
echo.

if not exist "venv\Scripts\activate.bat" (
    echo  [HATA] Once BASLAT.bat calistirin!
    pause
    exit /b 1
)
call venv\Scripts\activate.bat

set EXCEL=data\UNSPED_Karbon_Veri_Girisi_v3.xlsx

if not exist "%EXCEL%" (
    echo  [UYARI] Excel dosyasi bulunamadi: %EXCEL%
    echo  Sadece online kontrol yapilacak.
    echo  Bos yesil hucreler doldurulamaz ? Excel olmadan calisiliyor.
    echo.
    python run_update.py
) else (
    echo  Excel: %EXCEL%
    echo  Online kontrol + bos yesil hucreler doldurulacak...
    echo.
    python run_update.py --excel "%EXCEL%"
)

echo.
if errorlevel 1 (
    echo  [HATA] Guncelleme basarisiz!
    echo  Hata detayi icin yukaridaki mesajlara bakin.
    pause
    exit /b 1
)

echo  [OK] Guncelleme tamamlandi!
echo.
echo  Simdi ne yapmak istersiniz?
echo  1 - Import Et  (IMPORT_ET.bat)
echo  2 - Dashboard ac  (BASLAT.bat)
echo  3 - Cik
echo.
set /p SECIM="Seciminiz (1/2/3): "

if "%SECIM%"=="1" (
    call IMPORT_ET.bat
) else if "%SECIM%"=="2" (
    python dashboard.py
)

pause
@echo off
chcp 65001 >nul 2>&1
title UNSPED - Excel Import

echo.
echo  ================================================
echo   UNSPED Excel Import
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
    echo  [HATA] Excel dosyasi bulunamadi: %EXCEL%
    echo  Dosyayi su klasore koyun:
    echo  %CD%\data\
    pause
    exit /b 1
)

echo  Excel: %EXCEL%
echo  Import basliyor...
echo.

python pipeline\Importer.py "%EXCEL%"

echo.
if errorlevel 1 (
    echo  [HATA] Import basarisiz!
    echo  Hata detayi icin yukaridaki mesajlara bakin.
    pause
    exit /b 1
)

echo  [OK] Import tamamlandi!
echo.
set /p OPEN="Dashboard acilsin mi? (e/h): "
if /i "%OPEN%"=="e" (
    python dashboard.py
)
echo.
pause
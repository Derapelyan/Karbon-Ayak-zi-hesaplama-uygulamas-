@echo off
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
) else (
    echo  [OK] Import tamamlandi!
    echo  Sonuclar icin BASLAT.bat ile dashboard acin.
)
echo.
pause
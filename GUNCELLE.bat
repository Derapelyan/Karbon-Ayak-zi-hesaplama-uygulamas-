@echo off
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
    echo  [UYARI] Excel bulunamadi, sadece DB guncelleniyor...
    python run_update.py
) else (
    echo  Excel: %EXCEL%
    echo  Guncelleme basliyor...
    echo.
    python run_update.py --excel "%EXCEL%"
)

echo.
echo  Guncelleme tamamlandi!
echo.
pause
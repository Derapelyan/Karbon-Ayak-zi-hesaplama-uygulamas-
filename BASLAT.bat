@echo off
chcp 65001 >nul 2>&1
title UNSPED Karbon Dashboard

echo.
echo  ================================================
echo   UNSPED Karbon Ayak Izi Dashboard
echo  ================================================
echo.

:: Python kontrolu
python --version >nul 2>&1
if errorlevel 1 (
    echo  [HATA] Python bulunamadi!
    echo.
    echo  Lutfen once Python yukleyin:
    echo  https://www.python.org/downloads/
    echo.
    echo  Yuklerken "Add python.exe to PATH" secenegini
    echo  isaretle!
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b 1
)
echo  [OK] Python bulundu.

:: Sanal ortam olustur (yoksa)
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

:: Sanal ortami aktive et
call venv\Scripts\activate.bat
echo  [OK] Sanal ortam aktif.

:: Paketleri kur
echo  [KURULUM] Paketler kontrol ediliyor...
python -m pip install --upgrade pip --quiet --disable-pip-version-check

python -c "import sqlalchemy" >nul 2>&1
if errorlevel 1 (
    echo  [KURULUM] sqlalchemy kuruluyor...
    python -m pip install sqlalchemy --quiet
)

python -c "import pandas" >nul 2>&1
if errorlevel 1 (
    echo  [KURULUM] pandas kuruluyor...
    python -m pip install pandas --quiet
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo  [KURULUM] openpyxl kuruluyor...
    python -m pip install openpyxl --quiet
)

python -c "import requests" >nul 2>&1
if errorlevel 1 (
    echo  [KURULUM] requests kuruluyor...
    python -m pip install requests --quiet
)

python -c "import bs4" >nul 2>&1
if errorlevel 1 (
    echo  [KURULUM] beautifulsoup4 kuruluyor...
    python -m pip install beautifulsoup4 --quiet
)

python -c "import matplotlib" >nul 2>&1
if errorlevel 1 (
    echo  [KURULUM] matplotlib kuruluyor...
    python -m pip install matplotlib --quiet
)

echo  [OK] Tum paketler hazir.

:: DB ilk kurulum (yoksa)
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
)

:: Dashboard baslat
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
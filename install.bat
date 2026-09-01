@echo off
title Instalacja - PDF to XLSX
cd /d "%~dp0"

set "PY="
py -3 --version >nul 2>nul && set "PY=py -3"
if not defined PY (
    python --version >nul 2>nul && set "PY=python"
)
if not defined PY (
    echo BLAD: nie znaleziono Pythona.
    echo Zainstaluj Python 3.9+ z https://www.python.org/downloads/
    echo Podczas instalacji zaznacz "Add python.exe to PATH".
    pause
    exit /b 1
)
echo === Python ===
%PY% --version

echo.
echo === Instalacja bibliotek (pdfplumber, openpyxl) ===
%PY% -m pip install --upgrade pip
%PY% -m pip install --upgrade pdfplumber openpyxl
if errorlevel 1 (
    echo BLAD instalacji bibliotek.
    pause
    exit /b 1
)

echo.
echo === Test ===
%PY% -c "import pdfplumber, openpyxl, tkinter; print('Biblioteki OK')"
if errorlevel 1 (
    echo BLAD: brak modulu - jesli chodzi o tkinter, zainstaluj Pythona
    echo ponownie z zaznaczona opcja "tcl/tk and IDLE".
    pause
    exit /b 1
)
%PY% pdf_to_xlsx.py --selftest
if errorlevel 1 (
    echo BLAD testu programu.
    pause
    exit /b 1
)

echo.
echo Gotowe. Program uruchamiasz plikiem uruchom.bat
pause

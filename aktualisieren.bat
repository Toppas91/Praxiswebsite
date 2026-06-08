@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════╗
echo  ║   Praxis-Daten aktualisieren         ║
echo  ╚══════════════════════════════════════╝
echo.

python --version > nul 2>&1
if %errorlevel% neq 0 (
    py --version > nul 2>&1
    if %errorlevel% neq 0 (
        echo  FEHLER: Python wurde nicht gefunden.
        echo  Bitte Python installieren: https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    )
    py -m pip install openpyxl --quiet
    py aktualisieren.py
) else (
    python -m pip install openpyxl --quiet
    python aktualisieren.py
)

if %errorlevel% neq 0 (
    echo.
    echo  Fehler beim Aktualisieren. Bitte Ausgabe oben pruefen.
    echo.
    pause
) else (
    echo.
    echo  Nicht vergessen: Aenderungen per Git pushen.
    echo.
    timeout /t 5 > nul
)

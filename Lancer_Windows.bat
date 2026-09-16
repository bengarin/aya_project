@echo off
REM Double-cliquez sur ce fichier pour demarrer le logiciel sous Windows.
cd /d "%~dp0"
python main.py
if errorlevel 1 (
    echo.
    echo Si Python n'est pas reconnu, installez-le depuis https://www.python.org/downloads/
    echo puis relancez ce fichier.
    pause
)

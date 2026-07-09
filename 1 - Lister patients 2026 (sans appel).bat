@echo off
cd /d "%~dp0"
REM Verifie la lecture de PUBLIC.mdb : liste les patients 2026 (aucun appel).
python -u etude_insi.py --lister
echo.
pause

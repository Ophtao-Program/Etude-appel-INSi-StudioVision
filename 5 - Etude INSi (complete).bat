@echo off
cd /d "%~dp0"
REM Etude COMPLETE : tous les patients ayant consulte en 2026.
python -u etude_insi.py
echo.
pause

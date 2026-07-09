@echo off
cd /d "%~dp0"
REM Etude complete limitee a 10 patients (test).
python -u etude_insi.py --test
echo.
pause

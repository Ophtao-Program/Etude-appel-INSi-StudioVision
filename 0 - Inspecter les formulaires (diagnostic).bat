@echo off
cd /d "%~dp0"
REM DIAGNOSTIC : liste tous les formulaires ouverts dans StudioVision + leurs controles.
python -u etude_insi.py --inspecter
echo.
pause

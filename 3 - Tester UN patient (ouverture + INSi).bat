@echo off
cd /d "%~dp0"
REM Chaine complete sur UN patient : ouverture (RecordSource) + appel INSi, en detail.
set /p CODE=Entrez le code du patient a tester : 
python -u etude_insi.py --patient %CODE%
echo.
pause

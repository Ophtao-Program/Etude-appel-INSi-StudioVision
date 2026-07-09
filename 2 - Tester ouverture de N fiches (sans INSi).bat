@echo off
cd /d "%~dp0"
REM Teste UNIQUEMENT l'ouverture des fiches par reecriture du RecordSource,
REM SANS appel INSi (pour reperer les codes a valeurs speciales).
REM Prerequis : StudioVision ouvert avec une fiche patient affichee.
set /p N=Nombre de fiches a tester (Entree = toutes) : 
if "%N%"=="" (
  python -u etude_insi.py --test-ouverture
) else (
  python -u etude_insi.py --test-ouverture %N%
)
echo.
pause

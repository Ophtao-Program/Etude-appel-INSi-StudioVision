@echo off
cd /d "%~dp0"
REM Statistiques completes a partir de etude_insi_resultats.json.
REM Aucune connexion a StudioVision n'est necessaire : lecture du fichier seulement.
REM Produit : rapport a l'ecran + etude_insi_statistiques.json + etude_insi_rapport.md
REM           + CSV (01 non trouves, incidents a reprendre, sexes corriges).
REM
REM Option : indiquez le nombre total de patients de l'annee pour le taux de couverture.
set /p TOTAL=Nombre total de patients de l'annee (Entree = ignorer) : 
if "%TOTAL%"=="" (
  python -u analyse_statistique.py
) else (
  python -u analyse_statistique.py --total %TOTAL%
)
echo.
pause

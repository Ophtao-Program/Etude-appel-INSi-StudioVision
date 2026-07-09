@echo off
cd /d "%~dp0"
REM Applique la resolution du sexe aux patients laisses en "sans_reponse"
REM (sexe absent) dans etude_insi_resultats.json, et met le fichier a jour.
REM A lancer avant de relancer l'etude complete. Une sauvegarde .bak est creee.
python -u test_resolution_sexe.py
echo.
pause

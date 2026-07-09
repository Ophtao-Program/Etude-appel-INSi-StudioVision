@echo off
cd /d "%~dp0"
REM ===================================================================
REM  Rejoue l'appel INSi pour les PATIENTES en echec (reponse 01),
REM  en substituant temporairement leur NOM par leur NOM DE JEUNE FILLE.
REM
REM  Le nom legal est SYSTEMATIQUEMENT restaure apres chaque appel :
REM  ce programme MESURE, il ne modifie pas durablement les fiches.
REM
REM  Prerequis : StudioVision ouvert, avec une fiche patient affichee.
REM              etude_insi_resultats.json present.
REM ===================================================================
echo.
echo  Nombre de patientes a traiter (Entree = toutes).
echo  Conseil : commencez par 20 pour valider la methode.
set /p N=Nombre : 
echo.
if "%N%"=="" (
  python -u correction_nom_jeune_fille.py
) else (
  python -u correction_nom_jeune_fille.py --limit %N%
)
echo.
pause

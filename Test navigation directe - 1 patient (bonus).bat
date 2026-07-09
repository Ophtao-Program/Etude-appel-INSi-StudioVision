@echo off
cd /d "%~dp0"
REM Bonus : test interactif de la navigation directe sur un patient
REM (affiche l'ancien et le nouveau patient, restaure a la fin).
python -u test_charger_patient.py

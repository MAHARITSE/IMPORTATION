@echo off
rem ============================================================
rem  MCI_CARE_paiement.bat - Paiements de la societe MCI CARE
rem  Format : DECOMPTE DE REGLEMENT FACTURES
rem  Convertit les PDF de ce dossier en :
rem      MCI CARE <MOIS> <MONTANT>.xlsx   (ex : MCI CARE Mai 471 140.xlsx)
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "MCI_CARE_paiement_to_excel.py" %*
) else (
    python "MCI_CARE_paiement_to_excel.py" %*
)
echo.
pause

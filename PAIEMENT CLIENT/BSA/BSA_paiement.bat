@echo off
rem ============================================================
rem  BSA_paiement.bat - Paiements de la societe BSA
rem  Format : RELEVE DE REMBOURSEMENTS DES FRAIS DE SANTE
rem  Convertit les PDF de ce dossier en :
rem      BSA <MOIS> <MONTANT>.xlsx   (ex : BSA Avril 928 750.xlsx)
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "BSA_paiement_to_excel.py" %*
) else (
    python "BSA_paiement_to_excel.py" %*
)
echo.
pause

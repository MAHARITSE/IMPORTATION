@echo off
rem ============================================================
rem  MCI_CARE_paiement.bat - Paiements de la societe MCI CARE
rem  Format : DECOMPTE DE REGLEMENT FACTURES
rem  Les PDF a convertir se deposent dans le sous-dossier PDF :
rem      MCI CARE\PDF\....pdf
rem  Convertit les PDF en Excel, classes par annee de paiement
rem  puis par annee des soins (un paiement de cette annee peut
rem  regler des soins de l'annee derniere) :
rem      MCI CARE\<ANNEE_PAIEMENT>\<ANNEE_SOINS>\<DATE_PAIEMENT> MCI CARE <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx
rem      ex : MCI CARE\2026\2026\02-05-26 MCI CARE 2026 02-03-26 a 31-03-26 MONTANT 471 140Ar.xlsx
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

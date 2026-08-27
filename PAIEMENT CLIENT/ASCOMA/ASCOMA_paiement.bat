@echo off
rem ============================================================
rem  ASCOMA_paiement.bat - Paiements de la societe ASCOMA
rem  Format : DECOMPTE DE REGLEMENT TIERS PAYANT
rem  Les PDF a convertir se deposent dans le sous-dossier PDF :
rem      ASCOMA\PDF\....pdf
rem  Convertit les PDF en Excel, classes par annee :
rem      PAIEMENT CLIENT\ASCOMA\<ANNEE>\<DATE_PAIEMENT> ASCOMA <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx
rem      ex : ASCOMA\2025\09-01-25 ASCOMA 2025 13-05-24 a 31-08-24 MONTANT 7 035 543Ar.xlsx
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "ASCOMA_to_excel.py" %*
) else (
    python "ASCOMA_to_excel.py" %*
)
echo.
pause

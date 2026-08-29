@echo off
rem ============================================================
rem  BSA_paiement.bat - Paiements de la societe BSA
rem  Format : RELEVE DE REMBOURSEMENTS DES FRAIS DE SANTE
rem  Les PDF a convertir se deposent dans le sous-dossier PDF :
rem      BSA\PDF\....pdf
rem  Convertit les PDF en Excel, classes par annee de paiement
rem  puis par annee des soins (un paiement de cette annee peut
rem  regler des soins de l'annee derniere) :
rem      BSA\<ANNEE_PAIEMENT>\<ANNEE_SOINS>\<DATE_PAIEMENT> BSA <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx
rem      ex : BSA\2026\2026\17-04-26 BSA 2026 27-01-26 a 23-02-26 MONTANT 928 750Ar.xlsx
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

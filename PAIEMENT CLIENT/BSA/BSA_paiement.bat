@echo off
rem ============================================================
rem  BSA_paiement.bat - Paiements de la societe BSA
rem  Format : RELEVE DE REMBOURSEMENTS DES FRAIS DE SANTE
rem  Les PDF a convertir se deposent dans le sous-dossier PDF :
rem      BSA\PDF\....pdf
rem  Les relevés Excel BSA a convertir se deposent dans :
rem      BSA\Excel\....xlsx
rem  Convertit les PDF et les relevés Excel en fichiers d'importation.
rem  Les sorties PDF restent classees par annee de paiement puis par annee
rem  des soins ; les sorties issues d'Excel sont dans BSA\Excel\Import\ :
rem      BSA\Excel\Import\<DATE_PAIEMENT> BSA <ANNEE> <PERIODE> MONTANT <MONTANT>Ar.xlsx
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")
%PY% "BSA_paiement_to_excel.py" %*
echo.
echo --- Releves Excel BSA vers le modele d'importation ---
%PY% "BSA_excel_to_modele.py" %*
echo.
pause

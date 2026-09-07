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

rem --- Choix de l'interpreteur Python --------------------------
rem On prefere le lanceur "py -3" s'il existe, sinon "python".
rem ATTENTION : ne PAS mettre ce test dans un bloc entre
rem parentheses "if ... ( ) else ( )" : %errorlevel% y serait
rem evalue avant l'execution de "where", donc toujours 0, et le
rem script ne lancerait jamais python.
where py >nul 2>nul
if %errorlevel%==0 goto :avec_py

rem --- Python sans le lanceur "py" -----------------------------
where python >nul 2>nul
if %errorlevel%==0 goto :avec_python

rem --- Aucun interpreteur trouve ------------------------------
echo.
echo  [ERREUR] Python est introuvable.
echo  Installez Python (python.org) et cochez "Add Python to PATH",
echo  puis relancez ce fichier.
echo.
pause
exit /b 1

:avec_py
py -3 "BSA_paiement_to_excel.py" %*
goto :fin

:avec_python
python "BSA_paiement_to_excel.py" %*

:fin
echo.
pause

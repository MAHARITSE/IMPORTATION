@echo off
rem ============================================================
rem  BSA_paiement.bat - Paiements (BSA, ARO, SARO, ORANGE, ...)
rem  Format : RELEVE DE REMBOURSEMENTS DES FRAIS DE SANTE
rem  Deposez les PDF a convertir dans le sous-dossier PDF\
rem  Sortie : <ANNEE_PAIEMENT>\<ANNEE_SOINS>\<DATE> SOCIETE ... .xlsx
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
setlocal EnableExtensions
chcp 65001 >nul 2>nul
cd /d "%~dp0"

set "PYEXE="
rem --- 1) lanceur officiel Windows "py"
py -3 --version >nul 2>nul && set "PYEXE=py -3"
rem --- 2) sinon python.exe du PATH (on ignore l'alias Microsoft Store)
if not defined PYEXE (
    python --version >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
    echo.
    echo !! Python est introuvable sur cet ordinateur.
    echo    Installez Python 3 depuis https://www.python.org/downloads/
    echo    en cochant "Add python.exe to PATH", puis relancez ce fichier.
    echo.
    pause
    exit /b 9009
)

if not exist "BSA_paiement_to_excel.py" (
    echo.
    echo !! Fichier BSA_paiement_to_excel.py introuvable dans :
    echo    %CD%
    echo.
    pause
    exit /b 2
)

rem --- Verification / installation des bibliotheques necessaires
%PYEXE% -c "import pdfplumber, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo Installation des bibliotheques necessaires ^(pdfplumber, openpyxl^)...
    %PYEXE% -m pip install --upgrade pip >nul 2>nul
    %PYEXE% -m pip install pdfplumber openpyxl
    %PYEXE% -c "import pdfplumber, openpyxl" >nul 2>nul
    if errorlevel 1 (
        echo.
        echo !! Installation impossible ^(pas d'acces Internet ou droits insuffisants^).
        echo    Commande a lancer manuellement :
        echo        %PYEXE% -m pip install pdfplumber openpyxl
        echo.
        pause
        exit /b 3
    )
)

echo.
%PYEXE% "BSA_paiement_to_excel.py" %*
set "CODE=%errorlevel%"
echo.
if not "%CODE%"=="0" (
    echo == Traitement termine avec des erreurs ^(code %CODE%^) ==
) else (
    echo == Traitement termine ==
)
echo Appuyez sur une touche pour fermer cette fenetre.
pause >nul
exit /b %CODE%

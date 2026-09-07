@echo off
rem -------------------------------------------------------------------
rem  Lancer_BSA.bat  -  conversion des releves BSA (PDF) en Excel
rem  Double-cliquez sur ce fichier : la fenetre reste ouverte a la fin,
rem  meme en cas d'erreur (plus de fermeture automatique).
rem -------------------------------------------------------------------
setlocal
chcp 65001 >nul
cd /d "%~dp0"
title BSA - PDF vers Excel

rem --- 1. trouver Python -------------------------------------------------
set "PYTHON="
where py >nul 2>nul && set "PYTHON=py -3"
if not defined PYTHON (
    where python >nul 2>nul && set "PYTHON=python"
)
if not defined PYTHON (
    echo.
    echo !! Python n'est pas installe sur cet ordinateur.
    echo    Telechargez-le sur https://www.python.org/downloads/ puis,
    echo    pendant l'installation, cochez "Add python.exe to PATH".
    echo.
    pause
    endlocal
    exit /b 1
)

rem --- 2. verifier / installer les bibliotheques -------------------------
%PYTHON% -c "import pdfplumber, openpyxl" >nul 2>nul
if errorlevel 1 (
    echo.
    echo Bibliotheques manquantes : installation de pdfplumber et openpyxl...
    echo ^(une seule fois, patience pendant le telechargement^)
    echo.
    %PYTHON% -m pip install --disable-pip-version-check pdfplumber openpyxl
    if errorlevel 1 (
        echo.
        echo !! Installation impossible. Essayez vous-meme cette commande :
        echo        %PYTHON% -m pip install pdfplumber openpyxl
        echo.
        pause
        endlocal
        exit /b 1
    )
    echo.
)

rem --- 3. lancer la conversion -------------------------------------------
echo.
%PYTHON% BSA_PDF_vers_Excel.py %*
echo.
pause
endlocal

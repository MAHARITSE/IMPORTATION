@echo off
rem ============================================================
rem  CONVERTIR.bat - Conversion des factures PDF en Excel
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "FACTURE CLIENT\pdf_to_excel.py" %*
) else (
    python "FACTURE CLIENT\pdf_to_excel.py" %*
)
echo.
pause

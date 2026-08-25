@echo off
rem ============================================================
rem  CONVERTIR.bat - Conversion des PDF en Excel
rem  1) FACTURE CLIENT  : factures PDF        -> Excel
rem  2) PAIEMENT CLIENT : paiements (assureurs) -> Excel
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 "FACTURE CLIENT\pdf_to_excel.py" %*
    echo.
    py -3 "PAIEMENT CLIENT\pdf_paiement_to_excel.py" %*
) else (
    python "FACTURE CLIENT\pdf_to_excel.py" %*
    echo.
    python "PAIEMENT CLIENT\pdf_paiement_to_excel.py" %*
)
echo.
pause

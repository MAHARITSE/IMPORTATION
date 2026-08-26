@echo off
rem ============================================================
rem  CONVERTIR.bat - Conversion des PDF en Excel
rem  1) FACTURE CLIENT  : factures PDF        -> Excel
rem  2) PAIEMENT CLIENT : paiements (assureurs) -> Excel
rem     Un script par societe, adapte a son format de paiement :
rem       PAIEMENT CLIENT\BSA\BSA_paiement_to_excel.py
rem       PAIEMENT CLIENT\MCI CARE\MCI_CARE_paiement_to_excel.py
rem     (double-clic possible aussi sur le .bat de chaque dossier)
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    echo --- FACTURE CLIENT ---
    py -3 "FACTURE CLIENT\pdf_to_excel.py" %*
    echo.
    echo --- PAIEMENT CLIENT : BSA ---
    py -3 "PAIEMENT CLIENT\BSA\BSA_paiement_to_excel.py" %*
    echo.
    echo --- PAIEMENT CLIENT : MCI CARE ---
    py -3 "PAIEMENT CLIENT\MCI CARE\MCI_CARE_paiement_to_excel.py" %*
) else (
    echo --- FACTURE CLIENT ---
    python "FACTURE CLIENT\pdf_to_excel.py" %*
    echo.
    echo --- PAIEMENT CLIENT : BSA ---
    python "PAIEMENT CLIENT\BSA\BSA_paiement_to_excel.py" %*
    echo.
    echo --- PAIEMENT CLIENT : MCI CARE ---
    python "PAIEMENT CLIENT\MCI CARE\MCI_CARE_paiement_to_excel.py" %*
)
echo.
pause

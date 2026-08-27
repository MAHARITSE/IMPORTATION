@echo off
rem ============================================================
rem  CONVERTIR.bat - Conversion des PDF en Excel
rem  Au lancement, tapez :
rem    1 - pour convertir les FACTURES   (dossier FACTURE CLIENT)
rem    2 - pour convertir les PAIEMENTS  (dossier PAIEMENT CLIENT :
rem        BSA, MCI CARE et ASCOMA)
rem    0 - pour annuler
rem  Puis repondez a la question O/N :
rem    O - OUI  : ecraser les Excel existants et les regenerer (--force)
rem    N - NON  : les conserver (les Excel existants ne sont jamais ecrases)
rem
rem  1) FACTURE CLIENT  : factures PDF        -> Excel
rem  2) PAIEMENT CLIENT : paiements (assureurs) -> Excel
rem     Un script par societe, adapte a son format de paiement :
rem       PAIEMENT CLIENT\BSA\BSA_paiement_to_excel.py
rem       PAIEMENT CLIENT\MCI CARE\MCI_CARE_paiement_to_excel.py
rem       PAIEMENT CLIENT\ASCOMA\ASCOMA_to_excel.py  (PDF dans ASCOMA\PDF)
rem     (double-clic possible aussi sur le .bat de chaque dossier)
rem  Double-cliquez sur ce fichier (Windows) pour lancer.
rem ============================================================
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set "PY=py -3") else (set "PY=python")

:menu
echo.
echo  ==============================================
echo    Que voulez-vous convertir ?
echo  ==============================================
echo    1  -  les FACTURES   (dossier FACTURE CLIENT)
echo    2  -  les PAIEMENTS  (dossier PAIEMENT CLIENT)
echo    0  -  annuler
echo  ==============================================
set "choix="
set /p "choix= Votre choix (1, 2 ou 0) : "
if "%choix%"=="1" goto factures
if "%choix%"=="2" goto paiements
if "%choix%"=="0" goto annuler
echo.
echo  Choix non valide : tapez 1, 2 ou 0.
goto menu

:factures
call :demande_force
echo.
echo --- FACTURE CLIENT ---
%PY% "FACTURE CLIENT\pdf_to_excel.py" %ARGS% %*
goto fin

:paiements
call :demande_force
echo.
echo --- PAIEMENT CLIENT : BSA ---
%PY% "PAIEMENT CLIENT\BSA\BSA_paiement_to_excel.py" %ARGS% %*
echo.
echo --- PAIEMENT CLIENT : MCI CARE ---
%PY% "PAIEMENT CLIENT\MCI CARE\MCI_CARE_paiement_to_excel.py" %ARGS% %*
echo.
echo --- PAIEMENT CLIENT : ASCOMA ---
%PY% "PAIEMENT CLIENT\ASCOMA\ASCOMA_to_excel.py" %ARGS% %*
goto fin

:annuler
echo.
echo Operation annulee. Aucune conversion effectuee.

:fin
echo.
pause
exit /b 0

rem ------------------------------------------------------------
rem  Sous-routine : ecraser les Excel existants pour regenerer ?
rem  Reponse O/OUI -> les scripts recoivent --force (regeneration)
rem  Reponse N/NON -> les Excel existants sont conserves
rem ------------------------------------------------------------
:demande_force
echo.
set "force="
set /p "force= Ecraser les Excel existants pour les regenerer ? (O/N) : "
if /i "%force%"=="O" goto force_oui
if /i "%force%"=="OUI" goto force_oui
if /i "%force%"=="N" goto force_non
if /i "%force%"=="NON" goto force_non
echo  Repondez O (oui) ou N (non).
goto demande_force

:force_oui
set "ARGS=--force"
echo  OUI : les Excel existants seront ECRASES et regeneres.
goto :eof

:force_non
set "ARGS="
echo  NON : les Excel existants sont conserves, jamais ecrases.
goto :eof

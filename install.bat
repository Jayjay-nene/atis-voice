@echo off
:: ATIS Voice — Installeur Windows one-click
:: Double-cliquez sur ce fichier pour installer.

echo.
echo ============================================================
echo   ATIS Voice — Installation
echo ============================================================
echo.

:: --- Verifier les droits admin ---
net session >nul 2>&1
if %errorLevel% NEQ 0 (
    echo   Ce script necessite les droits administrateur.
    echo   Relancement en mode admin...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

:: --- Verifier Python ---
python --version >nul 2>&1
if %errorLevel% NEQ 0 (
    echo   [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo.
    echo   Installez Python 3.9+ depuis https://www.python.org/downloads/
    echo   IMPORTANT : cochez "Add Python to PATH" pendant l'installation.
    echo.
    pause
    exit /b 1
)

echo   Python detecte :
python --version
echo.

:: --- Installer les dependances ---
echo   Installation des dependances...
echo.
cd /d "%~dp0"
pip install -r requirements.txt
if %errorLevel% NEQ 0 (
    echo.
    echo   [ERREUR] L'installation des dependances a echoue.
    pause
    exit /b 1
)
echo.
echo   Dependances installees.
echo.

:: --- Lancer la configuration si pas de .env ---
if not exist "%~dp0.env" (
    echo   Premier lancement : configuration...
    echo.
    python voice_command.py --setup
    echo.
)

:: --- Creer la tache planifiee (demarrage auto) ---
echo   Configuration du demarrage automatique...

set TASK_NAME=ATIS Voice
set PYTHON_EXE=
for /f "delims=" %%i in ('python -c "import sys; print(sys.executable)"') do set PYTHON_EXE=%%i
set SCRIPT_PATH=%~dp0voice_command.py
set WORK_DIR=%~dp0

:: Supprimer l'ancienne tache si elle existe
schtasks /query /tn "%TASK_NAME%" >nul 2>&1
if %errorLevel% EQU 0 (
    schtasks /delete /tn "%TASK_NAME%" /f >nul 2>&1
)

:: Creer la nouvelle tache
schtasks /create /tn "%TASK_NAME%" /tr "\"%PYTHON_EXE%\" \"%SCRIPT_PATH%\"" /sc onlogon /rl highest /f >nul 2>&1
if %errorLevel% EQU 0 (
    echo   Demarrage automatique configure.
) else (
    echo   [ATTENTION] Impossible de creer la tache planifiee.
    echo   Vous pouvez lancer manuellement avec : python voice_command.py
)

echo.
echo ============================================================
echo   Installation terminee !
echo.
echo   Pour lancer maintenant :
echo     python voice_command.py
echo.
echo   L'outil se lancera automatiquement au demarrage de Windows.
echo   Pour reconfigurer : python voice_command.py --setup
echo ============================================================
echo.

:: --- Lancer immediatement ---
set /p LAUNCH="  Lancer ATIS Voice maintenant ? (O/n) : "
if /i "%LAUNCH%"=="n" goto :end
if /i "%LAUNCH%"=="N" goto :end

echo.
echo   Lancement...
start "" "%PYTHON_EXE%" "%SCRIPT_PATH%"
echo   ATIS Voice demarre. Vous pouvez fermer cette fenetre.

:end
echo.
pause

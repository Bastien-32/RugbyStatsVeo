@echo off
setlocal enabledelayedexpansion

rem ==========================================================
rem  VeoVideoControl - construction du moteur Windows
rem
rem  A lancer sur un VRAI PC Windows Intel/AMD 64 bits.
rem
rem  Un Windows ARM64, y compris celui d'une machine
rem  virtuelle sur Mac Apple Silicon, produit un executable
rem  ARM64 qui ne demarre pas sur les postes du club. Le
rem  script refuse donc de tourner dans ce cas.
rem
rem  Double-cliquer sur ce fichier suffit. A la fin, le
rem  dossier VeoVideoControlEngine est remplace par le
rem  moteur fraichement construit, l'ancien etant conserve
rem  sous le nom VeoVideoControlEngine-precedent.
rem ==========================================================

cd /d "%~dp0"

echo.
echo ==========================================================
echo   VeoVideoControl : construction du moteur Windows
echo ==========================================================
echo.

rem ----------------------------------------------------------
rem  1. Python
rem ----------------------------------------------------------

set PY_LANCEUR=py
%PY_LANCEUR% --version >nul 2>&1
if errorlevel 1 (
    set PY_LANCEUR=python
    !PY_LANCEUR! --version >nul 2>&1
    if errorlevel 1 (
        echo [ERREUR] Python est introuvable.
        echo.
        echo Installer Python 3 depuis python.org en cochant
        echo "Add python.exe to PATH", puis relancer ce script.
        echo.
        pause
        exit /b 1
    )
)

echo Python utilise :
%PY_LANCEUR% --version
echo.

rem ----------------------------------------------------------
rem  2. Architecture : le piege a eviter
rem ----------------------------------------------------------

for /f %%A in ('%PY_LANCEUR% -c "import platform; print(platform.machine())"') do set ARCHI=%%A

echo Architecture detectee : %ARCHI%

if /i not "%ARCHI%"=="AMD64" (
    echo.
    echo [ERREUR] Ce Python n'est pas en 64 bits Intel/AMD.
    echo.
    echo PyInstaller produit un executable pour l'architecture
    echo du Python qui le fait tourner. Construit ailleurs
    echo qu'en AMD64, le moteur ne demarrera pas sur les
    echo postes du club.
    echo.
    echo Utiliser un vrai PC Windows Intel ou AMD, avec un
    echo Python 64 bits.
    echo.
    pause
    exit /b 1
)

echo.

rem ----------------------------------------------------------
rem  3. Environnement isole
rem ----------------------------------------------------------

if not exist ".venv-windows\Scripts\python.exe" (
    echo Creation de l'environnement Python...
    %PY_LANCEUR% -m venv .venv-windows
    if errorlevel 1 (
        echo [ERREUR] La creation de l'environnement a echoue.
        pause
        exit /b 1
    )
)

set PYTHON=.venv-windows\Scripts\python.exe

echo Installation des dependances...
echo.

"%PYTHON%" -m pip install --upgrade pip --quiet
"%PYTHON%" -m pip install -r requirements-windows.txt --quiet
if errorlevel 1 (
    echo [ERREUR] L'installation des dependances a echoue.
    pause
    exit /b 1
)

rem ----------------------------------------------------------
rem  4. Construction
rem ----------------------------------------------------------

echo.
echo Construction en cours, compter une a deux minutes...
echo.

"%PYTHON%" -m PyInstaller --noconfirm --clean VeoVideoControl.spec
if errorlevel 1 (
    echo.
    echo [ERREUR] La construction a echoue.
    pause
    exit /b 1
)

if not exist "dist\VeoVideoControl\VeoVideoControl.exe" (
    echo.
    echo [ERREUR] L'executable attendu n'a pas ete produit.
    echo Cherche : dist\VeoVideoControl\VeoVideoControl.exe
    pause
    exit /b 1
)

rem ----------------------------------------------------------
rem  5. Mise en place
rem
rem  L'ancien moteur est mis de cote plutot que supprime :
rem  s'il faut revenir en arriere, il est encore la.
rem ----------------------------------------------------------

set CIBLE=..\VeoVideoControlEngine
set PRECEDENT=..\VeoVideoControlEngine-precedent

if exist "%PRECEDENT%" rmdir /s /q "%PRECEDENT%"
if exist "%CIBLE%" move "%CIBLE%" "%PRECEDENT%" >nul

xcopy /e /i /y /q "dist\VeoVideoControl" "%CIBLE%" >nul
if errorlevel 1 (
    echo.
    echo [ERREUR] La copie du moteur a echoue.
    pause
    exit /b 1
)

rem ----------------------------------------------------------
rem  6. Bilan
rem ----------------------------------------------------------

echo.
echo ==========================================================
echo   Moteur Windows construit.
echo ==========================================================
echo.
echo   Nouveau  : VeoVideoControl\VeoVideoControlEngine
echo   Ancien   : VeoVideoControl\VeoVideoControlEngine-precedent
echo.
echo   A faire ensuite :
echo     1. tester le bouton "Connecter lecteur video"
echo        depuis un fichier de match ;
echo     2. tester la lecture arriere continue ;
echo     3. rapporter le dossier VeoVideoControlEngine
echo        sur le Mac pour le versionner.
echo.
pause

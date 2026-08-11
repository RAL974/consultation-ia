@echo off
setlocal
cd /d "%~dp0"

set "JOURNAL=%~dp0lancement.log"

echo ============================================================> "%JOURNAL%"
echo Consultation AI - journal de lancement>> "%JOURNAL%"
echo Date : %DATE% %TIME%>> "%JOURNAL%"
echo Dossier : %CD%>> "%JOURNAL%"
echo ============================================================>> "%JOURNAL%"

echo ============================================================
echo   CONSULTATION AI - demarrage
echo ============================================================
echo Dossier : %CD%
echo.

if not exist "gui.py" goto PAS_DE_GUI

rem Le lanceur "py" (installe avec Python depuis python.org) est prefere :
rem il evite le raccourci Microsoft Store qui ne fait rien quand on double-clique.
set "PY_CMD="

py -3 --version >nul 2>&1
if not errorlevel 1 set "PY_CMD=py -3"

if not defined PY_CMD goto ESSAI_PYTHON
goto LANCER

:ESSAI_PYTHON
python --version >nul 2>&1
if not errorlevel 1 set "PY_CMD=python"

if not defined PY_CMD goto PAS_DE_PYTHON

:LANCER
echo Interpreteur utilise : %PY_CMD%
echo Interpreteur utilise : %PY_CMD%>> "%JOURNAL%"
%PY_CMD% -c "import sys; print('Chemin exact :', sys.executable)"
%PY_CMD% -c "import sys; print('Chemin exact :', sys.executable)">> "%JOURNAL%" 2>&1
echo.
echo Ouverture de la fenetre... (garde cette console ouverte)
echo.

%PY_CMD% gui.py
set "CODE=%ERRORLEVEL%"

echo.>> "%JOURNAL%"
echo Code de retour de gui.py : %CODE%>> "%JOURNAL%"

echo.
echo ------------------------------------------------------------
if not "%CODE%"=="0" goto ERREUR
echo Termine normalement.
goto FIN

:ERREUR
echo Le programme s'est arrete avec une erreur (code %CODE%).
echo Consulte le fichier  gui_erreur.log  a cote de ce lanceur.
goto FIN

:PAS_DE_GUI
echo /!\ Fichier gui.py introuvable dans ce dossier.
echo     Ce lanceur doit rester a cote de gui.py.
echo /!\ gui.py introuvable dans %CD%>> "%JOURNAL%"
goto FIN

:PAS_DE_PYTHON
echo /!\ Python n'a pas ete trouve sur ce poste.
echo     Installe Python depuis python.org en cochant
echo     "Add python.exe to PATH", puis relance ce fichier.
echo /!\ Aucun interpreteur Python trouve>> "%JOURNAL%"
goto FIN

:FIN
echo ------------------------------------------------------------
echo Un journal a ete ecrit ici : %JOURNAL%
echo.
echo Appuie sur une touche pour fermer cette fenetre.
pause >nul
endlocal

@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   Gutachten-Assistent - Einrichtung
echo ============================================
echo.
set "PYEXE="
set "PYW="
where py >nul 2>&1
if %errorlevel%==0 (py -3 -c "import sys" >nul 2>&1 && (set "PYEXE=py -3" & set "PYW=pyw -3"))
if defined PYEXE goto haspy
where python >nul 2>&1
if %errorlevel%==0 (python -c "import sys;sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1 && (set "PYEXE=python" & set "PYW=pythonw"))
if defined PYEXE goto haspy
echo Python ist noch nicht installiert. Es wird jetzt automatisch installiert.
echo Windows fragt eventuell um Erlaubnis - bitte mit "Ja" bestaetigen.
echo.
winget install -e --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
for %%P in ("%LocalAppData%\Programs\Python\Python312\python.exe" "%ProgramFiles%\Python312\python.exe") do if exist %%P (set "PYEXE=%%~P" & set "PYW=%%~dpPpythonw.exe")
if defined PYEXE (set PYEXE="%PYEXE%"& set PYW="%PYW%"& goto haspy)
echo.
echo Die automatische Installation von Python hat nicht funktioniert.
echo Bitte Python manuell installieren: Die Download-Seite oeffnet sich jetzt.
echo Wichtig: Beim Installieren unten "Add python.exe to PATH" anhaken. Danach diese Datei erneut doppelklicken.
start https://www.python.org/downloads/windows/
pause
exit /b 1
:haspy
echo Python gefunden. Die noetigen Programme werden installiert (einmalig, kann einige Minuten dauern) ...
%PYEXE% -m pip install --upgrade pip >nul 2>&1
%PYEXE% -m pip install --upgrade python-docx openpyxl pymupdf
if errorlevel 1 (
  echo.
  echo Die Installation ist fehlgeschlagen. Bitte pruefen, ob eine Internetverbindung besteht, und die Datei erneut doppelklicken.
  pause
  exit /b 1
)
echo.
echo Fertig. Der Assistent startet jetzt in deinem Browser.
echo Weitere Bausteine (Texterkennung, Transkription) richtest du dort im Schritt "Programme" ein.
start "" %PYW% "programm\web.py"
timeout /t 3 >nul
exit /b 0

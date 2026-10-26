@echo off
cd /d "%~dp0"
where pyw >nul 2>&1
if %errorlevel%==0 (start "" pyw "programm\web.py" & exit /b)
where pythonw >nul 2>&1
if %errorlevel%==0 (start "" pythonw "programm\web.py" & exit /b)
echo Python wurde nicht gefunden. Bitte zuerst "Einrichten-Windows.bat" doppelklicken.
pause

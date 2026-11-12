@echo off
rem Startet Bot + Dashboard (Python-Version) und startet sie nach einem Absturz automatisch neu.
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (set PY=.venv\Scripts\python.exe) else (set PY=py -3)
rem Dashboard im Browser oeffnen, sobald der Server laeuft
start "" /min cmd /c "timeout /t 5 /nobreak >nul & start http://localhost:5000"
:loop
echo [%date% %time%] Starte E-Mail-Bot...
%PY% -m emailbot %*
if errorlevel 1 (
  echo Bot wurde mit Fehler beendet - siehe email_processor.log. Neustart in 30 Sekunden ^(Strg+C zum Abbrechen^)...
  timeout /t 30 /nobreak >nul
  goto loop
)

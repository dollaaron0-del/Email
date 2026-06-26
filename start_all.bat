@echo off
cd /d "%~dp0"
echo Starte Email Bot...
start "Email Bot" powershell.exe -ExecutionPolicy Bypass -NoExit -File "%~dp0bot.ps1"
timeout /t 2 /nobreak >nul
echo Starte Dashboard...
start "Dashboard Server" powershell.exe -ExecutionPolicy Bypass -NoExit -File "%~dp0dashboard_server.ps1"
timeout /t 3 /nobreak >nul
start http://localhost:5000
echo.
echo Beide Fenster offen lassen!
echo Dashboard: http://localhost:5000

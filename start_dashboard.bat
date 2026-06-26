@echo off
cd /d "%~dp0"
start "Dashboard Server" powershell.exe -ExecutionPolicy Bypass -NoExit -File "%~dp0dashboard_server.ps1"
timeout /t 2 /nobreak >nul
start http://localhost:5000

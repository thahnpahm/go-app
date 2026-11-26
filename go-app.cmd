@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0go-app.ps1" %*
pause

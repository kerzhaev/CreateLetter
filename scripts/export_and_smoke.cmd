@echo off
setlocal
cd /d "%~dp0.."
powershell -ExecutionPolicy Bypass -File ".\scripts\export_and_smoke.ps1" %*
exit /b %errorlevel%

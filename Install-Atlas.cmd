@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Atlas.ps1"
echo.
if errorlevel 1 (echo Installation did not finish. Read the message above.) else (echo Atlas installed. You can open Excel now.)
pause

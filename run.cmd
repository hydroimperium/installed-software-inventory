@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Get-InstalledSoftware.ps1" -Open %*
if errorlevel 1 pause
endlocal

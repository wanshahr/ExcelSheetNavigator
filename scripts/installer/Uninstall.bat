@echo off
title Excel Sheet Navigator - Uninstall
echo Starting the Excel Sheet Navigator uninstaller...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall.ps1"
echo.
pause

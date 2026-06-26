@echo off
title Excel Sheet Navigator - Install
echo Starting the Excel Sheet Navigator installer...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install.ps1"
echo.
pause

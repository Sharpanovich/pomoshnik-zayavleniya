@echo off
title Sborka programmy
echo.
echo  ========================================================
echo   SBORKA "Pomoschnik po zayavleniyam" iz ishodnikov
echo  ========================================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0build.ps1"
echo.
pause

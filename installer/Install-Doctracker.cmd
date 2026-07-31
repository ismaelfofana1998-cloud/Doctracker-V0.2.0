@echo off
setlocal
title Installation de Doctracker

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-Doctracker.ps1"
set "DOCTRACKER_EXIT=%ERRORLEVEL%"

if not "%DOCTRACKER_EXIT%"=="0" (
  echo.
  echo L'installation de Doctracker n'a pas abouti.
  pause
)

exit /b %DOCTRACKER_EXIT%

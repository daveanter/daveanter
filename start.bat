@echo off
title Shift Scheduler
echo.
echo  Starting Shift Scheduler...
echo  (If other PCs can't reach the form, right-click this file and choose
echo   "Run as administrator")
echo.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0server.ps1"
echo.
echo  Server stopped.
pause

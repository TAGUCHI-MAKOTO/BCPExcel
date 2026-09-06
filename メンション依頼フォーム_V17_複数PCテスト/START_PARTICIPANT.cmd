@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Mention Request Form v26 - V17 Participant
echo ============================================================
echo  V17 Multi-PC Test - Participant
echo ============================================================
echo.
echo Enter the SAME shared parent folder and Session ID
echo shown on the coordinator PC.
echo.
set "SHARED_PARENT="
set /p "SHARED_PARENT=Shared parent folder: "
if not defined SHARED_PARENT exit /b

set "SESSION_ID="
set /p "SESSION_ID=Session ID: "
if not defined SESSION_ID exit /b

:PID
set "PARTICIPANT_ID="
set /p "PARTICIPANT_ID=Participant ID (PC01 etc.): "
if not defined PARTICIPANT_ID goto PID

echo %PARTICIPANT_ID%| findstr /R /X "[A-Za-z0-9_-][A-Za-z0-9_-]*" >nul
if errorlevel 1 (
  echo Use only A-Z, a-z, 0-9, _ or -.
  goto PID
)

call "%~dp0participant_core.cmd" "%SHARED_PARENT%" "%SESSION_ID%" "%PARTICIPANT_ID%" "INTERACTIVE"
echo.
pause

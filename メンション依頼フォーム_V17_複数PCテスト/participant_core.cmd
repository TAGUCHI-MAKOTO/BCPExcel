@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Mention Request Form v26 - V17 Participant

set "SHARED_PARENT=%~1"
set "SESSION_ID=%~2"
set "PARTICIPANT_ID=%~3"
set "AUTO_MODE=%~4"

if not defined SHARED_PARENT exit /b 2
if not defined SESSION_ID exit /b 2
if not defined PARTICIPANT_ID exit /b 2

set "TEST_ROOT=%SHARED_PARENT%\MentionRequest_V17_MultiPC_Test"
set "SESSION_ROOT=%TEST_ROOT%\sessions\%SESSION_ID%"
set "CONFIG=%SESSION_ROOT%\config.txt"

if not exist "%CONFIG%" (
  echo Session config not found:
  echo %CONFIG%
  exit /b 3
)

for /f "usebackq tokens=1,* delims==" %%A in ("%CONFIG%") do (
  set "CFG_%%A=%%B"
)

if /I not "!CFG_SESSION_ID!"=="%SESSION_ID%" (
  echo Session ID mismatch.
  exit /b 4
)

set "CSV_FOLDER=%SESSION_ROOT%\csv"
set "READY_FILE=%SESSION_ROOT%\ready\%PARTICIPANT_ID%.ready"
set "DONE_FILE=%SESSION_ROOT%\done\%PARTICIPANT_ID%.done"
set "REPORT_FILE=%SESSION_ROOT%\reports\%PARTICIPANT_ID%.txt"
set "MANIFEST_FILE=%SESSION_ROOT%\manifests\%PARTICIPANT_ID%.txt"

if exist "%READY_FILE%" (
  echo.
  echo Participant ID "%PARTICIPANT_ID%" is already in use.
  echo Use another ID such as PC01, PC02, PC03.
  exit /b 5
)

set "LOCAL_ROOT=%~dp0local_results\%SESSION_ID%\%PARTICIPANT_ID%"
set "PENDING_DIR=%LOCAL_ROOT%\Pending_TestOnly"
set "RESULT_DIR=%LOCAL_ROOT%\worker_results"

mkdir "%LOCAL_ROOT%" >nul 2>&1
mkdir "%PENDING_DIR%" >nul 2>&1
mkdir "%RESULT_DIR%" >nul 2>&1

echo.
echo ============================================================
echo  V17 Participant
echo ============================================================
echo Session     : %SESSION_ID%
echo Participant : %PARTICIPANT_ID%
echo Scenario    : !CFG_SCENARIO!
echo Packages    : !CFG_PACKAGES_PER_PARTICIPANT!
echo Rows/Pending: !CFG_ROWS_PER_PENDING!
echo Environment : !CFG_ENVIRONMENT!
echo ============================================================
echo.

cscript.exe //nologo "%~dp0create_v17_pending.js" "%PENDING_DIR%" "%CSV_FOLDER%" "%PARTICIPANT_ID%" "%SESSION_ID%" "!CFG_SCENARIO!" "!CFG_PACKAGES_PER_PARTICIPANT!" "!CFG_ROWS_PER_PENDING!" "!CFG_CREATED!" "%MANIFEST_FILE%"
if errorlevel 1 (
  echo Pending creation failed.
  exit /b 6
)

cscript.exe //nologo "%~dp0write_marker.js" "%READY_FILE%" "%COMPUTERNAME%|%PARTICIPANT_ID%"
echo READY - waiting for coordinator START...

cscript.exe //nologo "%~dp0wait_for_file.js" "%SESSION_ROOT%\control\start.flag" "1200"
if errorlevel 1 (
  echo START wait timeout.
  cscript.exe //nologo "%~dp0write_marker.js" "%DONE_FILE%" "START_TIMEOUT"
  exit /b 7
)

for /f "delims=" %%T in ('cscript.exe //nologo "%~dp0timestamp.js"') do set "STARTED=%%T"

echo START detected. Launching Workers...

for /L %%I in (1,1,!CFG_PACKAGES_PER_PARTICIPANT!) do (
  set "SHARED_FLAG=0"
  if /I "!CFG_SCENARIO!"=="DUPLICATE" set "SHARED_FLAG=1"
  if /I "!CFG_SCENARIO!"=="MIXED" (
    set /a ODD=%%I%%2
    if "!ODD!"=="1" set "SHARED_FLAG=1"
  )

  if "!SHARED_FLAG!"=="1" (
    set "RID=V17-%SESSION_ID%-SHARED-%%I"
  ) else (
    set "RID=V17-%SESSION_ID%-%PARTICIPANT_ID%-%%I"
  )

  set "PFILE=%PENDING_DIR%\!RID!.pending"
  set "RFILE=%RESULT_DIR%\%%I.txt"
  cscript.exe //nologo "%~dp0launch_worker.js" "!PFILE!" "%CSV_FOLDER%" "!RFILE!"
)

cscript.exe //nologo "%~dp0wait_for_count.js" "%RESULT_DIR%" "!CFG_PACKAGES_PER_PARTICIPANT!" ".txt" "360"
set "WAIT_RC=!ERRORLEVEL!"

cscript.exe //nologo "%~dp0participant_summary.js" "%RESULT_DIR%" "%PENDING_DIR%" "%REPORT_FILE%" "%PARTICIPANT_ID%" "%SESSION_ID%" "!CFG_PACKAGES_PER_PARTICIPANT!" "!STARTED!"
set "SUMMARY_RC=!ERRORLEVEL!"

if not "!SUMMARY_RC!"=="0" (
  echo.
  echo Participant summary creation FAILED. RC=!SUMMARY_RC!
  exit /b 8
)

if not exist "%REPORT_FILE%" (
  echo.
  echo Participant summary file was not created.
  exit /b 8
)

cscript.exe //nologo "%~dp0write_marker.js" "%DONE_FILE%" "%COMPUTERNAME%|%PARTICIPANT_ID%|RC=!WAIT_RC!"

echo.
echo Participant complete.
echo Report:
echo %REPORT_FILE%
exit /b 0

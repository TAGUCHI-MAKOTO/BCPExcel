@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Mention Request Form v26 - V17 Local Simulation

echo ============================================================
echo  V17 LOCAL 1-PC SIMULATION
echo ============================================================
echo.
echo This runs multiple virtual participants on ONE PC.
echo It checks V17 coordination, locking/dedupe logic and summary flow.
echo It does NOT prove SMB locking between physical PCs.
echo.

:PARTS
set "EXPECTED="
set /p "EXPECTED=Virtual participants 2-8 [3]: "
if not defined EXPECTED set "EXPECTED=3"
for /f "delims=0123456789" %%A in ("!EXPECTED!") do goto PARTS_BAD
if !EXPECTED! LSS 2 goto PARTS_BAD
if !EXPECTED! GTR 8 goto PARTS_BAD
goto SCENARIO
:PARTS_BAD
echo Enter 2-8.
goto PARTS

:SCENARIO
echo.
echo Scenario:
echo   1 = UNIQUE
echo   2 = DUPLICATE
echo   3 = MIXED
set "SCENARIO_NO="
set /p "SCENARIO_NO=Select [3]: "
if not defined SCENARIO_NO set "SCENARIO_NO=3"
if "!SCENARIO_NO!"=="1" set "SCENARIO=UNIQUE"&goto PACKAGES
if "!SCENARIO_NO!"=="2" set "SCENARIO=DUPLICATE"&goto PACKAGES
if "!SCENARIO_NO!"=="3" set "SCENARIO=MIXED"&goto PACKAGES
goto SCENARIO

:PACKAGES
set "PACKAGES="
set /p "PACKAGES=Pending packages per participant 1-10 [3]: "
if not defined PACKAGES set "PACKAGES=3"
for /f "delims=0123456789" %%A in ("!PACKAGES!") do goto PACKAGES_BAD
if !PACKAGES! LSS 1 goto PACKAGES_BAD
if !PACKAGES! GTR 10 goto PACKAGES_BAD
goto ROWS
:PACKAGES_BAD
echo Enter 1-10.
goto PACKAGES

:ROWS
set "ROWS="
set /p "ROWS=Requests in each Pending 1-3 [3]: "
if not defined ROWS set "ROWS=3"
if "!ROWS!"=="1" goto MAKE
if "!ROWS!"=="2" goto MAKE
if "!ROWS!"=="3" goto MAKE
goto ROWS

:MAKE
for /f "delims=" %%S in ('cscript.exe //nologo "%~dp0make_session_id.js"') do set "SESSION_ID=%%S"

set "SHARED_PARENT=%~dp0LOCAL_SIM_RESULTS\!SESSION_ID!"
set "TEST_ROOT=!SHARED_PARENT!\MentionRequest_V17_MultiPC_Test"
set "SESSION_ROOT=!TEST_ROOT!\sessions\!SESSION_ID!"

mkdir "!TEST_ROOT!\sessions" >nul 2>&1
cscript.exe //nologo "%~dp0init_session.js" "!SESSION_ROOT!" "!SESSION_ID!" "LOCAL_SIM" "!EXPECTED!" "!SCENARIO!" "!PACKAGES!" "!ROWS!"
if errorlevel 1 (
  echo Session init failed.
  exit /b
)

echo.
echo Launching !EXPECTED! virtual participants...

for /L %%I in (1,1,!EXPECTED!) do (
  set "NN=0%%I"
  set "PID=SIM!NN:~-2!"
  cscript.exe //nologo "%~dp0launch_participant.js" "%~dp0participant_core.cmd" "!SHARED_PARENT!" "!SESSION_ID!" "!PID!"
)

echo Waiting for READY...
set /a READY_WAIT=0
set /a READY_LAST=-1

:WAIT_READY
for /f %%C in ('cscript.exe //nologo "%~dp0count_files.js" "!SESSION_ROOT!\ready" ".ready"') do set "READY_COUNT=%%C"
if not "!READY_COUNT!"=="!READY_LAST!" (
  echo READY: !READY_COUNT!/!EXPECTED!
  set "READY_LAST=!READY_COUNT!"
)
if !READY_COUNT! GEQ !EXPECTED! goto GO
set /a READY_WAIT+=1
if !READY_WAIT! GEQ 120 (
  echo Local READY timeout.
  exit /b
)
>nul timeout /t 1 /nobreak
goto WAIT_READY

:GO
echo All virtual participants READY.
echo START in 2 seconds...
>nul timeout /t 2 /nobreak
cscript.exe //nologo "%~dp0write_marker.js" "!SESSION_ROOT!\control\start.flag" "START"

set /a DONE_WAIT=0
set /a DONE_LAST=-1

:WAIT_DONE
for /f %%C in ('cscript.exe //nologo "%~dp0count_files.js" "!SESSION_ROOT!\done" ".done"') do set "DONE_COUNT=%%C"
if not "!DONE_COUNT!"=="!DONE_LAST!" (
  echo DONE: !DONE_COUNT!/!EXPECTED!
  set "DONE_LAST=!DONE_COUNT!"
)
if !DONE_COUNT! GEQ !EXPECTED! goto SUMMARY
set /a DONE_WAIT+=1
if !DONE_WAIT! GEQ 420 (
  echo Local DONE timeout.
  goto SUMMARY
)
>nul timeout /t 1 /nobreak
goto WAIT_DONE

:SUMMARY
echo.
cscript.exe //nologo "%~dp0host_summary.js" "!SESSION_ROOT!" "!SESSION_ROOT!\V17_RESULT.txt"
echo.
echo RESULT:
echo !SESSION_ROOT!\V17_RESULT.txt
start "" notepad.exe "!SESSION_ROOT!\V17_RESULT.txt"
echo.
pause

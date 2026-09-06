@echo off
setlocal EnableExtensions EnableDelayedExpansion
title Mention Request Form v26 - V17 Coordinator

echo ============================================================
echo  V17 Multi-PC Shared Folder Test - Coordinator
echo ============================================================
echo.
echo IMPORTANT:
echo Choose a shared TEST parent folder accessible from every PC.
echo A child folder named MentionRequest_V17_MultiPC_Test is created.
echo Production CSV files are not used.
echo UNC paths are recommended.
echo.

set "SHARED_PARENT="
set /p "SHARED_PARENT=Shared TEST parent folder: "
if not defined SHARED_PARENT exit /b

set "TEST_ROOT=!SHARED_PARENT!\MentionRequest_V17_MultiPC_Test"
mkdir "!TEST_ROOT!" >nul 2>&1
if not exist "!TEST_ROOT!" (
  echo Cannot create/access test root:
  echo !TEST_ROOT!
  exit /b
)

:PARTICIPANTS
set "EXPECTED="
set /p "EXPECTED=Number of participant PCs 2-10 [3]: "
if not defined EXPECTED set "EXPECTED=3"
for /f "delims=0123456789" %%A in ("!EXPECTED!") do goto PARTICIPANTS_BAD
if !EXPECTED! LSS 2 goto PARTICIPANTS_BAD
if !EXPECTED! GTR 10 goto PARTICIPANTS_BAD
goto SCENARIO
:PARTICIPANTS_BAD
echo Enter 2-10.
goto PARTICIPANTS

:SCENARIO
echo.
echo Scenario:
echo   1 = UNIQUE    : every PC sends different RequestIDs
echo   2 = DUPLICATE : every PC sends the SAME RequestIDs
echo   3 = MIXED     : odd packages shared, even packages unique
set "SCENARIO_NO="
set /p "SCENARIO_NO=Select [3]: "
if not defined SCENARIO_NO set "SCENARIO_NO=3"
if "!SCENARIO_NO!"=="1" set "SCENARIO=UNIQUE"&goto PACKAGES
if "!SCENARIO_NO!"=="2" set "SCENARIO=DUPLICATE"&goto PACKAGES
if "!SCENARIO_NO!"=="3" set "SCENARIO=MIXED"&goto PACKAGES
echo Enter 1, 2, or 3.
goto SCENARIO

:PACKAGES
echo.
set "PACKAGES="
set /p "PACKAGES=Pending packages per PC 1-10 [3]: "
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
echo Enter 1, 2, or 3.
goto ROWS

:MAKE
for /f "delims=" %%S in ('cscript.exe //nologo "%~dp0make_session_id.js"') do set "SESSION_ID=%%S"
set "SESSION_ROOT=!TEST_ROOT!\sessions\!SESSION_ID!"

mkdir "!TEST_ROOT!\sessions" >nul 2>&1
cscript.exe //nologo "%~dp0init_session.js" "!SESSION_ROOT!" "!SESSION_ID!" "MULTI_PC" "!EXPECTED!" "!SCENARIO!" "!PACKAGES!" "!ROWS!"
if errorlevel 1 (
  echo Session initialization failed.
  exit /b
)

echo.
echo ============================================================
echo  SESSION CREATED
echo ============================================================
echo Shared parent : !SHARED_PARENT!
echo Session ID    : !SESSION_ID!
echo Participants  : !EXPECTED!
echo Scenario      : !SCENARIO!
echo Packages / PC : !PACKAGES!
echo Rows/Pending  : !ROWS!
echo ============================================================
echo.
echo On each participant PC:
echo   1. Copy/extract this V17 package locally.
echo   2. Run START_PARTICIPANT.cmd
echo   3. Enter the Shared parent and Session ID shown above.
echo   4. Use unique IDs such as PC01, PC02, PC03.
echo.
echo Waiting for READY participants...

set /a READY_WAIT=0
set /a READY_LAST=-1

:WAIT_READY
for /f %%C in ('cscript.exe //nologo "%~dp0count_files.js" "!SESSION_ROOT!\ready" ".ready"') do set "READY_COUNT=%%C"
if not "!READY_COUNT!"=="!READY_LAST!" (
  echo READY: !READY_COUNT!/!EXPECTED!
  set "READY_LAST=!READY_COUNT!"
)
if !READY_COUNT! GEQ !EXPECTED! goto ALL_READY
set /a READY_WAIT+=1
if !READY_WAIT! GEQ 1200 (
  echo READY timeout after 20 minutes.
  exit /b
)
>nul timeout /t 1 /nobreak
goto WAIT_READY

:ALL_READY
echo.
echo All participants READY.
echo START in 3 seconds...
>nul timeout /t 3 /nobreak
cscript.exe //nologo "%~dp0write_marker.js" "!SESSION_ROOT!\control\start.flag" "START"
echo START signal sent.
echo.
echo Waiting for participant completion...

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
if !DONE_WAIT! GEQ 900 (
  echo DONE timeout after 15 minutes.
  goto SUMMARY
)
>nul timeout /t 1 /nobreak
goto WAIT_DONE

:SUMMARY
echo.
echo Creating V17 summary...
cscript.exe //nologo "%~dp0host_summary.js" "!SESSION_ROOT!" "!SESSION_ROOT!\V17_RESULT.txt"
echo.
echo Result:
echo !SESSION_ROOT!\V17_RESULT.txt
start "" notepad.exe "!SESSION_ROOT!\V17_RESULT.txt"
echo.
pause

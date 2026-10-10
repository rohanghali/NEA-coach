@echo off
REM ============================================================
REM  NEA Coach - double-click this file to start.
REM
REM  What it does:
REM    1. Finds Node.js: the copy in the app folder, one on PATH, or
REM       one in the standard Windows install locations. If none is
REM       found, it explains how to install Node (one time).
REM    2. Starts the NEA Coach server.
REM    3. Opens your browser to it.
REM
REM  It never modifies your PowerPoint file. See README.md.
REM
REM  To stop the app: close this window, or press Control-C in it.
REM ============================================================

REM UTF-8 output, so a user name or folder with accented characters in it
REM does not turn into gibberish in this window.
chcp 65001 >nul

setlocal EnableExtensions

REM Always work relative to wherever this file actually lives, so it works from
REM the Desktop, a USB stick, or a folder whose name has spaces in it. %~dp0 is
REM the folder this script is in, and it always ends with a backslash; the
REM trailing dot keeps that backslash away from the closing quote, which
REM cmd.exe can otherwise mis-read as an escaped quote.
REM (Written as a goto rather than a bracketed block on purpose: a folder name
REM containing brackets, such as "NEA Coach (1)", would close the block early.)
pushd "%~dp0."
if not errorlevel 1 goto :in_folder
echo   Could not open the app folder itself:
echo   %~dp0
goto :finish
:in_folder

REM Mirror server.js and the Mac launcher: NEA_COACH_PORT wins if it is set,
REM otherwise 4173. An empty value counts as "not set", which is what makes
REM the "if not defined" test below work.
set "PORT=%NEA_COACH_PORT%"
if not defined PORT set "PORT=4173"

echo.
echo   NEA Coach - portfolio reviewer
echo   ------------------------------------------------------------
echo.

REM ------------------------------------------------------------ 1. find Node
set "NODE_BIN="

REM a) The copy inside the app folder, put there by an earlier run or by hand.
REM    On Windows that file is runtime\node.exe (the Mac copy is runtime\node).
if exist "%~dp0runtime\node.exe" set "NODE_BIN=%~dp0runtime\node.exe"

REM b) A Node already installed on this PC and visible on PATH.
if not defined NODE_BIN (
  where node >nul 2>nul
  if not errorlevel 1 set "NODE_BIN=node"
)

REM c) The standard install locations. This matters because a Node installed a
REM    minute ago is NOT on PATH until Windows is restarted, so without this
REM    check a student would be told Node is missing straight after installing
REM    it. The quotes around %ProgramFiles% are essential: that path contains a
REM    space, and on 32-bit Windows it also contains brackets.
if not defined NODE_BIN if exist "%ProgramFiles%\nodejs\node.exe" set "NODE_BIN=%ProgramFiles%\nodejs\node.exe"
if not defined NODE_BIN if exist "%LOCALAPPDATA%\Programs\nodejs\node.exe" set "NODE_BIN=%LOCALAPPDATA%\Programs\nodejs\node.exe"

REM Now prove that whatever was found actually runs and is new enough (the app
REM needs Node 18 or newer). Skipping this gives a confusing syntax error part
REM way through a scan instead of a clear message here.
if not defined NODE_BIN goto :skip_node_version
"%NODE_BIN%" -e "process.exit(process.versions.node.split('.')[0]>=18?0:1)" >nul 2>nul
if not errorlevel 1 goto :skip_node_version
echo   The Node.js found here will not run NEA Coach. It is either too old
echo   or damaged:
echo   %NODE_BIN%
echo.
echo   Install the current LTS from https://nodejs.org and try again.
echo.
goto :finish
:skip_node_version

if not defined NODE_BIN goto :no_node

REM ------------------------------------------------------------ 2. sanity check
if not exist "%~dp0server.js" goto :no_server
"%NODE_BIN%" --check "%~dp0server.js" >nul 2>nul
if not errorlevel 1 goto :server_ok
:no_server
echo   The application files look damaged, or the folder is incomplete:
echo   server.js is missing or will not parse.
echo.
echo   If the folder is still being copied or synced, wait for that to
echo   finish and double-click this file again.
echo.
goto :finish
:server_ok

REM ------------------------------------------------------------ 3. already running?
REM Ask the app's own /api/health endpoint before starting anything. This is
REM done with the Node we just found rather than curl or PowerShell, so it
REM behaves the same on every Windows version without assuming extra tools.
REM Exit code 0 means a NEA Coach server is already answering on the port.
"%NODE_BIN%" -e "const h=require('http');const r=h.get({host:'127.0.0.1',port:process.argv[1],path:'/api/health',timeout:1500},s=>process.exit(s.statusCode===200?0:1));r.on('error',()=>process.exit(1));r.on('timeout',()=>{r.destroy();process.exit(1);});" "%PORT%" >nul 2>nul
if errorlevel 1 goto :not_running
echo   NEA Coach is already running.
echo.
echo   Opening http://127.0.0.1:%PORT% in your browser.
echo   This window can be closed.
echo.
start "" "http://127.0.0.1:%PORT%"
goto :finish
:not_running

REM ------------------------------------------------------------ 4. start it
echo   Starting. Your browser will open in a moment.
echo.
echo   KEEP THIS WINDOW OPEN while you use NEA Coach. The server runs inside
echo   it, so closing the window stops the app. Press Control-C to stop it.
echo.

REM Open the browser first. "start" hands the browser off and returns straight
REM away, and a browser takes longer to appear and load than the server takes
REM to start listening, so the first page load finds it running. If the page
REM ever says it cannot connect, wait a second and refresh it.
start "" "http://127.0.0.1:%PORT%"

NEA_COACH_PORT=%PORT% "%NODE_BIN%" "%~dp0server.js"
set "STATUS=%ERRORLEVEL%"

REM ------------------------------------------------------------ 5. exit
echo.
if not "%STATUS%"=="0" goto :stopped_badly
echo   NEA Coach has stopped.
echo.
goto :finish

:stopped_badly
echo   NEA Coach stopped with an error (code %STATUS%).
echo   The lines above this message explain what happened.
echo.
echo   If they mention the port being busy, something is already using port
echo   %PORT%. Open http://127.0.0.1:%PORT% first, or run the app on a
echo   different port by double-clicking this file after typing:
echo     set NEA_COACH_PORT=4174
echo.
goto :finish

REM ------------------------------------------------------------ no Node found
:no_node
echo   NEA Coach needs Node.js. It is a free, one-time install.
echo.
echo     1. Go to https://nodejs.org
echo     2. Click the big button marked LTS, not the one marked Current.
echo     3. Run the installer and accept all the defaults.
echo     4. Close this window, then double-click this file again.
echo.
echo   If it still says Node is missing, restart the computer and try once
echo   more. Windows only notices a new PATH entry after a restart.
echo.
echo   The app deliberately does not download Node for you. A 45 MB download
echo   hidden inside a .bat file is unreliable and is flagged by antivirus,
echo   so it asks you to install Node yourself instead. Nothing has been
echo   changed on this computer.
echo.

REM ------------------------------------------------------------ finish
:finish
echo.
echo   Press any key to close this window.
pause >nul

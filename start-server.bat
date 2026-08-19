@echo off
REM ===========================================================================
REM  start-server.bat - launch the local dev server
REM
REM  Rebuilds src/data/projects.json from the Excel workbook, then starts Astro
REM  with hot reload. Edit the workbook and re-run this file to see changes.
REM ===========================================================================

setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js was not found on your PATH.
    echo         Install the LTS build from https://nodejs.org/ and try again.
    echo.
    pause
    exit /b 1
)

if not exist "node_modules" (
    echo [SETUP] First run - installing dependencies. This takes a minute...
    call npm install
    if errorlevel 1 (
        echo.
        echo [ERROR] npm install failed. See the messages above.
        pause
        exit /b 1
    )
    echo.
)

REM Astro's default port. Change here if 4321 is taken.
set "PORT=4321"

echo ===========================================================================
echo  Engineering Portfolio - development server
echo ===========================================================================
echo.
echo  Reading data\Engineering Portfolio.xlsx...
echo.

call npm run data
if errorlevel 1 (
    echo.
    echo [ERROR] Could not read the workbook.
    echo         Close the file in Excel if it is open, then try again.
    pause
    exit /b 1
)

echo.
echo  Starting server on http://localhost:%PORT%/
echo  A browser window will open shortly.
echo.
echo  Leave this window OPEN while you work.
echo  To stop: press Ctrl+C here, or run stop-server.bat.
echo.

REM Give Astro a moment to bind the port before the browser opens. Full path so
REM this resolves to Windows' timeout.exe rather than a same-named tool that may
REM appear earlier on PATH (Git Bash ships one).
start "" /b cmd /c ""%SystemRoot%\System32\timeout.exe" /t 4 /nobreak >nul & start http://localhost:%PORT%/"

call npx astro dev --port %PORT%

echo.
echo Server stopped.
pause

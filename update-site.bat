@echo off
REM ===========================================================================
REM  update-site.bat - rebuild the whole site from the Excel workbook
REM
REM  Run this after editing data\Engineering Portfolio.xlsx or dropping new
REM  files into images\. It does the full chain in one go:
REM
REM     1. read the workbook  -> src\data\projects.json
REM     2. mirror images\     -> public\images\  (adds, updates, removes)
REM     3. build project pages and the static site -> dist\
REM
REM  Pass /quiet to skip the final pause, so this can run from a scheduled task.
REM
REM  Related: start-server.bat  = live preview with hot reload
REM           publish-site.bat  = uploadable copy in publish\ for a web host
REM ===========================================================================

setlocal
cd /d "%~dp0"

set "QUIET="
if /i "%~1"=="/quiet" set "QUIET=1"

where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js was not found on your PATH.
    echo         Install the LTS build from https://nodejs.org/ and try again.
    echo.
    if not defined QUIET pause
    exit /b 1
)

if not exist "node_modules" (
    echo [SETUP] Installing dependencies. This takes a minute...
    call npm install
    if errorlevel 1 (
        echo.
        echo [ERROR] npm install failed. See the messages above.
        if not defined QUIET pause
        exit /b 1
    )
    echo.
)

if not exist "data\Engineering Portfolio.xlsx" (
    echo [ERROR] data\Engineering Portfolio.xlsx is missing.
    echo         The site is built from that workbook, so nothing can be updated.
    echo.
    if not defined QUIET pause
    exit /b 1
)

echo ===========================================================================
echo  Engineering Portfolio - updating site
echo ===========================================================================
echo.

echo [1/2] Reading data\Engineering Portfolio.xlsx and collecting images ...
echo.
call npm run data
if errorlevel 1 (
    echo.
    echo [ERROR] Could not read the workbook.
    echo         Close the file in Excel if it is open, then try again.
    if not defined QUIET pause
    exit /b 1
)

echo.
echo [2/2] Building project pages ...
echo.
call npx astro build
if errorlevel 1 (
    echo.
    echo [ERROR] The build failed. See the messages above.
    if not defined QUIET pause
    exit /b 1
)

REM Page count - one index.html per generated route. Counted in pure cmd rather
REM than piping to find, because Git ships a same-named find.exe that can win on
REM PATH and would walk the whole drive instead (same hazard start-server.bat
REM avoids by calling timeout.exe by full path).
set "PAGES=0"
for /f "delims=" %%f in ('dir /b /s "dist\index.html" 2^>nul') do set /a PAGES+=1

echo.
echo ===========================================================================
echo  DONE - %PAGES% page(s) rebuilt in dist\
echo ===========================================================================
echo.
echo  Any warnings above about missing or unused images come from the
echo  workbook: an "Image N" row points at a file that is not in images\,
echo  or a file in images\ is referenced by no sheet.
echo.
echo  Next:  start-server.bat   preview the site locally
echo         publish-site.bat   make the uploadable copy in publish\
echo.

if defined QUIET exit /b 0
pause

@echo off
REM ===========================================================================
REM  publish-site.bat - build the uploadable version of the site
REM
REM  Produces a self-contained folder of static HTML/CSS/JS in publish\ that can
REM  be uploaded to any web host (GitHub Pages, Netlify, cPanel, S3 - anything
REM  that serves files). No Node.js is needed on the server.
REM
REM  Set SITE_URL / SITE_BASE below once you know the final address.
REM ===========================================================================

setlocal
cd /d "%~dp0"

REM Pass /quiet to skip the Explorer window and the final pause, so this file
REM can also be run from a script or scheduled task.
set "QUIET="
if /i "%~1"=="/quiet" set "QUIET=1"

REM ---------------------------------------------------------------------------
REM  PUBLISH TARGET - edit these two lines when the hosting address is known.
REM
REM  Local / unknown host (default):
REM      SITE_URL=http://localhost:4321      SITE_BASE=/
REM
REM  GitHub Pages project site, e.g. https://nickparadizov.github.io/engineering-portfolio/
REM      SITE_URL=https://nickparadizov.github.io   SITE_BASE=/engineering-portfolio
REM
REM  Custom domain, e.g. https://nicholasparadizov.com/
REM      SITE_URL=https://nicholasparadizov.com     SITE_BASE=/
REM ---------------------------------------------------------------------------
set "SITE_URL=http://localhost:4321"
set "SITE_BASE=/"

where node >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Node.js was not found on your PATH.
    echo         Install the LTS build from https://nodejs.org/ and try again.
    echo.
    pause
    exit /b 1
)

if not exist "node_modules" (
    echo [SETUP] Installing dependencies. This takes a minute...
    call npm install
    if errorlevel 1 (
        echo.
        echo [ERROR] npm install failed. See the messages above.
        pause
        exit /b 1
    )
    echo.
)

echo ===========================================================================
echo  Engineering Portfolio - building publishable site
echo ===========================================================================
echo.
echo   Site URL  : %SITE_URL%
echo   Base path : %SITE_BASE%
echo.

echo [1/3] Reading data\Engineering Portfolio.xlsx ...
call npm run data
if errorlevel 1 (
    echo.
    echo [ERROR] Could not read the workbook.
    echo         Close the file in Excel if it is open, then try again.
    pause
    exit /b 1
)

echo.
echo [2/3] Building static site ...
call npx astro build
if errorlevel 1 (
    echo.
    echo [ERROR] The build failed. See the messages above.
    pause
    exit /b 1
)

echo.
echo [3/3] Copying output to publish\ ...
if exist "publish" rmdir /s /q "publish"
mkdir "publish"
xcopy "dist\*" "publish\" /e /i /q /y >nul
if errorlevel 1 (
    echo.
    echo [ERROR] Could not copy the build output.
    pause
    exit /b 1
)

REM GitHub Pages runs Jekyll by default, which skips files and folders starting
REM with an underscore. This empty file turns that off.
type nul > "publish\.nojekyll"

echo.
echo ===========================================================================
echo  DONE
echo ===========================================================================
echo.
echo  Your uploadable site is in:
echo      %CD%\publish\
echo.
echo  Upload the CONTENTS of that folder to your web host.
echo.
echo  Reminder: if the site will not live at %SITE_URL%%SITE_BASE%,
echo  edit SITE_URL and SITE_BASE near the top of this file and re-run it.
echo.

if defined QUIET exit /b 0

explorer "%CD%\publish"
pause

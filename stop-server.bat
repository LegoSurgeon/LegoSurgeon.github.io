@echo off
REM ===========================================================================
REM  stop-server.bat - shut down the local dev server
REM
REM  Finds whatever is listening on the dev port and ends it. Use this when the
REM  server window was closed without Ctrl+C and the port is still held.
REM ===========================================================================

setlocal enabledelayedexpansion
cd /d "%~dp0"

REM Must match the PORT in start-server.bat.
set "PORT=4321"

echo ===========================================================================
echo  Engineering Portfolio - stopping dev server on port %PORT%
echo ===========================================================================
echo.

set "FOUND="

REM netstat columns:  Proto  LocalAddress  ForeignAddress  State  PID
REM The local address may be 0.0.0.0:4321, 127.0.0.1:4321 or [::1]:4321, so we
REM match on ":<port> " (trailing space keeps :4321 from matching :43210).
REM findstr needs /c: for a literal containing a space - without it, the space
REM would be read as a separator between two alternative search strings.
for /f "tokens=5" %%P in ('netstat -ano ^| findstr "LISTENING" ^| findstr /c:":%PORT% "') do (
    if not "%%P"=="0" (
        REM One port can report several rows; only kill each PID once.
        if not defined KILLED_%%P (
            set "KILLED_%%P=1"
            set "FOUND=1"
            echo  Ending process with PID %%P ...
            taskkill /PID %%P /T /F >nul 2>nul
            if errorlevel 1 (
                echo  [WARN] Could not end PID %%P. Try running this file as Administrator.
            ) else (
                echo  [OK] Stopped PID %%P.
            )
        )
    )
)

if not defined FOUND (
    echo  Nothing was listening on port %PORT% - the server is already stopped.
)

echo.
echo Done.

REM Full path so this always resolves to Windows' timeout.exe, not a same-named
REM tool that may appear earlier on PATH (Git Bash ships one).
"%SystemRoot%\System32\timeout.exe" /t 3 /nobreak >nul 2>nul

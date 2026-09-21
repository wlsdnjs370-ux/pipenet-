@echo off
REM restart_5051.bat - restart the :5051 server on the current source.  ASCII ONLY.
REM Runs restart_5051.ps1 (stops the old server + launcher loop, then runs
REM start_server.bat).  If the old server needs elevation to stop, a UAC
REM prompt appears - answer Yes.
cd /d "%~dp0"
echo %date% %time% restart_5051.bat started>>"%~dp0restart_5051.log"
powershell -NoProfile -ExecutionPolicy Bypass -NoExit -File "%~dp0restart_5051.ps1"

@echo off
REM ---------------------------------------------------------------------
REM  fix_5051_autostart.cmd   ::  run ONCE as Administrator
REM
REM  *** ASCII ONLY *** (cmd reads this file in the OEM codepage)
REM
REM  WHY
REM    Port 5051 is started at boot by a SYSTEM-owned Task Scheduler task.
REM    A SYSTEM process cannot be stopped from a normal shell, so every time
REM    the code changes the server cannot be restarted without a UAC prompt -
REM    and a UAC prompt cannot be raised from an automated tool session.
REM    Measured 2026-09-11: that single fact cost an hour.
REM
REM  WHAT THIS DOES
REM    1) disables ONLY the 5051 boot task  (the 5052 domain task is LEFT
REM       ALONE - fncadnet.com depends on it)
REM    2) puts a normal (non-elevated) autostart shortcut in the user's
REM       Startup folder, so 5051 comes up at logon owned by the user
REM
REM  AFTER THIS
REM    start_server.bat needs NO administrator rights ever again.
REM
REM  TO UNDO
REM    schtasks /change /tn "<task name>" /enable
REM    del "%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\FNCADnet 5051.lnk"
REM ---------------------------------------------------------------------

setlocal
set "ROOT=%~dp0.."
pushd "%ROOT%"

net session >nul 2>&1
if errorlevel 1 (
    echo [ERROR] This window is NOT elevated.
    echo         Right click this file and choose "Run as administrator".
    pause
    exit /b 1
)
echo [OK] elevated
echo.

echo === FNCADnet tasks currently registered ===
schtasks /query /fo table /nh 2>nul | findstr /i "FNCADnet"
echo.

echo === disabling the 5051 boot task (5052 is left alone) ===
set "FOUND=0"
for /f "tokens=1 delims=," %%T in ('schtasks /query /fo csv /nh 2^>nul ^| findstr /i "FNCADnet"') do call :maybe %%T
if "%FOUND%"=="0" echo [WARN] no FNCADnet 5051 task found - nothing to disable
echo.

echo === installing user-level autostart (no admin needed at logon) ===
set "LNK=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\FNCADnet 5051.lnk"
powershell -NoProfile -Command ^
  "$w=New-Object -ComObject WScript.Shell;" ^
  "$s=$w.CreateShortcut('%LNK%');" ^
  "$s.TargetPath='%ROOT%\start_server.bat';" ^
  "$s.WorkingDirectory='%ROOT%';" ^
  "$s.Description='FNCADnet local server on :5051';" ^
  "$s.Save()"
if exist "%LNK%" (echo [OK] %LNK%) else (echo [ERROR] could not create the shortcut)
echo.

echo Done.  Reboot to verify: :5051 should come up owned by your user,
echo and start_server.bat will no longer need administrator rights.
popd
pause
exit /b 0

:maybe
set "TN=%~1"
echo %TN% | findstr /i "5051" >nul
if errorlevel 1 (
    echo   keep    : %TN%
    goto :eof
)
echo   disable : %TN%
schtasks /change /tn "%TN%" /disable
if errorlevel 1 (echo   [ERROR] could not disable %TN%) else (set "FOUND=1")
goto :eof

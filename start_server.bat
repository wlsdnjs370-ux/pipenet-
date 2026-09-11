@echo off
REM ---------------------------------------------------------------------
REM  CAD -> PIPENET server launcher (Windows)  ::  port 5051
REM
REM  *** ASCII ONLY ***
REM  Do NOT put Korean (or any non-ASCII) text in this file.
REM  The file is UTF-8 but cmd.exe reads it as the OEM codepage (949 here),
REM  so non-ASCII bytes are mangled; inside if/for blocks the mangled bytes
REM  break the block and the comment lines get EXECUTED as commands.
REM  Measured 2026-09-11: a Korean version produced
REM  "'nstall' is not recognized" and exit code 9009.
REM
REM  Usage:
REM    1) once: install_requirements.bat  (or pip install -r requirements)
REM    2) double click, or run from cmd
REM       If the port is held by a SYSTEM-owned process you must run this
REM       file as Administrator, otherwise taskkill is denied.
REM ---------------------------------------------------------------------

cd /d "%~dp0"

REM ---- port ----
REM  This file is the :5051 launcher (the taskkill below frees :5051), so the
REM  server must actually land on 5051. An inherited PORT from the parent
REM  environment silently moves it somewhere else.
REM  Measured 2026-09-11: the calling shell had PORT=0, so serve.py asked the
REM  OS for a free port and waitress bound a random one (61395). The server
REM  was perfectly healthy - it just was not on 5051, and :5051 stayed empty.
REM  That cost an hour of "the server is hung" misdiagnosis.
if not defined PORT set "PORT=5051"
if "%PORT%"=="0" set "PORT=5051"
echo [%date% %time%] port: %PORT%

REM ---- am I elevated? (net session needs admin) ----
net session >nul 2>&1
if errorlevel 1 (
    echo [%date% %time%] elevation: NO  ^(taskkill on a SYSTEM process will fail^)
) else (
    echo [%date% %time%] elevation: YES
)

REM ---- pick a python that can actually import waitress ----
REM  Do not blindly activate .venv: this project's .venv has no waitress
REM  (measured 2026-09-11), so serve.py dies instantly with
REM  ModuleNotFoundError and the loop below spins every 5 seconds with no
REM  visible reason. This server has always run on the anaconda python.
set "PY="
if exist ".venv\Scripts\python.exe" call :trypy ".venv\Scripts\python.exe"
if not defined PY if exist "venv\Scripts\python.exe" call :trypy "venv\Scripts\python.exe"
if not defined PY call :trypy "python"
if not defined PY call :trypy "C:\Users\admin\anaconda3\python.exe"
if not defined PY goto nopy
echo [%date% %time%] python: %PY%
goto run

:trypy
if defined PY goto :eof
"%~1" -c "import waitress" >nul 2>&1
if errorlevel 1 goto :eof
set "PY=%~1"
goto :eof

:nopy
echo [ERROR] no python with waitress found.  pip install waitress
pause
exit /b 1

REM ---- run waitress in a loop so a crash restarts the server ----
:run
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":%PORT% " ^| findstr "LISTENING"') do call :freeport %%P
echo [%date% %time%] starting serve.py ...
"%PY%" -u serve.py
echo [%date% %time%] serve.py exited with code %errorlevel%, restarting in 5s ...
REM  Use the full path: a PATH that has a unix-like "timeout" (git bash) would
REM  otherwise win and fail with "invalid time interval", making the loop spin
REM  with no delay. Measured 2026-09-11.
"%SystemRoot%\System32\timeout.exe" /t 5 /nobreak >nul
goto run

:freeport
REM  Never hide the taskkill result. The old version had ">nul 2>&1" here,
REM  so when the kill was denied the window said nothing and just spun.
echo [%date% %time%] port %PORT% is held by PID %1 - killing it
taskkill /F /PID %1
if errorlevel 1 echo [%date% %time%] *** taskkill FAILED for PID %1 - not elevated, or owner is SYSTEM
goto :eof

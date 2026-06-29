@echo off
REM ─────────────────────────────────────────────────────────────────────
REM   CAD → PIPENET 자동 도식 변환 시스템 — 서버 자동 실행 (Windows)
REM
REM   사용법:
REM     1) 처음 한 번: install_requirements.bat 또는 수동으로 pip install
REM     2) 더블 클릭 또는 cmd 에서 start_server.bat 실행
REM     3) 부팅 시 자동 실행하려면 작업 스케줄러에 이 .bat 을 등록
REM        (README.md 의 "자동 시작 등록" 섹션 참조)
REM ─────────────────────────────────────────────────────────────────────

cd /d "%~dp0"

REM venv 가 있으면 활성화 (.venv 또는 venv)
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) else if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

REM waitress 로 production 서버 실행 (loop 안에서 — 크래시 시 자동 재시작)
:run
REM ── 좀비 정리 — :5051 을 점유한 잔존 인스턴스가 있으면 그 PID 만 종료.
REM    (PyCharm 콘솔 등 다른 python 은 건드리지 않음) 이게 없으면 좀비가 포트를
REM    쥔 채 죽어 새 인스턴스가 영원히 바인드 실패 → 무한 재시작 루프에 빠진다.
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":5051 " ^| findstr "LISTENING"') do (
    echo [%date% %time%] freeing port 5051 — killing stale PID %%P
    taskkill /F /PID %%P >nul 2>&1
)
echo [%date% %time%] starting serve.py ...
python serve.py
echo [%date% %time%] serve.py exited with code %errorlevel%, restarting in 5s ...
timeout /t 5 /nobreak >nul
goto run

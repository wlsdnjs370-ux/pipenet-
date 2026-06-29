"""Production WSGI 진입점 — waitress 로 Flask 앱 서빙.

Flask 의 개발 서버 (app.run) 는 단일 스레드 + 안정성 한계 → 24/7 외부 노출에는 부적합.
waitress 는 Windows-friendly production WSGI 서버 (멀티 스레드, 안정적, 신호 처리).

실행::

    python serve.py
        ─ host/port 는 환경변수 HOST, PORT 또는 기본 (0.0.0.0:5051)
        ─ threads 는 환경변수 WAITRESS_THREADS 또는 기본 16

환경변수 (.env 파일 또는 OS env):
    HOST                기본 0.0.0.0  (모든 인터페이스)
    PORT                기본 5051
    WAITRESS_THREADS    기본 16       (동시 요청 처리 워커)
    FLASK_SECRET_KEY    필수 (대조 서버.py 에서 세션용)
    LOGIN_PASSWORD      기본 "5361"   (로그인 게이트 비번)
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# .env 자동 로드 (대조 서버.py 보다 먼저)
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# pipenet_converter src 를 PYTHONPATH 에 (대조 서버.py 와 동일 패턴)
_HERE = Path(__file__).resolve().parent
_PC_SRC = _HERE / "pipenet_converter" / "src"
if _PC_SRC.is_dir() and str(_PC_SRC) not in sys.path:
    sys.path.insert(0, str(_PC_SRC))

# 대조 서버.py 의 app 객체 import (파일명에 공백/한글이 있어 importlib 사용)
import importlib.util
_SPEC = importlib.util.spec_from_file_location("server_app", str(_HERE / "대조 서버.py"))
_MODULE = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(_MODULE)
app = _MODULE.app


def _port_already_bound(host: str, port: int) -> bool:
    """이미 다른 프로세스가 이 포트를 점유 중인가.

    Windows 의 waitress 는 SO_REUSEADDR 를 켜기 때문에, 좀비 인스턴스가 떠 있어도
    두 번째 serve() 가 OSError 없이 같은 포트에 바인드돼 버린다(요청이 어느 쪽으로
    갈지 비결정적 → 멈춘 인스턴스로 가면 무한로딩). 이를 막기 위해 SO_EXCLUSIVEADDRUSE
    로 사전 바인드를 시도한다 — 다른 프로세스가 이미 점유 중이면 실패한다.
    """
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        if hasattr(socket, "SO_EXCLUSIVEADDRUSE"):  # Windows
            s.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        s.bind(("" if host == "0.0.0.0" else host, port))
        return False
    except OSError:
        return True
    finally:
        s.close()


def main() -> None:
    from waitress import serve
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5051"))
    threads = int(os.environ.get("WAITRESS_THREADS", "16"))
    if _port_already_bound(host, port):
        print(f"[serve.py] ERROR: 포트 {port} 이미 사용 중 — 다른 serve.py 인스턴스가 "
              f"떠 있습니다. 중복 실행을 거부하고 종료합니다. (start_server.bat 은 재실행 시 "
              f":{port} 점유 PID 를 먼저 정리합니다.)", file=sys.stderr)
        raise SystemExit(1)
    print(f"[serve.py] waitress serving on http://{host}:{port}  (threads={threads})")
    if not os.environ.get("FLASK_SECRET_KEY"):
        print("[serve.py] WARNING: FLASK_SECRET_KEY 가 .env/env 에 없음 — "
              "서버 재시작 시 세션이 모두 만료됩니다. .env 에 한 줄 추가 권장.")
    try:
        serve(app, host=host, port=port, threads=threads, ident="cad-pipenet-server")
    except OSError as exc:
        # 포트 점유(좀비 인스턴스 등) → raw traceback 대신 원인을 명확히 알리고 종료.
        # start_server.bat 은 재실행 시 :PORT 점유 PID 를 먼저 정리하므로 다음 루프에서 회복.
        print(f"[serve.py] ERROR: 포트 {port} 바인드 실패 — 이미 다른 인스턴스가 점유 "
              f"중일 수 있습니다. ({exc})", file=sys.stderr)
        raise SystemExit(1)


if __name__ == "__main__":
    main()

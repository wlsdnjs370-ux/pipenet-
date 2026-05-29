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


def main() -> None:
    from waitress import serve
    host = os.environ.get("HOST", "0.0.0.0")
    port = int(os.environ.get("PORT", "5051"))
    threads = int(os.environ.get("WAITRESS_THREADS", "16"))
    print(f"[serve.py] waitress serving on http://{host}:{port}  (threads={threads})")
    if not os.environ.get("FLASK_SECRET_KEY"):
        print("[serve.py] WARNING: FLASK_SECRET_KEY 가 .env/env 에 없음 — "
              "서버 재시작 시 세션이 모두 만료됩니다. .env 에 한 줄 추가 권장.")
    serve(app, host=host, port=port, threads=threads, ident="cad-pipenet-server")


if __name__ == "__main__":
    main()

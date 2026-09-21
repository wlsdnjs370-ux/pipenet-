# -*- coding: utf-8 -*-
"""F-3 브라우저 검증용 임시 서버 — 5065 (크로미움이 5060/5061 을 막는다)."""
import importlib.util
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
sys.stdout.reconfigure(errors="replace")
spec = importlib.util.spec_from_file_location("daejo", str(ROOT / "대조 서버.py"))
srv = importlib.util.module_from_spec(spec)
spec.loader.exec_module(srv)
srv.app.run(host="127.0.0.1", port=5065, debug=False, use_reloader=False)

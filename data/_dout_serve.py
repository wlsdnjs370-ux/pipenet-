# -*- coding: utf-8 -*-
"""브라우저 검증용 임시 서버 — 127.0.0.1:5057. 공개 서버(5051)는 건드리지 않는다."""
import importlib.util
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("daejo", ROOT / "대조 서버.py")
srv = importlib.util.module_from_spec(spec)
spec.loader.exec_module(srv)
srv.app.run(host="127.0.0.1", port=5057, debug=False, threaded=True)

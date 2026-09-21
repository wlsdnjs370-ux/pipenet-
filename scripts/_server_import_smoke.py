# -*- coding: utf-8 -*-
"""서버 앱 import 스모크 — 편집된 코드로 Flask app 이 깨끗이 로드되는지."""
import importlib.util
import sys
from pathlib import Path

here = Path(__file__).resolve().parent.parent
pc = here / "pipenet_converter" / "src"
if pc.is_dir():
    sys.path.insert(0, str(pc))
if str(here) not in sys.path:
    sys.path.insert(0, str(here))

spec = importlib.util.spec_from_file_location("server_app", str(here / "대조 서버.py"))
m = importlib.util.module_from_spec(spec)
spec.loader.exec_module(m)
n = len(list(m.app.url_map.iter_rules()))
print(f"server app import OK, routes={n}")

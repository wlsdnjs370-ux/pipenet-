from __future__ import annotations

import runpy
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
TARGET = BASE_DIR / "대조 서버.py"

runpy.run_path(str(TARGET), run_name="__main__")

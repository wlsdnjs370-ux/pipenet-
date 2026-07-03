# -*- coding: utf-8 -*-
"""pytest 부트스트랩 — 이 디렉토리를 sys.path 에 올려 golden_cases 임포트 허용."""
import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

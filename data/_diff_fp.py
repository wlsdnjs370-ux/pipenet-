# -*- coding: utf-8 -*-
"""지문 두 벌 대조 — 무엇이 어떻게 달라졌나."""
from __future__ import annotations

import json
import sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

a = json.load(open(sys.argv[1], encoding="utf-8"))
b = json.load(open(sys.argv[2], encoding="utf-8"))

for k in a:
    ca = Counter(tuple(x) if isinstance(x, list) else x for x in a[k])
    cb = Counter(tuple(x) if isinstance(x, list) else x for x in b[k])
    if ca == cb:
        print(f"{k:9} 동일 ({sum(cb.values())}개)")
        continue
    oa, ob = ca - cb, cb - ca
    print(f"{k:9} ★다름 — 전에만 {sum(oa.values())} · 후에만 {sum(ob.values())}")
    for t, c in list(oa.items())[:6]:
        print(f"     전에만 ×{c}: {t}")
    for t, c in list(ob.items())[:6]:
        print(f"     후에만 ×{c}: {t}")

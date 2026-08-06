# -*- coding: utf-8 -*-
"""PR-B 검증 준비 — inspect 스트림이 방출하는 엔티티 타입 히스토그램."""
from __future__ import annotations
import io, json, sys, collections
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import importlib
srv = importlib.import_module("대조 서버")
app = srv.app
app.before_request_funcs[None] = [
    f for f in app.before_request_funcs.get(None, [])
    if f.__name__ != "_require_login_gate"
]

for name in sys.argv[1:]:
    p = BASE / name
    if not p.is_file():
        print(f"없음: {name}"); continue
    c = app.test_client()
    r = c.post("/api/remote30/inspect",
               data={"dxf_file": (io.BytesIO(p.read_bytes()), p.name)},
               content_type="multipart/form-data")
    hist = collections.Counter()
    layers = collections.Counter()
    for line in r.get_data(as_text=True).splitlines():
        if not line.strip():
            continue
        try:
            msg = json.loads(line)
        except Exception:
            continue
        for e in msg.get("entities", ()) or ():
            hist[e.get("t")] += 1
            if e.get("t") in ("A", "H", "S"):
                layers[(e.get("t"), e.get("l"))] += 1
    print(f"\n=== {p.name}  (HTTP {r.status_code}) ===")
    print("  타입:", dict(sorted(hist.items())))
    if layers:
        print("  A/H/S 상위 레이어:", layers.most_common(8))

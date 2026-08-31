# -*- coding: utf-8 -*-
"""정찰만 — 띠 분포가 실행마다 같은가.

F-11c 의 B1F 실측이 F-11a 때와 두 배 어긋났다:

    F-11a  띠 {높음 72, 중간 3163, 낮음 103} → 채택 3,235 · 최원 416.85 m
    F-11c  채택 예정 6,688                    · 최원 851.35 m

정찰은 DXF 만 읽는 순수 함수여야 한다 — 같은 파일이면 같은 수가 나와야 한다.
두 배가 나온다면 «도면이 두 겹으로 들어왔다» 는 뜻이고, 그러면 그 위에서 잰
F-11a 의 수치도 F-11c 의 수치도 못 믿는다. 45분짜리 전 구간 실측을 다시 돌리기
전에 여기서 먼저 가른다 — 열기+정찰만이라 몇 분이면 된다.

    python scripts/_probe_recon_only.py [도면.dxf ...]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = [ROOT / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"]


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def once(c, dxf, tag):
    t0 = time.time()
    with open(dxf, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                   content_type="multipart/form-data")
    sid = (r.get_json() or {})["sid"]
    j = wait(c, sid)
    rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
           .get("recon") or {})
    ad = rec.get("adopt") or {}
    b = rec.get("bands") or {}
    print(f"    {tag}  후보 {rec.get('n')} · 띠 {b} · 규칙 {ad.get('rule')} "
          f"· 채택 예정 {ad.get('n')} · {time.time() - t0:.0f}초"
          + (f" · 열기 {j.get('state')}" if j.get("state") != "done" else ""))
    return sid, ad.get("n"), rec


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    files = [Path(x) for x in sys.argv[1:]] or DEF
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        for f in files:
            if not f.is_file():
                print(f"\n■ {f.name} — 파일 없음")
                continue
            print(f"\n■ {f.name} ({f.stat().st_size / 1048576:.0f} MB)")
            # ★두 번 연다. 한 번만 재면 «이 값이 원래 값인지» 를 못 가린다 —
            #   두 번이 같으면 적어도 한 세션 안에서는 결정적이다.
            _sid1, n1, _r1 = once(c, f, "1회")
            _sid2, n2, _r2 = once(c, f, "2회")
            print(f"    → 두 번이 {'같다' if n1 == n2 else '★다르다'} "
                  f"({n1} vs {n2})")
    return 0


if __name__ == "__main__":
    sys.exit(main())

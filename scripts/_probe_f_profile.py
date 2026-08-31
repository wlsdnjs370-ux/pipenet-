# -*- coding: utf-8 -*-
"""모듈 F 완주를 프로파일링 — «F 자신의 코드» 가 얼마나 쓰나.

「최적화」를 하려면 시간이 어디로 가는지부터 알아야 한다. 이 저장소에서 이미
아는 것: B1F 완주 51분 중 50분이 표 확정(엔진)이다(BLOCKED §21). 그렇다면
모듈 F 자신의 라우트·직렬화 코드는 얼마나 쓰는가? 그 답에 따라 리팩터링의
성격이 갈린다 — 느리면 «최적화», 안 느리면 «구조·정직성» 이 목표가 된다.

    python scripts/_probe_f_profile.py [도면.dxf]
    → data/_f_profile.txt
"""
from __future__ import annotations

import cProfile
import io
import os
import pstats
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.05)
    return {"state": "timeout"}


def lane(c, dxf):
    with open(dxf, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                   content_type="multipart/form-data")
    sid = (r.get_json() or {})["sid"]
    wait(c, sid)
    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
           .get("recon") or {})
    ad = rec.get("adopt") or {}
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True,
                 "heads": {"conf_min": ad.get("conf_min")}})
    wait(c, sid)
    c.post("/api/module-f/pick/commit", json={"sid": sid})
    wait(c, sid)
    st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    heads = [(float(h[0]), float(h[1])) for h in (st["heads"] or ())]
    hs = heads[::max(1, len(heads) // 60)][:60]
    for s2 in sorted((g.get("segs") or [] for g in st["body_groups"]),
                     key=len, reverse=True)[:18]:
        pts = [(float(s2[i]), float(s2[i + 1]))
               for i in range(0, len(s2) - 3, 4)]
        if len(pts) > 4000:
            pts = pts[::len(pts) // 4000]
        if not (pts and hs):
            continue
        p0, d0 = None, None
        for hx, hy in hs:
            p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
            d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
            if d0 is None or d < d0:
                p0, d0 = p, d
        if p0 is None or d0 > 2000.0:
            continue
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": p0[0], "y": p0[1]})
        wait(c, sid)
        if (c.get(f"/api/module-f/edit/state?sid={sid}")
                .get_json()["state"].get("worst") or {}).get("k"):
            break
    c.post("/api/module-f/design/build", json={"sid": sid})
    wait(c, sid)
    # 화면이 실제로 부르는 조회들 — 직렬화 비용이 여기 있다.
    c.get(f"/api/module-f/world?sid={sid}")
    c.get(f"/api/module-f/design/preview?sid={sid}")
    c.get(f"/api/module-f/edit/state?sid={sid}")
    c.post("/api/module-f/design/emit", json={"sid": sid})
    return sid


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        t0 = time.time()
        pr = cProfile.Profile()
        pr.enable()
        lane(c, dxf)
        pr.disable()
        wall = time.time() - t0

    buf = io.StringIO()
    st = pstats.Stats(pr, stream=buf).sort_stats("cumulative")
    st.print_stats(60)
    text = buf.getvalue()
    out = ROOT / "data" / "_f_profile.txt"
    out.write_text(text, encoding="utf-8")

    print(f"\n■ {dxf.name} 완주 {wall:.1f}초 (프로파일러 포함이라 실제보다 느리다)")
    # 모듈 F 자신의 코드만 골라 «누적» 이 아니라 «자기 시간(tottime)» 으로 본다.
    buf2 = io.StringIO()
    pstats.Stats(pr, stream=buf2).sort_stats("tottime").print_stats(400)
    mine, engine, other = 0.0, 0.0, 0.0
    rows = []
    for ln in buf2.getvalue().splitlines():
        parts = ln.split(None, 5)
        if len(parts) < 6 or not parts[1].replace(".", "").isdigit():
            continue
        try:
            tot = float(parts[1])
        except ValueError:
            continue
        where = parts[5]
        if "routes\\module_f" in where or "routes/module_f" in where:
            mine += tot
            rows.append((tot, where))
        elif "cad_project_editor_g" in where or "remote30_prototype" in where:
            engine += tot
        else:
            other += tot
    tot_all = mine + engine + other
    print(f"\n   자기 시간 합계 {tot_all:,.1f}초")
    print(f"     모듈 F 라우트   {mine:>7,.1f}초  ({mine / max(1e-9, tot_all) * 100:>4.1f}%)")
    print(f"     엔진(G·A)       {engine:>7,.1f}초  ({engine / max(1e-9, tot_all) * 100:>4.1f}%)")
    print(f"     그 외(표준·flask){other:>7,.1f}초  ({other / max(1e-9, tot_all) * 100:>4.1f}%)")
    print("\n   모듈 F 안에서 무거운 순서")
    for tot, where in sorted(rows, reverse=True)[:12]:
        print(f"     {tot:>7.3f}초  {where}")
    print(f"\n   전체 프로파일 → {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

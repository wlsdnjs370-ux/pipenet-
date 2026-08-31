# -*- coding: utf-8 -*-
"""«앵커 노드» meta 가 비는가 — 미리보기의 최원 유하거리 경로가 안 그려지는 이유.

라우트 시험을 깊게 넣다가 `design/preview` 의 `view.anchor` 가 **None** 인 것을
봤다. 그러면 「최원 유하거리가 어느 줄인가」를 화면이 못 그린다 —
`_worst_view` 도크스트링이 그 줄을 따로 싣는 이유로 적어 둔 바로 그 기능이다.

시험의 단정이 과했던 것인지, 진짜 결함인지 가른다. 표의 meta 를 직접 본다.

    python scripts/_probe_anchor_meta.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def wait(c, sid, limit=9000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


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
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": (rec.get("adopt") or {})
                               .get("conf_min") or 0.75}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        heads = [(float(h[0]), float(h[1])) for h in (st["heads"] or ())]
        hs = heads[::max(1, len(heads) // 40)][:40]
        for s2 in sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[:12]:
            pts = [(float(s2[i]), float(s2[i + 1]))
                   for i in range(0, len(s2) - 3, 4)]
            if not (pts and hs):
                continue
            best, bd = None, None
            for hx, hy in hs:
                p = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
                d = ((p[0] - hx) ** 2 + (p[1] - hy) ** 2) ** 0.5
                if bd is None or d < bd:
                    best, bd = p, d
            if best is None or bd > 2000.0:
                continue
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": best[0], "y": best[1]})
            wait(c, sid)
            w = ((c.get(f"/api/module-f/edit/state?sid={sid}").get_json()
                  or {}).get("state") or {}).get("worst") or {}
            if w.get("k"):
                print(f"■ 원클릭 — 최불리 {w['k']}개 · 최원 {w.get('far_m')} m")
                print(f"   worst.anchor = {w.get('anchor')!r} "
                      f"(board 헤드 번호)")
                break
        c.post("/api/module-f/design/build", json={"sid": sid})
        wait(c, sid)
        d = c.get(f"/api/module-f/design/preview?sid={sid}").get_json() or {}
        t = d.get("tables") or {}
        meta = dict((k, v) for k, v in (t.get("meta") or []))
        v = d.get("view") or {}
        print(f"\n■ 표 meta")
        for k in ("기준개수 K", "앵커 노드", "최원 유하거리 (m)",
                  "설계면적 폭 (m)", "corridor 총연장 (m)"):
            print(f"   {k:<22} {meta.get(k)!r}")
        print(f"\n■ 미리보기 view")
        print(f"   anchor         {v.get('anchor')!r}")
        print(f"   anchor_path    {len(v.get('anchor_path') or [])}점 · "
              f"{v.get('anchor_path_m')} m")
        print(f"   nodes 중 anchor 표시 "
              f"{sum(1 for n in (v.get('nodes') or []) if n.get('anchor'))}개")
        if meta.get("앵커 노드") in (None, "?"):
            print("\n  ★meta 의 «앵커 노드» 가 비었다 — 화면이 최원 유하거리")
            print("    경로를 못 그린다. `_anchor_node()` 가 최불리 앵커를 kfp")
            print("    노드로 못 옮긴 것이다.")
        else:
            print("\n  [OK] 앵커가 표와 화면에 다 있다")
    return 0


if __name__ == "__main__":
    sys.exit(main())

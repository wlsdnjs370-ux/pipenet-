# -*- coding: utf-8 -*-
"""[B1F] 붙는 헤드 걸러 고르기가 큰 도면에서 얼마나 걸리나 + 선정==표 인가."""
from __future__ import annotations
import math, os, sys, time
from pathlib import Path
ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

def wait(c, sid, limit=200000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}

def main() -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        j0 = c.post("/api/module-f/reopen",
                    json={"key": "B1F 현장조사 소화설비 평면도"}).get_json() or {}
        if not j0.get("ok"):
            print("★이어서 열기 실패", str(j0)[:200]); return 1
        sid = j0["sid"]
        if wait(c, sid).get("state") != "done":
            print("★열기 잡 실패"); return 1
        from routes.module_f.jobs import _sess
        sess = _sess(sid); b = sess["edit"].board
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        print(f"\n■ B1F · 절점 {len(b.pts):,} · 헤드 {len(b.disks):,}")
        t0 = time.perf_counter()
        jw = c.post("/api/module-f/edit/worst",
                    json={"sid": sid, "k": 30}).get_json()
        t_sel = time.perf_counter() - t0
        if not jw.get("ok"):
            print("★최불리 실패", str(jw)[:200]); return 1
        picked = [int(i) for i in sess["worst"]["heads"]]
        print(f"  최불리 선정 {len(picked)}개 · **{t_sel:.1f}s** (첫 회 — 탐침 포함)")
        t0 = time.perf_counter()
        jw2 = c.post("/api/module-f/edit/worst",
                     json={"sid": sid, "k": 30}).get_json()
        print(f"  다시 선정        · **{time.perf_counter()-t0:.1f}s** (캐시)")
        t0 = time.perf_counter()
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패"); return 1
        print(f"  표 확정          · {time.perf_counter()-t0:.1f}s")
        got, tbl = sess["design"]["got"], sess["design"]["tables"]
        o = got.get("origin_mm") or (0.0, 0.0)
        ox, oy = float(o[0]) - 1000.0, float(o[1]) - 1000.0
        at = {str(n.get("label")): (float(n.get("x") or 0)+ox,
                                    float(n.get("y") or 0)+oy) for n in tbl.nodes}
        A = [(float(b.disks[i][0]), float(b.disks[i][1]), float(b.disks[i][2]))
             for i in [int(i) for i in sess["worst"]["heads"]]]
        B = [at[str(z.get("in"))] for z in tbl.nozzles if str(z.get("in")) in at]
        used, gaps, miss = set(), [], 0
        for (x, y, r) in A:
            best, bd = None, 1e18
            for t, p in enumerate(B):
                if t in used: continue
                dd = math.dist(p, (x, y))
                if dd < bd: best, bd = t, dd
            if best is None or bd > r + 60.0: miss += 1; continue
            used.add(best); gaps.append(bd)
        gaps.sort()
        print(f"\n  ㉠ 선정 {len(A)} · ㉡ 표 {len(B)}")
        print(f"  ★1:1 {len(gaps)}/{len(A)}"
              + (f" · 최대 오차 {gaps[-1]:.1f} mm" if gaps else "")
              + f" · 선정에만 {miss} · 표에만 {len(B)-len(used)}")
        ok = len(A) == len(B) and miss == 0 and len(used) == len(B)
        print("\n  " + ("★개수·좌표가 정확히 같다" if ok else "★★아직 다르다"))
        return 0 if ok else 2

if __name__ == "__main__":
    raise SystemExit(main())

# -*- coding: utf-8 -*-
"""[신축배관 접기] 손질판에 실제로 접히는가 — 라우트를 붙이기 전에 재 본다.

`routes/module_f/flexfold.py` 를 그대로 불러, 대명동 손질판에서

    · 가닥 몇 개가 접히고 몇 개가 못 접히는가 · 왜
    · 간선이 몇 개 줄어드는가
    · 선언 길이 합이 접기 전 총연장과 같은가 (§5 기준 3)

를 낸다. 접기를 화면에 붙이기 전에 여기서 걸러야, 라우트가 조용히 틀린 채로
돌아가는 일이 없다.

    python scripts/_probe_flex_fold_dry.py
"""
from __future__ import annotations

import os
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _why(board, chains, ff):
    """★못 접은 가닥 하나를 «점 단위로» 펴 본다 — 셈만 보면 원인을 못 짚는다."""
    import math
    pts = list(board.pts)
    edges = {tuple(sorted(e)) for e in board.edges}
    deg: dict = {}
    nb: dict = {}
    for i, j in edges:
        deg[i] = deg.get(i, 0) + 1
        deg[j] = deg.get(j, 0) + 1
        nb.setdefault(i, set()).add(j)
        nb.setdefault(j, set()).add(i)

    def near(p, eps=ff.SNAP_EPS):
        best, bd = None, eps
        for i, q in enumerate(pts):
            d = math.hypot(q[0] - p[0], q[1] - p[1])
            if d <= bd:
                best, bd = i, d
        return best, bd

    print("  [진단] 못 접은 가닥 2개를 점 단위로")
    shown = 0
    for path in chains:
        rows = [(near(p), p) for p in path]
        ns = [r[0][0] for r in rows]
        if None in (ns[0], ns[-1]):
            continue
        mine = {tuple(sorted((a, b))) for a, b in zip(ns, ns[1:])
                if a is not None and b is not None and a != b}
        mine &= edges
        bad = [n for n in set(ns[1:-1])
               if n is not None and deg.get(n, 0) > sum(1 for e in mine
                                                        if n in e)]
        if not bad:
            continue
        print(f"    가닥 {len(path)}점 · 손질판 간선 {len(mine)}개")
        for (n, d), p in rows:
            mark = " ←★가닥 밖 배관" if n in bad else ""
            print(f"      ({p[0]:9.1f},{p[1]:9.1f}) → 노드 {n}"
                  f" (거리 {d:.2f} · 차수 {deg.get(n, 0)}){mark}")
        shown += 1
        if shown >= 2:
            break


def main() -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    if not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    from routes.module_f import flexfold as ff

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(PLAN, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": (rec.get("adopt") or {})
                               .get("conf_min")}})
        wait(c, sid)
        from routes.module_f.jobs import _sess
        ps = _sess(sid)["pick"]

        layers = sorted({str(k[0]) for k in ps.board.mat if ff.is_flex(k[0])})
        pieces = ff.flex_pieces(ps.board.w, layers)
        chains, tangled, rest = ff.strands(pieces)
        tot = sum(ff.run_length(p) for p in chains)
        print(f"\n■ 접기 예행 · 레이어 {layers}")
        print(f"  조각 {len(pieces)} → 가닥 {len(chains)}"
              f" · 사슬아님 {len(tangled)} · 미포함 {len(rest)}")
        print(f"  가닥당 조각 {dict(sorted(Counter(len(p) - 1 for p in chains).items()))}")
        print(f"  ★접기 전 총연장 {tot:,.0f} mm")

        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        es = _sess(sid)["edit"]
        n0 = len(es.board.edges)
        got = ff.fold_board(es.board, chains)
        print(f"  간선 {got['before']} → {got['after']}"
              f"  ({got['after'] - got['before']:+d})  · 손질판 {n0}")
        print(f"  ★가닥 {got['strands']}/{len(chains)} 처리 · 접힌 구간"
              f" {got['folded']} · 분기에서 자른 곳 {got.get('cuts', 0)}"
              f" · 못 접음 {got['skipped']} {got['why']}")
        dec = sum(got["edge_len_mm"].values())
        print(f"  ★선언 길이 합 {dec:,.0f} mm"
              f" · 접기 전 총연장 대비 {dec - tot:+,.0f} mm")

        # 접은 뒤 물닿음 — 하나라도 줄면 멈춰야 한다(§5 기준 2)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        from services.cad_import.design.restrict import attachable_heads
        pay = es.convert_payload()
        pay["edge_len_mm"] = got["edge_len_mm"]
        wet = attachable_heads(pay)
        print(f"  ★접은 뒤 물닿음 헤드 {len(wet['wet'])} / {wet['total']}"
              f" · 떨어짐 {wet['dropped']}   (접기 전 111)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

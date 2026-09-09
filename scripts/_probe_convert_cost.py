# -*- coding: utf-8 -*-
"""[변환 비용] 큰 도면에서 무엇이 시간을 먹는가 — 전체망을 몇 번 도는가.

사용자 물음: 「최불리 배관망뿐 아니라 **전체 도면까지 다 파싱**하느라 오래
걸리는 것 아니냐」.

코드 흐름은 이렇다::

    4.수리계산 「표 확정」  select_and_expand()
        ① attachable_heads()  → build_planar_graph(**전체망**)
             최불리를 고르기 «전» 에 전체를 한 번 전개한다 — 어느 헤드가 전개에
             붙을 수 있는지 알아야 후보를 좁힐 수 있기 때문이다(주석: 선정은
             board 도달로 세지만 전개는 더 엄격해 B1F 실측 868 대 619).
        ② worst_k_heads()     → board 그래프(planar 아님)
        ③ expand_worst()      → build_planar_graph(제한망 K개)

    5.입력 변환  convert/run
        outputs 기본값이 full_kfp=True 다 → **전체망**을 또 변환한다

여기서 재는 것: `build_planar_graph` 호출마다 **입력 크기와 소요**. 그러면
「전체망이 몇 번 · 몇 초」가 숫자로 나온다.

    python scripts/_probe_convert_cost.py [--key 저장본이름] [--k 30]
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

CALLS: list = []


def wait(c, sid, limit=200000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _wrap():
    """`build_planar_graph` 를 감싸 호출마다 (입력 크기, 소요) 를 적는다."""
    from services.cad_import.convert import planar
    from services.cad_import.design import restrict

    real = planar.build_planar_graph

    def spy(key, out=None, write=False, **graph):
        n_pts = len(graph.get("pts") or ())
        n_edg = len(graph.get("edges") or ())
        n_hcov = len(graph.get("hcov") or ())
        t0 = time.perf_counter()
        got = real(key, out, write, **graph)
        dt = time.perf_counter() - t0
        CALLS.append({"pts": n_pts, "edges": n_edg, "hcov": n_hcov,
                      "sec": round(dt, 2),
                      "ok": bool((got or {}).get("ok"))})
        return got

    planar.build_planar_graph = spy
    # ★`restrict` 는 함수 안에서 import 하므로 모듈 속성을 갈아도 그때 다시
    #   집어 온다 — 그래서 원본 모듈만 갈면 된다.
    _ = restrict
    return real


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default="B1F 현장조사 소화설비 평면도")
    ap.add_argument("--k", type=int, default=30)
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    _wrap()

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        t0 = time.perf_counter()
        r = c.post("/api/module-f/reopen", json={"key": args.key})
        j0 = r.get_json() or {}
        if not j0.get("ok"):
            print(f"★이어서 열기 실패 — {str(j0)[:200]}")
            return 1
        sid = j0["sid"]
        if wait(c, sid).get("state") != "done":
            print("★이어서 열기 잡 실패")
            return 1
        t_open = time.perf_counter() - t0
        from routes.module_f.jobs import _sess
        es = _sess(sid)["edit"]
        print(f"\n■ 변환 비용 · {args.key} · K={args.k}")
        print(f"  손질판 점 {len(es.board.pts):,} · 간선 {len(es.board.edges):,}"
              f" · 헤드 {len(es.board.disks):,}")
        print(f"  [이어서 열기] {t_open:.1f}s")

        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)

        CALLS.clear()
        t0 = time.perf_counter()
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": args.k})
        t_worst = time.perf_counter() - t0
        n1 = len(CALLS)
        print(f"  [최불리 선정] {t_worst:.1f}s · planar {n1}회")

        t0 = time.perf_counter()
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        j = wait(c, sid)
        t_build = time.perf_counter() - t0
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {str(j)[:200]}")
            return 1
        print(f"  [표 확정] {t_build:.1f}s · planar {len(CALLS) - n1}회")
        n2 = len(CALLS)

        for name, outs in (("최불리 SDF 만", {"worst_sdf": True}),
                           ("전체망 kfp 포함", {"full_kfp": True,
                                             "worst_kfp": True,
                                             "worst_sdf": True})):
            before = len(CALLS)
            t0 = time.perf_counter()
            c.post("/api/module-f/convert/run",
                   json={"sid": sid, "dto": {}, "outputs": outs})
            jj = wait(c, sid)
            dt = time.perf_counter() - t0
            print(f"  [변환 · {name}] {dt:.1f}s"
                  f" · planar {len(CALLS) - before}회"
                  f" · {jj.get('state')}")
        _ = n2

        print(f"\n  ── `build_planar_graph` 호출 전수 {len(CALLS)}")
        tot = 0.0
        for i, r in enumerate(CALLS, 1):
            tot += r["sec"]
            tag = ("★전체망" if r["pts"] >= len(es.board.pts) * 0.99
                   else "제한망")
            print(f"    {i:2}. {tag} · 점 {r['pts']:>7,} · 간선"
                  f" {r['edges']:>7,} · 헤드 {r['hcov']:>5,}"
                  f" · **{r['sec']:>7.1f}s**")
        print(f"    합 {tot:.1f}s")
        full = [r for r in CALLS if r["pts"] >= len(es.board.pts) * 0.99]
        print(f"    ★그중 «전체망» {len(full)}회 ·"
              f" {sum(r['sec'] for r in full):.1f}s"
              f" ({sum(r['sec'] for r in full) / max(tot, 1e-9) * 100:.0f} %)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

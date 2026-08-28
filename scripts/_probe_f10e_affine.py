# -*- coding: utf-8 -*-
"""[F-10e] board → 설계 표 좌표가 «전역 변환» 으로 이어지는가.

앞선 실측(`_probe_f10e_coords.py`)에서 두 좌표계가 다르다는 것이 드러났다:

    설계 표  x 1,000~22,650   y 1,000~17,150
    board   x 242,770~301,570 y -256,650~-215,100   (일치 0/178)

엔진이 표를 만들기 전에 노드를 옮긴다(로그 「직선 위치 복원: 109노드」).
그러면 밑그림(board 전체)을 아이소 좌표계에 «어긋남 없이» 깔 수 있는지가
전적으로 «그 옮김이 전역 변환인가» 에 달린다.

대응은 추측하지 않는다 — `got["edge_ref"]` 가 설계 배관 ↔ board 간선을 이미
들고 있다. 그것으로 (board 점, 설계 점) 짝을 만들어 최소제곱 아핀을 맞추고
**잔차** 를 본다. 잔차가 크면 전역 변환이 없다는 뜻이고, 그러면 밑그림은
지시서가 요구한 «1픽셀도 안 어긋남» 을 만족할 수 없다.

    python scripts/_probe_f10e_affine.py [도면.dxf]
"""
from __future__ import annotations

import os
import statistics
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=3000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def fit_affine(src, dst):
    """최소제곱 아핀 (x,y,1)→(X,Y). numpy 없이 정규방정식 6x6 을 푼다."""
    n = len(src)
    S = [[0.0] * 3 for _ in range(3)]
    tx = [0.0] * 3
    ty = [0.0] * 3
    for (x, y), (X, Y) in zip(src, dst):
        v = (x, y, 1.0)
        for i in range(3):
            for j in range(3):
                S[i][j] += v[i] * v[j]
            tx[i] += v[i] * X
            ty[i] += v[i] * Y

    def solve(M, b):
        A = [row[:] + [b[i]] for i, row in enumerate(M)]
        for col in range(3):
            p = max(range(col, 3), key=lambda r: abs(A[r][col]))
            if abs(A[p][col]) < 1e-12:
                return None
            A[col], A[p] = A[p], A[col]
            for r in range(3):
                if r == col:
                    continue
                f = A[r][col] / A[col][col]
                for k in range(col, 4):
                    A[r][k] -= f * A[col][k]
        return [A[i][3] / A[i][i] for i in range(3)]

    a = solve([r[:] for r in S], tx)
    b2 = solve([r[:] for r in S], ty)
    if a is None or b2 is None:
        return None, None
    return a, b2


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
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True, "heads": {"conf_min": 0.9}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        pipe = []
        for g in st["body_groups"]:
            s2 = g["segs"]
            for i in range(0, len(s2) - 3, 4):
                pipe.append((float(s2[i]), float(s2[i + 1])))
        hx, hy = float(st["heads"][0][0]), float(st["heads"][0][1])
        px, py = min(pipe, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": px, "y": py})
        wait(c, sid)
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        if wait(c, sid)["state"] != "done":
            print("표 확정 실패")
            return 2

        from routes.module_f.jobs import _SESSIONS
        sess = _SESSIONS[sid]
        tbl = sess["design"]["tables"]
        got = sess["design"]["got"]
        board = sess["edit"].board
        ref = got.get("edge_ref") or {}
        at = {str(n.get("label")): n for n in tbl.nodes}

        print("\n■ got 의 열쇠들")
        for k, v in sorted(got.items()):
            kind = type(v).__name__
            n = ""
            try:
                n = f" · {len(v)}" if hasattr(v, "__len__") else ""
            except Exception:
                pass
            extra = ""
            if hasattr(v, "nodes"):
                extra = (f"  ← nodes {len(v.nodes)} · pipes "
                         f"{len(getattr(v, 'pipes', []) or [])}")
            print(f"    {k:<20} {kind}{n}{extra}")
        print(f"\n■ 대응 재료 — edge_ref {len(ref)}건 · 설계 배관 {len(tbl.pipes)}")
        # 설계 배관 라벨 → (in, out)
        ends = {str(p.get("label")): (str(p.get("in")), str(p.get("out")))
                for p in tbl.pipes}
        # (설계 라벨 → board 절점) 표를 짝수 대응으로 모은다. 방향이 뒤집힐 수
        # 있으므로 두 조합 중 거리가 작은 쪽을 고른다.
        pairs: dict = {}
        for pid, e in ref.items():
            lab = ends.get(str(pid))
            if not lab:
                continue
            try:
                i, j = int(e[0]), int(e[1])
            except (TypeError, ValueError, IndexError):
                continue
            if not (0 <= i < len(board.pts) and 0 <= j < len(board.pts)):
                continue
            a_lab, b_lab = lab
            if a_lab not in at or b_lab not in at:
                continue
            pairs.setdefault(a_lab, []).append(i)
            pairs.setdefault(b_lab, []).append(j)

        # 라벨마다 가장 자주 나온 board 절점을 그 짝으로 본다.
        from collections import Counter
        src, dst, labs = [], [], []
        for lab, idxs in pairs.items():
            bi = Counter(idxs).most_common(1)[0][0]
            bp = board.pts[bi]
            n = at[lab]
            src.append((float(bp[0]), float(bp[1])))
            dst.append((float(n.get("x", 0) or 0), float(n.get("y", 0) or 0)))
            labs.append(lab)
        print(f"    짝지어진 노드 {len(src)}개")
        if len(src) < 6:
            print("    ★대응이 너무 적다 — 판단 보류")
            return 3

        a, b2 = fit_affine(src, dst)
        if a is None:
            print("    ★아핀을 못 맞춘다(퇴화)")
            return 3
        errs = []
        for (x, y), (X, Y) in zip(src, dst):
            ex = a[0] * x + a[1] * y + a[2] - X
            ey = b2[0] * x + b2[1] * y + b2[2] - Y
            errs.append((ex * ex + ey * ey) ** 0.5)
        errs.sort()
        print(f"\n■ 최소제곱 아핀 잔차 (설계 표 좌표 단위)")
        print(f"    중앙값 {statistics.median(errs):,.1f} · "
              f"p90 {errs[int(len(errs) * .9)]:,.1f} · 최대 {errs[-1]:,.1f}")
        span = max(max(d[0] for d in dst) - min(d[0] for d in dst),
                   max(d[1] for d in dst) - min(d[1] for d in dst))
        print(f"    설계 도면 한 변 {span:,.0f} · 최대 잔차가 "
              f"{errs[-1] / max(1e-9, span) * 100:.1f}%")
        ok = errs[-1] <= span * 0.005
        print(f"\n  {'전역 변환이 있다 — 밑그림을 어긋남 없이 깔 수 있다'  if ok else '★전역 변환이 없다 — 엔진이 노드를 개별로 옮겼다'}")
        print("  (엔진 로그의 「직선 위치 복원」이 그 옮김이다.)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

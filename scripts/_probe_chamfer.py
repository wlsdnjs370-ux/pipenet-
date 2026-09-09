# -*- coding: utf-8 -*-
"""[챔퍼] 45° 모서리 조각이 배관을 자르고 있다 — 손질판에서 전수로 잰다.

■ 왜 재는가

  사용자 목적: **평면도와 아이소의 위상이 1:1 · 누락·휨 없이.**

  등각에서 나올 수 있는 각도는 셋뿐이다 — 평면 가로 → +30° · 평면 세로 →
  +150° · 헤드 스텁 → 수직. 평면의 45° 대각은 그 격자를 벗어나 «휘어» 보인다.

  표 단계 실측(대명동 K=30 · `_probe_pipe_placement.py`)::

      배관 262 중 0.3 m 미만 조각 140 (53 %)
      그 조각의 각도  45° 대각 91 · 기타 26 · 세로 13 · 가로 10
      조각 140 중 «양 끝 차수 2»(한 배관을 자른 것) 105
      관경 역전 26건 중 «조각이 낀» 것 12

  ★긴 45° 30개는 **실제 대각 주행**이라 건드리면 안 된다(접기 지시서 §6).
    챔퍼와 주행은 **길이로 갈린다.**

■ 여기서 재는 것 (손질판 = 접기가 도는 그 자리)

    · 「직선 – 짧은 45° – 직선」 꼴이 몇 곳인가
    · 그 양옆 두 간선이 **직교**인가 (직교라야 모서리로 복원할 수 있다)
    · 복원하면 길이가 얼마나 느는가 (모서리를 도는 만큼 — 보수측)
    · 직교가 아닌 것은 몇 개인가 (그대로 둬야 한다)

    python scripts/_probe_chamfer.py [--key 저장본이름] [--max-mm 300]
"""
from __future__ import annotations

import argparse
import math
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


def wait(c, sid, limit=40000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _ang(a, b):
    """0~180. 가로 0 · 세로 90."""
    return math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])) % 180.0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default="")
    ap.add_argument("--max-mm", type=float, default=300.0)
    ap.add_argument("--tol", type=float, default=3.0)
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        if args.key:
            r = c.post("/api/module-f/reopen", json={"key": args.key})
            j0 = r.get_json() or {}
            if not j0.get("ok"):
                print(f"★이어서 열기 실패 — {str(j0)[:200]}")
                return 1
            sid = j0["sid"]
            wait(c, sid)
        else:
            if not PLAN.is_file():
                print(f"표본 없음: {PLAN}")
                return 0
            with open(PLAN, "rb") as fh:
                r = c.post("/api/module-f/slot/open",
                           data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
                           content_type="multipart/form-data")
            sid = (r.get_json() or {})["sid"]
            wait(c, sid)
            c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "manual"})
            rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
                   .get("recon") or {})
            c.post("/api/module-f/pick/adopt",
                   json={"sid": sid, "materials": True,
                         "heads": {"conf_min": (rec.get("adopt") or {})
                                   .get("conf_min")}})
            wait(c, sid)
            c.post("/api/module-f/pick/commit", json={"sid": sid})
            wait(c, sid)

        from routes.module_f.jobs import _sess
        es = _sess(sid)["edit"]
        _report(es.board, args)
    return 0


def _report(board, args):
    pts = list(board.pts)
    edges = {tuple(sorted(e)) for e in board.edges}
    nb: dict = {}
    for i, j in edges:
        nb.setdefault(i, []).append(j)
        nb.setdefault(j, []).append(i)

    print(f"\n■ 손질판 챔퍼 계측 · 점 {len(pts)} · 간선 {len(edges)}"
          f" · 자 {args.max_mm:.0f} mm · ±{args.tol}°")

    lens = [math.dist(pts[i], pts[j]) for i, j in edges]
    short = [(i, j) for (i, j) in edges
             if math.dist(pts[i], pts[j]) < args.max_mm]
    print(f"  간선 길이 중앙값 {sorted(lens)[len(lens) // 2]:.0f} mm"
          f" · {args.max_mm:.0f} mm 미만 {len(short)}")

    def is45(a, b):
        v = _ang(pts[a], pts[b])
        return abs(v - 45) <= args.tol or abs(v - 135) <= args.tol

    def is_ortho(a, b):
        v = _ang(pts[a], pts[b])
        return (v <= args.tol or v >= 180 - args.tol
                or abs(v - 90) <= args.tol)

    kinds = Counter()
    ok, grow = [], []
    for (i, j) in short:
        if not is45(i, j):
            kinds["45°가 아님"] += 1
            continue
        if len(nb.get(i, ())) != 2 or len(nb.get(j, ())) != 2:
            kinds["양 끝 차수≠2"] += 1
            continue
        pi = [n for n in nb[i] if n != j][0]
        pj = [n for n in nb[j] if n != i][0]
        # 양옆 두 간선이 **직교** 라야 모서리로 복원할 수 있다.
        if not (is_ortho(pi, i) and is_ortho(j, pj)):
            kinds["양옆이 직교가 아님"] += 1
            continue
        a1 = _ang(pts[pi], pts[i])
        a2 = _ang(pts[j], pts[pj])
        if abs(a1 - a2) < args.tol or abs(abs(a1 - a2) - 180) < args.tol:
            kinds["양옆이 서로 평행(모서리 아님)"] += 1
            continue
        # 교점 — 두 직선을 연장해 만나는 자리
        x = _meet(pts[pi], pts[i], pts[j], pts[pj])
        if x is None:
            kinds["교점을 못 구함"] += 1
            continue
        before = (math.dist(pts[pi], pts[i]) + math.dist(pts[i], pts[j])
                  + math.dist(pts[j], pts[pj]))
        after = math.dist(pts[pi], x) + math.dist(x, pts[pj])
        kinds["★모서리로 복원 가능"] += 1
        ok.append((i, j, pi, pj, x))
        grow.append(after - before)

    print(f"  {dict(kinds)}")
    if grow:
        grow.sort()
        print(f"  ★복원 가능 {len(ok)}곳 · 길이 변화 중앙값"
              f" {grow[len(grow) // 2]:+.1f} mm"
              f" (최소 {grow[0]:+.1f} · 최대 {grow[-1]:+.1f})"
              f" · 합 {sum(grow):+,.0f} mm")
        print("  (모서리를 도는 만큼 늘어난다 — 줄지 않으므로 보수측이다)")
    # ★가장 큰 벽 — 「양 끝 차수≠2」 가 무엇인지 펴 본다.
    stuck = []
    for (i, j) in short:
        if not is45(i, j):
            continue
        if len(nb.get(i, ())) == 2 and len(nb.get(j, ())) == 2:
            continue
        stuck.append((i, j))
    print(f"\n  ── 「양 끝 차수≠2」 {len(stuck)}곳의 정체")
    degpair = Counter((len(nb.get(i, ())), len(nb.get(j, ())))
                      for i, j in stuck)
    print(f"    차수쌍 분포 {dict(sorted(degpair.items()))}")
    # 그 셋째 이웃이 무엇인가 — 또 짧은 45°(=두 줄 그림)인가, 긴 배관인가
    third = Counter()
    for i, j in stuck:
        for n, other in ((i, j), (j, i)):
            for m in nb.get(n, ()):
                if m == other:
                    continue
                d = math.dist(pts[n], pts[m])
                if d >= args.max_mm:
                    third["긴 배관"] += 1
                elif is45(n, m):
                    third["★또 짧은 45°"] += 1
                elif is_ortho(n, m):
                    third["짧은 직교"] += 1
                else:
                    third["짧은 기타"] += 1
    print(f"    그 자리에 함께 붙은 간선 {dict(third)}")
    for i, j in stuck[:4]:
        print(f"      {i}({pts[i][0]:.0f},{pts[i][1]:.0f}) 차수 {len(nb[i])}"
              f" – {j}({pts[j][0]:.0f},{pts[j][1]:.0f}) 차수 {len(nb[j])}"
              f" · 챔퍼 {math.dist(pts[i], pts[j]):.0f} mm")
        for n in (i, j):
            for m in nb[n]:
                print(f"          {n}–{m} {math.dist(pts[n], pts[m]):7.1f} mm"
                      f" · {_ang(pts[n], pts[m]):5.1f}°")

    for i, j, pi, pj, x in ok[:5]:
        print(f"      {pi}({pts[pi][0]:.0f},{pts[pi][1]:.0f})"
              f" – {i} – {j} – {pj}({pts[pj][0]:.0f},{pts[pj][1]:.0f})"
              f"  →  모서리 ({x[0]:.0f},{x[1]:.0f})"
              f" · 챔퍼 {math.dist(pts[i], pts[j]):.0f} mm")
    return ok


def _meet(a1, a2, b1, b2):
    """두 직선의 교점. 평행이면 None."""
    d1 = (a2[0] - a1[0], a2[1] - a1[1])
    d2 = (b2[0] - b1[0], b2[1] - b1[1])
    den = d1[0] * d2[1] - d1[1] * d2[0]
    if abs(den) < 1e-9:
        return None
    t = ((b1[0] - a1[0]) * d2[1] - (b1[1] - a1[1]) * d2[0]) / den
    return (a1[0] + t * d1[0], a1[1] + t * d1[1])


if __name__ == "__main__":
    raise SystemExit(main())

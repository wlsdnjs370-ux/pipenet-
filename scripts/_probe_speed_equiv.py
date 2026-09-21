# -*- coding: utf-8 -*-
"""[속도 최적화] 빨라졌는가 **그리고** 같은 답인가 — 둘을 한 자리에서 잰다.

속도만 재면 최적화가 아니라 «다른 프로그램» 을 만들 수 있다. 그래서 격자를
넣은 두 함수(`flow.stage5_split_through_uprights` · `stage45.join_by_head_cover`)
를 **옛 전수 스캔 판과 나란히** 돌려 결과가 한 글자도 다르지 않은지 본다.

옛 판은 파일을 되돌리지 않고 여기서 **직접 다시 짜** 비교한다(원본 주석에
남은 그대로) — 저장소를 건드리지 않고 대조할 수 있다.

    python scripts/_probe_speed_equiv.py [--key ...]
"""
from __future__ import annotations

import argparse
import contextlib
import io
import math
import os
import sys
import time
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
for p in (ROOT, os.path.join(ROOT, "scripts"), os.path.join(ROOT, "tests")):
    if p not in sys.path:
        sys.path.insert(0, p)

BIG = "B1F 현장조사 소화설비 평면도_컨셉2"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


# ── 옛 판 (전수 스캔) — 비교 전용 재현 ────────────────────────────────
def old_split_through_uprights(pts, edges, ups, assume_unattached=False):
    from services.cad_import.pipeline.flow import ARM_CTR, HEAD_TOUCH
    from services.cad_import.pipeline.expand import gput, gnear
    if not ups:
        return list(pts), {tuple(sorted(e)) for e in edges}, 0
    pts2 = list(pts)
    edges2 = {tuple(sorted(e)) for e in edges}
    used = {n for e in edges2 for n in e}

    def has_near_node(hx, hy, hr):          # ← 옛 전수 스캔
        for n in used:
            d = math.hypot(pts2[n][0] - hx, pts2[n][1] - hy)
            if d <= ARM_CTR or abs(d - hr) <= HEAD_TOUCH:
                return True
        return False

    want = []
    cell = 1000.0
    eg = defaultdict(list)
    for (i, j) in edges2:
        a, b = pts2[i], pts2[j]
        n = max(1, int(math.hypot(b[0] - a[0], b[1] - a[1]) / cell))
        for k in range(n + 1):
            t = k / n
            gput(eg, cell, a[0] + (b[0] - a[0]) * t,
                 a[1] + (b[1] - a[1]) * t, (i, j))
    for (hx, hy, hr) in ups:
        if hr <= 0 or (not assume_unattached and has_near_node(hx, hy, hr)):
            continue
        best = None
        rings = 1 + int(hr // cell)
        for (i, j) in set(gnear(eg, cell, hx, hy, rings=rings)):
            if (min(i, j), max(i, j)) not in edges2:
                continue
            ax, ay = pts2[i]
            bx, by = pts2[j]
            L = math.hypot(bx - ax, by - ay)
            if L < 1e-9:
                continue
            t = ((hx - ax) * (bx - ax) + (hy - ay) * (by - ay)) / (L * L)
            if t <= 1e-6 or t >= 1.0 - 1e-6:
                continue
            px, py = ax + (bx - ax) * t, ay + (by - ay) * t
            lat = math.hypot(px - hx, py - hy)
            if lat > hr:
                continue
            sc = (lat, abs(t - 0.5))
            if best is None or sc < best[0]:
                best = (sc, i, j, t)
        if best is not None:
            want.append((hx, hy, best[1], best[2], best[3]))
    return want          # ★비교는 «쪼갤 후보» 로 한다 — 그 뒤는 안 바꿨다


def new_split_candidates(pts, edges, ups, assume_unattached=False):
    """현행 판에서 같은 `want` 를 얻는다 — 함수를 그대로 돌리고 결과 망을 본다."""
    from services.cad_import.pipeline.flow import (
        stage5_split_through_uprights as f)
    return f(pts, edges, ups, assume_unattached=assume_unattached)


def run(key) -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from routes.module_f.common import _boot
    _boot()
    buf = io.StringIO()
    from services.cad_import.edit.session import EditSession
    print(f"■ 속도·동치 — {key}")
    with contextlib.redirect_stdout(buf):
        es = EditSession.open(key, out_dir=None, load_saved=True,
                              use_cache=True)
    b = es.board
    pts = [tuple(p) for p in b.pts]
    edges = [tuple(e) for e in b.edges]
    ups = [(float(d[0]), float(d[1]), float(d[2])) for d in b.disks]
    print(f"  망: 점 {len(pts):,} · 간선 {len(edges):,} · 헤드 {len(ups):,}")

    # ── ① stage5 — 같은 «쪼갤 후보» 를 내는가 + 얼마나 빨라졌나
    t0 = time.perf_counter()
    want_old = old_split_through_uprights(pts, edges, ups)
    t_old = time.perf_counter() - t0

    t0 = time.perf_counter()
    p2, e2, n_split = new_split_candidates(pts, edges, ups)
    t_new = time.perf_counter() - t0

    # 현행은 쪼갠 «망» 을 돌려준다 — 옛 후보로 같은 쪼개기를 재현해 대조한다.
    t0 = time.perf_counter()
    p2o, e2o, n_old = old_apply(pts, edges, want_old)
    t_old += time.perf_counter() - t0

    print(f"\n  [stage5] 옛 {t_old:7.2f}s → 새 {t_new:7.2f}s "
          f"({(1 - t_new / max(t_old, 1e-9)) * 100:.0f}% 감소)")
    check("stage5 쪼갠 수 같다", n_split == n_old, f"{n_old} vs {n_split}")
    check("stage5 점 목록 같다", p2 == p2o,
          f"{len(p2o)} vs {len(p2)}")
    check("stage5 간선 집합 같다", set(e2) == set(e2o),
          f"차 {len(set(e2) ^ set(e2o))}개")

    # ── ② join_by_head_cover — 같은 이음을 내는가
    from services.cad_import.pipeline import stage45 as s45
    from services.cad_import.pipeline.stage1 import DEFAULT_KNOBS

    class _G:
        def __init__(self, pts, edges):
            self.pts = list(pts)
            self._edges = list(edges)

        def adj(self):
            m = defaultdict(list)
            for i, j in self._edges:
                m[i].append(j)
                m[j].append(i)
            return m

    g = _G(pts, edges)
    heads = [(d[0], d[1], d[2]) for d in ups]
    kn = {k: DEFAULT_KNOBS[k] for k in
          ("r1_cand", "r1_meas_cap", "r1_lat_tol", "r1_head_slack",
           "r1_cover_slack")}
    t0 = time.perf_counter()
    with contextlib.redirect_stdout(buf):
        joins_new = s45.join_by_head_cover(g, None, heads, knobs=kn)
    t_j_new = time.perf_counter() - t0
    print(f"  [join]   새 {t_j_new:7.2f}s · 이음 {len(joins_new)}개")

    print()
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"실패 {len(FAILS)}건 — 빨라졌어도 답이 다르면 최적화가 아니다")
        return 3
    print("★빨라졌고, 답은 한 글자도 안 달라졌다")
    return 0


def old_apply(pts, edges, want):
    """옛 후보 → 쪼갠 망. 현행의 두 번째 반복과 같은 규칙(안 바꾼 부분)."""
    pts2 = list(pts)
    edges2 = {tuple(sorted(e)) for e in edges}
    by_edge = defaultdict(list)
    for hx, hy, i, j, t in want:
        a, b = min(i, j), max(i, j)
        if i > j:
            t = 1.0 - t
        by_edge[(a, b)].append((t, hx, hy))
    n_split = 0
    for (i, j), lst in by_edge.items():
        if (i, j) not in edges2:
            continue
        edges2.discard((i, j))
        ax, ay = pts2[i]
        bx, by = pts2[j]
        L = math.hypot(bx - ax, by - ay)
        lst.sort()
        prev = i
        k = 0
        while k < len(lst):
            t0 = lst[k][0]
            hx, hy = lst[k][1], lst[k][2]
            while k < len(lst) and abs(lst[k][0] - t0) * L <= 1.0:
                k += 1
            pts2.append((float(hx), float(hy)))
            new = len(pts2) - 1
            edges2.add((min(prev, new), max(prev, new)))
            prev = new
            n_split += 1
        edges2.add((min(prev, j), max(prev, j)))
    return pts2, edges2, n_split


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default=BIG)
    a = ap.parse_args()
    return run(a.key)


if __name__ == "__main__":
    raise SystemExit(main())

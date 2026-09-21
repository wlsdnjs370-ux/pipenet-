"""평면도 교차점 실태조사 — 각 X교차가 "접속/통과" 중 무엇인지 판정할 증거가
도면에 실제로 얼마나 남아 있는지 세는 일회성 측정 스크립트.

usage:  python scripts/_crossing_survey.py [dxf ...]
"""
from __future__ import annotations

import math
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from remote30_prototype import (  # noqa: E402
    MIN_PIPE_EDGE_MM,
    SNAP_TOL_MM,
    CLOSED_PL_TOL_MM,
    KNOWN_HEAD_BLOCKS,
    parse_dxf_bundle,
)

SYM_R_MM = 100.0        # 이음쇠 기호 게이트 반경 (POC3 gate_r=60 보다 관대하게)
BREAK_MAX_MM = 600.0    # "끊어 그리기" 로 인정할 최대 갭
BREAK_OFF_MM = 40.0     # 공선 판정 수직 허용오차
BREAK_ANG_DEG = 2.0     # 공선 판정 각도 허용오차
PARALLEL_SIN = 0.087    # sin 5° — 이보다 평행하면 교차 판정 제외
HEAD_NEAR_MM = 1500.0   # 선분이 "헤드를 달고 있다" 고 볼 거리

_NON_PIPE_CATS = {"HEAD", "TEXT", "ALARM", "ARCH", "EXCLUDE"}


def pipe_segments(entities, layer_cat, broad: bool):
    """_build_graph 와 같은 필터로 raw 선분 목록을 뽑는다 (그래프 병합 전)."""
    def is_pipe(layer: str) -> bool:
        cat = layer_cat.get(layer, "OTHER")
        return cat not in _NON_PIPE_CATS if broad else cat == "PIPE"

    segs = []
    min_sq = MIN_PIPE_EDGE_MM ** 2
    for en in entities:
        ly = en.get("l", "")
        if not is_pipe(ly):
            continue
        t = en["t"]
        if t == "L":
            x1, y1, x2, y2 = en["p"]
            if (x2 - x1) ** 2 + (y2 - y1) ** 2 >= min_sq:
                segs.append((x1, y1, x2, y2, ly))
        elif t == "PL":
            pts = en["p"]
            if len(pts) < 2:
                continue
            if len(pts) >= 3 and math.hypot(pts[0][0] - pts[-1][0],
                                            pts[0][1] - pts[-1][1]) <= CLOSED_PL_TOL_MM:
                continue
            for p0, p1 in zip(pts, pts[1:]):
                if (p1[0] - p0[0]) ** 2 + (p1[1] - p0[1]) ** 2 >= min_sq:
                    segs.append((p0[0], p0[1], p1[0], p1[1], ly))
    return segs


def grid_index(segs, cell):
    inv = 1.0 / cell
    grid = defaultdict(list)
    for si, (ax, ay, bx, by, _ly) in enumerate(segs):
        n = int(math.hypot(bx - ax, by - ay) * inv) + 1
        for i in range(n + 1):
            t = i / n
            grid[(int(math.floor((ax + t * (bx - ax)) * inv)),
                  int(math.floor((ay + t * (by - ay)) * inv)))].append(si)
    return grid, inv


def candidate_pairs(segs, cell=2000.0):
    grid, _inv = grid_index(segs, cell)
    seen = set()
    for bucket in grid.values():
        for i in range(len(bucket)):
            for j in range(i + 1, len(bucket)):
                a, b = bucket[i], bucket[j]
                key = (a, b) if a < b else (b, a)
                if key not in seen:
                    seen.add(key)
                    yield key


def intersect(s1, s2):
    """교차점 (x, y, t, u) 또는 None. 평행/거의평행은 None."""
    ax, ay, bx, by = s1[:4]
    cx, cy, dx_, dy_ = s2[:4]
    r = (bx - ax, by - ay)
    s = (dx_ - cx, dy_ - cy)
    den = r[0] * s[1] - r[1] * s[0]
    lr = math.hypot(*r)
    ls = math.hypot(*s)
    if lr < 1e-9 or ls < 1e-9:
        return None
    if abs(den) / (lr * ls) < PARALLEL_SIN:
        return None
    t = ((cx - ax) * s[1] - (cy - ay) * s[0]) / den
    u = ((cx - ax) * r[1] - (cy - ay) * r[0]) / den
    if not (0.0 <= t <= 1.0 and 0.0 <= u <= 1.0):
        return None
    return ax + t * r[0], ay + t * r[1], t, u, lr, ls


def classify_hit(t, u, lr, ls, tol=SNAP_TOL_MM):
    """끝점 근처(=노드 접속)인지, 양쪽 모두 내부 관통(=X교차)인지."""
    e1 = min(t * lr, (1 - t) * lr) <= tol
    e2 = min(u * ls, (1 - u) * ls) <= tol
    if e1 and e2:
        return "node"        # 끝점끼리 — 그냥 이어진 곳
    if e1 or e2:
        return "tee"         # 한쪽 끝점이 다른 선 중간에 — T접속
    return "cross"           # 양쪽 다 내부 — 문제의 X교차


def collinear_breaks(segs):
    """공선 두 선분 사이의 갭(= 끊어 그리기 후보) 목록 [(gx1,gy1,gx2,gy2,w)]."""
    out = []
    for i, j in candidate_pairs(segs, cell=BREAK_MAX_MM * 2):
        a = segs[i]; b = segs[j]
        d1 = (a[2] - a[0], a[3] - a[1])
        d2 = (b[2] - b[0], b[3] - b[1])
        l1 = math.hypot(*d1); l2 = math.hypot(*d2)
        if l1 < 1e-9 or l2 < 1e-9:
            continue
        sin_ = abs(d1[0] * d2[1] - d1[1] * d2[0]) / (l1 * l2)
        if sin_ > math.sin(math.radians(BREAK_ANG_DEG)):
            continue
        # 가장 가까운 끝점 쌍
        best = None
        for p in ((a[0], a[1]), (a[2], a[3])):
            for q in ((b[0], b[1]), (b[2], b[3])):
                d = math.hypot(p[0] - q[0], p[1] - q[1])
                if best is None or d < best[0]:
                    best = (d, p, q)
        w, p, q = best
        if not (1.0 < w <= BREAK_MAX_MM):
            continue
        # 갭 방향이 선분 축과 같아야 공선 단절 (옆으로 나란한 두 줄 배제)
        gv = (q[0] - p[0], q[1] - p[1])
        cos_ = abs(gv[0] * d1[0] + gv[1] * d1[1]) / (w * l1)
        if cos_ < math.cos(math.radians(20.0)):
            continue
        # 수직 옵셋 (두 선의 축이 정말 같은 직선인지)
        off = abs((b[0] - a[0]) * d1[1] - (b[1] - a[1]) * d1[0]) / l1
        if off > BREAK_OFF_MM:
            continue
        out.append((p[0], p[1], q[0], q[1], w))
    return out


def survey(dxf: Path):
    bundle = parse_dxf_bundle(dxf)
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    segs = pipe_segments(bundle.entities, layer_cat, broad=False)
    mode = "PIPE-strict"
    if not segs:
        segs = pipe_segments(bundle.entities, layer_cat, broad=True)
        mode = "broad-fallback"

    # 기호(이음쇠) 후보 = INSERT + 작은 CIRCLE
    syms = []
    heads = []
    for en in bundle.entities:
        if en["t"] == "I":
            p = en.get("p") or en.get("c")
            if p:
                xy = (p[0], p[1])
                name = en.get("n", "") or en.get("b", "")
                cat = layer_cat.get(en.get("l", ""), "OTHER")
                (heads if (name in KNOWN_HEAD_BLOCKS or cat == "HEAD") else syms).append(xy)
        elif en["t"] == "C":
            r = float(en.get("r", 0) or 0)
            if r <= 300.0:
                c = en.get("c")
                if c:
                    (heads if layer_cat.get(en.get("l", ""), "OTHER") == "HEAD"
                     else syms).append((c[0], c[1]))

    sym_grid, sym_inv = grid_index([(x, y, x, y, "") for x, y in syms], SYM_R_MM * 4) \
        if syms else (defaultdict(list), 1.0 / (SYM_R_MM * 4))
    head_pts = heads

    counts = defaultdict(int)
    crosses = []
    for i, j in candidate_pairs(segs):
        hit = intersect(segs[i], segs[j])
        if hit is None:
            continue
        x, y, t, u, lr, ls = hit
        kind = classify_hit(t, u, lr, ls)
        counts[kind] += 1
        if kind == "cross":
            crosses.append((x, y, i, j))

    # 중복 제거 (같은 교차점이 여러 번 잡히는 경우)
    uniq = {}
    for x, y, i, j in crosses:
        uniq[(round(x / SNAP_TOL_MM), round(y / SNAP_TOL_MM))] = (x, y, i, j)
    crosses = list(uniq.values())

    # X교차별 증거 분류
    ev = defaultdict(int)
    for x, y, i, j in crosses:
        has_sym = False
        if syms:
            cg = (int(math.floor(x * sym_inv)), int(math.floor(y * sym_inv)))
            for dgx in (-1, 0, 1):
                for dgy in (-1, 0, 1):
                    for si in sym_grid.get((cg[0] + dgx, cg[1] + dgy), ()):
                        sx, sy = syms[si]
                        if math.hypot(sx - x, sy - y) <= SYM_R_MM:
                            has_sym = True
                            break
                    if has_sym:
                        break
                if has_sym:
                    break
        diff_layer = segs[i][4] != segs[j][4]
        near_head = any(math.hypot(hx - x, hy - y) <= HEAD_NEAR_MM for hx, hy in head_pts)
        tag = ("SYM" if has_sym else "") + ("|LAYER" if diff_layer else "") + \
              ("|HEAD" if near_head else "")
        ev[tag or "무증거"] += 1

    breaks = collinear_breaks(segs)
    # 갭을 관통하는 다른 선분이 있는가 = 끊어그린 통과 표기
    break_hits = 0
    gap_widths = []
    for gx1, gy1, gx2, gy2, w in breaks:
        gseg = (gx1, gy1, gx2, gy2, "")
        crossed = False
        for s in segs:
            if max(abs(s[0] - gx1), abs(s[1] - gy1)) > 20000:
                continue
            hit = intersect(gseg, s)
            if hit is None:
                continue
            _x, _y, t, u, lr, ls = hit
            if 0.05 < t < 0.95 and min(u * ls, (1 - u) * ls) > SNAP_TOL_MM:
                crossed = True
                break
        if crossed:
            break_hits += 1
            gap_widths.append(w)

    print(f"\n=== {dxf.name} ===")
    print(f"  선분 {len(segs)}개 ({mode}) · 기호후보 {len(syms)} · 헤드후보 {len(head_pts)}")
    print(f"  끝점-끝점 접속(node) : {counts['node']}")
    print(f"  T접속(끝점→중간)     : {counts['tee']}")
    print(f"  X교차(양쪽 관통)     : {len(crosses)}   ← 접속/통과 미상")
    for k, v in sorted(ev.items(), key=lambda kv: -kv[1]):
        print(f"      {k:<20} {v}")
    print(f"  공선 갭(≤{BREAK_MAX_MM:.0f}mm) : {len(breaks)}")
    print(f"  그중 다른 배관이 관통 = 끊어그린 통과 표기 : {break_hits}")
    if gap_widths:
        gap_widths.sort()
        n = len(gap_widths)
        print(f"      갭 폭 min/중앙/max = {gap_widths[0]:.0f} / "
              f"{gap_widths[n // 2]:.0f} / {gap_widths[-1]:.0f} mm")


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        args = [
            "samples/dxf/대명동201동 단위세대_layer정리.dxf",
            "samples/dxf/B1F 현장조사 소화설비 평면도.dxf",
            "samples/dxf/LH306동_평면도.dxf",
        ]
    for a in args:
        p = (ROOT / a) if not Path(a).is_absolute() else Path(a)
        if not p.exists():
            print(f"  (없음) {a}")
            continue
        try:
            survey(p)
        except Exception as exc:  # noqa: BLE001
            print(f"  (실패) {p.name}: {type(exc).__name__}: {exc}")

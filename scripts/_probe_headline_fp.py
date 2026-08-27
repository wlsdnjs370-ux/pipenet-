# -*- coding: utf-8 -*-
"""「헤드가 줄줄이 꿰인 선」 지문 — 배관을 이름이 아니라 «생김새» 로 찾는다.

R10 의 지문은 «헤드 틈»(런이 헤드마다 끊겨 그려진 자국)이다. 그런데 끊기지
않고 «옆으로 나란히» 헤드를 매단 배관도 있다 — B1F 의 `현장조사#셔터` 가 그렇다.
사람 눈에는 명백한 배관인데 지문이 안 잡혀 OTHER 로 남는다.

그래서 두 번째 지문을 잰다. 배관에 헤드가 달리면 세 가지가 한꺼번에 성립한다:

    ① 헤드가 여럿이다                     — 우연이 아니려면 수가 있어야 한다
    ② 선에서 «같은 거리» 에 있다           — 접속관 길이가 일정하다
    ③ 선을 따라 «규칙적인 간격» 으로 선다   — 헤드 간격은 설계값이다

건축선 위에 헤드가 우연히 놓이는 것과 갈리는 자리가 ②③ 이다. 벽선은 헤드가
몇 개 걸쳐도 거리도 간격도 제멋대로다.

    python scripts/_probe_headline_fp.py [도면.dxf] [--all]
"""
from __future__ import annotations

import argparse
import math
import statistics
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"

# 지문 문턱 — 아래 «판정» 에서 재는 값들.
MIN_HEADS = 5           # 이보다 적으면 우연을 못 가린다
OFF_MIN, OFF_MAX = 40.0, 400.0   # 접속관이 있을 법한 거리대
OFF_SPREAD_MAX = 40.0   # 거리가 이보다 흩어지면 «같은 거리» 가 아니다
PITCH_CV_MAX = 0.35     # 간격의 변동계수 — 이보다 크면 규칙적이지 않다
AXIS_TOL = 1.0


def _segs(e):
    t = str(e.get("t") or "")
    if t not in ("L", "PL"):
        return []
    p = e.get("p") or []
    if not p:
        return []
    if isinstance(p[0], (list, tuple)):
        pts = [(float(q[0]), float(q[1])) for q in p if len(q) >= 2]
    else:
        if len(p) < 4:
            return []
        pts = [(float(p[i]), float(p[i + 1])) for i in range(0, len(p) - 1, 2)]
    return [(pts[i], pts[i + 1]) for i in range(len(pts) - 1)]


def runs_of(segs):
    """축평행 선분을 «같은 축선» 끼리 묶는다 — 한 줄이 곧 한 배관 후보다."""
    band = defaultdict(list)
    for p, q in segs:
        if abs(p[0] - q[0]) <= AXIS_TOL and abs(p[1] - q[1]) > AXIS_TOL:
            band[(0, round(p[0], 1))].append((min(p[1], q[1]), max(p[1], q[1])))
        elif abs(p[1] - q[1]) <= AXIS_TOL and abs(p[0] - q[0]) > AXIS_TOL:
            band[(1, round(p[1], 1))].append((min(p[0], q[0]), max(p[0], q[0])))
    out = []
    for (ax, fixed), spans in band.items():
        spans.sort()
        lo, hi = spans[0]
        for s, e in spans[1:]:
            if s <= hi + 500.0:            # 한 줄로 보는 틈 (헤드 틈 포함)
                hi = max(hi, e)
            else:
                out.append((ax, fixed, lo, hi))
                lo, hi = s, e
        out.append((ax, fixed, lo, hi))
    return out


def score(run, hgrid, cell):
    """이 줄이 «헤드를 꿴 선» 인가 — (판정, 근거) 를 낸다."""
    ax, fixed, lo, hi = run
    L = hi - lo
    if L < 3000.0:
        return None
    picked = []
    if ax == 0:      # 세로줄 — 헤드의 x 가 fixed 근처
        g0, g1 = int((fixed - OFF_MAX) // cell), int((fixed + OFF_MAX) // cell)
        for gx in range(g0, g1 + 1):
            for gy in range(int(lo // cell), int(hi // cell) + 1):
                for hp in hgrid.get((gx, gy), ()):
                    off = abs(hp[0] - fixed)
                    if OFF_MIN <= off <= OFF_MAX and lo <= hp[1] <= hi:
                        picked.append((hp[1], hp[0] - fixed))
    else:            # 가로줄
        g0, g1 = int((fixed - OFF_MAX) // cell), int((fixed + OFF_MAX) // cell)
        for gy in range(g0, g1 + 1):
            for gx in range(int(lo // cell), int(hi // cell) + 1):
                for hp in hgrid.get((gx, gy), ()):
                    off = abs(hp[1] - fixed)
                    if OFF_MIN <= off <= OFF_MAX and lo <= hp[0] <= hi:
                        picked.append((hp[0], hp[1] - fixed))
    if len(picked) < MIN_HEADS:
        return None
    picked.sort()
    offs = [abs(o) for _, o in picked]
    spread = max(offs) - min(offs)
    alongs = [a for a, _ in picked]
    pitches = [b - a for a, b in zip(alongs, alongs[1:])]
    if not pitches:
        return None
    mean = statistics.fmean(pitches)
    cv = (statistics.pstdev(pitches) / mean) if mean > 1e-6 else 9.9
    ok = spread <= OFF_SPREAD_MAX and cv <= PITCH_CV_MAX
    return {"ok": ok, "n": len(picked), "len_m": L / 1000.0,
            "off": statistics.fmean(offs), "spread": spread,
            "pitch": mean, "cv": cv, "ax": ax, "fixed": fixed,
            "lo": lo, "hi": hi}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--all", action="store_true",
                    help="배관 아닌 레이어 전부 (기본은 셔터만)")
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1
    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    cat = {}
    for nm in {str(e.get("l") or "0") for e in ents}:
        try:
            cat[nm] = A._categorize_layer(nm)
        except Exception:  # noqa: BLE001
            cat[nm] = "OTHER"
    heads = A.detect_heads(ents, cat)
    CELL = 2000.0
    hgrid = defaultdict(list)
    for h in heads:
        hgrid[(int(h.pos[0] // CELL), int(h.pos[1] // CELL))].append(
            (h.pos[0], h.pos[1]))

    print(f"{dxf.name} · 헤드 {len(heads):,}")
    print(f"문턱 — 헤드 ≥{MIN_HEADS} · 거리 {OFF_MIN:.0f}~{OFF_MAX:.0f}mm · "
          f"거리흩어짐 ≤{OFF_SPREAD_MAX:.0f}mm · 간격변동 ≤{PITCH_CV_MAX}\n")

    by_layer = defaultdict(list)
    for e in ents:
        ly = str(e.get("l") or "0")
        if cat.get(ly) == "PIPE":
            continue
        if not a.all and "셔터" not in ly:
            continue
        by_layer[ly].extend(_segs(e))

    hits = defaultdict(list)
    for ly, segs in by_layer.items():
        for run in runs_of(segs):
            s = score(run, hgrid, CELL)
            if s and s["ok"]:
                hits[ly].append(s)

    if not hits:
        print("지문에 걸리는 줄이 없다.")
    else:
        print(f"{'헤드':>5} {'연장m':>8} {'거리mm':>8} {'흩어짐':>7} "
              f"{'간격mm':>8} {'변동':>6}  레이어")
        print("-" * 74)
        for ly, rows in sorted(hits.items(),
                               key=lambda kv: -sum(r["n"] for r in kv[1])):
            for r in sorted(rows, key=lambda r: -r["n"])[:6]:
                print(f"{r['n']:>5} {r['len_m']:>8.1f} {r['off']:>8.0f} "
                      f"{r['spread']:>7.0f} {r['pitch']:>8.0f} "
                      f"{r['cv']:>6.2f}  {ly}")
            if len(rows) > 6:
                print(f"{'':>5} … 외 {len(rows) - 6}줄  {ly}")
        tot = sum(r["n"] for rows in hits.values() for r in rows)
        print(f"\n  지문에 걸린 줄 {sum(len(v) for v in hits.values())}개 · "
              f"헤드 {tot}개 · 레이어 {len(hits)}종")

    # 이미 배관인 줄에도 같은 지문이 걸리는지 — 자가 맞는지 대조한다.
    pipe_segs = []
    for e in ents:
        if cat.get(str(e.get("l") or "0")) == "PIPE":
            pipe_segs.extend(_segs(e))
    ok = sum(1 for run in runs_of(pipe_segs)
             if (score(run, hgrid, CELL) or {}).get("ok"))
    print(f"  (대조) 이미 배관인 줄 중 같은 지문에 걸리는 것 {ok}개")
    return 0


if __name__ == "__main__":
    sys.exit(main())

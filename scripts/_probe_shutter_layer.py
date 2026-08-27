# -*- coding: utf-8 -*-
"""`현장조사#셔터` 를 레이어×색으로 갈라 «헤드가 붙는 선» 을 찾는다.

R10b 는 이 레이어에서 17 entity 만 배관으로 승격하고 4개를 OTHER 로 남겼다.
승격 근거는 «헤드 틈 지문»(런이 헤드마다 끊겨 그려진 자국)이다. 그런데 그
지문이 안 잡히는 배관도 있을 수 있다 — 헤드가 달려 있는데도.

그래서 다른 자로 재 본다: **그 선에 헤드가 몇 개나 붙는가.**
찍기(E)가 재료를 «레이어×색» 단위로 고르므로 여기서도 같은 단위로 가른다.

    python scripts/_probe_shutter_layer.py [도면.dxf] [--layer 셔터]
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


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


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--layer", default="셔터",
                    help="이 문자열이 든 레이어만 본다")
    ap.add_argument("--all", action="store_true",
                    help="배관으로 «안» 잡힌 레이어 전부를 훑는다")
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import remote30_prototype as A
    from remote30_graph import _point_to_segment_dist

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
    hpts = [(h.pos[0], h.pos[1]) for h in heads]
    print(f"{dxf.name} · 헤드 {len(hpts):,}\n")

    CELL = 2000.0
    hgrid = defaultdict(list)
    for hp in hpts:
        hgrid[(int(hp[0] // CELL), int(hp[1] // CELL))].append(hp)

    DROP = A.HEAD_DROP_MAX_MM

    def heads_on(segs):
        """이 선분들에 «달린» 헤드 수 — 결합선 상한(HEAD_DROP_MAX_MM) 기준."""
        got = set()
        for p, q in segs:
            x0, x1 = sorted((p[0], q[0]))
            y0, y1 = sorted((p[1], q[1]))
            for gx in range(int((x0 - DROP) // CELL), int((x1 + DROP) // CELL) + 1):
                for gy in range(int((y0 - DROP) // CELL),
                                int((y1 + DROP) // CELL) + 1):
                    for hp in hgrid.get((gx, gy), ()):
                        if hp in got:
                            continue
                        if _point_to_segment_dist(hp[0], hp[1], p[0], p[1],
                                                  q[0], q[1]) <= DROP:
                            got.add(hp)
        return len(got)

    # 레이어×색으로 가른다 — 찍기(E)가 재료를 고르는 그 단위다.
    groups = defaultdict(list)
    for e in ents:
        ly = str(e.get("l") or "0")
        if a.all:
            if cat.get(ly) == "PIPE":
                continue
        elif a.layer not in ly:
            continue
        c = e.get("c")
        if isinstance(c, list):          # 일부 entity 는 색이 목록으로 온다
            c = tuple(c)
        groups[(ly, c)].extend(_segs(e))

    if not groups:
        print(f"«{a.layer}» 가 든 레이어가 없습니다.")
        return 1

    rows = []
    for (ly, c), segs in groups.items():
        if not segs:
            continue
        L = sum(math.hypot(q[0] - p[0], q[1] - p[1]) for p, q in segs)
        rows.append((heads_on(segs), len(segs), L / 1000.0, ly, c,
                     cat.get(ly, "?")))
    rows.sort(reverse=True)

    print(f"{'헤드':>6} {'선분':>6} {'연장m':>9}  {'분류':<8} 레이어 × 색")
    print("-" * 78)
    for n, ns, lm, ly, c, ct in rows[:30]:
        star = "  ★헤드가 달렸는데 배관이 아니다" if (n >= 3 and ct != "PIPE") else ""
        print(f"{n:>6,} {ns:>6,} {lm:>9,.1f}  {ct:<8} {ly}×{c}{star}")

    tot = sum(r[0] for r in rows if r[5] != "PIPE")
    print(f"\n  배관이 아닌 묶음에 달린 헤드 합계 {tot:,}")
    print(f"  (헤드 결합선 상한 {DROP:.0f}mm 기준)")
    return 0


if __name__ == "__main__":
    sys.exit(main())

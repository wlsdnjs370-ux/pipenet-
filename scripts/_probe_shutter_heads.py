# -*- coding: utf-8 -*-
"""셔터선 둘레에 «검출기가 못 잡은» 헤드 기호가 있나.

「분명 이 선에도 헤드가 있다」는 말을 받았다. 지금까지 나는 `detect_heads` 가
낸 목록으로만 쟀는데, 그 목록에 안 든 기호가 있다면 내 측정이 통째로 헛것이다.

그래서 도형을 직접 센다 — 헤드 크기의 원, 짧은 십자 획, 삼각형 변. 검출된
헤드와 나란히 놓고 «검출기가 놓친 것» 이 있는지 본다.

    python scripts/_probe_shutter_heads.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import Counter, defaultdict
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
    ap.add_argument("--band", type=float, default=500.0)
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

    # 두 줄의 y 대역을 잡는다.
    prom_y, raw_y, xs = [], [], []
    for e in ents:
        ly = str(e.get("l") or "0")
        if "셔터" not in ly:
            continue
        for p, q in _segs(e):
            (prom_y if "승격" in ly else raw_y).extend([p[1], q[1]])
            xs.extend([p[0], q[0]])
    if not prom_y or not raw_y:
        print("두 줄이 다 있어야 비교가 된다.")
        return 1
    y_prom = sum(prom_y) / len(prom_y)
    y_raw = sum(raw_y) / len(raw_y)
    x0, x1 = min(xs), max(xs)
    print(f"승격분 y≈{y_prom:.0f} · 셔터선 y≈{y_raw:.0f} "
          f"(간격 {abs(y_prom - y_raw):.0f}mm)")
    print(f"x {x0:.0f}~{x1:.0f}\n")

    def near(y, py):
        return abs(py - y) <= a.band

    # ① 검출된 헤드
    det = Counter()
    for h in heads:
        if not (x0 <= h.pos[0] <= x1):
            continue
        if near(y_prom, h.pos[1]):
            det["승격분 쪽"] += 1
        elif near(y_raw, h.pos[1]):
            det["셔터선 쪽"] += 1

    # ② 도형 — 헤드 크기의 원 / 짧은 획 / 호
    circ = Counter()
    stroke = Counter()
    arc = Counter()
    for e in ents:
        t = str(e.get("t") or "")
        ly = str(e.get("l") or "0")
        p = e.get("p") or []
        if t == "C" and len(p) >= 2:
            cx, cy = float(p[0]), float(p[1])
            r = float(e.get("r") or 0)
            if not (x0 <= cx <= x1) or not (20 <= r <= 400):
                continue
            key = "승격분 쪽" if near(y_prom, cy) else (
                "셔터선 쪽" if near(y_raw, cy) else None)
            if key:
                circ[f"{key} · {ly}"] += 1
        elif t == "A" and len(p) >= 2:
            cx, cy = float(p[0]), float(p[1])
            if not (x0 <= cx <= x1):
                continue
            key = "승격분 쪽" if near(y_prom, cy) else (
                "셔터선 쪽" if near(y_raw, cy) else None)
            if key:
                arc[f"{key} · {ly}"] += 1
        else:
            for (sp, sq) in _segs(e):
                L = math.hypot(sq[0] - sp[0], sq[1] - sp[1])
                if not (30 <= L <= 400):
                    continue
                mx, my = (sp[0] + sq[0]) / 2, (sp[1] + sq[1]) / 2
                if not (x0 <= mx <= x1):
                    continue
                key = "승격분 쪽" if near(y_prom, my) else (
                    "셔터선 쪽" if near(y_raw, my) else None)
                if key:
                    stroke[f"{key} · {ly}"] += 1

    print(f"■ 검출된 헤드 (±{a.band:.0f}mm 대역)")
    for k, v in det.most_common():
        print(f"    {k:<10} {v:>5}")
    print(f"\n■ 헤드 크기 원 (r 20~400mm)")
    for k, v in circ.most_common(10):
        print(f"    {v:>5}  {k}")
    if not circ:
        print("    없음")
    print(f"\n■ 짧은 획 (30~400mm) — 십자·삼각 기호의 재료")
    for k, v in stroke.most_common(12):
        print(f"    {v:>5}  {k}")
    print(f"\n■ 호(ARC)")
    for k, v in arc.most_common(8):
        print(f"    {v:>5}  {k}")
    if not arc:
        print("    없음")
    return 0


if __name__ == "__main__":
    sys.exit(main())

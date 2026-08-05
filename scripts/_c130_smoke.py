# -*- coding: utf-8 -*-
"""C130 지문 수집 실도면 스모크 — 일회용.

단위 테스트는 합성 도형이라 성능을 못 잡는다. 평행쌍 탐색이 O(n^2) 로 퇴화하면
합성 테스트는 그대로 통과하고 실 도면에서만 멈춘다.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import ezdxf

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.design.recognize import geom_stats as G  # noqa: E402


def entities_of(path: Path) -> tuple[list, dict]:
    doc = ezdxf.readfile(str(path))
    out: list[dict] = []
    minx = miny = float("inf")
    maxx = maxy = float("-inf")

    def upd(x, y):
        nonlocal minx, miny, maxx, maxy
        minx, miny = min(minx, x), min(miny, y)
        maxx, maxy = max(maxx, x), max(maxy, y)

    for e in doc.modelspace():
        layer = e.dxf.layer
        t = e.dxftype()
        if t == "LINE":
            p = [e.dxf.start.x, e.dxf.start.y, e.dxf.end.x, e.dxf.end.y]
            out.append({"t": "L", "l": layer, "p": p})
            upd(p[0], p[1]); upd(p[2], p[3])
        elif t == "ARC":
            out.append({"t": "A", "l": layer, "c": [e.dxf.center.x, e.dxf.center.y],
                        "r": float(e.dxf.radius),
                        "a": [float(e.dxf.start_angle), float(e.dxf.end_angle)]})
            upd(e.dxf.center.x, e.dxf.center.y)
        elif t == "CIRCLE":
            out.append({"t": "C", "l": layer, "c": [e.dxf.center.x, e.dxf.center.y],
                        "r": float(e.dxf.radius)})
            upd(e.dxf.center.x, e.dxf.center.y)
        elif t == "LWPOLYLINE":
            pts = [[p[0], p[1]] for p in e.get_points()]
            if pts:
                out.append({"t": "PL", "l": layer, "p": pts})
                for x, y in pts:
                    upd(x, y)
        elif t in ("TEXT", "MTEXT"):
            raw = str(getattr(e, "text", "") or getattr(e.dxf, "text", ""))[:60]
            out.append({"t": "T", "l": layer, "p": [e.dxf.insert.x, e.dxf.insert.y], "v": raw})
            upd(e.dxf.insert.x, e.dxf.insert.y)
        elif t == "INSERT":
            out.append({"t": "I", "l": layer, "p": [e.dxf.insert.x, e.dxf.insert.y],
                        "n": str(e.dxf.name)})
            upd(e.dxf.insert.x, e.dxf.insert.y)
    return out, {"minx": minx, "miny": miny, "maxx": maxx, "maxy": maxy}


def main(path: Path) -> None:
    t0 = time.perf_counter()
    ents, bbox = entities_of(path)
    t1 = time.perf_counter()
    unit = G.suggest_unit_to_mm(bbox)
    print(f"{path.name}: entity {len(ents)}, 파싱 {t1 - t0:.1f}s, 단위제안 {unit}")

    fps = G.fingerprints(ents, unit_to_mm=(unit or {}).get("unit_to_mm") or 1.0, bbox=bbox)
    t2 = time.perf_counter()
    print(f"지문 {len(fps)}개, {t2 - t1:.1f}s\n")

    header = ("레이어", "n", "par", "peaks", "len중앙", "closed", "rep", "grid",
              "arc붙", "문반경", "계단", "숫자", "장선")
    print("{:<28}{:>7}{:>6}{:>16}{:>9}{:>8}{:>6}{:>6}{:>6}{:>7}{:>6}{:>6}{:>6}".format(*header))
    for fp in sorted(fps, key=lambda f: -f.n_entities)[:20]:
        print("{:<28}{:>7}{:>6.2f}{:>16}{:>9.0f}{:>8}{:>6.2f}{:>6.2f}{:>6.2f}{:>7.2f}{:>6}{:>6.2f}{:>6.2f}".format(
            fp.name[:27], fp.n_entities, fp.parallel_pair_ratio,
            str(fp.offset_peaks_mm)[:15], fp.len_median_mm, fp.closed_shape_count,
            fp.closed_repeat_score, fp.grid_alignment_score, fp.arc_attach_ratio,
            fp.door_radius_ratio, fp.stair_bundle_max, fp.text_numeric_ratio,
            fp.long_line_ratio))


if __name__ == "__main__":
    main(Path(sys.argv[1]))

# -*- coding: utf-8 -*-
"""C130 지문 + C140 판정 실도면 스모크 — 일회용.

단위 테스트는 합성 도형이라 두 가지를 못 잡는다. 성능(평행쌍 탐색이 O(n^2) 로
퇴화하면 합성 테스트는 그대로 통과하고 실 도면에서만 멈춘다)과, 실 도면 레이어가
판정표의 "레이어=한 카테고리" 가정을 지키지 않는다는 사실이다.
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

import ezdxf

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.design.recognize import arch_category as C  # noqa: E402
from core.design.recognize import core_detect as K  # noqa: E402
from core.design.recognize import geom_stats as G  # noqa: E402
from core.design.recognize import opening_close as O  # noqa: E402
from core.design.recognize import room_faces as F  # noqa: E402
from core.design.recognize import room_label as L  # noqa: E402
from core.design.recognize import wall_centerline as W  # noqa: E402


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


def main(path: Path, wall_override: set = frozenset()) -> None:
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

    print("\n── C140 판정 ────────────────────────────────────────────────")
    verdicts = {}
    for fp in sorted(fps, key=lambda f: -f.n_entities):
        verdicts[fp.name] = C.evaluate(fp)
    for fp in sorted(fps, key=lambda f: -f.n_entities)[:20]:
        v = verdicts[fp.name]
        alt = " / 대안 " + ", ".join(f"{c} {s:.2f}" for c, s in v.alternatives) if v.alternatives else ""
        print(f"{fp.name[:27]:<28}{v.category:<11}{v.confidence:.2f}{alt}")
        for text in v.provenance:
            print(f"      + {text}")
        for text in v.near_misses:
            print(f"      - {text}")

    unit_to_mm = (unit or {}).get("unit_to_mm") or 1.0
    pair = _c150_c160(ents, fps, verdicts, unit_to_mm, t2, wall_override)
    if pair is not None:
        _c170_c190(ents, fps, verdicts, unit_to_mm, bbox, *pair)


def _c150_c160(ents, fps, verdicts, unit_to_mm, t2, wall_override):
    """C140 이 WALL 로 본 레이어를 중심선화하고, DOOR 레이어의 호로 간극을 닫는다.

    `wall_override` 는 GATE 가 사람 손으로 WALL 을 확정하는 자리를 대신한다. 실
    도면에서 WALL 이 한 곳도 안 나오는 경우가 있어(§3.2 판정표의 한계) C150 을
    실 기하로 돌려보려면 이 통로가 필요하다.
    """
    walls = ({fp.name for fp in fps if verdicts[fp.name].category == C.WALL}
             | {fp.name for fp in fps if fp.name in wall_override})
    doors = {fp.name for fp in fps if verdicts[fp.name].category == C.DOOR}
    print(f"\n── C150/C160 ─ WALL {sorted(walls)} / DOOR {sorted(doors)}")
    if not walls:
        print("   WALL 레이어가 없다 — 중심선화할 대상이 없다")
        return None

    peaks: list[float] = []
    for fp in fps:
        if fp.name in walls:
            peaks += fp.offset_peaks_mm
    lines = [e["p"] for e in ents if e["t"] == "L" and e["l"] in walls]
    arcs = [e for e in ents if e["t"] == "A" and e["l"] in doors]

    result = W.build_centerlines(lines, offset_peaks_mm=peaks, unit_to_mm=unit_to_mm)
    t3 = time.perf_counter()
    paired = sum(1 for c in result.centerlines if not c.unpaired)
    print(f"   중심선 {len(result.centerlines)}개 (짝지음 {paired}), 노드 {len(result.nodes)}, "
          f"{t3 - t2:.1f}s")
    for line in result.provenance:
        print(f"      + {line}")

    closure = O.close_openings(result, arcs, unit_to_mm=unit_to_mm)
    t4 = time.perf_counter()
    kinds: dict = {}
    for edge in closure.virtual_edges:
        kinds[edge.kind] = kinds.get(edge.kind, 0) + 1
    print(f"   가상 간선 {len(closure.virtual_edges)}개 {kinds}, {t4 - t3:.1f}s")
    for line in closure.provenance:
        print(f"      + {line}")
    for edge in closure.virtual_edges[:3]:
        print(f"      · {edge.kind} {edge.confidence:.2f} gap {edge.gap_mm:.0f}mm — "
              f"{edge.evidence[0]}")
    return result, closure, t4


def _c170_c190(ents, fps, verdicts, unit_to_mm, bbox, result, closure, t4) -> None:
    """실 폴리곤 → 실명 → 코어. 단층 도면이라 C190 은 이름 힌트 경로만 돈다."""
    bbox_area = ((bbox["maxx"] - bbox["minx"]) * (bbox["maxy"] - bbox["miny"])
                 * unit_to_mm * unit_to_mm)
    faces = F.build_faces(result, closure, bbox_area_mm2=bbox_area)
    t5 = time.perf_counter()
    print(f"\n── C170 ─ 실 {len(faces.faces)}개, {t5 - t4:.1f}s")
    for line in faces.provenance:
        print(f"      + {line}")
    for face in faces.faces[:5]:
        flags = (" " + ",".join(face.flags)) if face.flags else ""
        print(f"      · {face.area_m2:8.1f}㎡  변 {face.edge_count:3d}  "
              f"conf {face.confidence:.2f}  가상 {face.virtual_ratio:.0%}{flags}")

    rooms = {fp.name for fp in fps if verdicts[fp.name].category == C.ROOM_TEXT}
    texts = [e for e in ents if e["t"] == "T" and e["l"] in rooms]
    lines = [e for e in ents if e["t"] == "L"]
    labels = L.assign_labels(faces.faces, texts, lines, unit_to_mm=unit_to_mm)
    t6 = time.perf_counter()
    print(f"\n── C180 ─ ROOM_TEXT {sorted(rooms)} / 텍스트 {len(texts)}개, "
          f"{t6 - t5:.1f}s")
    for line in labels.provenance:
        print(f"      + {line}")
    named = [lb for lb in labels.labels if lb.name]
    for lb in named[:5]:
        print(f"      · face {lb.face_index} '{lb.name}' ({lb.source} "
              f"{lb.confidence:.2f}) 용도 {lb.use_hint}")

    names = [[lb.name for lb in labels.labels]]
    cores = K.detect_cores([faces.faces], names=names)
    t7 = time.perf_counter()
    print(f"\n── C190 ─ 코어 후보 {len(cores.candidates)}개, {t7 - t6:.1f}s")
    for line in cores.provenance:
        print(f"      + {line}")
    for cand in cores.candidates[:5]:
        print(f"      · face {cand.face_index} {cand.kind} {cand.confidence:.2f} "
              f"— {cand.evidence[0]}")


if __name__ == "__main__":
    main(Path(sys.argv[1]), set(sys.argv[2:]))

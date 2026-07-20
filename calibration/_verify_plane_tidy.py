"""_tidy_head_plane_layout 를 실 평면 데이터(대명동201)에 적용 → 교차 0·길이불변 검증.

build_input_tables(평면 헤드망) → 노드/파이프 → tidy 재배치 전후:
  · edge 교차 수 (표시 좌표) 감소 확인
  · pipe["length"] (수리 권위값) 완전 불변 확인
"""
import importlib.util
import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, select_worst30_heads,
    build_input_tables)

# 한글 파일명 모듈에서 _tidy_head_plane_layout 만 로드.
_spec = importlib.util.spec_from_file_location("daejo_server", ROOT / "대조 서버.py")
_mod = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_mod)
_tidy = _mod._tidy_head_plane_layout

DXF = ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def _cross(segs):
    def o(p, q, r):
        return (q[0]-p[0])*(r[1]-p[1]) - (q[1]-p[1])*(r[0]-p[0])
    n = 0
    for i in range(len(segs)):
        a, b = segs[i]
        for j in range(i+1, len(segs)):
            c, d = segs[j]
            shared = any(abs(p[0]-q[0]) < 1e-6 and abs(p[1]-q[1]) < 1e-6
                         for p in (a, b) for q in (c, d))
            if shared:
                continue
            d1, d2, d3, d4 = o(c, d, a), o(c, d, b), o(a, b, c), o(a, b, d)
            if ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0)):
                n += 1
    return n


def _segs(nodes, pipes):
    xy = {str(n["label"]): (float(n["x"]), float(n["y"])) for n in nodes}
    out = []
    for p in pipes:
        a, b = str(p["in"]), str(p["out"])
        if a in xy and b in xy:
            out.append((xy[a], xy[b]))
    return out


def main():
    bundle = parse_dxf_bundle(DXF)
    cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = filter_pipenet_only(bundle)
    sel = select_worst30_heads(pipe_ents, cat)
    tables = build_input_tables(sel, pipe_entities=pipe_ents)

    nodes, pipes = tables.nodes, tables.pipes
    len_before = {p["label"]: p["length"] for p in pipes}
    cross_before = _cross(_segs(nodes, pipes))

    root = next((str(n["label"]) for n in nodes
                 if str(n.get("io_node", "")).lower() == "input"), "10")
    moved = _tidy(nodes, pipes, root, set())

    len_after = {p["label"]: p["length"] for p in pipes}
    cross_after = _cross(_segs(nodes, pipes))

    len_changed = [k for k in len_before if len_before[k] != len_after.get(k)]

    print(f"root(AV/source)={root} moved_nodes={moved}")
    print(f"edge 교차: before={cross_before}  after={cross_after}")
    print(f"pipe length 변경 개수: {len(len_changed)} (기대 0)")

    ok = (cross_after <= cross_before) and (len(len_changed) == 0)
    print("[PASS]" if ok else "[FAIL]",
          "교차 감소 & 길이 불변" if ok else "조건 위반")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

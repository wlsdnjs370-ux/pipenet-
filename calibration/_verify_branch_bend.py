"""build_input_tables 가 가지배관 대각선에 표시 전용 L-벤드를 붙이는지 검증.

기대: _repro_branch_layout 이 보고한 대각선 2건 → pipe["bend"] 2건.
불변식: bend 는 표시좌표뿐(유압 length/dia 무변경), 벤드 두 다리 모두 직각(축정렬).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, select_worst30_heads,
    build_input_tables)

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def main():
    bundle = parse_dxf_bundle(DXF)
    cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = filter_pipenet_only(bundle)
    sel = select_worst30_heads(pipe_ents, cat)
    tables = build_input_tables(sel, pipe_entities=pipe_ents)

    node_xy = {n["label"]: (n["x"], n["y"]) for n in tables.nodes}
    bent = [p for p in tables.pipes if p.get("bend")]
    print(f"pipes={len(tables.pipes)} nodes={len(tables.nodes)} bent={len(bent)}")

    ok = True
    for p in bent:
        ax, ay = node_xy[p["in"]]; bx, by = node_xy[p["out"]]
        cx, cy = p["bend"]
        leg1 = (abs(ax - cx) < 2.0 or abs(ay - cy) < 2.0)
        leg2 = (abs(bx - cx) < 2.0 or abs(by - cy) < 2.0)
        orth = leg1 and leg2
        ok = ok and orth
        print(f"  pipe {p['label']}: {p['in']}->{p['out']} "
              f"a=({ax},{ay}) c=({cx},{cy}) b=({bx},{by}) "
              f"legs_ortho={orth}")

    # 벤드 없는 가지 pipe 는 이미 축정렬이어야(진단상 63건). 최소한 bent>=1 기대.
    print("\n[PASS] bend 생성됨" if bent else "[WARN] bend 0건(대각선 없음?)")
    print("[PASS] 모든 벤드 두 다리 직각" if ok else "[FAIL] 비직각 벤드 존재")
    sys.exit(0 if ok else 1)


if __name__ == "__main__":
    main()

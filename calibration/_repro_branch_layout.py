"""가지배관 표시 레이아웃 진단 — 직각화 후 남는 대각선/겹침을 수치화.

파이프라인: parse → filter → select_worst30 → orthogonalize_edge_positions.
분류:
  - branch edge 중 축정렬 안 된 것(대각선) 개수 + 각도
  - 축선 공유 겹침(collision) 잔존 개수
대각선이 de-overlap peel-off(고정노드 연결)인지, 미분류 가지인지 구분.
"""
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, select_worst30_heads,
    orthogonalize_edge_positions, _classify_branch_edges)

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def _key(p):
    return (round(float(p[0]), 3), round(float(p[1]), 3))


def main():
    bundle = parse_dxf_bundle(DXF)
    cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = filter_pipenet_only(bundle)
    sel = select_worst30_heads(pipe_ents, cat)
    print(f"heads={len(sel.heads)} edges={len(sel.edges)} source={sel.source_pos is not None}")

    head_pts = [h.pos for h in sel.heads]
    branch_keys, keyfn = _classify_branch_edges(sel.edges, head_pts, sel.source_pos)
    print(f"branch edges classified: {len(branch_keys)} / total {len(sel.edges)}")

    ortho = orthogonalize_edge_positions(
        sel.edges, head_points=head_pts, source_point=sel.source_pos)

    def oxy(p):
        return ortho.get(_key(p), (float(p[0]), float(p[1])))

    # 대각선 분석 (branch edge 만)
    diag = []
    axis_aligned = 0
    for e in sel.edges:
        ka, kb = keyfn(e[0]), keyfn(e[1])
        if ka == kb or frozenset((ka, kb)) not in branch_keys:
            continue
        (ax, ay), (bx, by) = oxy(e[0]), oxy(e[1])
        dx, dy = abs(ax - bx), abs(ay - by)
        if dx < 1.0 or dy < 1.0:
            axis_aligned += 1
        else:
            ang = math.degrees(math.atan2(dy, dx))
            off = min(ang, abs(90 - ang))  # 축에서 벗어난 각
            diag.append((round(off, 1), round(dx), round(dy), ka in _fixedset, kb in _fixedset))

    print(f"\n[branch edges] axis-aligned={axis_aligned}  DIAGONAL={len(diag)}")
    diag.sort(reverse=True)
    for off, dx, dy, fa, fb in diag[:20]:
        print(f"   off={off:5.1f}deg dx={dx:>7} dy={dy:>7} fixedA={fa} fixedB={fb}")

    # 겹침 잔존
    segs = []
    for e in sel.edges:
        ka, kb = keyfn(e[0]), keyfn(e[1])
        if ka == kb or frozenset((ka, kb)) not in branch_keys:
            continue
        (ax, ay), (bx, by) = oxy(e[0]), oxy(e[1])
        if abs(ax - bx) < 1.0 and abs(ay - by) >= 1.0:
            segs.append(('V', (ax + bx) / 2, min(ay, by), max(ay, by)))
        elif abs(ay - by) < 1.0 and abs(ax - bx) >= 1.0:
            segs.append(('H', (ay + by) / 2, min(ax, bx), max(ax, bx)))
    overlaps = 0
    worst = 0.0
    for i in range(len(segs)):
        oi, li, loi, hii = segs[i]
        for j in range(i + 1, len(segs)):
            oj, lj, loj, hij = segs[j]
            if oi != oj or abs(li - lj) >= 1.0:
                continue
            ov = min(hii, hij) - max(loi, loj)
            if ov > 150.0:
                overlaps += 1
                worst = max(worst, ov)
    print(f"\n[overlaps] count={overlaps}  worst={round(worst)}mm")


# _classify 로 fixed 재도출 (진단용)
def _compute_fixed(edges, branch_keys, keyfn, source_point):
    fixed = set()
    for e in edges:
        ka, kb = keyfn(e[0]), keyfn(e[1])
        if ka == kb:
            continue
        if frozenset((ka, kb)) not in branch_keys:
            fixed.add(ka); fixed.add(kb)
    if source_point is not None:
        fixed.add(keyfn(source_point))
    return fixed


if __name__ == "__main__":
    _b = parse_dxf_bundle(DXF)
    _c = {ly["name"]: ly["auto_category"] for ly in _b.layers}
    _pe = filter_pipenet_only(_b)
    _sel = select_worst30_heads(_pe, _c)
    _bk, _kf = _classify_branch_edges(_sel.edges, [h.pos for h in _sel.heads], _sel.source_pos)
    _fixedset = _compute_fixed(_sel.edges, _bk, _kf, _sel.source_pos)
    main()

"""평면도 배관망 '꼬임/우회' 진단 — 표시(display) vs 토폴로지(추출) 층위 분리.

측정:
  A) 추출 트리 총 길이 vs 노드집합 MST(유클리드) 총 길이 → 우회율(>1.3 이면 토폴로지 우회)
  B) 각 헤드까지 트리경로 길이 vs 소스-헤드 직선거리 → detour ratio 분포
  C) 배관 edge 교차 개수 — raw 좌표 / display(ortho) 좌표 각각
     (raw 교차 큼 = 실제 추출 꼬임 / display 만 큼 = 표시 레이아웃 꼬임)
"""
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, select_worst30_heads,
    orthogonalize_edge_positions)

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"


def _key(p):
    return (round(float(p[0]), 3), round(float(p[1]), 3))


def _seg_cross(a, b, c, d):
    """선분 ab, cd 가 (끝점 공유 제외) 실제 교차하면 True."""
    def o(p, q, r):
        return (q[0]-p[0])*(r[1]-p[1]) - (q[1]-p[1])*(r[0]-p[0])
    # 끝점 공유는 교차 아님(트리 분기)
    for p in (a, b):
        for q in (c, d):
            if abs(p[0]-q[0]) < 1e-6 and abs(p[1]-q[1]) < 1e-6:
                return False
    d1, d2, d3, d4 = o(c, d, a), o(c, d, b), o(a, b, c), o(a, b, d)
    return ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0))


def _count_cross(segs):
    n = 0
    for i in range(len(segs)):
        a, b = segs[i]
        for j in range(i+1, len(segs)):
            c, d = segs[j]
            if _seg_cross(a, b, c, d):
                n += 1
    return n


def _mst_len(nodes):
    """노드집합 유클리드 MST 총 길이(Prim)."""
    if len(nodes) < 2:
        return 0.0
    pts = list(nodes)
    inq = [False]*len(pts)
    dist = [float('inf')]*len(pts)
    dist[0] = 0.0
    tot = 0.0
    for _ in range(len(pts)):
        u = min((i for i in range(len(pts)) if not inq[i]), key=lambda i: dist[i])
        inq[u] = True
        tot += dist[u]
        for v in range(len(pts)):
            if not inq[v]:
                dv = math.hypot(pts[u][0]-pts[v][0], pts[u][1]-pts[v][1])
                if dv < dist[v]:
                    dist[v] = dv
    return tot


def main():
    bundle = parse_dxf_bundle(DXF)
    cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = filter_pipenet_only(bundle)
    sel = select_worst30_heads(pipe_ents, cat)
    print(f"heads={len(sel.heads)} edges={len(sel.edges)} source={sel.source_pos is not None}")

    # 노드집합
    node_set = set()
    for a, b, _ in sel.edges:
        node_set.add(_key(a)); node_set.add(_key(b))

    # A) 추출 총 길이 vs MST
    tot_extracted = sum(L for _, _, L in sel.edges)
    tot_mst = _mst_len(list(node_set))
    print(f"\n[A] 추출 총길이={tot_extracted/1000:.1f}m  MST(유클리드)={tot_mst/1000:.1f}m  "
          f"우회율={tot_extracted/tot_mst:.2f}x")

    # B) detour ratio (트리경로/직선)
    from collections import defaultdict, deque
    adj = defaultdict(list)
    for a, b, L in sel.edges:
        ka, kb = _key(a), _key(b)
        adj[ka].append((kb, L)); adj[kb].append((ka, L))
    src = _key(sel.source_pos)
    pathlen = {src: 0.0}
    q = deque([src])
    while q:
        u = q.popleft()
        for v, L in adj[u]:
            if v not in pathlen:
                pathlen[v] = pathlen[u] + L
                q.append(v)
    ratios = []
    for h in sel.heads:
        kh = _key(h.pos)
        if kh in pathlen:
            straight = math.hypot(kh[0]-src[0], kh[1]-src[1])
            if straight > 1.0:
                ratios.append(pathlen[kh]/straight)
    if ratios:
        ratios.sort()
        print(f"[B] 헤드 detour(트리경로/직선): median={ratios[len(ratios)//2]:.2f}  "
              f"max={ratios[-1]:.2f}  n={len(ratios)}")

    # C) 교차 — raw vs display
    raw_segs = [((float(a[0]), float(a[1])), (float(b[0]), float(b[1])))
                for a, b, _ in sel.edges]
    ortho = orthogonalize_edge_positions(
        sel.edges, head_points=[h.pos for h in sel.heads], source_point=sel.source_pos)

    def oxy(p):
        return ortho.get(_key(p), (float(p[0]), float(p[1])))
    disp_segs = [(oxy(a), oxy(b)) for a, b, _ in sel.edges]
    print(f"[C] edge 교차: raw={_count_cross(raw_segs)}  display(ortho)={_count_cross(disp_segs)}")

    print("\n해석: 우회율>1.3 또는 raw교차 큼 → 토폴로지(추출) 꼬임 / "
          "raw는 낮은데 display만 큼 → 표시 레이아웃 꼬임")


if __name__ == "__main__":
    main()

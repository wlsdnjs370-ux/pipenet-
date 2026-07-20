"""Reproduce Stage-3 '전체 배관망 그래프 인식' graph build and check for loops.

A spanning tree must satisfy: edges == nodes - components (acyclic). Report
whether the emitted `graph` (post force_spanning_tree) is truly acyclic, and
how many removed-cycle edges / head-drop / bridge edges remain (those get drawn
and can LOOK like rings).
"""
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import remote30_prototype as R  # noqa: E402

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / \
    "대명동201동 단위세대_layer정리.dxf"


def _count(graph):
    nodes = len(graph)
    edges = sum(len(set(v)) for v in graph.values()) // 2
    comps = len(R._connected_components(graph))
    return nodes, edges, comps


def main():
    print("DXF=%s" % DXF.name)
    bundle = R.parse_dxf_bundle(DXF)
    cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = R.filter_pipenet_only(bundle)
    dets = R.detect_heads(pipe_ents, cat)

    ni = R._NodeIndex()
    graph, edge_len = R._build_graph(pipe_ents, node_index=ni, layer_categories=cat)
    R.collapse_parallel_ladders(graph, edge_len)

    bridge_edges = set()
    for tol in (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0):
        R._bridge_components(graph, edge_len, max_bridge_mm=tol, bridge_edges_out=bridge_edges)

    head_drop = set()
    for h in dets:
        hp = ni.canonical(h.pos[0], h.pos[1])
        if hp in graph:
            continue
        nn = R._nearest_graph_node(graph, hp)
        if nn is None or hp == nn:
            continue
        d = math.hypot(hp[0]-nn[0], hp[1]-nn[1])
        if 1e-3 < d <= R.HEAD_BRIDGE_MAX_MM:
            graph.setdefault(hp, set()).add(nn)
            graph[nn].add(hp)
            edge_len[(min(hp, nn), max(hp, nn))] = d
            head_drop.add((min(hp, nn), max(hp, nn)))

    n0, e0, c0 = _count(graph)
    print("BEFORE SPT : nodes=%d edges=%d comps=%d  (cycles=%d)"
          % (n0, e0, c0, e0 - (n0 - c0)))

    penalty = set(bridge_edges) | set(head_drop)
    tree_edges, removed = R.force_spanning_tree(graph, edge_len, source=None,
                                                penalty_keys=penalty)
    n1, e1, c1 = _count(graph)
    print("AFTER  SPT : nodes=%d edges=%d comps=%d  (cycles=%d)  removed=%d"
          % (n1, e1, c1, e1 - (n1 - c1), len(removed)))
    print("tree acyclic? %s" % (e1 == n1 - c1))
    print("bridge in tree=%d  head_drop in tree=%d  (penalty applied)"
          % (len(bridge_edges & tree_edges), len(head_drop & tree_edges)))

    def _len(k):
        (a, b) = k
        return math.hypot(a[0]-b[0], a[1]-b[1])
    br = sorted((_len(k) for k in (bridge_edges & tree_edges)), reverse=True)
    rc = sorted((_len(k) for k in removed), reverse=True)
    if br:
        print("bridge len(mm): max=%.0f  >2000mm=%d  >5000mm=%d  <500mm=%d"
              % (br[0], sum(1 for x in br if x > 2000),
                 sum(1 for x in br if x > 5000), sum(1 for x in br if x < 500)))
    if rc:
        print("removed_cycle len(mm): " + " ".join("%.0f" % x for x in rc))


if __name__ == "__main__":
    main()

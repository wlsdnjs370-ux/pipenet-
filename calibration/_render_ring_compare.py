"""Two-panel render: LEFT = full tree (cyan real, amber bridge, slate head_drop),
RIGHT = same tree but head_drop edges DROPPED. If the visual rings vanish on the
right, the head-tip->backward-branch drop lines are the ring-closers.
"""
import math
import sys
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import remote30_prototype as R  # noqa: E402

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / \
    "대명동201동 단위세대_layer정리.dxf"
OUT = Path(__file__).resolve().parent / "_ring_compare.png"


def build():
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
    heads = []
    for h in dets:
        hp = ni.canonical(h.pos[0], h.pos[1])
        heads.append(hp)
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
    tree, removed = R.force_spanning_tree(
        graph, edge_len, source=None,
        penalty_keys=(bridge_edges | head_drop))
    bridge_edges &= tree
    head_drop &= tree
    return graph, bridge_edges, head_drop, heads


def draw(ax, graph, bridge_edges, head_drop, heads, skip_drop):
    ax.set_facecolor("#0f172a")
    seen = set()
    for u, nbrs in graph.items():
        for v in nbrs:
            k = (min(u, v), max(u, v))
            if k in seen:
                continue
            seen.add(k)
            if k in head_drop:
                if skip_drop:
                    continue
                c, lw, ls = "#94a3b8", 1.0, (0, (2, 3))
            elif k in bridge_edges:
                c, lw, ls = "#f59e0b", 1.2, (0, (6, 4))
            else:
                c, lw, ls = "#06b6d4", 1.4, "-"
            ax.plot([u[0], v[0]], [u[1], v[1]], color=c, lw=lw, ls=ls)
    ax.scatter([h[0] for h in heads], [h[1] for h in heads],
               s=14, c="#ef4444", zorder=5)
    ax.set_aspect("equal")


def main():
    graph, bridge_edges, head_drop, heads = build()
    print("head_drop edges in tree: %d" % len(head_drop))
    lens = sorted((math.hypot(a[0]-b[0], a[1]-b[1]) for (a, b) in head_drop),
                  reverse=True)
    print("head_drop lengths mm: " + " ".join("%.0f" % x for x in lens[:20]))
    fig, (axL, axR) = plt.subplots(1, 2, figsize=(28, 14))
    draw(axL, graph, bridge_edges, head_drop, heads, skip_drop=False)
    axL.set_title("WITH head_drop (slate)", color="w")
    draw(axR, graph, bridge_edges, head_drop, heads, skip_drop=True)
    axR.set_title("head_drop REMOVED", color="w")
    fig.savefig(OUT, dpi=80, facecolor="#0f172a", bbox_inches="tight")
    print("saved", OUT)


if __name__ == "__main__":
    main()

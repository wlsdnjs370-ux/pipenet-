"""Test: render estimated connectors (bridge amber, head_drop slate) as
ORTHOGONAL L-paths instead of diagonals. LEFT=diagonal (current), RIGHT=ortho.
Does the ortho routing remove the triangular 'ring' look?
Elbow orientation: go along the axis of the nearest REAL edge at the attach
node first (hug the branch), then perpendicular to reach the free end.
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
OUT = Path(__file__).resolve().parent / "_ortho_test.png"
CROP = None
if len(sys.argv) > 2:
    CROP = [float(x) for x in sys.argv[2].split(",")]


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


def _real_axis_at(graph, bridge_edges, head_drop, node):
    """Dominant axis (dx,dy unit) of a REAL edge incident to node, or None."""
    best = None
    bestlen = 0.0
    for nb in graph.get(node, ()):
        k = (min(node, nb), max(node, nb))
        if k in bridge_edges or k in head_drop:
            continue
        L = math.hypot(nb[0]-node[0], nb[1]-node[1])
        if L > bestlen:
            bestlen = L
            best = (nb[0]-node[0], nb[1]-node[1])
    if best is None:
        return None
    ax = 'h' if abs(best[0]) >= abs(best[1]) else 'v'
    return ax


def _elbow(graph, bridge_edges, head_drop, u, v):
    """Return elbow point for L from u->v. Attach at the node that has a real
    edge; hug that node's axis first."""
    # decide which endpoint is the 'attach on real branch' end
    au = _real_axis_at(graph, bridge_edges, head_drop, u)
    av = _real_axis_at(graph, bridge_edges, head_drop, v)
    if au is not None:
        anchor, free, ax = u, v, au
    elif av is not None:
        anchor, free, ax = v, u, av
    else:
        anchor, free, ax = u, v, ('h' if abs(v[0]-u[0]) >= abs(v[1]-u[1]) else 'v')
    if ax == 'h':
        elbow = (free[0], anchor[1])
    else:
        elbow = (anchor[0], free[1])
    return anchor, elbow, free


def draw(ax, graph, bridge_edges, head_drop, heads, ortho):
    ax.set_facecolor("#0f172a")
    seen = set()
    for u, nbrs in graph.items():
        for v in nbrs:
            k = (min(u, v), max(u, v))
            if k in seen:
                continue
            seen.add(k)
            est = k in bridge_edges or k in head_drop
            if k in bridge_edges:
                c, lw, ls = "#f59e0b", 1.3, (0, (6, 4))
            elif k in head_drop:
                c, lw, ls = "#94a3b8", 1.1, (0, (2, 3))
            else:
                c, lw, ls = "#06b6d4", 1.4, "-"
            if est and ortho:
                a, e, f = _elbow(graph, bridge_edges, head_drop, u, v)
                ax.plot([a[0], e[0], f[0]], [a[1], e[1], f[1]], color=c, lw=lw, ls=ls)
            else:
                ax.plot([u[0], v[0]], [u[1], v[1]], color=c, lw=lw, ls=ls)
    ax.scatter([h[0] for h in heads], [h[1] for h in heads],
               s=14, c="#ef4444", zorder=5)
    ax.set_aspect("equal")
    if CROP:
        ax.set_xlim(CROP[0], CROP[1]); ax.set_ylim(CROP[2], CROP[3])


def main():
    graph, bridge_edges, head_drop, heads = build()
    fig, (axL, axR) = plt.subplots(1, 2, figsize=(28, 14))
    draw(axL, graph, bridge_edges, head_drop, heads, ortho=False)
    axL.set_title("DIAGONAL (current)", color="w")
    draw(axR, graph, bridge_edges, head_drop, heads, ortho=True)
    axR.set_title("ORTHO L", color="w")
    fig.savefig(OUT, dpi=80, facecolor="#0f172a", bbox_inches="tight")
    print("saved", OUT)


if __name__ == "__main__":
    main()

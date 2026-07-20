"""Verify estimated (forced) bridges are excluded from network generation.

Monkeypatch _bridge_components to capture the exact bridge-edge keys added
during select_worst30_heads, then check how many of those keys survive into
selection.edges (i.e. the extraction routed THROUGH a bridge). Compare penalty
ON vs OFF. Heads must not drop.
"""
import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import remote30_prototype as R  # noqa: E402

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / \
    "대명동201동 단위세대_layer정리.dxf"

_captured = {"keys": set()}
_orig_bridge = R._bridge_components


def _spy_bridge(graph, edge_len, max_bridge_mm=500.0, bridge_edges_out=None):
    local = set()
    n = _orig_bridge(graph, edge_len, max_bridge_mm=max_bridge_mm,
                     bridge_edges_out=local)
    _captured["keys"] |= local
    if bridge_edges_out is not None:
        bridge_edges_out |= local
    return n


def _cross(segs):
    def o(p, q, r):
        return (q[0]-p[0])*(r[1]-p[1]) - (q[1]-p[1])*(r[0]-p[0])
    n = 0
    for i in range(len(segs)):
        a, b = segs[i]
        for j in range(i+1, len(segs)):
            c, d = segs[j]
            if any(abs(p[0]-q[0]) < 1e-6 and abs(p[1]-q[1]) < 1e-6
                   for p in (a, b) for q in (c, d)):
                continue
            d1, d2, d3, d4 = o(c, d, a), o(c, d, b), o(a, b, c), o(a, b, d)
            if ((d1 > 0) != (d2 > 0)) and ((d3 > 0) != (d4 > 0)):
                n += 1
    return n


def _bridge_in_edges(sel):
    """selection.edges among captured bridge keys (int-rounded compare)."""
    bkeys = {(tuple(int(round(c)) for c in u), tuple(int(round(c)) for c in v))
             for (u, v) in _captured["keys"]}
    bkeys = {(min(a, b), max(a, b)) for (a, b) in bkeys}
    hit = []
    for a, b, L in sel.edges:
        ra = tuple(int(round(c)) for c in a)
        rb = tuple(int(round(c)) for c in b)
        if (min(ra, rb), max(ra, rb)) in bkeys:
            hit.append((a, b, L))
    return hit


def run(label):
    _captured["keys"] = set()
    bundle = R.parse_dxf_bundle(DXF)
    cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = R.filter_pipenet_only(bundle)
    sel = R.select_worst30_heads(pipe_ents, cat)
    est = [k for k in _captured["keys"]
           if math.hypot(k[0][0]-k[1][0], k[0][1]-k[1][1]) > R.ESTIMATED_BRIDGE_MM]
    hit = _bridge_in_edges(sel)
    esthit = [e for e in hit
              if math.hypot(e[0][0]-e[1][0], e[0][1]-e[1][1]) > R.ESTIMATED_BRIDGE_MM]
    segs = [((a[0], a[1]), (b[0], b[1])) for a, b, _ in sel.edges]
    print("[%s] heads=%d edges=%d bridges_total=%d est_bridges=%d "
          "bridge_in_network=%d est_bridge_in_network=%d cross=%d"
          % (label, len(sel.heads), len(sel.edges), len(_captured["keys"]),
             len(est), len(hit), len(esthit), _cross(segs)))
    return sel, len(esthit)


def main():
    print("DXF=%s" % DXF.name)
    R._bridge_components = _spy_bridge
    try:
        _, on = run("penalty ON")
        _orig_sp = R._shortest_path

        def _no_pen(g, el, s, t, penalty_keys=None, penalty_mm=1.0e9):
            return _orig_sp(g, el, s, t, penalty_keys=None)
        R._shortest_path = _no_pen
        try:
            _, off = run("penalty OFF")
        finally:
            R._shortest_path = _orig_sp
    finally:
        R._bridge_components = _orig_bridge
    print("RESULT: est_bridge_in_network  ON=%d  OFF=%d  -> %s"
          % (on, off, "REDUCED" if on < off else ("none present" if off == 0 else "no change")))


if __name__ == "__main__":
    main()

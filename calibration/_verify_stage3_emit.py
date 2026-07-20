"""Drive the REAL production emit (run_stages_0_2) and render the stage-3
entities exactly as sent to the browser. Confirms estimated connectors now
emit as orthogonal L 2-segments (no diagonal ring cut-across).
"""
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
OUT = Path(__file__).resolve().parent / "_stage3_emit.png"
CROP = [float(x) for x in sys.argv[2].split(",")] if len(sys.argv) > 2 else None

COLORS = {
    "_graph_edge": ("#06b6d4", 1.4, "-"),
    "_graph_bridge": ("#f59e0b", 1.3, (0, (6, 4))),
    "_graph_head_drop": ("#94a3b8", 1.1, (0, (2, 3))),
}


def main():
    ents = None
    summary = None
    for ev in R.run_stages_0_2(DXF, job_id="verify"):
        if ev.get("type") == "entities" and ev.get("stage") == 3:
            ents = ev["entities"]
            summary = ev.get("summary")
            break
    assert ents is not None, "no stage-3 entities emitted"
    diag = sum(1 for e in ents if e.get("t") == "L"
               and e["l"] in ("_graph_bridge", "_graph_head_drop"))
    print("stage3 entities=%d  est-connector segments=%d" % (len(ents), diag))
    if summary:
        print("summary:", {k: summary[k] for k in summary
                           if k in ("node_count", "edge_count", "real_edge_count",
                                    "bridge_edge_count", "head_drop_edge_count",
                                    "removed_cycle_count")})
    fig, ax = plt.subplots(figsize=(18, 14))
    ax.set_facecolor("#0f172a")
    for e in ents:
        if e.get("t") != "L":
            continue
        c, lw, ls = COLORS.get(e["l"], ("#334155", 0.8, "-"))
        p = e["p"]
        ax.plot([p[0], p[2]], [p[1], p[3]], color=c, lw=lw, ls=ls)
    for e in ents:
        if e.get("t") == "C" and e["l"] == "_alarm_source":
            ax.plot(e["c"][0], e["c"][1], "*", ms=18, c="#22c55e", zorder=6)
    ax.set_aspect("equal")
    ax.set_title("PROD stage-3 emit  cyan=real amber=bridge slate=drop", color="w")
    if CROP:
        ax.set_xlim(CROP[0], CROP[1]); ax.set_ylim(CROP[2], CROP[3])
    fig.savefig(OUT, dpi=80, facecolor="#0f172a", bbox_inches="tight")
    print("saved", OUT)


if __name__ == "__main__":
    main()

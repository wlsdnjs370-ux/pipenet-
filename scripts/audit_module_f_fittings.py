"""Read-only audit of a saved plan board and its fitting classification."""
from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main() -> None:
    """Rebuild in memory; never replace drawing, picks, cache or user edits."""
    sys.stdout.reconfigure(encoding="utf-8")
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.io import _board_from_data, load_edits
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.tables import build_design_tables
    key = sys.argv[1]
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline, ho_from_spots
    result = pipeline(stage1_body(key))
    result["ho"] = ho_from_spots(result.get("spots"))
    board = _board_from_data(key, result)
    load_edits(board)
    es = EditSession(board)
    if not board.sources:
        print("AUDIT BLOCKED: saved board has no supply/AV point; no location guessed.")
        return
    got = select_and_expand(es.convert_payload(), board, k=30)
    if not got.get("ok"):
        raise RuntimeError(got)
    tbl = build_design_tables(got["kfp"], got["worst"], got["edge_ref"], [],
        board_pts=board.pts, tree_loads=got.get("tree_loads"),
        origin_mm=got.get("origin_mm"), phys=got.get("phys"),
        interior_junctions=got.get("interior_junctions"),
        interior_unresolved=got.get("interior_unresolved"))
    print("FITTINGS", dict(Counter(f["type"] for f in tbl.fittings)))
    print("BOARD", len(board.pts), len(board.edges), "HEADS", len(tbl.nozzles))
    from routes.module_f.api_design import _label_keys
    keys = _label_keys({"edit": es}, got, tbl)
    for nid, degree in list((got.get("phys") or {}).items()):
        if degree < 4:
            continue
        vid = got["node_ref"][nid]
        meta = got["kfp"]["nodes_meta_runtime"]
        links = []
        for pid,p in got["kfp"]["pipe_data"].items():
            if nid in (p["start"],p["end"]):
                other = p["end"] if p["start"] == nid else p["start"]
                links.append((pid, pid in got["edge_ref"],meta[nid]["coords"],meta[other]["coords"]))
        print('PHYS_SAMPLE',nid,vid,degree, 'raw',sum(vid in e for e in board.edges),links)
        print('SOURCE_PORTS', board.pts[vid],
              [board.pts[b if a == vid else a] for a,b in board.edges if vid in (a,b)])
        break
    for label in sys.argv[2:] or ["196", "200", "202", "204"]:
        print("NODE", label, [n for n in tbl.nodes if str(n["label"]) == label])
        print("PIPES", [p for p in tbl.pipes if label in (str(p["in"]), str(p["out"]))])
        print("FITS", [f for f in tbl.fittings if str(f["in"]) == label])
        nid = keys["nid"].get(label)
        vid = (got.get("node_ref") or {}).get(nid)
        print("SOURCE", label, nid, vid, "phys", (got.get("phys") or {}).get(nid),
              "xy", board.pts[vid] if vid is not None else None,
              "neighbors", [(b if a == vid else a, board.pts[b if a == vid else a])
                             for a,b in board.edges if vid in (a,b)])
    print("UNRESOLVED", len(tbl.unresolved.get("kind_items") or []))
    for row in tbl.unresolved.get("kind_items") or []:
        vid = (got.get("node_ref") or {}).get(row.get("node"))
        print("REVIEW", row, "CAD_XY_MM", board.pts[vid] if vid is not None else None)


if __name__ == "__main__":
    main()

"""Read-only live-session capture and isolated fitting reproduction.

Writes diagnostic artifacts only; never commits picks, edits or server state.
Only sessions without manual joins/deletions are replayed from the pipeline.
"""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main() -> None:
    """Capture GET responses and rebuild the same head selection in memory."""
    import requests
    from dotenv import dotenv_values

    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("sid")
    parser.add_argument("--offline", action="store_true")
    args = parser.parse_args()
    if not (len(args.sid) == 16 and all(c in "0123456789abcdef" for c in args.sid)):
        parser.error("Expected the exact 16-character session ID")
    path = ROOT / "data" / f"fitting_audit_{args.sid}.json"
    if args.offline:
        snapshot = json.loads(path.read_text(encoding="utf-8"))
    else:
        client = requests.Session()
        base = "http://127.0.0.1:5051"
        client.post(base + "/login", data={
            "password": dotenv_values(ROOT / ".env")["LOGIN_PASSWORD"],
            "next": "/module-f"}, timeout=15).raise_for_status()
        snapshot = {}
        for name in ("design/preview", "edit/state"):
            response = client.get(base + "/api/module-f/" + name,
                                  params={"sid": args.sid}, timeout=30)
            response.raise_for_status()
            snapshot[name] = response.json()
        path.write_text(json.dumps(snapshot, ensure_ascii=False), encoding="utf-8")
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline, ho_from_spots
    from services.cad_import.edit.io import _board_from_data
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.restrict import select_and_expand, board_port_neighbors
    from services.cad_import.design.tables import build_design_tables
    from routes.module_f.api_design import _label_keys, _dia_texts

    live = snapshot["design/preview"]
    edit = snapshot["edit/state"]
    state = edit["state"]
    assert not state["counts"]["joins"] and not state["counts"]["deletes"], "Manual edits need an exact replay"
    result = pipeline(stage1_body(edit["key"]))
    result["ho"] = ho_from_spots(result.get("spots"))
    board = _board_from_data(edit["key"], result)
    from routes.module_f.api_pick import _apply_chamfer
    _apply_chamfer({}, EditSession(board))
    for xy in state["valves"]:
        node, added = board.toggle_valve(*xy, max_d=1.0)
        assert added
        board.set_source_nodes([node])
    assert len(board.pts) == state["counts"]["pts"]
    assert len(board.edges) == state["counts"]["edges"]
    selected = []
    for disk in state["worst"]["heads"]:
        matches = [i for i, source in enumerate(board.disks)
                   if math.dist(source[:2], disk[:2]) < .1]
        assert len(matches) == 1, (disk, matches)
        selected.extend(matches)
    es = EditSession(board)
    got = select_and_expand(es.convert_payload(), board, k=len(selected), only_heads=set(selected))
    assert got["ok"], got
    tbl = build_design_tables(got["kfp"], got["worst"], got["edge_ref"], _dia_texts({"key": edit["key"]}),
        board_pts=board.pts, tree_loads=got.get("tree_loads"), origin_mm=got.get("origin_mm"),
        phys=got.get("phys"), interior_junctions=got.get("interior_junctions"),
        interior_unresolved=got.get("interior_unresolved"))
    keys = _label_keys({"edit": es}, got, tbl)
    print("REPLAY", len(tbl.nodes), "live", len(live["tables"]["nodes"]))
    assert tbl.nodes == live["tables"]["nodes"], "Node identity/geometry changed"
    fields = ("label", "in", "out", "length", "elev", "dia", "type", "c")
    assert [[p.get(k) for k in fields] for p in tbl.pipes] == [
        [p.get(k) for k in fields] for p in live["tables"]["pipes"]], "Pipe geometry/material changed"
    from routes.module_f.fitting_inspection import build_inspection
    inspection = build_inspection(tbl, got, keys, live["view"]["nodes"], board=board,
                                  transform=live["view"]["underlay"])
    adj = {}
    for a, b in board.edges:
        adj.setdefault(a, set()).add(b)
        adj.setdefault(b, set()).add(a)
    reports = []
    for label in ["23", "30", "38", "46", "53", "59", "90", "76"]:
        nid = keys["nid"].get(label)
        node = next((n for n in tbl.nodes if str(n["label"]) == label), None)
        live_node = next((n for n in live["tables"]["nodes"] if str(n["label"]) == label), None)
        assert node == live_node, (label, node, live_node)
        vid = got.get("fitting_node_ref", got["node_ref"]).get(nid)
        meta = got["kfp"]["nodes_meta_runtime"].get(nid)
        row = dict(label=label, nid=nid, node=node, source_node=vid,
                   metadata=meta, phys=got.get("phys", {}).get(nid),
                   fits=[f for f in tbl.fittings if str(f.get("in")) == label])
        if vid is not None:
            row["xy"] = board.pts[vid]
            row["ports"] = [board.pts[v] for v in board_port_neighbors(board.pts, adj, vid)]
        else:
            # Diagnose peeled head connections without guessing by coordinates.
            adjacent = [p["end"] if p["start"] == nid else p["start"]
                        for p in got["kfp"]["pipe_data"].values()
                        if nid in (p["start"], p["end"])]
            for other in adjacent:
                m = got["kfp"]["nodes_meta_runtime"][other]
                v = got["node_ref"].get(other)
                if m.get("type_id") == "head" and v is not None:
                    row["head_origin"] = dict(nid=other, vid=v, xy=board.pts[v],
                        ports=[board.pts[n] for n in board_port_neighbors(board.pts, adj, v)])
        reports.append(row)
        print("SITE", label, "source", vid, "ports", row["phys"], "kinds", [f["type"] for f in row["fits"]])
        if node is not None:
            glyph = next(f for f in inspection["fittings"] if f["node"] == label)
            assert glyph["kind"] == "tee" and glyph["original_degree"] == 3
            assert len(glyph["arms"]) == 3 and not glyph["symbolic"]
            row["inspection"] = glyph
    changed = []
    old = {(str(f["node"]), f["pipe"]): f for f in live["view"]["inspection"]["fittings"] if f["node"]}
    for f in inspection["fittings"]:
        previous = old.get((str(f["node"]), f["pipe"]))
        if previous and previous["kind"] != f["kind"]:
            changed.append(dict(node=f["node"], before=previous["kind"], after=f["kind"],
                                eq_before=previous["eq_m"], eq_after=f["eq_m"]))
    print("CHANGED", changed)
    out = ROOT / "data" / f"fitting_replay_{args.sid}.json"
    out.write_text(json.dumps(dict(reports=reports, tables=tbl.as_dict(), changes=changed, inspection=inspection),
                             ensure_ascii=False, default=list), encoding="utf-8")


if __name__ == "__main__":
    main()

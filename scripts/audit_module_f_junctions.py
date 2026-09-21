"""Read-only live capture and isolated junction replay for a Module F session."""
from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def replay(sid: str, *, offline: bool = False) -> dict:
    """Reconstruct an unedited session, retaining inspectable source identities."""
    path = ROOT / "data" / f"junction_audit_{sid}.json"
    if offline:
        if not path.exists():
            path = ROOT / "data" / f"fitting_audit_{sid}.json"
        snapshot = json.loads(path.read_text(encoding="utf-8"))
    else:
        import requests
        from dotenv import dotenv_values
        client = requests.Session()
        base = "http://127.0.0.1:5051"
        client.post(base + "/login", data={
            "password": dotenv_values(ROOT / ".env")["LOGIN_PASSWORD"],
            "next": "/module-f"}, timeout=15).raise_for_status()
        snapshot = {}
        for name in ("design/preview", "edit/state"):
            response = client.get(base + "/api/module-f/" + name,
                                  params={"sid": sid}, timeout=30)
            response.raise_for_status()
            snapshot[name] = response.json()
        path.write_text(json.dumps(snapshot, ensure_ascii=False), encoding="utf-8")
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline, ho_from_spots
    from services.cad_import.edit.io import _board_from_data
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.tables import build_design_tables
    from routes.module_f.api_design import _label_keys, _dia_texts
    from routes.module_f.api_pick import _apply_chamfer

    edit, live = snapshot["edit/state"], snapshot["design/preview"]
    state = edit["state"]
    assert not state["counts"]["joins"] and not state["counts"]["deletes"]
    result = pipeline(stage1_body(edit["key"]))
    result["ho"] = ho_from_spots(result.get("spots"))
    board = _board_from_data(edit["key"], result)
    es = EditSession(board)
    _apply_chamfer({}, es)
    for xy in state["valves"]:
        node, added = board.toggle_valve(*xy, max_d=1)
        assert added
        board.set_source_nodes([node])
    selected = {i for disk in state["worst"]["heads"] for i, head in enumerate(board.disks)
                if math.dist(disk[:2], head[:2]) < .1}
    # Duplicate CAD head symbols share one physical position; keep the live K.
    got = select_and_expand(es.convert_payload(), board,
                            k=len(state["worst"]["heads"]), only_heads=selected)
    assert got["ok"], got
    tbl = build_design_tables(got["kfp"], got["worst"], got["edge_ref"], _dia_texts({"key": edit["key"]}),
        board_pts=board.pts, tree_loads=got.get("tree_loads"), origin_mm=got.get("origin_mm"),
        phys=got["phys"], interior_junctions=got["interior_junctions"],
        interior_unresolved=got["interior_unresolved"])
    keys = _label_keys({"edit": es}, got, tbl)
    print("REPLAY", len(tbl.nodes), "live", len(live["tables"]["nodes"]),
          "same nodes", tbl.nodes == live["tables"]["nodes"])
    return dict(board=board, got=got, tbl=tbl, keys=keys, live=live, result=result)


def main() -> None:
    """Print the requested nodes, source arms and nearby original vertices."""
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("sid")
    parser.add_argument("--offline", action="store_true")
    parser.add_argument("--nodes", nargs="+", default=["8", "9", "72"])
    args = parser.parse_args()
    if len(args.sid) != 16 or any(c not in "0123456789abcdef" for c in args.sid):
        parser.error("Expected a 16-character session ID")
    context = replay(args.sid, offline=args.offline)
    board, got, tbl, keys = (context[k] for k in ("board", "got", "tbl", "keys"))
    adj: dict[int, set[int]] = {}
    for a, b in board.edges:
        adj.setdefault(a, set()).add(b)
        adj.setdefault(b, set()).add(a)
    for label in args.nodes:
        nid = keys["nid"].get(label)
        vid = got["fitting_node_ref"].get(nid)
        print("NODE", label, "identity", nid, "source", vid, "phys", got["phys"].get(nid))
        print("TABLE", [n for n in tbl.nodes if n["label"] == label])
        print("FITTINGS", [f for f in tbl.fittings if f.get("in") == label])
        print("PIPES", [p for p in tbl.pipes if label in (p["in"], p["out"])])
        if vid is None:
            continue
        for v in [vid] + sorted(adj.get(vid, ())):
            print("SOURCE", v, board.pts[v], "neighbors", [
                (n, board.pts[n], round(math.dist(board.pts[v], board.pts[n]), 3))
                for n in sorted(adj.get(v, ()))])


if __name__ == "__main__":
    main()

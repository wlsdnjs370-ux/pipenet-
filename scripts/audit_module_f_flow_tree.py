"""Offline B1F calculation-tree replay; never changes saved handwork or server."""
from __future__ import annotations

import json
import math
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def main() -> None:
    """Check all-head, selected corridor, fittings and table invariants on B1F."""
    sys.stdout.reconfigure(encoding="utf-8")
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline, ho_from_spots
    from services.cad_import.edit.io import _board_from_data
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.flow import flow_for_board
    from services.cad_import.design.worst import worst_k_heads
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.design.tables import build_design_tables
    from routes.module_f.api_pick import _apply_chamfer
    from routes.module_f.api_design import _dia_texts
    from routes.module_f.attach import picked_heads_wet

    snapshot = json.loads((ROOT / "data/fitting_audit_698e0f100c264a61.json").read_text(encoding="utf-8"))
    edit = snapshot["edit/state"]
    state = edit["state"]
    result = pipeline(stage1_body(edit["key"]))
    result["ho"] = ho_from_spots(result.get("spots"))
    b = _board_from_data(edit["key"], result)
    es = EditSession(b)
    _apply_chamfer({}, es)
    for xy in state["valves"]:
        node, added = b.toggle_valve(*xy, max_d=1)
        assert added
        b.set_source_nodes([node])
    original = (list(b.pts), set(b.edges))
    selected = {i for disk in state["worst"]["heads"] for i, d in enumerate(b.disks)
                if math.dist(disk[:2], d[:2]) < .1}
    flow = flow_for_board(b)
    print("FLOW", flow.report())
    worst = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, len(selected),
                          only_heads=selected, head_xy=b.disks, flow_tree=flow)
    assert worst["edges"] <= flow.edges
    probe = picked_heads_wet(es, selected, selected_source="Z1")
    assert probe["ok"] and selected <= probe["wet"], probe
    got = select_and_expand(es.convert_payload(), b, k=len(selected), only_heads=selected,
                            selected_source="Z1", probe=probe)
    assert got["ok"], got
    assert set(got["worst"]["heads"]) == selected
    assert got["worst"]["edges"] == worst["edges"]
    net = got["kfp"]
    assert len(net["pipe_data"]) == len(net["nodes_meta_runtime"]) - 1
    assert len([n for n in net["nodes_meta_runtime"].values() if n.get("type_id") == "head"]) == len(selected)
    assert (b.pts, b.edges) == original
    tbl = build_design_tables(net, got["worst"], got["edge_ref"], _dia_texts({"key": edit["key"]}),
        board_pts=b.pts, tree_loads=got.get("tree_loads"), origin_mm=got.get("origin_mm"),
        phys=got["phys"], interior_junctions=got["interior_junctions"],
        interior_unresolved=got["interior_unresolved"])
    audit = {"flow": flow.report(), "selected_heads": len(selected),
             "selected_edges": len(worst["edges"]), "far_m": worst["far_m"],
             "nodes": len(tbl.nodes), "pipes": len(tbl.pipes),
             "full_load_max": max(got["physical_pipe_loads"].values()),
             "selected_load_max": max(got["tree_loads"].values()),
             "tables": tbl.as_dict(), "fitting_phys": got["phys"],
             "pipe_full_loads": got["physical_pipe_loads"],
             "excluded_edges": [[a, z, why] for (a, z), why in sorted(flow.excluded.items())]}
    path = ROOT / "data/module_f_flow_tree_b1f_audit.json"
    path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    print("PASS", {k: v for k, v in audit.items() if k not in
                   {"tables", "fitting_phys", "pipe_full_loads", "excluded_edges"}})


if __name__ == "__main__":
    main()

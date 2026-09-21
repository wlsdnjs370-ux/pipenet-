"""Merge at actual connection nodes, independently of clicks and row order."""
from copy import deepcopy
import math

import pytest

from test_module_f_system_layout import heads, system
from test_module_f_network_editor import session
from routes.module_f.merge import MergeError, bake_combined_iso, merge_network


def machine_room():
    """A measured L-shaped path in a different drawing coordinate frame."""
    return {
        "nodes": [
            dict(label="m1", x=4000, y=9000, elevation=0, io_node="Input"),
            dict(label="m2", x=7000, y=9000, elevation=0, io_node="No"),
            dict(label="m3", x=7000, y=6000, elevation=0, io_node="No"),
        ],
        "pipes": [
            {"label": "m1", "in": "m1", "out": "m2", "length": 3, "elev": 0, "dia": 100, "c": 120},
            {"label": "m2", "in": "m2", "out": "m3", "length": 3, "elev": 0, "dia": 100, "c": 120},
        ],
        "conn_node_label": "m3", "conn_xy": [7000, 6000],
        "plan_edges": [[4000, 9000, 7000, 9000], [7000, 9000, 7000, 6000]],
        "extracted_from": "dxf",
    }


@pytest.mark.parametrize("source_kind", ["dxf", "dxf_clean_network", "template"])
def test_plan_shared_node_uses_label_and_exact_coordinates_not_row_order(source_kind):
    plan = heads()
    for n in plan.nodes:
        n["x"] += .375
        n["y"] += .125
    raw = system()
    raw["extracted_from"] = source_kind
    # Clean-network results need not end with the AV row.
    raw["nodes"] = raw["nodes"][::-1]
    before = deepcopy((vars(plan), raw))
    got = merge_network(plan, riser=raw, mode="lsp_gravity")
    combined = got["combined"]
    at = {n["label"]: n for n in combined.nodes}
    assert (at["10"]["x"], at["10"]["y"]) == (1000.375, 1000.125)
    assert at["11"]["x"] - at["10"]["x"] == pytest.approx(1000)
    assert at["11"]["y"] == at["10"]["y"]
    assert len(combined.nodes) == len(plan.nodes) + len(raw["nodes"]) - 1
    assert len(combined.pipes) == len(plan.pipes) + len(raw["pipes"])
    assert (vars(plan), raw) == before


@pytest.mark.parametrize("click", [[7500, 6400], None])
def test_machine_shared_node_uses_snapped_node_not_click_or_bbox(click):
    room = machine_room()
    expected = merge_network(heads(), riser=system(), machineroom=room, mode="lsp_gravity")
    if click is None:
        room.pop("conn_xy")
    else:
        room["conn_xy"] = click
    before = deepcopy(room)
    actual = merge_network(heads(), riser=system(), machineroom=room, mode="lsp_gravity")
    assert actual["combined"].nodes == expected["combined"].nodes
    assert actual["combined"].pipes == expected["combined"].pipes
    assert actual["combined"].machine_room_plan_edges == expected["combined"].machine_room_plan_edges
    assert room == before


def test_missing_declared_connection_is_reported_instead_of_adding_a_long_join():
    room = machine_room()
    room["conn_node_label"] = "missing"
    with pytest.raises(MergeError, match="기계실.*접속"):
        merge_network(heads(), riser=system(), machineroom=room, mode="lsp_gravity")


@pytest.mark.parametrize("mode", ["lsp_gravity", "lsp_1stage", "llsp_2stage", "hsp_pump"])
def test_physical_machine_path_has_no_stretched_seam_or_rescaled_lengths(mode):
    raw, room, plan = system(True), machine_room(), heads()
    got = merge_network(plan, riser=raw, machineroom=room, mode=mode)
    combined = got["combined"]
    at = {n["label"]: n for n in combined.nodes}
    assert "m3" not in at  # same physical node as the system input, not another pipe
    assert len(combined.nodes) == len(plan.nodes) + len(raw["nodes"]) + len(room["nodes"]) - 2
    assert len(combined.pipes) == len(plan.pipes) + len(raw["pipes"]) + len(room["pipes"])
    for pipe in combined.pipes:
        a, b = (at[pipe[key]] for key in ("in", "out"))
        distance = math.dist((a["x"] / 1000, a["y"] / 1000, a["elevation"]),
                             (b["x"] / 1000, b["y"] / 1000, b["elevation"]))
        assert distance == pytest.approx(pipe["length"], abs=.001), pipe["label"]


def test_machine_connection_elevation_rebases_whole_path_not_only_last_pipe():
    raw, room = system(True), machine_room()
    # A source datum 5m above its connection. Retain the known drops in both
    # existing pipes; merging must not introduce a third, artificial drop.
    room["nodes"][0]["elevation"] = 7
    room["nodes"][1]["elevation"] = 4
    room["nodes"][2]["elevation"] = 2
    room["pipes"][0]["elev"] = -3
    room["pipes"][1]["elev"] = -2
    before = deepcopy((raw, room))
    got = merge_network(heads(), riser=raw, machineroom=room, mode="lsp_gravity")
    at = {n["label"]: n for n in got["combined"].nodes}
    assert at["m1"]["elevation"] == 5
    assert at["m2"]["elevation"] == 2
    assert at["1"]["elevation"] == 0
    assert at["12"]["elevation"] == pytest.approx(1.7)
    for pipe in got["combined"].pipes:
        a, b = (at[pipe[k]] for k in ("in", "out"))
        assert b["elevation"] - a["elevation"] == pytest.approx(pipe["elev"])
    assert (raw, room) == before


def test_conflicting_explicit_source_drop_is_not_silently_stretched_or_erased():
    with pytest.raises(MergeError, match="수원 낙차"):
        merge_network(heads(), riser=system(), machineroom=machine_room(),
                      mode="hsp_pump", source_drop_m=3)


def test_consistent_source_drop_preserves_both_known_elevations_and_pipe_lengths():
    room = machine_room()
    room["nodes"][0]["elevation"] = -3
    room["pipes"][0]["elev"] = 3
    got = merge_network(heads(), riser=system(), machineroom=room,
                        mode="hsp_pump", source_drop_m=3)
    at = {n["label"]: n for n in got["combined"].nodes}
    assert at["m1"]["elevation"] == -3
    assert at["1"]["elevation"] == 0
    assert at["m1"]["x"] == at["m2"]["x"]
    assert at["m1"]["y"] == at["m2"]["y"]


def test_pump_display_offset_does_not_move_a_physical_pipe():
    raw, room = system(), machine_room()
    before = deepcopy((raw, room))
    got = merge_network(heads(), riser=raw, machineroom=room,
                        mode="hsp_pump", pump={"rated_q_lpm": 800, "rated_h_m": 60})
    at = {n["label"]: n for n in got["combined"].nodes}
    pump = got["combined"].pumps[0]
    assert tuple(at[pump["in"]][k] for k in ("x", "y", "elevation")) == tuple(
        at[pump["out"]][k] for k in ("x", "y", "elevation"))
    for pipe in got["combined"].pipes:
        a, b = (at[pipe[k]] for k in ("in", "out"))
        assert math.dist((a["x"] / 1000, a["y"] / 1000, a["elevation"]),
                         (b["x"] / 1000, b["y"] / 1000, b["elevation"])) == pytest.approx(pipe["length"])
    assert (raw, room) == before
    again = merge_network(heads(), riser=raw, machineroom=room,
                          mode="hsp_pump", pump={"rated_q_lpm": 800, "rated_h_m": 60})
    assert again["combined"].nodes == got["combined"].nodes
    assert again["combined"].pipes == got["combined"].pipes


def test_both_shared_nodes_and_original_pipes_survive_sdf_kfp_export(tmp_path):
    import json
    import xml.etree.ElementTree as ET
    from pathlib import Path
    from routes.module_f.emit import emit_merged

    room = machine_room()
    room["conn_xy"] = [7500, 6400]
    raw = system(True)
    raw["nodes"].reverse()
    got = merge_network(heads(), riser=raw, machineroom=room, mode="lsp_gravity")
    combined = got["combined"]
    files = emit_merged(combined, tmp_path, iso_nodes=bake_combined_iso(got)[0])
    assert files.get("kfp"), files["warnings"]
    data = json.loads(Path(files["kfp"]).read_text(encoding="utf-8"))
    assert len(data["pipe_data"]) == len(combined.pipes) == 7
    for path in (files["sdf"], files["sdf_iso"]):
        root = ET.parse(path).getroot()
        nodes = [n.get("label") for n in root.iter("Node")]
        assert nodes.count("10") == nodes.count("1") == 1
        assert "m3" not in nodes
        pipes = {p.get("label"): p for p in root.iter("Pipe")}
        assert len(pipes) == 7
        assert pipes["m2"].get("output") == pipes["r1"].get("input") == "1"
        assert pipes["r3"].get("output") == pipes["P1"].get("input") == "10"
        for pipe in combined.pipes:
            assert float(pipes[pipe["label"]].get("length")) == pytest.approx(pipe["length"])
    coords = {k: n["coords"] for k, n in data["nodes_meta_runtime"].items()}
    for pipe in data["pipe_data"].values():
        assert math.dist(coords[pipe["start"]], coords[pipe["end"]]) == pytest.approx(pipe["length_m"], abs=.001)
    assert Path(files["kfp"]).read_bytes() == Path(files["kfp_iso"]).read_bytes()


@pytest.mark.parametrize("iso", [False, True])
def test_preview_and_underlay_share_both_joints_after_off_pipe_click(session, tmp_path, iso):
    from flask import Flask
    from routes.module_f import api_merge
    from test_module_f_merge_underlays import drawings, transform

    drawings(session)
    session["slots"]["system"]["riser"] = system(True)
    room = session["slots"]["machineroom"]["machineroom"]
    room["conn_xy"] = [7500, 6400]  # old saved extraction with pre-snap coords
    api_merge.rebuild_merged(session, persist_overrides=False)
    app = Flask(__name__)
    api_merge.register(app, UPLOAD_DIR=tmp_path)
    response = app.test_client().get("/api/module-f/merge/preview", query_string={
        "sid": session["id"], "iso": int(iso), "underlays": "1"})
    assert response.status_code == 200, response.json
    data = response.json
    at = {n["label"]: (n["x"], n["y"]) for n in data["view"]["nodes"]}
    refs = {r["kind"]: r for r in data["view"]["underlays"]}
    assert set(data["counts"]["anchor"]) == {"10", "1"}
    assert transform(refs["machineroom"]["matrix"], (7000, 6000)) == pytest.approx(at["1"])
    assert transform(refs["plan"]["matrix"], (90100, 37154.0934468307)) == pytest.approx(at["10"])
    assert len(data["view"]["pipes"]) == 7


def test_extract_route_records_snapped_connection_separately_from_click():
    from flask import Flask
    from routes.module_f import api_sub, jobs

    app = Flask(__name__)
    app.testing = True
    api_sub.register(app)
    sess = jobs._new_session(key="shared-node-regression")
    sess.update(active="machineroom", entities=[
        dict(t="L", l="PIPE", p=[0, 0, 3000, 0]),
        dict(t="L", l="PIPE", p=[3000, 0, 3000, 4000]),
    ], sub_layers=["PIPE"], sub_layers_auto=False)
    try:
        response = app.test_client().post("/api/module-f/machineroom/extract", json={
            "sid": sess["id"], "source_x": 0, "source_y": 0,
            "conn_x": 3200, "conn_y": 4100, "ceiling_m": 0,
        })
        assert response.status_code == 200, response.json
        assert sess["machineroom"]["conn_xy"] == [3000, 4000]
        assert sess["machineroom"]["conn_pick_xy"] == [3200, 4100]
    finally:
        jobs._SESSIONS.pop(sess["id"], None)

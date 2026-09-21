"""Regressions for displaced tees and duplicate arc connections before flow."""
from __future__ import annotations

import copy
from pathlib import Path

import pytest

from src.pipenet_converter.graph.junctions import arc_connection_pairs, normalize_junctions
from src.pipenet_converter.graph.flow import build_flow_tree


def test_explicit_zero_link_moves_ports_to_actual_tee_and_splits_bypass():
    # u---h---v plus duplicate u-v. h and branch tip are explicitly zero-linked.
    pts = [(262430, -232385), (262575, -232385), (262502, -232385),
           (262502, -232385), (262502, -232000), (263000, -232385)]
    edges = {(0, 1), (0, 2), (1, 2), (2, 3), (3, 4), (1, 5)}
    before = copy.deepcopy((pts, edges))
    got = normalize_junctions(pts, edges)
    assert got.aliases == {3: 2}
    assert got.edges == {(0, 2), (1, 2), (2, 4), (1, 5)}
    assert got.splits == ((0, 1, 2),)
    assert sum(1 in e for e in got.edges) == 2  # old node 8: plain pipe
    assert sum(2 in e for e in got.edges) == 3  # old node 9: actual tee
    assert (pts, edges) == before
    twice = normalize_junctions(pts, got.edges)
    assert twice.edges == got.edges and not twice.aliases and not twice.splits
    tree = build_flow_tree(pts, got.edges, [{0}, {4}], [5])
    assert tree.path(4) == [5, 1, 2, 4]
    assert tree.path(0) == [5, 1, 2, 0]
    assert tree.loads[(1, 2)] == 2  # shared run counted once, two actual heads


def test_unconnected_crossing_and_coincident_points_remain_separate():
    pts = [(-10, 0), (10, 0), (0, 0), (0, 10), (0, 0), (0, -10)]
    edges = {(0, 1), (2, 3), (4, 5)}
    assert normalize_junctions(pts, edges).edges == edges


@pytest.mark.parametrize("hub", [(50, .001), (101, 0), (-1, 0)])
def test_nearby_or_outside_hub_does_not_split(hub):
    pts = [(0, 0), (100, 0), hub, (50, 20)]
    edges = {(0, 1), (0, 2), (2, 3)}
    assert normalize_junctions(pts, edges).edges == edges


def test_large_coordinate_diagonal_overlap_uses_stable_distance():
    pts = [(300000, -250000), (303000, -247000), (301000, -249000), (301000, -248000)]
    got = normalize_junctions(pts, {(0, 1), (0, 2), (2, 3)})
    assert got.edges == {(0, 2), (1, 2), (2, 3)}


def test_multiple_inline_junctions_are_order_independent():
    pts = [(0, 0), (100, 0), (25, 0), (75, 0), (25, 20), (75, 20)]
    edges = [(0, 1), (0, 2), (0, 3), (2, 4), (3, 5)]
    expected = {(0, 2), (2, 3), (1, 3), (2, 4), (3, 5)}
    assert normalize_junctions(pts, edges).edges == expected
    assert normalize_junctions(pts, reversed(edges)).edges == expected


def test_duplicate_hubs_exposed_by_split_are_idempotent():
    pts = [(0, 0), (100, 0), (50, 0), (50, 0), (50, 20), (50, -20)]
    edges = {(0, 1), (0, 2), (0, 3), (2, 4), (3, 5)}
    got = normalize_junctions(pts, edges)
    assert got.edges == {(0, 2), (1, 2), (2, 4), (2, 5)}
    assert normalize_junctions(pts, got.edges).edges == got.edges


def test_arc_join_keeps_existing_arm_and_one_connection_not_triangle():
    # Old node 72: two endpoints of a 3.55mm arm must not become two branches.
    pts = [(256967.651899, -226746.610794), (256971.205929, -226746.610794),
           (257071.205929, -226835.195413)]
    got = arc_connection_pairs(pts, [0, 1, 2], {(0, 1)}, [(2, 0), (2, 1)])
    assert got == [(1, 2)]
    assert arc_connection_pairs(pts, [0, 1, 2], {(0, 1), (1, 2)}, [(2, 0), (2, 1)]) == []


def test_arc_does_not_remove_a_real_loop_outside_symbol():
    pts = [(0, 0), (100, 100), (-100, 0), (-100, 100)]
    edges = {(0, 2), (2, 3), (1, 3)}
    assert arc_connection_pairs(pts, [0, 1], edges, [(0, 1)]) == [(0, 1)]


@pytest.mark.parametrize("value", [0, -1, float("nan"), float("inf")])
def test_invalid_tolerance_rejected(value):
    with pytest.raises(ValueError):
        normalize_junctions([], [], tolerance_mm=value)


def test_pipeline_arc_hook_and_head_exclusion():
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline.flow import join_all
    pts = [(0, 0), (3.55, 0), (100, -90)]
    spots = [{"k": "호", "cx": 10, "cy": -90}]
    out, joins, _ = join_all(pts, {(0, 1)}, spots, [[0, 1, 2]])
    assert len(out) == 2 and len(joins) == 1
    spots[0]["k"] = "헤드"
    assert join_all(pts, {(0, 1)}, spots, [[0, 1, 2]])[0] == {(0, 1)}


def test_captured_daemyeong_sites_replay_by_source_identity():
    # Optional local source fixture; never touches the live server or saved edits.
    from scripts.audit_module_f_junctions import replay
    capture = Path(__file__).resolve().parents[1] / "data/junction_audit_6626f0c826164ca4.json"
    if not capture.exists():
        pytest.skip("Local Daemyeong capture is not installed")
    context = replay("6626f0c826164ca4", offline=True)
    b, got, tbl = (context[k] for k in ("board", "got", "tbl"))
    labels = {nid: label for label, nid in context["keys"]["nid"].items()}
    origin = got["fitting_node_ref"]
    assert sum(1193 in e for e in b.edges) == 2
    assert 1193 not in origin.values()  # old 8 absorbed into straight run
    tee = next(nid for nid, vid in origin.items() if vid == 1103)
    assert got["phys"][tee] == 3
    assert "tee" in {f["type"] for f in tbl.fittings if f["in"] == labels[tee]}
    assert not any(f["type"].startswith("elbow") for f in tbl.fittings if f["in"] == labels[tee])
    assert sum(1528 in e for e in b.edges) == 2
    assert 1528 not in origin.values()  # old 72's spurious branch is gone
    bend = next(nid for nid, vid in origin.items() if vid == 1529)
    assert [f["type"] for f in tbl.fittings if f["in"] == labels[bend]] == ["elbow-45"]
    assert len(tbl.nozzles) == 30 and got["flow_report"]["heads"] == 109
    assert len(tbl.pipes) == len(tbl.nodes) - 1
    assert len(got["worst"]["heads"]) == 30


def test_b1f_keeps_geometry_and_reads_arc_marked_four_port_site_as_pass_through_tee(monkeypatch):
    from scripts.audit_module_f_junctions import replay
    from src.pipenet_converter.graph import junctions
    capture = Path(__file__).resolve().parents[1] / "data/fitting_audit_698e0f100c264a61.json"
    if not capture.exists():
        pytest.skip("Local B1F capture is not installed")
    after = replay("698e0f100c264a61", offline=True)
    monkeypatch.setattr(junctions, "normalize_junctions", lambda pts, edges:
                        junctions.JunctionNormalization(frozenset(edges), {}, ()))
    monkeypatch.setattr(junctions, "arc_connection_pairs",
                        lambda pts, arms, existing, candidates: list(candidates))
    before = replay("698e0f100c264a61", offline=True)
    new, old = after["tbl"], before["tbl"]
    assert new.nodes == old.nodes and new.nozzles == old.nozzles
    physical_fields = ("label", "in", "out", "length", "elev", "dia", "type", "c")
    assert [[p[k] for k in physical_fields] for p in new.pipes] == [
        [p[k] for k in physical_fields] for p in old.pipes]
    assert len(new.nozzles) == 28 and after["got"]["flow_report"]["heads"] == 264
    # 이 자리(평면 4방향, 접속 교차 괄호 «호» 가 앉은 곳)는 2026-09-20 판에서
    # «4방향 미해결» 로 남겼다. 오너 확정(2026-09-21, 덱 「z축 및 피팅류」 3장
    # 그림 2 «통과 갈래»): 호가 앉은 4방향 교차는 가지관이 주배관 위 0.5 m 로
    # 넘어가는 자리다 — 주배관 접속점은 분류티, 세로관 꼭대기는 티 한 번.
    # 호 없는 4방향은 여전히 미해결이다(tests/test_module_f_arc_symbols.py).
    kfp = after["got"]["kfp"]
    arcs = kfp["arc_junctions"]
    assert len(arcs) == 1 and all(v["kind"] == "T" for v in arcs.values()), arcs
    main_id, top_id = next(iter(arcs)), next(iter(arcs.values()))["top"]

    def label_of(nid):
        x, y, z = (float(v) for v in kfp["nodes_meta_runtime"][nid]["coords"][:3])
        hits = [n for n in new.nodes if abs(float(n["x"]) - x * 1000) <= 1
                and abs(float(n["y"]) - y * 1000) <= 1 and abs(float(n["elevation"]) - z) < 1e-6]
        assert len(hits) == 1, (nid, hits)
        return str(hits[0]["label"])

    main_label, top_label = label_of(main_id), label_of(top_id)
    unresolved = new.as_dict()["unresolved"]["kind_items"]
    assert not any(str(i.get("node_label")) == main_label for i in unresolved), unresolved
    assert not any(i.get("ports") == 4 for i in unresolved), unresolved
    fits = {}
    for f in new.fittings:
        fits.setdefault(f["pipe"], []).append(f["type"])
    riser = [p for p in new.pipes if {p["in"], p["out"]} == {main_label, top_label}]
    assert len(riser) == 1 and fits.get(riser[0]["label"]) == ["tee"], riser
    run = [p for p in new.pipes if p["in"] == main_label and p["out"] != top_label]
    assert all(fits.get(p["label"]) == ["tee-run"] for p in run), run
    tops = [p for p in new.pipes if p["in"] == top_label]
    assert tops and sum((fits.get(p["label"], []) for p in tops), []) == ["tee"], tops
    assert not any(f["type"].startswith("cross") for f in new.fittings)

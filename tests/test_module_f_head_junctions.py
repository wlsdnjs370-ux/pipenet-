"""Preserve mapped tee ports without changing the established tree graph."""
from __future__ import annotations

import copy
from dataclasses import asdict

import pytest

from src.pipenet_converter.graph.cycle_closures import ClosureResult
from src.pipenet_converter.graph.flow import edge_key
from src.pipenet_converter.graph.head_junctions import (
    HeadJunction, collect_head_junctions, preserve_head_junctions,
)

POINTS = [(-150., 0.), (150., 0.), (0., -150.),
          (-1000., 0.), (1000., 0.), (0., -1000.), (0., 0.)]
DRAWN = {(0, 3), (1, 4), (2, 5)}
MATERIALS = {e: ('picked-pipe', 5) for e in DRAWN}
HEADS = [(0., 0., 150.)]


def proof():
    return collect_head_junctions(POINTS, DRAWN, HEADS, MATERIALS)[0]


def project(edges, candidates=None, points=POINTS, **kwargs):
    return preserve_head_junctions(points, ClosureResult(frozenset(edges), frozenset(), ()),
                                   [proof()] if candidates is None else candidates,
                                   head_centers=[6], **kwargs)


def test_source_ports_are_captured_before_tree_and_not_from_unpicked_geometry():
    before = copy.deepcopy((POINTS, DRAWN, MATERIALS))
    candidate = proof()
    assert {p.node for p in candidate.ports} == {0, 1, 2}
    assert HeadJunction.from_record(asdict(candidate)) == candidate
    assert not collect_head_junctions(POINTS, DRAWN, HEADS, {(0, 3): 'pipe', (1, 4): 'pipe'})
    assert (POINTS, DRAWN, MATERIALS) == before


def test_existing_run_is_split_not_duplicated_and_center_is_reused():
    edges = DRAWN | {(0, 1), (2, 6)}
    before = copy.deepcopy((POINTS, edges))
    result = project(edges)
    assert result.edges == DRAWN | {(0, 6), (1, 6), (2, 6)}
    assert result.added == {(0, 6), (1, 6)}
    assert result.removed == {(0, 1)}
    assert (POINTS, edges) == before
    import math
    assert sum(math.dist(POINTS[a], POINTS[b]) for a, b in edges) == pytest.approx(
        sum(math.dist(POINTS[a], POINTS[b]) for a, b in result.edges))
    again = project(result.edges)
    assert not again.added and not again.removed
    assert again.records[0]['status'] == 'already_connected'


def test_third_port_can_attach_without_an_existing_return_route():
    result = project(DRAWN | {(0, 6), (1, 6)})
    assert result.added == {(2, 6)}
    assert result.records[0]['status'] == 'restored'


@pytest.mark.parametrize('deleted', [((0, -150), (0, 0)), ((0, -70), (0, 0))])
def test_manual_full_or_partial_port_deletion_blocks_projection(deleted):
    result = project(DRAWN | {(0, 6), (1, 6)}, forbidden_segments=[deleted])
    assert not result.added and result.records[0]['status'] == 'user_deleted'


def test_deleted_split_half_is_not_bypassed_by_original_unsplit_run():
    result = project(DRAWN | {(0, 1), (2, 6)},
                     forbidden_segments=[(POINTS[0], POINTS[6])])
    assert (0, 1) not in result.edges and (0, 6) not in result.edges
    assert (1, 6) in result.edges and (2, 6) in result.edges
    assert result.records[0]['status'] == 'user_deleted'


def test_material_conflict_does_not_join_independent_systems():
    materials = dict(MATERIALS)
    materials[(2, 5)] = ('other-system', 5)
    candidates = collect_head_junctions(POINTS, DRAWN, HEADS, materials)
    result = project(DRAWN | {(0, 6), (1, 6)}, candidates)
    assert not result.added and result.records[0]['status'] == 'material_conflict'


def test_four_ports_are_reviewed_not_invented_as_cross_fittings():
    pts = POINTS + [(0, 150), (0, 1000)]
    edges = DRAWN | {(7, 8)}
    candidates = collect_head_junctions(pts, edges, HEADS, MATERIALS | {(7, 8): ('picked-pipe', 5)})
    result = project(edges | {(0, 6), (1, 6)}, candidates, pts)
    assert not result.added and result.records[0]['status'] == 'four_ports_need_review'


def test_nearby_symbol_or_tangent_arm_is_not_connection_evidence():
    assert not collect_head_junctions(POINTS, DRAWN, [(10, 10, 150)], MATERIALS)
    pts = POINTS.copy()
    pts[5] = (1000, -150)
    assert not collect_head_junctions(pts, DRAWN, HEADS, MATERIALS)


def test_deleted_or_moved_exterior_does_not_reappear():
    result = project((DRAWN - {(2, 5)}) | {(0, 6), (1, 6)})
    assert not result.added and result.records[0]['status'] == 'exterior_arm_changed'
    pts = POINTS.copy()
    pts[2] = (0, -160)
    result = project(DRAWN | {(0, 6), (1, 6)}, points=pts)
    assert not result.added and result.records[0]['status'] == 'source_changed'


@pytest.mark.parametrize('angle', [30, 45, 120, 215])
def test_junction_rule_is_rotation_independent(angle):
    import math
    c, s = math.cos(math.radians(angle)), math.sin(math.radians(angle))
    points = [(700000 + c*x - s*y, 200000 + s*x + c*y) for x, y in POINTS]
    candidates = collect_head_junctions(points, DRAWN, [(700000, 200000, 150)], MATERIALS)
    assert len(candidates) == 1
    result = project(DRAWN | {(0, 1), (2, 6)}, candidates, points)
    assert result.edges == DRAWN | {(0, 6), (1, 6), (2, 6)}


@pytest.mark.parametrize('change', [dict(radius_mm=0), dict(radius_mm=float('nan')),
                                    dict(head_xy=[float('inf'), 0]), dict(ports=[])])
def test_invalid_head_port_cache_fails_closed(change):
    with pytest.raises(ValueError):
        HeadJunction.from_record(asdict(proof()) | change)


def junction_board():
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.board import EditBoard
    b = EditBoard('head-junction', POINTS, DRAWN | {(0, 1), (2, 6)}, HEADS, ups=HEADS)
    b.head_junctions = [asdict(proof())]
    b.sources = [3]
    b.valves = [3]
    b.set_head_kind(b.disks[0], '상향식')
    return b


@pytest.mark.parametrize('mode', ['loop', 'grid'])
def test_board_modes_delete_undo_and_saved_edits_keep_tree_unchanged(mode, tmp_path):
    b = junction_board()
    from services.cad_import.design.flow import flow_for_board, network_for_board
    from services.cad_import.design.cycle_source import cycle_source
    from services.cad_import.edit.io import load_edits, write_edits
    original = copy.deepcopy((b.pts, b.edges))
    tree = flow_for_board(b)
    b.network_mode = mode
    source = cycle_source(b)
    assert source.report()['restored'] == 1
    assert 0 in network_for_board(b).reference.head_node
    assert 0 not in tree.head_node
    deleted_edge = next(iter(source.added))
    segment = tuple(b.pts[n] for n in deleted_edge)
    assert b.delete(segment)[1] > 0
    assert deleted_edge not in cycle_source(b).edges
    assert (0, 1) not in cycle_source(b).edges
    write_edits(b, str(tmp_path))
    reopened = junction_board()
    assert load_edits(reopened, str(tmp_path))
    assert reopened.network_mode == mode
    assert deleted_edge not in cycle_source(reopened).edges
    assert (0, 1) not in cycle_source(reopened).edges
    assert b.undo()
    assert cycle_source(b).edges == source.edges
    b.network_mode = 'tree'
    assert (b.pts, b.edges) == original
    assert flow_for_board(b).report() == tree.report()


def test_source_cache_requires_and_roundtrips_head_port_evidence(tmp_path, monkeypatch):
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import disp_cache as cache
    from services.cad_import.edit.io import _board_from_data
    import json
    stamp = {'ver': cache._DISP_CACHE_VER, 'dxf': {'mtime_ns': 1}, 'spec': {'mtime_ns': 1}}
    path = tmp_path / 'cache.json'
    monkeypatch.setattr(cache, '_disp_cache_inputs', lambda key: stamp)
    monkeypatch.setattr(cache, '_disp_cache_path', lambda key: str(path))
    data = {'pts': POINTS, 'edges': DRAWN, 'edges1': DRAWN, 'hcov': HEADS,
            'hnodes': [{6}], 'ups': HEADS, 'head_kinds': [], 'head_junctions': [asdict(proof())]}
    cache._disp_cache_save('fixture', data)
    loaded = cache._disp_cache_load('fixture')
    assert HeadJunction.from_record(loaded['head_junctions'][0]) == proof()
    assert _board_from_data('fixture', loaded).head_junctions == loaded['head_junctions']
    blob = json.loads(path.read_text(encoding='utf-8'))
    blob['data'].pop('head_junctions')
    path.write_text(json.dumps(blob), encoding='utf-8')
    assert cache._disp_cache_load('fixture') is None


def test_split_run_preserves_declared_length_without_mutating_source():
    b = junction_board()
    from services.cad_import.design.cycle_source import projected_board
    from services.cad_import.design.flow import flow_for_board
    b.edge_len_mm = {'0|1': 900.0}
    original = copy.deepcopy(b.edge_len_mm)
    b.network_mode = 'grid'
    view = projected_board(b)
    tree = flow_for_board(view)
    assert tree.lengths_mm[(0, 6)] == tree.lengths_mm[(1, 6)] == 450.0
    assert b.edge_len_mm == original
    assert view.edge_len_mm is not b.edge_len_mm


from test_module_f_flow_integration import flow_client


@pytest.mark.parametrize('mode', ['loop', 'grid'])
def test_http_display_flow_candidate_and_audit_use_same_head_junction(flow_client, mode):
    from services.cad_import.edit.session import EditSession
    from routes.module_f.cancellation import install
    c, sess = flow_client
    install(c.application)
    b = junction_board()
    sess['edit'] = EditSession(b, key=b.key)
    original = copy.deepcopy((b.pts, b.edges))
    changed = c.post('/api/module-f/edit/network-mode', json={'sid': sess['id'], 'network_mode': mode})
    assert changed.status_code == 200, changed.json
    state = changed.json['state']
    assert state['counts']['edges'] == len(b.edges) + 1
    assert state['cycle_recovery']['restored'] == 1
    assert state['cycle_recovery']['split_edges'] == 1
    flowed = c.post('/api/module-f/edit/flow', json={'sid': sess['id']})
    assert flowed.status_code == 200 and flowed.json['state']['flow_report']['heads'] == 1
    selected = c.post('/api/module-f/edit/worst', json={'sid': sess['id'], 'k': 1})
    assert selected.status_code == 200, selected.json
    assert selected.json['summary']['rank_invariant']['ok']
    audit = c.get('/api/module-f/edit/flow/report', query_string={'sid': sess['id']})
    assert audit.status_code == 200, audit.json
    assert audit.json['cycle_recovery'][0]['evidence'] == 'drawn_head_tee_ports'
    assert not audit.json['network']['calculation_ready']
    assert (0, 6) in sess['worst']['edges'] and (0, 1) not in sess['worst']['edges']
    assert (sess['edit'].board.pts, sess['edit'].board.edges) == original


def test_b1f_actual_supply_and_all_three_return_branches_read_only(tmp_path):
    """Opt-in replay includes saved valve/deletion and exact pictured pipe ports."""
    import os
    if os.environ.get('MODULE_F_REAL_CYCLE_CHECK') != '1':
        pytest.skip('Local DXF replay is opt-in')
    import hashlib
    import json
    from pathlib import Path
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline, ho_from_spots
    from services.cad_import.edit.io import _board_from_data, load_edits
    from services.cad_import.design.cycle_source import cycle_source
    from services.cad_import.design.flow import flow_for_board, network_for_board

    key = 'B1F 현장조사 소화설비 평면도_컨셉2-수정본 (1)'
    folder = Path('cad_project_editor_g/docs/import')
    cache = folder / ('_edit_disp_cache_' + key + '.json')
    saved = folder / 'DWG' / (key + '_유저손질.json')
    fingerprints = {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in (cache, saved)}
    old = json.loads(cache.read_text(encoding='utf-8'))['data']
    result = pipeline(stage1_body(key))
    assert set(result['edges']) == {tuple(e) for e in old['edges']}
    assert list(map(tuple, result['pts'])) == list(map(tuple, old['pts']))
    result['ho'] = ho_from_spots(result.get('spots'))
    before = _board_from_data(key, dict(old, head_junctions=[]))
    after = _board_from_data(key, result)
    for b in (before, after):
        assert load_edits(b, str(folder / 'DWG'))
    original = copy.deepcopy((after.pts, after.edges))
    assert flow_for_board(before).report() == flow_for_board(after).report()
    before.network_mode = 'loop'
    old_net = network_for_board(before)
    for mode in ('loop', 'grid'):
        after.network_mode = mode
        network = network_for_board(after)
        recovery = cycle_source(after)
        pictured = [r for r in recovery.records if r['evidence'] == 'drawn_head_tee_ports'
                    and 725500 < r['head_xy'][0] < 751000
                    and 209000 < r['head_xy'][1] < 218000]
        assert len(pictured) == 6, pictured  # top/bottom of each of three middle branches
        for row in pictured:
            assert row['status'] == 'restored', row
            assert {edge_key(n, row['center']) for n in row['ports']} <= network.edges
        assert network.report()['cycle_rank'] >= old_net.report()['cycle_rank'] + 3
        assert set(old_net.reference.head_node) <= set(network.reference.head_node)
        assert len(network.reference.head_node) > len(old_net.reference.head_node)
        assert (after.pts, after.edges) == original
    audit = {'before': old_net.report(), 'after': network.report(),
             'recovery': recovery.report(), 'pictured_junctions': pictured}
    (tmp_path / 'b1f-loop-grid-audit.json').write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding='utf-8')
    print('B1F_SOURCE_AUDIT', json.dumps(audit, ensure_ascii=False))
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.collections import LineCollection
    fig, axes = plt.subplots(1, 2, figsize=(16, 6), layout='constrained')
    for ax, net, title in zip(axes, (old_net, network), ('Before: preserved run, isolated middle pipe', 'After: preserved run AND all tee ports')):
        ax.set_facecolor('#071014')
        ax.add_collection(LineCollection([[after.pts[a], after.pts[b]] for a, b in net.edges], colors='#00cfa0', linewidths=1.5))
        hs = [after.disks[h] for h in net.reference.head_node]
        ax.scatter([h[0] for h in hs], [h[1] for h in hs], s=16, color='#ffaa44', zorder=3)
        ax.set(xlim=(725500, 751000), ylim=(209000, 218000), title=title, aspect='equal')
        ax.tick_params(labelbottom=False, labelleft=False)
    fig.savefig(tmp_path / 'b1f-loop-grid-ports.png', dpi=140)
    plt.close(fig)
    print('B1F_VISUAL_CHECK', tmp_path / 'b1f-loop-grid-ports.png')
    assert fingerprints == {p: hashlib.sha256(p.read_bytes()).hexdigest() for p in fingerprints}

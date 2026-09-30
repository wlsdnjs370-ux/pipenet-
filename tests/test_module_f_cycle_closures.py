"""Evidence-backed cycle repair without mutating the legacy tree source."""
from __future__ import annotations

import copy
from dataclasses import asdict

import pytest

from src.pipenet_converter.graph.cycle_closures import (
    ClosureCandidate, recover_cycle_closures,
)

POINTS = [(-150, 0), (150, 0), (-1000, 0), (-1000, -1000),
          (1000, -1000), (1000, 0), (0, 0)]
RETURN = {(0, 2), (2, 3), (3, 4), (4, 5), (1, 5)}
CANDIDATE = ClosureCandidate(0, 1, POINTS[0], POINTS[1], (0, 0), 100)


@pytest.mark.parametrize('half', [False, True])
def test_closes_proven_gap_preserving_existing_center_and_source(half):
    edges = RETURN | ({(0, 6)} if half else set())
    before = copy.deepcopy((POINTS, edges))
    result = recover_cycle_closures(POINTS, edges, [CANDIDATE], head_centers=[6])
    assert result.added == ({(1, 6)} if half else {(0, 1)})
    assert result.report()['restored'] == 1
    assert (POINTS, edges) == before
    again = recover_cycle_closures(POINTS, result.edges, [CANDIDATE], head_centers=[6])
    assert not again.added and again.report()['already_connected'] == 1
    assert len(result.edges) == len(edges) + 1  # one cycle, never a bypass triangle


def test_no_evidence_means_no_guessed_connection():
    assert recover_cycle_closures(POINTS, RETURN, []).edges == RETURN


@pytest.mark.parametrize('segment', [((-150, 0), (150, 0)), ((0, 0), (150, 0))])
def test_explicit_full_or_half_deletion_blocks_repair(segment):
    r = recover_cycle_closures(POINTS, RETURN | {(0, 6)}, [CANDIDATE],
                               head_centers=[6], forbidden_segments=[segment])
    assert not r.added and r.records[0]['status'] == 'user_deleted'


def test_deleted_return_route_is_not_silently_repaired():
    r = recover_cycle_closures(POINTS, RETURN - {(3, 4)}, [CANDIDATE])
    assert not r.added and r.records[0]['status'] == 'return_route_missing'


def test_changed_source_and_exterior_arms_are_rejected():
    pts = copy.deepcopy(POINTS)
    pts[0] = (-151, 0)
    assert recover_cycle_closures(pts, RETURN, [CANDIDATE]).records[0]['status'] == 'source_changed'
    pts[0], pts[2] = POINTS[0], (-150, 1000)
    assert recover_cycle_closures(pts, RETURN, [CANDIDATE]).records[0]['status'] == 'exterior_arm_changed'


def test_ambiguous_head_centers_fail_closed():
    pts = POINTS + [(0, 0)]
    r = recover_cycle_closures(pts, RETURN | {(0, 6), (1, 7)}, [CANDIDATE], head_centers=[6, 7])
    assert not r.added and r.records[0]['status'] == 'ambiguous_centers'


@pytest.mark.parametrize('change', [dict(evidence='nearest_point'), dict(head_radius_mm=0),
                                    dict(head_xy=[float('nan'), 0])])
def test_invalid_cache_evidence_is_rejected(change):
    with pytest.raises(ValueError):
        ClosureCandidate.from_record(dict(asdict(CANDIDATE), **change))


def closure_board():
    """Exercise real legacy head completion (one half-neck), not a mocked view."""
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.board import EditBoard
    b = EditBoard('cycle-gap', POINTS, RETURN | {(0, 6)}, [(0, 0, 100)], ups=[(0, 0, 100)])
    b.set_head_kind(b.disks[0], '상향식')
    b.sources = [2]
    b.valves = [2]
    b.cycle_candidates = [asdict(CANDIDATE)]
    return b


def test_board_projection_delete_undo_save_load_and_tree_isolation(tmp_path):
    from services.cad_import.design.cycle_source import cycle_source
    from services.cad_import.design.flow import flow_for_board, network_for_board
    from services.cad_import.edit.io import write_edits, load_edits
    b = closure_board()
    original = copy.deepcopy((b.pts, b.edges))
    tree = flow_for_board(b)
    b.network_mode = 'loop'
    result = cycle_source(b)
    assert result.report()['restored'] == 1 and len(result.added) == 1
    assert network_for_board(b).report()['cycle_rank'] == 1
    edge = next(iter(result.added))
    segment = tuple(b.pts[n] for n in edge)
    _, count = b.delete(segment)
    assert count == 1 and not cycle_source(b).added
    assert network_for_board(b).report()['cycle_rank'] == 0
    write_edits(b, str(tmp_path))
    restored = closure_board()
    assert load_edits(restored, str(tmp_path))
    assert restored.network_mode == 'loop' and not cycle_source(restored).added
    assert b.undo()
    assert cycle_source(b).edges == result.edges
    b.network_mode = 'tree'
    assert (b.pts, b.edges) == original
    assert flow_for_board(b).report() == tree.report()


from test_module_f_flow_integration import flow_client


@pytest.mark.parametrize('mode', ['loop', 'grid'])
def test_http_display_flow_scenario_and_json_share_recovered_edges(flow_client, mode):
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.cycle_source import cycle_source
    from routes.module_f.cancellation import install
    c, sess = flow_client
    install(c.application)
    b = closure_board()
    sess['edit'] = EditSession(b, key=b.key)
    before = copy.deepcopy((b.pts, b.edges))
    tree_report = c.post('/api/module-f/edit/flow', json={'sid':sess['id']}).json['state']['flow_report']
    changed = c.post('/api/module-f/edit/network-mode', json={'sid':sess['id'], 'network_mode':mode})
    assert changed.status_code == 200, changed.json
    assert changed.json['state']['cycle_recovery']['restored'] == 1
    assert changed.json['state']['counts']['edges'] == len(b.edges) + 1
    water = c.post('/api/module-f/edit/flow', json={'sid':sess['id']})
    assert water.status_code == 200, water.json
    assert water.json['state']['flow_report']['cycle_rank'] == 1
    selected = c.post('/api/module-f/edit/worst', json={'sid':sess['id'], 'k':1})
    assert selected.status_code == 200, selected.json
    assert selected.json['summary']['rank_invariant']['ok']
    assert cycle_source(sess['edit'].board).added <= sess['worst']['edges']
    audit = c.get('/api/module-f/edit/flow/report', query_string={'sid':sess['id']})
    assert audit.status_code == 200, audit.json
    assert audit.json['cycle_recovery'][0]['status'] == 'restored'
    assert audit.json['network']['cycle_rank'] == audit.json['scenario']['cycle_rank'] == 1
    assert (sess['edit'].board.pts, sess['edit'].board.edges) == before
    # Delete the projected half-neck using the real UI hit-test route.
    added = next(iter(cycle_source(sess['edit'].board).added))
    xy = [sum(sess['edit'].board.pts[n][i] for n in added) / 2 for i in (0, 1)]
    c.post('/api/module-f/edit/mode', json={'sid':sess['id'], 'mode':'삭제'})
    deleted = c.post('/api/module-f/edit/click', json={'sid':sess['id'], 'x':xy[0], 'y':xy[1], 'max_d':5})
    assert deleted.status_code == 200, deleted.json
    assert deleted.json['report']['n'] == 1
    assert deleted.json['state']['cycle_recovery']['restored'] == 0
    assert deleted.json['state']['flow_report'] is None
    restored = c.post('/api/module-f/edit/undo', json={'sid':sess['id']})
    assert restored.json['undone'] and restored.json['state']['cycle_recovery']['restored'] == 1
    c.post('/api/module-f/edit/network-mode', json={'sid':sess['id'], 'network_mode':'tree'})
    again = c.post('/api/module-f/edit/flow', json={'sid':sess['id']})
    assert again.json['state']['flow_report'] == tree_report


def test_source_evidence_survives_cache_and_old_caches_rebuild(tmp_path, monkeypatch):
    from services.cad_import.pipeline import disp_cache as cache
    from services.cad_import.edit.io import _board_from_data
    import json
    stamp = {'ver': cache._DISP_CACHE_VER, 'dxf': {'mtime_ns':1}, 'spec': {'mtime_ns':1}}
    monkeypatch.setattr(cache, '_disp_cache_inputs', lambda key: stamp)
    monkeypatch.setattr(cache, '_disp_cache_path', lambda key: str(tmp_path / 'cache.json'))
    data = {'pts':POINTS, 'edges': RETURN, 'edges1': RETURN,
            'hcov':[(0, 0, 100)], 'hnodes':[{6}], 'ups':[(0, 0, 100)],
            'head_kinds': [], 'cycle_candidates':[asdict(CANDIDATE)]}
    cache._disp_cache_save('fixture', data)
    loaded = cache._disp_cache_load('fixture')
    assert ClosureCandidate.from_record(loaded['cycle_candidates'][0]) == CANDIDATE
    assert _board_from_data('fixture', loaded).cycle_candidates == loaded['cycle_candidates']
    blob = json.loads((tmp_path / 'cache.json').read_text(encoding='utf-8'))
    blob['data'].pop('cycle_candidates')
    (tmp_path / 'cache.json').write_text(json.dumps(blob), encoding='utf-8')
    assert cache._disp_cache_load('fixture') is None


def test_browser_shows_recovered_connection_and_review_count(flow_client, tmp_path):
    from test_module_f_viewport_browser import install_page, pw_api, ROOT
    from services.cad_import.edit.session import EditSession
    from urllib.parse import urlparse
    c, sess = flow_client
    sess['edit'] = EditSession(closure_board(), key='cycle-gap')
    source = (ROOT / 'static/module_f.js').read_text(encoding='utf-8').replace(
        '  setStage("open");\n  loadSaved();',
        '  window.__cycleTest={setEdit,renderEdit};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={'width':1440, 'height':1000})
        errors = install_page(page, source)
        def route(req):
            url = urlparse(req.request.url)
            r = c.open(url.path + ('?' + url.query if url.query else ''), method=req.request.method,
                       data=req.request.post_data, content_type='application/json')
            if r.status_code == 404:
                req.fulfill(json={'ok':True, 'rows':[], 'items':[], 'fields':{}})
            else:
                req.fulfill(status=r.status_code, body=r.data, content_type=r.content_type)
        page.route('http://module-f.test/api/module-f/**', route)
        state = c.get('/api/module-f/edit/state', query_string={'sid':sess['id']}).json['state']
        page.evaluate('''({sid,state})=>{
          Object.assign(__mf,{sid,slot:'plan',method:'manual'});
          __cycleTest.setEdit(state);__viewTest.setStage('edit');__cycleTest.renderEdit();
        }''', {'sid':sess['id'], 'state':state})
        page.select_option('#ed-network-mode', 'loop')
        page.wait_for_function("__mf.edit.network_mode==='loop'")
        assert '이음 복원 1곳' in page.locator('#ed-flow-summary').inner_text()
        page.click('#ed-flow')
        page.wait_for_function('!!__mf.edit.flow_report')
        assert page.evaluate('__mf.edit.flow_report.cycle_rank') == 1
        assert '이음 복원 1곳' in page.locator('#ed-flow-summary').inner_text()
        page.screenshot(path=str(tmp_path / 'cycle-recovery-ui.png'))
        page.select_option('#ed-network-mode', 'tree')
        page.wait_for_function("__mf.edit.network_mode==='tree'")
        assert '이음 복원' not in page.locator('#ed-flow-summary').inner_text()
        assert not errors, errors
        browser.close()


def test_confirmed_b1f_drawing_read_only(tmp_path):
    """Opt-in DXF replay. Never overwrite live caches, saved edits or sessions."""
    import os
    if os.environ.get('MODULE_F_REAL_CYCLE_CHECK') != '1':
        pytest.skip('Local DXF replay is opt-in')
    import json
    from pathlib import Path
    import networkx as nx
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline, ho_from_spots
    from services.cad_import.edit.io import _board_from_data
    from services.cad_import.design.cycle_source import cycle_source
    from services.cad_import.design.flow import flow_for_board, network_for_board
    key = 'B1F 현장조사 소화설비 평면도_컨셉2-수정본 (1)'
    old = json.loads(Path('cad_project_editor_g/docs/import/_edit_disp_cache_' + key + '.json').read_text(encoding='utf-8'))['data']
    result = pipeline(stage1_body(key))
    assert set(result['edges']) == {tuple(e) for e in old['edges']}
    assert list(map(tuple,result['pts'])) == list(map(tuple,old['pts']))
    result['ho'] = ho_from_spots(result.get('spots'))
    b = _board_from_data(key, result)
    original = copy.deepcopy((b.pts, b.edges))
    recovered = cycle_source(b)
    gap_records = [r for r in recovered.records if r['evidence'] == 'same_material_head_cover']
    assert sum(r['status'] == 'restored' for r in gap_records) == len(result['cycle_candidates']) > 0
    assert (b.pts, b.edges) == original
    # Diagnostic supply only, on the isolated board: exercise the pictured region.
    candidate = next(r for r in result['cycle_candidates'] if 744000 < r['head_xy'][0] < 745000)
    b.sources = [candidate['a']]
    tree = flow_for_board(b)
    b.network_mode = 'loop'
    network = network_for_board(b)
    original_graph = nx.Graph(list(b.edges))
    component = nx.node_connected_component(original_graph, candidate['a'])
    component_added = {e for e in recovered.added if set(e) <= component}
    assert component_added <= network.edges
    assert network.report()['cycle_rank'] >= len(component_added)
    assert set(tree.head_node) <= set(network.reference.head_node)
    assert (b.pts, b.edges) == original
    b.network_mode = 'tree'
    assert flow_for_board(b).report() == tree.report()
    print('REAL_CYCLE_CHECK', json.dumps({'source': recovered.report(),
        'pictured_component_added':len(component_added), 'network':network.report()}, ensure_ascii=False))
    # Technical visual check: same coordinates, old head-neck vs restored edge.
    import matplotlib
    matplotlib.use('Agg')
    import matplotlib.pyplot as plt
    from matplotlib.collections import LineCollection
    fig, axes = plt.subplots(1, 2, figsize=(16, 7), layout='constrained')
    for ax, edges, title in zip(axes, (b.edges, recovered.edges), ('Before: legacy tree source', 'After: loop / grid source')):
        ax.set_facecolor('#071014')
        ax.add_collection(LineCollection([[b.pts[a], b.pts[z]] for a, z in edges], colors='#00cfa0', linewidths=1.1))
        disks = [d for d in b.disks if 725500 < d[0] < 751000 and 209000 < d[1] < 218000]
        ax.scatter([d[0] for d in disks], [d[1] for d in disks], s=18, color='#ffaa44', zorder=3)
        if title.startswith('After'):
            ax.add_collection(LineCollection([[b.pts[a],b.pts[z]] for a,z in recovered.added], colors='#ff4165', linewidths=4, zorder=4))
        ax.set(xlim=(725500,751000), ylim=(209000,218000), title=title, aspect='equal')
        ax.tick_params(labelbottom=False,labelleft=False)
    fig.savefig(tmp_path / 'b1f-cycle-recovery.png', dpi=140)
    plt.close(fig)
    (tmp_path / 'cycle-recovery.json').write_text(json.dumps(list(recovered.records), ensure_ascii=False, indent=2), encoding='utf-8')
    print('VISUAL_CHECK', tmp_path / 'b1f-cycle-recovery.png')

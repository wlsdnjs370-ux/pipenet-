"""All-area selection and table handoff, with F quota behaviour unchanged."""
from __future__ import annotations

import pytest

from src.pipenet_converter.graph.area_selection import area_candidates, select_area_heads, area_corridor
from src.pipenet_converter.graph.flow import build_flow_tree
from test_module_f_flow_integration import flow_client


def test_region_required_union_boundary_and_concave_polygon():
    disks = [(0,0), (10,0), (10,10), (6,6), (50,50)]
    with pytest.raises(ValueError, match="영역"):
        area_candidates(disks, [])
    polygon = {"type":"polygon", "points":[[0,0],[10,0],[10,4],[4,4],[4,10],[0,10]]}
    assert area_candidates(disks, [polygon]) == (0,1)
    assert area_candidates(disks, [polygon, [5,5,12,12], [0,0,10,10]]) == (0,1,2,3)


def test_no_quota_over_200_and_aliases_disconnection_explicit():
    pts = [(i*1000,0) for i in range(260)]
    disks = pts[1:] + [pts[1], (999999,0)]
    flow = build_flow_tree(pts, [(i,i+1) for i in range(259)],
                           [{i} for i in range(1,260)] + [{1},set()], [0], head_xy=disks)
    selected = select_area_heads(disks, [[500,-1,300000,1]], flow)
    assert len(selected.heads) == 259
    assert selected.aliases == ((259,0),)
    result = area_corridor(selected, disks, flow)
    assert len(result["heads"]) == 259 and (0,1) in result["edges"]  # source outside region
    selected = select_area_heads(disks, [[0,-1,1000000,1]], flow)
    assert selected.disconnected == (260,)
    with pytest.raises(ValueError, match="연결"):
        area_corridor(selected, disks, flow)


@pytest.mark.parametrize("mode", ["tree", "loop", "grid"])
def test_all_heads_reach_tables_and_k_is_ignored(flow_client, mode):
    client, sess = flow_client
    sess["edit"].board.network_mode = mode
    body = {"sid":sess["id"], "selection_mode":"area_all", "k":1, "sheet":99,
            "zones":[[1900,1900,3100,2100]]}
    response = client.post('/api/module-f/edit/worst', json=body)
    assert response.status_code == 200, response.json
    assert response.json["summary"]["k"] == 2
    assert sess["worst"]["heads"] == [0,1]
    assert response.json["state"]["worst"]["selection_mode"] == "area_all"
    if mode != "tree":
        assert {(3,5),(4,5),(2,4)} <= sess["worst"]["edges"]
        assert (2,3) not in sess["worst"]["edges"]  # outside the thin selected region
        assert response.json["summary"]["area_scope"]["policy"] == 'region_and_supply_v1'
    response = client.post('/api/module-f/design/build', json={"sid":sess["id"],
        "selection_mode":"area_all", "k":"ignored", "review_default_mm":65})
    assert response.status_code == 200, response.json
    assert sess["job"]["result"]["ok"], sess["job"]["result"]
    assert sess["design_settings"]["k"] == 2
    assert len(sess["design"]["tables"].nozzles) == 2
    assert sess["design"]["got"]["worst"]["heads"] == [0,1]
    assert "기준개수 K" not in dict(sess["design"]["tables"].meta)
    if mode != 'tree':
        got = sess['design']['got']
        assert set(got['edge_ref'].values()) == sess['worst']['edges']
        assert got['cycle_rank'] == 0
        report = client.get('/api/module-f/edit/flow/report', query_string={'sid':sess['id']}).json
        assert report['scenario']['scope'] == 'region_and_supply'
        assert len(report['scenario']['pipes']) == len(sess['worst']['edges'])
        from src.pipenet_converter.graph.flow import edge_key
        assert {edge_key(a,b) for a,b in zip(sess['worst']['worst_path'], sess['worst']['worst_path'][1:])} <= sess['worst']['edges']
        emitted = client.post('/api/module-f/design/emit', json={"sid":sess["id"]})
        assert emitted.status_code == 200, emitted.json
        from services.cad_import.design.emit import _pc_models
        _pc_models()
        from pipenet_converter.sdf_parser import parse_sdf
        parsed = parse_sdf(sess['design_sdf_path'])
        assert len(parsed.nozzles) == 2


def test_old_all_path_area_result_must_be_reextracted(flow_client):
    from services.cad_import.design.flow import network_for_board
    from routes.module_f.area_selection import validate_area_selection
    client, sess = flow_client
    sess['edit'].board.network_mode = 'loop'
    client.post('/api/module-f/edit/worst', json={'sid':sess['id'], 'selection_mode':'area_all',
        'zones':[[1900,1900,3100,2100]]})
    sess['worst']['edges'] = set(network_for_board(sess['edit'].board).selected_edges([0,1]))
    with pytest.raises(ValueError, match='추출 범위'):
        validate_area_selection(sess)


def test_no_area_no_fallback_and_changed_area_invalidates(flow_client):
    client, sess = flow_client
    sid = {"sid":sess["id"]}
    assert client.post('/api/module-f/edit/worst', json={**sid,"k":1}).status_code == 200
    assert len(sess["worst"]["heads"]) == 1  # F unchanged
    missing = client.post('/api/module-f/edit/worst', json={**sid,"selection_mode":"area_all"})
    assert missing.status_code == 400
    assert not sess["worst"]
    assert client.post('/api/module-f/design/build', json=sid).status_code == 409
    region = [[1900,1900,2100,2100]]
    response = client.post('/api/module-f/edit/worst', json={**sid,"selection_mode":"area_all","zones":region})
    assert response.status_code == 200, response.json
    assert sess["worst"]["heads"] == [0]
    client.post('/api/module-f/edit/zones', json={**sid,"zones":[[2900,1900,3100,2100]]})
    assert client.post('/api/module-f/design/build', json=sid).status_code == 409


def test_disconnected_head_is_not_silently_dropped(flow_client):
    client, sess = flow_client
    b = sess["edit"].board
    b.edges = b.edges - {(3,5),(4,5)}
    response = client.post('/api/module-f/edit/worst', json={"sid":sess["id"],
        "selection_mode":"area_all", "zones":[[1900,1900,3100,2100]]})
    assert response.status_code == 400, response.json
    assert response.json["not_attached"]["items"][0]["disk"] == 1
    assert not sess["worst"]


def test_selection_survives_transaction_and_area_save_load(flow_client, tmp_path):
    import copy
    from routes.module_f.cancellation import install
    from services.cad_import.edit.io import write_edits, load_edits
    client, sess = flow_client
    install(client.application)
    before = sess['edit']
    zones = [[1900,1900,3100,2100]]
    response = client.post('/api/module-f/edit/worst', json={"sid":sess["id"],
        "selection_mode":"area_all", "zones":zones})
    assert response.status_code == 200, response.json
    assert sess['edit'] is not before and sess['selection_mode'] == 'area_all'
    assert getattr(before.board, 'selection_zones', []) == []
    write_edits(sess['edit'].board, str(tmp_path))
    restored = copy.deepcopy(before.board)
    assert load_edits(restored, str(tmp_path))
    assert restored.selection_zones == zones


def test_large_area_http_and_settings_have_no_200_limit(flow_client):
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import EditSession
    from routes.module_f.api_design import _settings
    client, sess = flow_client
    count = 211
    points = [(x*1000,0) for x in range(count+1)]
    points += [(x*1000,2000) for x in range(1,count+1)]
    edges = {(i,i+1) for i in range(count)} | {(i,count+i) for i in range(1,count+1)}
    board = EditBoard('area-many',points,edges,[(x*1000,2000,30) for x in range(1,count+1)])
    board.sources = [0]
    board.network_mode = 'grid'
    for disk in board.disks:
        board.set_head_kind(disk,'상향식')
    sess['edit'] = EditSession(board,key=board.key)
    response = client.post('/api/module-f/edit/worst', json={'sid':sess['id'],
        'selection_mode':'area_all','k':13,'zones':[[500,1900,212000,2100]]})
    assert response.status_code == 200, response.json
    assert response.json['summary']['k'] == count
    assert _settings(sess, {'k':13})['k'] == count
    built = client.post('/api/module-f/design/build', json={'sid':sess['id'],
        'selection_mode':'area_all','k':13,'review_default_mm':65})
    assert built.status_code == 200, built.json
    assert sess['job']['result']['ok'], sess['job']['result']
    assert len(sess['design']['tables'].nozzles) == count
    emitted = client.post('/api/module-f/design/emit', json={'sid':sess['id']})
    assert emitted.status_code == 200, emitted.json
    from services.cad_import.design.emit import _pc_models
    _pc_models()
    from pipenet_converter.sdf_parser import parse_sdf
    assert len(parse_sdf(sess['design_sdf_path']).nozzles) == count

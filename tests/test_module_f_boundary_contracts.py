"""Shared AV points must not produce blind branches or fictional seam pipes."""
from copy import deepcopy

import pytest

from src.pipenet_converter.graph.boundaries import separate_valve_picks
from test_module_f_flow_integration import flow_client


def test_exact_source_pick_is_shared_but_nearby_valve_is_independent():
    assert separate_valve_picks([(200000., -300000.), (200001., -300000.)],
                               [(200000., -300000.)]) == ([1], [0])
    assert separate_valve_picks([(0., 0.)], []) == ([0], [])
    with pytest.raises(ValueError):
        separate_valve_picks([(float('nan'), 0.)], [])


def test_corridor_root_keeps_alarm_loss_without_default_blind_pipes(flow_client):
    client, sess = flow_client
    board = sess['edit'].board
    before = deepcopy(sess['edit'].convert_payload())
    client.post('/api/module-f/edit/flow', json={'sid': sess['id']})
    client.post('/api/module-f/edit/worst', json={'sid': sess['id'], 'k': 1})
    client.post('/api/module-f/design/build', json={'sid': sess['id'], 'k': 1})
    assert sess['job']['result']['ok'], sess['job']['result']
    tables = sess['design']['tables']
    assert any(e['desc'] == 'A/V' for e in tables.equipment)
    assert all(p['head_count_selected'] > 0 for p in tables.pipes)
    assert not any(p['length'] == 2.5 for p in tables.pipes)
    assert sess['edit'].convert_payload() == before
    assert board.valves == [0]


def test_only_calculation_source_valve_geometry_is_suppressed(monkeypatch):
    from services.cad_import.design.restrict import apply_vertical
    from services.cad_import.convert import engine
    payload = {'valve_picks': [{'xy': [12., 34.], 'note': 'source'},
                               {'xy': [12.1, 34.], 'note': 'other'}],
               'sources': [{'xy': [12., 34.]}]}
    before = deepcopy(payload)
    captured = {}
    def convert(sub, out, **kwargs):
        captured.update(sub)
        return {'ok': True, 'kfp': sub['kfp']}
    monkeypatch.setattr(engine, 'convert_to_kfp', convert)
    apply_vertical(payload, {'kfp': {}}, convert_kwargs={'valve_1_m': 2.5})
    assert captured['valve_picks'] == [payload['valve_picks'][1]]
    assert payload == before


def test_original_pipe_provenance_survives_shared_nodes_and_label_collisions():
    from test_module_f_system_layout import heads, system
    from routes.module_f.merge import merge_network
    plan, riser = heads(), system(True)
    plan.pipes[0]['label'] = 'r1'  # Same original label in two source drawings.
    got = merge_network(plan, riser=riser, mode='lsp_gravity')
    assert got['pipe_parts']['r1'] == 'system'
    assert got['pipe_parts']['r1_2'] == 'plan'
    assert got['pipe_parts']['r3'] == 'system'  # Its output is the shared AV.
    assert len(got['combined'].pipes) == len(plan.pipes) + len(riser['pipes'])
    assert sum(n['label'] == '10' for n in got['combined'].nodes) == 1

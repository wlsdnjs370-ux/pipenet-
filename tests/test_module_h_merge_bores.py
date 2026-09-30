"""Drawing-first integration preserves values and source identity, without guessing."""
from copy import deepcopy

from test_module_f_merge import _sample, _riser
from test_module_f_network_editor import client, session, post
from routes.module_f.merge import merge_network
from routes.module_f import api_merge, network_edit as ne


def test_preserve_defined_and_unknown_bores_without_mutating_source():
    table = _sample()
    table.pipes[0]['dia'] = 32
    table.pipes[1]['dia'] = 0
    before = deepcopy(table.pipes)
    riser = _riser()
    riser['pipes'][0]['label'] = 'P1'
    got = merge_network(table, riser=riser, mode='lsp_gravity', preserve_defined_bores=True)
    mapped = {got['plan_pipe_sources'][p['label']]: p for p in got['combined'].pipes
              if p['label'] in got['plan_pipe_sources']}
    assert {k: p['dia'] for k, p in mapped.items()} == {'P1': 32, 'P2': 0}
    assert mapped['P1']['label'] != 'P1'
    assert table.pipes == before
    legacy = merge_network(table, riser=riser, mode='lsp_gravity')
    assert next(p for p in legacy['combined'].pipes if p['label']=='P2')['dia'] == 25


def test_h_merge_and_edit_follow_colliding_plan_label(client, session):
    session['design_settings'] = {'diameter_policy': 'drawing_first_v1'}
    table = session['design']['tables']
    table.equipment = []  # This test is about identity, not an unspecified valve library.
    table.pipes[0].update(dia=32, inner_mm=35.7,
        bore_provenance={'policy':'drawing_first_v1','text_mm':32,'text_raw':'32A'})
    table.pipes[1].update(dia=0)
    riser = _riser()
    riser['pipes'][0]['label'] = 'P1'
    session['slots'] = {'system': {'riser': riser}}
    session['supply_mode'] = 'lsp_gravity'
    api_merge.rebuild_merged(session, persist_overrides=False)
    got = session['merged']
    mapped = next(k for k,v in got['plan_pipe_sources'].items() if v=='P1')
    values = {p['label']:p['dia'] for p in got['combined'].pipes}
    assert values['P1'] == 100 and values[mapped] == 32 and values['P2'] == 0
    response = post(client,session,dict(op='pipe',target=mapped,schedule='KSD 3507',dn=50),scope='merge')
    assert response.status_code == 200, response.json
    assert response.json['selection']['label'] == mapped
    assert session['design']['tables'].pipes[0]['dia'] == 50
    net = ne.ensure(session,'merge')['current']
    assert net.pipes['P1'].row['dia'] == 100
    assert net.pipes[mapped].row['dia'] == 50
    assert net.pipes[mapped].row['bore_provenance']['text_raw'] == '32A'


def test_h_policy_follows_inactive_plan_slot(session):
    from routes.module_f.slots import _slot_switch
    session['design_settings'] = {'diameter_policy':'drawing_first_v1'}
    session['design']['tables'].pipes[0]['dia'] = 40
    session['design']['tables'].pipes[1]['dia'] = 0
    session['slots']['system']['riser'] = _riser()
    session['supply_mode'] = 'lsp_gravity'
    _slot_switch(session, 'system')
    api_merge.rebuild_merged(session, persist_overrides=False)
    assert {p['label']:p['dia'] for p in session['merged']['combined'].pipes
            if p['label'] in ('P1','P2')} == {'P1':40,'P2':0}

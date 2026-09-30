"""Explicit endpoint fitting edits resolve only their own review spot."""
from copy import deepcopy

import pytest

from core.network_editor import Network, apply_edit
from routes.module_f import network_edit as ne, api_merge
from routes.module_f.fitting_inspection import build_inspection, build_merged_inspection
from routes.module_f.hydraulic_sizing import prepare, fingerprint
from routes.module_f.merge import merge_network
from src.pipenet_converter.graph.fitting_review import fitting_review, sync_fitting_review
from test_module_f_network_editor import table, client, session, post, state
from test_module_f_merge import _riser


def unresolved(tbl):
    tbl.meta = [('부속 판정 불가', '1')]
    tbl.unresolved = dict(kind_items=[dict(pipe='KP2',node='N2',pipe_label='P2',
        node_label='2',n=1,reason='original source ambiguity')], applied=[])
    return tbl


def add(tbl, **kw):
    command = dict(op='fitting', target='P2', node='2', fitting='ELBOW_90_STD', count=1)
    command.update(kw)
    return apply_edit(Network.from_tables(tbl), command, ne.catalog())[0].to_tables()


def test_exact_resolution_updates_count_keeps_evidence_and_counts_loss_once():
    base = unresolved(table()); before = deepcopy(base)
    tbl = add(base)
    assert base == before
    review = fitting_review(tbl)
    assert review.count == 0 and len(review.resolved) == 1
    assert tbl.unresolved['kind_items'] == base.unresolved['kind_items']
    assert dict(tbl.meta)['부속 판정 불가'] == '0'
    net, _, warnings = prepare(tbl, {})
    p = next(p for p in net.pipes if p.label == 'P2')
    assert p.sizes[0].equivalent_m == tbl.equipment[-1]['eq_len']
    assert not p.locked and len(p.sizes) > 1
    definition = next(f for f in ne.catalog()['fittings'] if f['id']=='ELBOW_90_STD')
    assert all(s.equivalent_m == definition['lengths'][str(int(s.nominal_mm))] for s in p.sizes)
    assert any('사용자 지정 부속' in w for w in warnings)
    rows = build_inspection(tbl, {}, {}, tbl.nodes)['fittings']
    assert not any(r['kind'] == 'unresolved' for r in rows)
    assert sum(r['kind'] == 'ELBOW_90_STD' for r in rows) == 1
    from routes.module_f.api_design import _tables_for_screen
    before_view = deepcopy(tbl)
    screen = _tables_for_screen(tbl)
    assert screen['unresolved']['kind_items'] == []
    assert len(screen['unresolved']['editor_resolved']) == 1
    assert tbl == before_view
    # Deleting the user's fitting restores the unresolved spot and sizing guard.
    restored, _ = apply_edit(Network.from_tables(tbl), dict(op='remove_fitting',
        target='P2',equipment=tbl.equipment[-1]['label']),ne.catalog())
    assert fitting_review(restored.tables).count == 1
    assert dict(restored.tables.meta)['부속 판정 불가'] == '1'
    with pytest.raises(ValueError, match='노드 2 / 배관 P2'):
        prepare(restored.tables, {})


@pytest.mark.parametrize('kw', [dict(node='3'),dict(node=''),
    dict(target='P1'),dict(fitting='VALVE_SWING_CHECK')])
def test_other_endpoint_pipe_unbound_fitting_or_valve_does_not_resolve(kw):
    tbl = add(unresolved(table()), **kw)
    assert fitting_review(tbl).count == 1
    with pytest.raises(ValueError, match='부속 판정 미확정'):
        prepare(tbl, {})


def test_partial_resolution_and_unknown_legacy_count_are_never_lost():
    tbl = unresolved(table())
    tbl.unresolved['kind_items'].append(dict(pipe_label='P1', node_label='1', n=1))
    tbl.meta = [('부속 판정 불가', '3')]  # one legacy issue without a position
    tbl = add(tbl)
    for _ in range(3):
        sync_fitting_review(tbl)
        assert fitting_review(tbl).count == 2
        assert dict(tbl.meta)['부속 판정 불가'] == '2'
    assert fitting_review(tbl).unlocated_count == 1


def test_legacy_stale_count_is_reconciled_read_only_and_fingerprint_tracks_issues():
    tbl = add(unresolved(table()))
    tbl.meta = [('부속 판정 불가', '1')]
    before = deepcopy(tbl)
    assert fitting_review(tbl).count == 0
    prepare(tbl, {})
    assert tbl == before
    original = fingerprint(tbl)
    tbl.unresolved['kind_items'].append(dict(pipe_label='P1', node_label='1'))
    assert fingerprint(tbl) != original
    with pytest.raises(ValueError, match='부속 판정 미확정'):
        prepare(tbl, {})


def test_merge_shifts_resolution_and_renames_pipe_without_resolving_riser():
    tbl = unresolved(table())
    # Intentional collision with the riser's r1 pipe.
    tbl.pipes[1]['label'] = 'r1'
    tbl.unresolved['kind_items'][0]['pipe_label'] = 'r1'
    tbl = add(tbl, target='r1')
    before = deepcopy(tbl)
    got = merge_network(tbl, riser=_riser(), mode='lsp_gravity')
    c = got['combined']
    review = fitting_review(c)
    assert review.count == 0
    assert review.resolved[0]['node_label'] == '11'
    assert review.resolved[0]['pipe_label'] == 'r1_2'
    assert review.resolved[0]['node'] == 'N2'  # source identity is retained
    row = next(r for r in c.equipment if r.get('editor_library'))
    assert (row['editor_node'],row['pipe']) == ('11','r1_2')
    assert tbl == before
    rows = build_merged_inspection(c,dict(tables=tbl),c.nodes,offset=9)['fittings']
    assert not any(r['kind'] == 'unresolved' for r in rows)


def test_edit_from_merged_view_then_undo_redo_updates_both_views_and_sizing(client,session):
    unresolved(session['design']['tables'])
    # The synthetic alarm-valve host must have library-supported valve DN.
    session['design']['tables'].pipes[0]['dia'] = 65
    riser = _riser()
    for row in riser['pipes']:
        row.update(type='KSD 3507', c=120)
    session.update(slots={'system':{'riser':riser}}, supply_mode='lsp_gravity')
    api_merge.rebuild_merged(session,persist_overrides=False)
    ne.ensure(session,'merge')
    result = post(client,session,dict(op='fitting',target='P2',node='11',
        fitting='ELBOW_90_STD',count=1),scope='merge')
    assert result.status_code == 200, result.json
    for scope in ('design','merge'):
        assert fitting_review(ne.ensure(session,scope)['current'].tables).count == 0
    prepare(session['merged']['combined'], {})
    result = post(client,session,action='undo',scope='merge')
    assert result.status_code == 200, result.json
    for scope in ('design','merge'):
        assert fitting_review(ne.ensure(session,scope)['current'].tables).count == 1
    with pytest.raises(ValueError,match='노드 11 / 배관 P2'):
        prepare(session['merged']['combined'], {})
    result = post(client,session,action='redo',scope='merge')
    assert result.status_code == 200, result.json
    prepare(session['merged']['combined'], {})

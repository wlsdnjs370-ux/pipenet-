"""Native text reference edits use authoritative libraries and canonical history."""
from copy import deepcopy
from dataclasses import asdict

import pytest

from core.network_editor import Network, apply_edit
from routes.module_f import network_edit as ne
from src.pipenet_converter.graph.diameter_inference import DiameterAnnotation
from test_module_f_network_editor import session, client, post, state, table


def annotation(identity='dia-50', nominal=50):
    return DiameterAnnotation(identity, 12500, 8400, nominal, 90, 120, 'FIRE-DIA', f'SP {nominal}', 'AB12')


def test_reference_changes_library_and_audit_without_geometry():
    base = Network.from_tables(table())
    base.pipes['P2'].row['bore_provenance'] = dict(text_mm=25, text_raw='25A', text_xy_mm=[1, 2])
    ref = annotation()
    changed, _ = apply_edit(base, dict(op='pipe_reference', target='P2', annotation=asdict(ref)), ne.catalog())
    pipe = changed.pipes['P2'].row
    spec = next(p for p in ne.catalog()['pipes'] if p['type'] == pipe['type'] and p['dia'] == 50)
    assert pipe['dia'] == 50 and pipe['inner_mm'] == spec['inner_mm']
    assert pipe['bore_provenance']['reference_annotation'] == asdict(ref)
    assert pipe['bore_provenance']['text_raw'] == '25A'  # Original automatic audit is retained.
    assert pipe['length'] == base.pipes['P2'].row['length']
    assert changed.nodes == base.nodes
    assert base.pipes['P2'].row['dia'] == 25
    direct, _ = apply_edit(changed, dict(op='pipe', target='P2', schedule=pipe['type'], dn=80), ne.catalog())
    assert 'reference_annotation' not in direct.pipes['P2'].row['bore_provenance']


def test_api_references_resolved_on_server_saved_replayed_and_undoable(client, session):
    ref = annotation()
    from routes.module_f.api_design import _DEFAULT_SETTINGS
    session['design_settings'] = dict(_DEFAULT_SETTINGS, diameter_policy='drawing_first_v1')
    session['_h_diameter_annotations'] = [ref, annotation('same-dn-other-position')]
    session['_h_diameter_annotations'][1] = DiameterAnnotation('other', 15000, 2000, 50, raw_text='50')
    before = deepcopy(session['design']['tables'])
    command = dict(op='pipe_reference', target='P2', annotation_id=ref.identity, dn=999,
                   annotation=asdict(annotation(nominal=999)))
    response = post(client, session, command, action='preview')
    assert response.status_code == 200, response.json
    assert session['design']['tables'].pipes == before.pipes
    response = post(client, session, command)
    assert response.status_code == 200, response.json
    saved = state(client, session)
    row = next(p for p in saved['pipes'] if p['label'] == 'P2')
    assert row['bore_provenance']['reference_annotation'] == asdict(ref)
    assert row['dia'] == 50 and saved['undo'] == 1
    # Switching the reference at the same DN is still a real undoable operation.
    assert post(client, session, dict(op='pipe_reference', target='P2', annotation_id='other')).status_code == 200
    assert post(client, session, action='undo').status_code == 200
    assert state(client, session)['pipes'][1]['bore_provenance']['reference_annotation']['identity'] == ref.identity
    assert post(client, session, action='undo').status_code == 200
    assert state(client, session)['pipes'][1]['dia'] == 25
    assert post(client, session, action='redo').status_code == 200
    session.pop('network_editor')
    restored = state(client, session)
    assert restored['pipes'][1]['bore_provenance']['reference_annotation'] == asdict(ref)
    assert [(p['in'], p['out'], p['length']) for p in restored['pipes']] == [(p['in'], p['out'], p['length']) for p in before.pipes]
    for identity in ['foreign-id', None]:
        rejected = post(client, session, dict(op='pipe_reference', target='P2', annotation_id=identity))
        assert rejected.status_code == 409
    assert state(client, session)['revision'] == restored['revision']


def test_invalid_text_or_library_size_leaves_network_unchanged():
    net = Network.from_tables(table())
    for raw in [dict(asdict(annotation()), raw_text=''), asdict(annotation(nominal=999))]:
        with pytest.raises(ValueError):
            apply_edit(net, dict(op='pipe_reference', target='P2', annotation=raw), ne.catalog())
    assert net.pipes['P2'].row['dia'] == 25


def test_reference_patch_is_one_commit_and_preserves_revision_guards(client, session, monkeypatch):
    """Patch responses skip full inspection/projection, not validation/history."""
    session['design_settings'] = dict(diameter_policy='drawing_first_v1')
    session['_h_diameter_annotations'] = [annotation()]
    revision = state(client, session)['revision']
    monkeypatch.setattr(ne, 'project', lambda *_args: pytest.fail('reference projected entire graph'))
    monkeypatch.setattr(ne, 'public_state', lambda *_args: pytest.fail('reference rebuilt full inspector'))
    body = dict(sid=session['id'], scope='design', action='apply', revision=revision,
                response_mode='reference_patch', command=dict(op='pipe_reference', target='P2', annotation_id='dia-50'))
    response = client.post('/api/module-f/network-editor', json=body)
    assert response.status_code == 200, response.json
    patch = response.json['reference_patch']
    assert patch['pipe']['dia'] == 50 and patch['pipe']['inner_mm'] > 50
    assert patch['undo'] == 1 and patch['redo'] == 0 and patch['revision'] != revision
    assert patch['pipe'] == session['design']['tables'].pipes[1]
    assert patch['equipment'] == session['design']['tables'].equipment
    # A duplicate/stale request cannot silently create a second history entry.
    assert client.post('/api/module-f/network-editor', json=body).status_code == 409
    assert session['network_editor']['cursor'] == 1
    body.update(revision=patch['revision'])
    body['command']['annotation_id'] = 'foreign'
    assert client.post('/api/module-f/network-editor', json=body).status_code == 409
    assert session['network_editor']['cursor'] == 1


def test_attributes_show_current_reference_not_older_stable_override(monkeypatch, tmp_path):
    from types import SimpleNamespace
    from routes import module_h_attributes as attributes
    from src.pipenet_converter.graph.drawing_reference import record_drawing_reference

    tbl = table()
    row = next(p for p in tbl.pipes if p['label'] == 'P2')
    record_drawing_reference(row, annotation())
    row['dia'] = 50
    board = SimpleNamespace(pts=[], edges=[], disk_kinds=[], sources=[], disks=[])
    context = SimpleNamespace(flow=SimpleNamespace(between=lambda *args: []), decisions={}, summary={})
    sess = dict(edit=SimpleNamespace(board=board), _h_diameter_context=context,
                design_settings={'diameter_policy': 'drawing_first_v1'},
                design=dict(got={}, tables=tbl, editor_canonical=True))
    monkeypatch.setattr(attributes.ov, 'label_keys', lambda *args: {'pipe': {'P2': ['pipe', 0, 1]}, 'node': {}})
    monkeypatch.setattr(attributes.ov, 'ensure_loaded', lambda s: [dict(key=['pipe', 0, 1], field='dia', new=25)])
    monkeypatch.setattr(attributes.ov, 'path_for', lambda key: str(tmp_path / 'overrides.json'))
    result = next(r for r in attributes.state(sess)['records'] if r['label'] == 'P2')
    assert result['nominal_mm'] == 50
    assert result['source'] == 'user'
    assert result['evidence']['reference_annotation']['identity'] == annotation().identity
    assert result['inner_mm'] > 50

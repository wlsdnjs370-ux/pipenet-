"""Direct node menu actions preserve geometry, losses and reversible exports."""
from copy import deepcopy

import pytest

from core.network_editor import Network, EditError, apply_edit, merge_reason
from routes.module_f import network_edit as ne
from test_module_f_network_editor import table, cmd, client, session, state, post


def test_merge_intermediate_node_preserves_pipe_length_and_equipment_position():
    n=Network.from_tables(table())
    original=sum(p.row['length'] for p in n.pipes.values())
    position=n.tables.equipment[0]['rel_pos']*n.pipes['P1'].row['length']
    merged,r=apply_edit(n,dict(op='merge_node',target='2'),ne.catalog())
    assert r['label']=='P1'
    assert set(merged.nodes)=={'1','3'} and len(merged.pipes)==1
    assert merged.pipes['P1'].row['length']==pytest.approx(original)
    assert merged.tables.equipment[0]['rel_pos']*original==pytest.approx(position)
    assert merged.pipes['P1'].b=='3' and merged.tables.nozzles==n.tables.nozzles


@pytest.mark.parametrize('cause',['diameter','bend','head','branch','fitting','root'])
def test_meaningful_nodes_cannot_be_merged(cause):
    n=Network.from_tables(table())
    if cause=='diameter': n.pipes['P2'].row['dia']=50
    if cause=='bend': n.nodes['3'].xyz=(93.1,40,0)
    if cause=='head': n.tables.nozzles[0]['in']='2'
    if cause=='branch': n,_=apply_edit(n,cmd(),ne.catalog())
    if cause=='fitting': n,_=apply_edit(n,dict(op='fitting',target='P2',node='2',fitting='ELBOW_45'),ne.catalog())
    if cause=='root': n.protected.add('2')
    assert merge_reason(n,'2')
    with pytest.raises(EditError):apply_edit(n,dict(op='merge_node',target='2'),ne.catalog())


def test_node_delete_and_merge_roundtrip_exports(client,session):
    before=Network.from_tables(session['design']['tables']).to_tables().as_dict()
    assert post(client,session,dict(op='merge_node',target='2')).status_code==200
    kfp=session['design']['got']['kfp']
    assert set(kfp['nodes_meta_runtime'])=={'N1','N3'}
    assert kfp['pipe_data']['KP1']['end']=='N3'
    assert post(client,session,action='undo').status_code==200
    assert session['design']['tables'].as_dict()==before
    assert post(client,session,dict(op='delete_node',target='3')).status_code==200
    assert not session['design']['tables'].nozzles
    assert post(client,session,action='undo').status_code==200
    assert session['design']['tables'].as_dict()==before


@pytest.mark.parametrize('cut',[False,True])
def test_clipboard_creates_axis_branch_and_carries_head(cut):
    n=Network.from_tables(table())
    command=dict(op='paste',target='2',source='3',source_pipe='P2',axis='-Y',
                 length=3,schedule='CPVC2',dn=25,cut=cut)
    edited,r=apply_edit(n,command,ne.catalog())
    end=edited.nodes[r['label']].xyz
    assert end==(n.nodes['2'].xyz[0],n.nodes['2'].xyz[1]-3,n.nodes['2'].xyz[2])
    assert len(edited.tables.nozzles)==(1 if cut else 2)
    assert len(edited.pipes)==(2 if cut else 3)
    assert edited.pipes[r['pipe']].row['inner_mm']==28.02
    assert len(n.nodes)==3 and 'P2' in n.pipes


def test_clipboard_old_revision_rejected_and_undo_restores_cut(client,session):
    revision=state(client,session)['revision']
    c=dict(op='paste',target='2',source='3',source_pipe='P2',axis='-Y',
           length=3,schedule='KSD 3507',dn=25,cut=True,copied_revision=revision)
    assert post(client,session,c,action='preview').status_code==200
    assert len(session['design']['tables'].pipes)==2
    assert post(client,session,c).status_code==200
    assert post(client,session,c).status_code==409
    assert post(client,session,action='undo').status_code==200
    assert 'KP2' in session['design']['got']['kfp']['pipe_data']
    assert post(client,session,action='redo').status_code==200


def test_move_node_updates_length_without_moving_neighbors():
    n=Network.from_tables(table())
    xyz=n.nodes['3'].xyz
    out,_=apply_edit(n,dict(op='move_node',target='3',x=xyz[0]+2,y=xyz[1],z=xyz[2]),ne.catalog())
    assert out.pipes['P2'].row['length']==pytest.approx(n.pipes['P2'].row['length']+2)
    assert out.nodes['2'].xyz==n.nodes['2'].xyz
    with pytest.raises(EditError):
        apply_edit(n,dict(op='move_node',target='3',x=xyz[0],y=xyz[1]+2,z=0),ne.catalog())


def test_node_fitting_tracks_its_endpoint_when_pipe_is_split():
    n=Network.from_tables(table())
    n,_=apply_edit(n,dict(op='fitting',target='P2',node='3',fitting='ELBOW_45'),ne.catalog())
    n,_=apply_edit(n,dict(op='split',target='P2',distance=.5),ne.catalog())
    f=n.tables.equipment[-1]
    assert f['pipe']=='EP1' and f['editor_node']=='3' and f['rel_pos']==1


def test_minus_y_length_in_isometric_and_kfp(client,session):
    c=cmd();c.update(axis='-Y',length=3)
    assert post(client,session,c).status_code==200
    d=session['design']
    a=d['got']['kfp']['nodes_meta_runtime']['N2']['coords']
    p=d['got']['kfp']['pipe_data']['EP1']
    b=d['got']['kfp']['nodes_meta_runtime'][p['end']]['coords']
    assert b==[a[0],a[1]-3,a[2]] and p['length_m']==3
    from services.cad_import.design.emit import display_tables
    view,_=display_tables(d['tables'],iso=True)
    at={n['label']:n for n in view.nodes}
    scale=view.norm['scale']*1000
    assert at['4']['x']-at['2']['x']==pytest.approx(3*scale*3**.5/2)
    assert at['4']['y']-at['2']['y']==pytest.approx(-3*scale/2)


def test_copy_preserves_pipe_equipment_and_native_fittings():
    t=table();t.pipes[1]['eq_len']=.6
    t.fittings=[dict(label='F1',pipe='P2',type='elbow',count=1)]
    t.equipment.append(dict(label='V1',pipe='P2',rel_pos=.4,eq_len=1.2,desc='valve',op_id='existing'))
    n=Network.from_tables(t)
    edited,r=apply_edit(n,dict(op='paste',target='2',source='3',source_pipe='P2',axis='-Y',length=2,
                              schedule='KSD 3507',dn=25),ne.catalog())
    assert edited.pipes[r['pipe']].row['eq_len']==.6
    eq=next(e for e in edited.tables.equipment if e['pipe']==r['pipe'])
    assert eq['label']!='V1' and eq['eq_len']==1.2 and eq['rel_pos']==.4
    assert any(f['pipe']==r['pipe'] and f['count']==1 for f in edited.tables.fittings)


def test_merged_plan_properties_use_plan_coordinate_origin(client,session):
    from test_module_f_merge import _riser
    from routes.module_f.api_merge import rebuild_merged
    session['slots']['system']['riser']=_riser();session['supply_mode']='lsp_gravity'
    rebuild_merged(session)
    xyz=next(n['xyz'] for n in state(client,session,'merge')['nodes'] if n['label']=='12')
    old=ne.ensure(session,'design')['current'].nodes['3'].xyz
    r=post(client,session,dict(op='move_node',target='12',x=xyz[0]+1,y=xyz[1],z=xyz[2]),scope='merge')
    assert r.status_code==200,r.json
    assert ne.ensure(session,'design')['current'].nodes['3'].xyz==pytest.approx((old[0]+1,old[1],old[2]))


def test_merged_system_z_extension_and_split_visible_in_iso(client,session):
    from test_module_f_merge import _riser
    from routes.module_f.api_merge import rebuild_merged
    from routes.module_f.merge import bake_combined_iso
    session['slots']['system']['riser']=_riser();session['supply_mode']='lsp_gravity'
    rebuild_merged(session)
    p=next(p for p in session['merged']['combined'].pipes if str(p['label']).startswith('r'))
    r=post(client,session,dict(op='split',target=p['label'],distance=float(p['length'])/2),scope='merge')
    assert r.status_code==200,r.json
    start=r.json['selection']['label']
    c=cmd(target=start);c.update(axis='+Z',length=.75)
    r=post(client,session,c,scope='merge');assert r.status_code==200,r.json
    end=r.json['selection']['label'];pipe=r.json['selection']['pipe']
    rows={n['label']:n for n in bake_combined_iso(session['merged'])[0]}
    assert rows[end]['x']==pytest.approx(rows[start]['x'])
    assert rows[end]['y']-rows[start]['y']==pytest.approx(750)
    r=post(client,session,dict(op='split',target=pipe,distance=.25),scope='merge')
    assert r.status_code==200,r.json
    mid=r.json['selection']['label']
    rows={n['label']:n for n in bake_combined_iso(session['merged'])[0]}
    assert rows[mid]['y']-rows[start]['y']==pytest.approx(250)

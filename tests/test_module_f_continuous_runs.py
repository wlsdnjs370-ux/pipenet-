"""One canonical run for selection, editing, tables, history and SDF."""
from copy import deepcopy
import json
import xml.etree.ElementTree as ET

import pytest

from core.network_editor import Network, apply_edit
from routes.module_f import network_edit as ne, api_merge, api_design
from src.pipenet_converter.graph.continuous_runs import compact_continuous
from src.pipenet_converter.graph.bore_provenance import describe_bore
from test_module_f_network_editor import client, session, post, state
from test_module_f_export_compaction import chain


def design_chain():
    """Straight subdivided pipe with a final head connection, like the screenshots."""
    from services.cad_import.design.tables import PipeTablesG
    t=chain()
    # Plan input convention is label 1; merged plan becomes 10 onward.
    for row in t.nodes:
        row['label']=str(int(row['label'])+1)
    for row in t.pipes+t.nozzles:
        for key in ('in','out'):
            if str(row[key]).isdigit(): row[key]=str(int(row[key])+1)
    t=PipeTablesG(nodes=t.nodes,pipes=t.pipes,nozzles=t.nozzles,
                  pipe_labels={f'KP{i}':f'P{i}' for i in range(5)})
    for p in t.pipes:
        p['bore_provenance']=dict(version=2,source='review_default',auto_mm=65,
                                  review_default_mm=65,review_only=True)
    nodes={f'N{n["label"]}':dict(coords=[n['x']/1000,n['y']/1000,0],elevation_m=0,
                               type_id='head' if n['label']=='6' else 'base') for n in t.nodes}
    pipes={f'KP{i}':dict(start='N'+p['in'],end='N'+p['out'],length_m=p['length'],
                        nominal_mm=65,diameter=67.9,C=120,equivalent_length=0)
           for i,p in enumerate(t.pipes)}
    return dict(tables=t,method='manual',got={'kfp':dict(nodes_meta_runtime=nodes,pipe_data=pipes)},
                keys=dict(node={},pipe={},nid={n['label']:'N'+n['label'] for n in t.nodes}))


def ok(r):
    assert r.status_code==200,r.json
    return r.json


def test_screenshot_shaped_segments_become_one_real_pipe():
    t=chain();t.nodes=t.nodes[:3];t.pipes=t.pipes[:2];t.nozzles=[]
    labels={'0':'121','1':'130','2':'137'}
    for n in t.nodes:n['label']=labels[n['label']]
    for p,label,length in zip(t.pipes,('P129','P136'),(2,3)):
        p.update(label=label,length=length)
        p['in'],p['out']=labels[p['in']],labels[p['out']]
    compact,_=compact_continuous(t)
    assert [n['label'] for n in compact.nodes]==['121','137']
    assert len(compact.pipes)==1
    p=compact.pipes[0]
    assert (p['label'],p['in'],p['out'],p['length'],p['dia'])==('P129','121','137',5,65)


def test_compact_preserves_all_evidence_and_uncertainty_without_mutation():
    t=chain()
    for p in t.pipes:
        p['bore_provenance']=dict(source='text',auto_mm=65,text_mm=65,text_xy_mm=[int(p['in']),0])
    t.pipes[1]['bore_provenance']=dict(source='review_default',auto_mm=65,review_only=True)
    before=deepcopy(t)
    compact,audit=compact_continuous(t)
    assert t==before
    row=compact.pipes[0];info=describe_bore(row)
    assert len(info['continuous_sources'])==4 and info['category']=='review' and info['review_only']
    assert info['continuous_sources'][0]['bore_provenance']['text_xy_mm']==[0,0]
    assert compact_continuous(compact)[0]==compact
    assert audit.removed_nodes==('1','2','3')


def test_repeated_compaction_flattens_lineage_and_keeps_deliberate_split():
    t=chain();first,_=compact_continuous(t,keep_nodes=['2'])
    second,_=compact_continuous(first)
    info=describe_bore(second.pipes[0])
    assert [r['label'] for r in info['continuous_sources']]==['P0','P1','P2','P3']
    net=Network.from_tables(second)
    split,res=apply_edit(net,dict(op='split',target='P0',distance=1.2),ne.catalog())
    compact,_=apply_edit(split,dict(op='compact_runs',version=1),ne.catalog())
    assert res['label'] in compact.nodes


def test_existing_crossing_does_not_become_a_connection_or_block_contraction():
    t=chain();t.nodes.extend([dict(label='a',x=2500,y=-1000,elevation=0),
                             dict(label='b',x=2500,y=1000,elevation=0)])
    t.pipes.append(dict(t.pipes[0],label='crossing',**{'in':'a','out':'b'},length=2))
    net=Network.from_tables(t)
    result,_=apply_edit(net,dict(op='compact_runs',version=1),ne.catalog())
    assert len(result.pipes)==3 and result.component('0')=={'0','4','5'}
    assert result.component('a')=={'a','b'}


def test_confirm_then_edit_whole_run_undo_redo_and_reload(client,session,tmp_path,monkeypatch):
    session['design']=design_chain();base=deepcopy(session['design'])
    editor=ne.accept_rebuilt(session)
    assert editor['commands'][0]['op']=='compact_runs'
    assert len(editor['current'].pipes)==len(session['design']['tables'].pipes)==2
    assert session['design']['tables'].pipes[0]['in']=='1'
    assert session['design']['tables'].pipes[0]['out']=='5'
    kfp=session['design']['got']['kfp']
    assert len(kfp['pipe_data'])==2 and len(kfp['nodes_meta_runtime'])==3
    ok(post(client,session,dict(op='pipe',target='P0',schedule='CPVC2',dn=65)))
    row=session['design']['tables'].pipes[0]
    assert row['type']=='CPVC2' and row['length']==4 and row['out']=='5'
    assert len(row['bore_provenance']['continuous_sources'])==4
    ok(post(client,session,action='undo'))
    assert session['design']['tables'].pipes[0]['type']=='KSD 3507'
    ok(post(client,session,action='undo'))
    assert len(session['design']['tables'].pipes)==5
    # Ordinary SDF must agree with the shown (explicitly un-compacted) graph too.
    monkeypatch.setattr(api_design,'_design_stale',lambda s:None)
    out,error=api_design.emit_design_files(session,tmp_path)
    assert error is None,error
    assert len(list(ET.parse(out).getroot().iter('Pipe')))==5
    ok(post(client,session,action='redo'));ok(post(client,session,action='redo'))
    session.pop('network_editor');session['design']=deepcopy(base)
    restored=state(client,session)
    assert restored['undo']==2 and len(restored['pipes'])==2
    assert restored['pipes'][0]['type']=='CPVC2'
    out,error=api_design.emit_design_files(session,tmp_path)
    assert error is None,error
    sdf=[(r.get('label'),r.get('input'),r.get('output'),float(r.get('length'))) for r in ET.parse(out).getroot().iter('Pipe')]
    assert sorted(sdf)==sorted((p['label'],p['in'],p['out'],p['length']) for p in session['design']['tables'].pipes)


def test_old_segment_edits_replay_before_compaction_and_redo_survives(client,session):
    session['design']=design_chain();base=deepcopy(session['design'])
    ok(post(client,session,dict(op='pipe',target='P2',schedule='KSD 3507',dn=80)))
    session['design']=deepcopy(base)
    ne.accept_rebuilt(session)
    assert next(p for p in session['design']['tables'].pipes if p['label']=='P2')['dia']==80
    saved=json.loads(ne._path(session,'design').read_text(encoding='utf-8'))
    assert saved['commands'][0]['target']=='P2' and saved['commands'][-1]['op']=='compact_runs'
    ok(post(client,session,action='undo'))
    before=ne._path(session,'design').read_bytes()
    session['design']=deepcopy(base)
    editor=ne.accept_rebuilt(session)
    assert editor['cursor']==1 and len(editor['commands'])==2
    assert ne._path(session,'design').read_bytes()==before
    ok(post(client,session,action='redo'))
    assert len(session['design']['tables'].pipes)==4


def test_plan_and_merge_keep_same_run_identity_across_edits(client,session):
    from test_module_f_merge import _riser
    session['design']=design_chain();ne.accept_rebuilt(session)
    riser=_riser(5)
    for row in riser['pipes']:
        row.update(type='KSD 3507',c=120,eq_len=0)
    session['slots']['system']['riser']=riser;session['supply_mode']='lsp_gravity'
    api_merge.rebuild_merged(session);ne.accept_rebuilt(session,'merge')
    assert ne.ensure(session,'merge')['commands'][0]['op']=='compact_runs'
    merged=next(p for p in session['merged']['combined'].pipes if p['label']=='P0')
    assert (merged['in'],merged['out'])==('10','14')
    ok(post(client,session,dict(op='pipe',target='P0',schedule='CPVC2',dn=65),scope='merge'))
    assert session['design']['tables'].pipes[0]['type']=='CPVC2'
    updated=next(p for p in session['merged']['combined'].pipes if p['label']=='P0')
    assert updated['type']=='CPVC2' and updated['length']==4 and updated['out']=='14'
    ok(post(client,session,action='undo',scope='merge'))
    assert session['design']['tables'].pipes[0]['type']=='KSD 3507'


def test_combined_marker_rebases_after_generated_node_id_changes():
    from core.network_edit_replay import rebase_commands
    base=Network.from_tables(chain())
    commands=[dict(op='compact_runs',version=1,keep=[]),
              dict(op='pipe',target='P0',schedule='CPVC2',dn=65)]
    current,translated=rebase_commands(base,deepcopy(base),commands,2,ne.catalog())
    assert len(current.pipes)==2 and current.pipes['P0'].row['type']=='CPVC2'
    assert translated==commands


def test_endpoint_losses_survive_contraction_and_recalculate_as_one_pipe(client,session):
    session['design']=design_chain()
    t=session['design']['tables']
    t.fittings=[dict(pipe='P1',node='2',type='elbow',count=1),
                dict(pipe='P3',node='5',type='elbow',count=1)]
    for p in t.pipes:
        if p['label'] in ('P1','P3'):p['eq_len']=1.8
    ne.accept_rebuilt(session)
    row=next(p for p in session['design']['tables'].pipes if p['label']=='P1')
    assert row['in']=='2' and row['out']=='5' and row['eq_len']==3.6
    assert session['design']['got']['kfp']['pipe_data']['KP1']['equivalent_length']==3.6
    ok(post(client,session,dict(op='pipe',target='P1',schedule='KSD 3507',dn=80)))
    row=next(p for p in session['design']['tables'].pipes if p['label']=='P1')
    definition=next(f for f in ne.catalog()['fittings'] if f['id']=='ELBOW_90_STD')
    assert row['eq_len']==2*definition['lengths']['80']
    assert len(session['design']['tables'].fittings)==2


def test_compaction_preview_pure_and_save_failure_rolls_back(client,session,monkeypatch):
    session['design']=design_chain()
    before=deepcopy(session['design'])
    r=ok(post(client,session,dict(op='compact_runs'),action='preview'))
    assert r['counts']['pipes']==2 and session['design']==before
    def fail(*args):raise OSError('disk failure')
    monkeypatch.setattr(ne,'save_history',fail)
    assert post(client,session,dict(op='compact_runs')).status_code==409
    assert session['design']==before and len(ne.ensure(session,'design')['current'].pipes)==5


def test_compact_button_in_merge_updates_plan_owner_not_a_separate_copy(client,session):
    from test_module_f_merge import _riser
    session['design']=design_chain()
    session['slots']['system']['riser']=_riser();session['supply_mode']='lsp_gravity'
    api_merge.rebuild_merged(session)
    before=len(session['merged']['combined'].pipes)
    r=ok(post(client,session,dict(op='compact_runs'),action='preview',scope='merge'))
    assert r['effective_scope']=='design' and r['counts']['pipes']==before-3
    assert len(session['design']['tables'].pipes)==5
    ok(post(client,session,dict(op='compact_runs'),scope='merge'))
    assert len(session['design']['tables'].pipes)==2
    assert next(p for p in session['merged']['combined'].pipes if p['label']=='P0')['out']=='14'

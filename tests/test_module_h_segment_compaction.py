"""H draft spans reduce without inventing bores or breaking pipe identity."""
from copy import deepcopy
from collections import Counter
import json
import xml.etree.ElementTree as ET

import pytest

from core.network_editor import Network, apply_edit
from routes.module_f import network_edit as ne, api_design
from src.pipenet_converter.graph.continuous_runs import compact_continuous
from src.pipenet_converter.validate.diameter_evidence import require_resolved_diameters
from test_module_f_export_compaction import chain
from test_module_f_continuous_runs import design_chain
from test_module_f_network_editor import client, session, post


def unassigned(tables):
    """H's explicitly blocked, zero-DN draft convention."""
    for index, p in enumerate(tables.pipes):
        p.update(dia=0, dia_src='unresolved')
        p.pop('inner_mm', None)
        p['bore_provenance'] = dict(policy='drawing_first_v1', source='unresolved',
            auto_mm=None, block_export=True, review_only=True, reason='no_match',
            stable_key=['pipe', index, index+1], source_edges=[[index, index+1]])
    return tables


def test_unknown_runs_remain_unknown_and_cover_each_original_once():
    t=unassigned(chain());before=deepcopy(t)
    result,audit=compact_continuous(t,allow_unassigned=True)
    assert t==before and len(result.pipes)==2
    assert Counter(p['label'] for rows in audit.source_pipes.values() for p in rows)==Counter(p['label'] for p in t.pipes)
    merged=result.pipes[0]
    assert merged['dia']==0 and merged['length']==4
    assert merged['bore_provenance']['block_export'] is True
    assert merged['bore_provenance']['source']=='unresolved'
    assert len(merged['bore_provenance']['source_edges'])==4
    assert len(merged['bore_provenance']['continuous_sources'])==4
    with pytest.raises(ValueError):
        require_resolved_diameters(result.pipes)
    assert compact_continuous(result,allow_unassigned=True)[0]==result


@pytest.mark.parametrize('case',['known','conflict','partial','candidate','generated','manual','reference','loss','editor'])
def test_review_and_physical_boundaries_cannot_be_swallowed(case):
    t=unassigned(chain());p=t.pipes[1];rec=p['bore_provenance']
    if case=='known':p.update(dia=40);rec.update(block_export=False,text_mm=40)
    elif case=='conflict':rec['reason']='conflicting_texts'
    elif case=='partial':rec['reason']='partial_or_conflicting_path'
    elif case=='candidate':rec['candidate_mm']=[40,50]
    elif case=='generated':rec['reason']='generated_pipe'
    elif case=='manual':rec['manual']={'value_mm':0}
    elif case=='reference':rec['reference_annotation']={'identity':'ref'}
    elif case=='loss':p['eq_len']=None
    elif case=='editor':t.nodes[1]['editor_added']=True;t.nodes[2]['editor_added']=True
    result,_=compact_continuous(t,allow_unassigned=True)
    assert {'1','2'} <= {n['label'] for n in result.nodes}


def test_version_one_replay_unchanged_and_h_requests_version_two(client,session):
    t=unassigned(chain());net=Network.from_tables(t)
    old,_=apply_edit(net,dict(op='compact_runs',version=1),ne.catalog())
    new,_=apply_edit(net,dict(op='compact_runs',version=2),ne.catalog())
    assert len(old.pipes)==5 and len(new.pipes)==2
    session['design']=design_chain()
    assert ne.compaction_command(session,'design')['version']==1
    session['design_settings']={'diameter_policy':'drawing_first_v1'}
    assert ne.compaction_command(session,'design')['version']==2


def test_accept_new_basis_normalizes_once_atomically_and_preserves_archived_edits(client,session,tmp_path,monkeypatch):
    session['design_settings']=dict(api_design._DEFAULT_SETTINGS,diameter_policy='drawing_first_v1')
    session['design']=design_chain()
    assert post(client,session,dict(op='pipe_c',target='P2',c=135)).status_code==200
    saved=ne._path(session,'design').read_bytes()
    session['design']=design_chain()
    unassigned(session['design']['tables'])
    assert ne.accept_rebuilt(session)['conflict']
    response=post(client,session,action='accept_basis')
    assert response.status_code==200,response.json
    fresh=ne.ensure(session,'design')
    assert len(fresh['current'].pipes)==len(session['design']['tables'].pipes)==2
    assert fresh['commands'][0]['op']=='compact_runs' and fresh['commands'][0]['version']==2
    assert any(path.read_bytes()==saved for path,_ in ne._history_versions(session,'design'))
    assert all(p.row['c']==120 and p.row['dia']==0 for p in fresh['current'].pipes.values())
    monkeypatch.setattr(api_design,'_design_stale',lambda _:None)
    path,error=api_design.emit_design_files(session,tmp_path)
    assert error is None,error
    sdf=list(ET.parse(path).getroot().iter('Pipe'))
    assert len(sdf)==2
    assert {p.get('label') for p in sdf}==set(fresh['current'].pipes)
    review=json.loads(path.with_suffix('.review.json').read_text(encoding='utf8'))
    assert review['calculation_ready'] is False
    assert len(session['design']['got']['kfp']['pipe_data'])==2
    assert post(client,session,action='undo').status_code==200
    assert len(session['design']['tables'].pipes)==5
    assert post(client,session,action='redo').status_code==200
    assert len(session['design']['tables'].pipes)==2

"""Evidence cannot change hydraulic choices or disguise an unknown source."""
from copy import deepcopy
from pathlib import Path

import pytest
from flask import Flask

from src.pipenet_converter.graph.bore_provenance import (
    describe_bore, record_adjustment, record_manual,
)
from routes.module_f import api_design, api_merge, jobs, network_edit as ne
from core.network_editor import Network, apply_edit
from test_module_f_network_editor import design, table


def test_nearest_match_and_minimum_policy_unchanged_with_evidence():
    ne._boot()
    from services.cad_import.design.bore import decide_bores, match_diameter_for_segment
    pts=[(0,0),(0,6200)]
    # This is an implementation invariant, not a claim about legal applicability.
    got=decide_bores({'pipe_data':{'K1':{}}},{'K1':(0,1)},{(0,1):4},
                     [(100,3000,32)],pts=pts)
    assert got['K1']==(40,'nfpc_min')
    assert got.evidence['K1']['text_mm']==32
    assert got.evidence['K1']['rule_mm']==40
    assert got.evidence['K1']['text_xy_mm']==[100,3000]
    assert got.evidence['K1']['distance_mm']==100
    assert match_diameter_for_segment(*pts,[(100,3000,32),(50,500,100)])==100
    assert match_diameter_for_segment(*pts,[(1500,3000,32)]) is None


@pytest.mark.parametrize('source,category',[('text','drawing'),('nfpc_min','review'),
    ('nfpc_fallback','rule'),('default','unknown'),(None,'unknown')])
def test_source_categories_read_only(source,category):
    row={'dia':40,'dia_src':source};before=deepcopy(row)
    assert describe_bore(row)['category']==category
    assert row==before


def test_repeated_manual_changes_retain_original_evidence():
    row={'dia':40,'dia_src':'nfpc_fallback'}
    record_manual(row,32,'drawing checked');row['dia']=32
    record_manual(row,50,'second check');row['dia']=50
    info=describe_bore(row)
    assert info['category']=='rule' and info['edited']
    assert info['auto_mm']==40
    assert info['manual']==dict(previous_mm=32,value_mm=50,note='second check')


def test_normalization_and_unrecorded_changes_need_review():
    row={'dia':100,'dia_src':'text'}
    record_adjustment(row,40,'merge normalization')
    assert describe_bore(row)['category']=='review'
    assert describe_bore(row)['auto_mm']==40
    row={'dia':40,'bore_provenance':dict(source='text',auto_mm=32)}
    assert describe_bore(row)['untracked_change']


def test_display_metadata_preserves_old_history_fingerprint():
    n=Network.from_tables(table());before=ne.fingerprint(n)
    n.pipes['P1'].row['bore_provenance']={'source':'text','auto_mm':25}
    assert ne.fingerprint(n)==before
    n.pipes['P1'].row['dia']=32
    assert ne.fingerprint(n)!=before


def test_direct_library_edit_retains_origin_and_exact_inner_bore():
    n=Network.from_tables(table());n.pipes['P2'].row['dia_src']='nfpc_fallback'
    out,_=apply_edit(n,dict(op='pipe',target='P2',schedule='CPVC2',dn=25,note='test'),ne.catalog())
    r=out.pipes['P2'].row
    assert r['inner_mm']==28.02 and r['dia_src']=='user'
    assert describe_bore(r)['category']=='rule'
    assert describe_bore(r)['edited']
    assert 'bore_provenance' not in n.pipes['P2'].row


def test_display_evidence_does_not_change_sdf_or_kfp_exports(tmp_path):
    d=design();before=deepcopy(d['tables'])
    for row in d['tables'].pipes:
        row['bore_provenance']=dict(source='text',auto_mm=25,text_mm=25,
            text_xy_mm=[100,200],distance_mm=10)
    from services.cad_import.design.emit import emit_design_kfp, emit_design_sdf
    assert emit_design_kfp(before,d['got'],None)['kfp']==emit_design_kfp(d['tables'],d['got'],None)['kfp']
    a=tmp_path/'before'/'network.sdf'; b=tmp_path/'after'/'network.sdf'
    emit_design_sdf(before,a,iso=True)
    emit_design_sdf(d['tables'],b,iso=True)
    assert a.read_bytes()==b.read_bytes()


def test_missing_coordinates_does_not_claim_drawing_has_no_diameter_text():
    ne._boot()
    from services.cad_import.design.bore import decide_bores
    got=decide_bores({'pipe_data':{'K1':{}}},{'K1':(0,1)}, {},[(100,200,65)])
    assert got['K1']==(25,'nfpc_fallback')
    assert got.evidence['K1']['reason']=='no_coordinates'


def test_evidence_survives_plan_and_merged_previews(tmp_path, monkeypatch):
    monkeypatch.syspath_prepend(str(Path(ne.__file__).resolve().parents[2] / 'core'))
    from routes.module_f.merge import merge_network
    from test_module_f_merge_view import _riser
    d=design();row=d['tables'].pipes[1]
    row.update(dia_src='text',bore_provenance=dict(version=1,source='text',auto_mm=25,
               text_mm=25,text_xy_mm=[100,200],distance_mm=10))
    sess=jobs._new_session(key='bore-api-test');sess['design']=d
    sess['merged']=merge_network(d['tables'],riser=_riser(),mode='lsp_gravity')
    app=Flask(__name__);app.testing=True
    api_design.register(app,UPLOAD_DIR=tmp_path);api_merge.register(app,UPLOAD_DIR=tmp_path)
    client=app.test_client()
    try:
        for url in ('/api/module-f/design/preview','/api/module-f/merge/preview'):
            r=client.get(url,query_string=dict(sid=sess['id'],iso=1))
            assert r.status_code==200,r.json
            assert r.json['view'] is not None,(url,r.json)
            p=next(p for p in r.json['view']['pipes'] if p['label']=='P2')
            assert p['bore_info']['text_mm']==25
            assert p['bore_info']['category']=='drawing'
    finally:
        jobs._SESSIONS.pop(sess['id'],None)

"""Topology ownership, conservative propagation, and visible export conflicts."""
from copy import deepcopy
import math

import pytest

from src.pipenet_converter.graph.diameter_inference import (
    DiameterAnnotation as Text, InferenceConfig, infer_diameter_annotations,
)
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.graph.bore_provenance import describe_bore, record_manual
from src.pipenet_converter.validate.diameter_evidence import require_resolved_diameters


def tee():
    pts = [(0,0),(6000,0),(12000,0),(6000,4000),(12000,4000)]
    edges = [(0,1),(1,2),(1,3),(2,4)]
    flow = build_flow_tree(pts,edges,[[3],[4]],[0])
    return pts,edges,flow


def test_vertical_25_is_not_given_to_closer_horizontal_main():
    pts,edges,flow = tee()
    before = deepcopy((pts,edges,flow))
    texts = [Text('main',2000,200,65,0),Text('branch',6100,5,25,90)]
    ctx = infer_diameter_annotations(pts,edges,flow,texts)
    assert ctx.for_path((0,1))['text_mm'] == 65
    assert ctx.for_path((1,2))['text_mm'] == 65
    assert ctx.for_path((1,2))['inferred']
    assert ctx.for_path((1,2))['excluded_nearest']['text_mm'] == 25
    assert ctx.for_path((1,3))['text_mm'] == 25
    assert ctx.for_path((0,2))['inferred']
    assert ctx.for_path((0,2))['excluded_nearest']['text_mm'] == 25
    assert ctx.for_path((0,1))['full_head_count'] == 2
    assert (pts,edges,flow) == before
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.design.bore import decide_bores
    # Passing the tempting nearest text again must not bypass topology.
    out = decide_bores({'pipe_data':{'p':{}}},{'p':(1,2)},flow.loads,
                       [(6100,5,25)],pts=pts,context=ctx)
    assert out['p'] == (65,'text')
    assert describe_bore({'dia':65,'bore_provenance':out.evidence['p']})['category'] == 'review'


def test_rotation_of_entire_drawing_does_not_change_ownership():
    pts,edges,flow = tee()
    rad = math.radians(33)
    def rotate(x,y):
        return (x*math.cos(rad)-y*math.sin(rad),x*math.sin(rad)+y*math.cos(rad))
    points = [rotate(*p) for p in pts]
    texts = [Text('main',*rotate(2000,200),65,33),Text('branch',*rotate(6100,5),25,123)]
    ctx = infer_diameter_annotations(points,edges,flow,texts)
    assert ctx.for_path((0,1))['text_mm'] == 65
    assert ctx.for_path((1,3))['text_mm'] == 25


def test_node_numbering_does_not_choose_diameter():
    pts,edges,_ = tee()
    order=[3,0,4,2,1]
    new_id={old:new for new,old in enumerate(order)}
    points=[pts[old] for old in order]
    links=[(new_id[a],new_id[b]) for a,b in edges]
    flow=build_flow_tree(points,links,[[new_id[3]],[new_id[4]]],[new_id[0]])
    ctx=infer_diameter_annotations(points,links,flow,
            [Text('main',2000,200,65,0),Text('branch',6100,5,25,90)])
    assert ctx.for_path((new_id[0],new_id[2]))['text_mm']==65
    assert ctx.for_path((new_id[1],new_id[3]))['text_mm']==25


def test_parallel_runs_are_not_joined_or_tiebroken():
    pts = [(0,0),(5000,0),(0,400),(5000,400)]
    edges = [(0,1),(2,3)]
    flow = build_flow_tree(pts,edges,[[1]],[0])
    ctx = infer_diameter_annotations(pts,edges,flow,[Text('unclear',2500,200,65,0)])
    assert ctx.for_path((0,1))['block_export']
    assert 'text_mm' not in ctx.for_path((0,1))
    ctx = infer_diameter_annotations(pts,edges,flow,[Text('other',2500,390,65,0)])
    assert 'text_mm' not in ctx.for_path((0,1))


@pytest.mark.parametrize('stop', ['valve','bend'])
def test_propagation_does_not_cross_semantic_or_geometric_boundary(stop):
    pts = [(0,0),(5000,0),(10000,5000 if stop=='bend' else 0)]
    edges = [(0,1),(1,2)]
    flow = build_flow_tree(pts,edges,[[2]],[0])
    ctx = infer_diameter_annotations(pts,edges,flow,[Text('main',2000,100,65,0)],
                                     barriers=[1] if stop=='valve' else [])
    assert ctx.for_path((0,1))['text_mm'] == 65
    assert 'text_mm' not in ctx.for_path((1,2))


def test_inline_head_does_not_break_continuous_main_annotation():
    pts = [(0,0),(5000,0),(10000,0)]
    edges = [(0,1),(1,2)]
    flow = build_flow_tree(pts,edges,[[1],[2]],[0])
    ctx = infer_diameter_annotations(pts,edges,flow,[Text('main',2000,100,65,0)])
    assert ctx.for_path((1,2))['text_mm']==65
    assert ctx.for_path((1,2))['inferred']


def test_distinct_sizes_in_one_pipe_require_review_not_invented_reducer():
    pts = [(0,0),(5000,0),(10000,0),(15000,0)]
    edges = [(0,1),(1,2),(2,3)]
    flow = build_flow_tree(pts,edges,[[3]],[0])
    ctx = infer_diameter_annotations(pts,edges,flow,
                                    [Text('a',1000,100,65,0),Text('b',13000,100,40,0)])
    assert ctx.for_path((1,2))['block_export']
    merged = ctx.for_path((0,3))
    assert merged['block_export'] and merged['candidate_mm'] == [40,65]
    one = infer_diameter_annotations(pts,[(0,3)],build_flow_tree(pts,[(0,3)],[[3]],[0]),
                                    [Text('a',1000,100,65,0),Text('b',13000,100,40,0)])
    assert one.for_path((0,3))['block_export']


def test_upstream_inversion_warns_without_rewriting_drawing_values():
    pts,edges,flow = tee()
    ctx = infer_diameter_annotations(pts,edges,flow,
          [Text('main',2000,200,25,0),Text('branch',6100,1000,65,90)])
    assert ctx.for_path((0,1))['text_mm'] == 25
    assert ctx.for_path((1,3))['text_mm'] == 65
    assert any('상류 표기 25 → 하류 표기 65' in w for w in ctx.for_path((0,1))['review_reasons'])


def test_missing_rotation_is_reviewed_and_inference_has_finite_range():
    pts,edges,flow = tee()
    ctx = infer_diameter_annotations(pts,edges,flow,[Text('old',2000,200,65)],
                                    config=InferenceConfig(max_propagation_mm=1000))
    assert ctx.for_path((0,1))['review_reasons']
    assert 'text_mm' not in ctx.for_path((1,2))


def test_explicit_manual_choice_resolves_export_hold_but_keeps_evidence():
    row = dict(label='P66',dia=65,bore_provenance=dict(version=2,source='nfpc_fallback',
        auto_mm=65,block_export=True,candidate_mm=[25,65],review_reasons=['소속 충돌']))
    with pytest.raises(ValueError,match='P66'):
        require_resolved_diameters([row])
    assert describe_bore(row)['category'] == 'review'
    record_manual(row,65,'원본 확인 완료')
    require_resolved_diameters([row])
    assert row['bore_provenance']['candidate_mm'] == [25,65]
    assert describe_bore(row)['edited']


@pytest.mark.parametrize('writer',['sdf','kfp','merged'])
def test_conflicted_export_writes_nothing(tmp_path,writer):
    from test_module_f_network_editor import design
    from services.cad_import.design.emit import emit_design_sdf, emit_design_kfp
    from routes.module_f.emit import emit_merged
    d = design()
    d['tables'].pipes[0]['bore_provenance'] = dict(block_export=True)
    with pytest.raises(ValueError,match='관경 표기 충돌'):
        if writer=='sdf':
            emit_design_sdf(d['tables'],tmp_path/'blocked.sdf')
        elif writer=='kfp':
            emit_design_kfp(d['tables'],d['got'],tmp_path/'blocked.kfp')
        else:
            emit_merged(d['tables'],tmp_path/'blocked')
    assert not list(tmp_path.iterdir())


def test_native_dxf_metadata_respects_mapping_and_transforms(tmp_path):
    import ezdxf
    from src.pipenet_converter.dxf.diameter_annotations import enrich_diameter_annotations
    doc = ezdxf.new()
    ms = doc.modelspace()
    ms.add_text('25',dxfattribs=dict(insert=(100,200),height=50,rotation=90))
    ms.add_mtext('65',dxfattribs=dict(insert=(1000,200),char_height=50,text_direction=(0,1,0)))
    block = doc.blocks.new('DN')
    block.add_text('32',dxfattribs=dict(insert=(0,0),height=50))
    ms.add_blockref('DN',(2000,200),dxfattribs=dict(rotation=30))
    doc.layers.new('HIDDEN'); doc.layers.get('HIDDEN').off()
    ms.add_text('100',dxfattribs=dict(insert=(3000,200),layer='HIDDEN',height=50,rotation=0))
    ms.add_text('200',dxfattribs=dict(insert=(4000,200),height=50))
    path = tmp_path/'mapped.dxf'; doc.saveas(path)
    before = path.read_bytes()
    got = enrich_diameter_annotations(path,[(100,200,25),(1000,200,65),
          (2000,200,32),(3000,200,100),(5000,200,40)])
    assert [t.rotation_deg for t in got] == [90,90,pytest.approx(30),None,None]
    assert len(got) == 5 and all(t.nominal_mm!=200 for t in got)
    assert path.read_bytes() == before


def test_conflicts_are_reported_as_actionable_api_response(tmp_path):
    from flask import Flask
    from routes.module_f import api_design,api_merge,jobs
    from test_module_f_network_editor import design
    sess=jobs._new_session(key='diameter-conflict-test')
    sess['design']=design()
    sess['design']['tables'].pipes[0]['bore_provenance']=dict(block_export=True)
    sess['merged']={'combined':sess['design']['tables']}
    app=Flask(__name__);app.testing=True
    api_design.register(app,UPLOAD_DIR=tmp_path)
    api_merge.register(app,UPLOAD_DIR=tmp_path)
    try:
        for endpoint in ('design','merge'):
            response=app.test_client().post(f'/api/module-f/{endpoint}/emit',json={'sid':sess['id']})
            assert response.status_code==409,response.json
            assert '관경 표기 충돌' in response.json['message']
    finally:
        jobs._SESSIONS.pop(sess['id'],None)

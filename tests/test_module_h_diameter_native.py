"""H native ownership anchors and configured grammar; F stays opt-out."""
import json
from pathlib import Path

import ezdxf
import pytest

from src.pipenet_converter.dxf.diameter_annotations import read_native_texts, enrich_diameter_annotations
from src.pipenet_converter.dxf.diameter_labels import DiameterLabelConfig, extract_mapped_diameter_points
from src.pipenet_converter.graph.diameter_definition import define_diameters
from src.pipenet_converter.graph.diameter_inference import DiameterAnnotation, infer_diameter_annotations
from src.pipenet_converter.graph.flow import build_flow_tree


def grammar():
    return DiameterLabelConfig(**json.loads((Path(__file__).resolve().parents[1]/'configs/h_diameter_labels.json').read_text()))


@pytest.mark.parametrize('raw,value',[('SP 150',150),('H\n100',100),('SP-65A',65),('DN32',32),('ø 40',40),('50mm',50),('25',25),('90A',90)])
def test_configured_grammar_accepts_unambiguous_native_labels(raw,value):
    assert extract_mapped_diameter_points([('mapped',7,10,20,50,raw)],grammar())==[(10.,20.,value)]


@pytest.mark.parametrize('raw',['150 x 100','SP 150 / 65','1500','25 EA','ROOM 32','PUMP 100A','65A 2EA','100㎡','99A'])
def test_mixed_sizes_quantities_and_unmapped_prefixes_are_not_first_number_guesses(raw):
    assert extract_mapped_diameter_points([('mapped',7,10,20,50,raw)],grammar())==[]


def test_nested_attributes_and_original_bounds_are_read_without_source_mutation(tmp_path):
    doc=ezdxf.new();doc.layers.new('MAPPED')
    base=doc.blocks.new('LEVEL0');base.add_text('25A',dxfattribs={'height':100})
    for i in range(1,4):doc.blocks.new(f'LEVEL{i}').add_blockref(f'LEVEL{i-1}',(100,0))
    ref=doc.modelspace().add_blockref('LEVEL3',(1000,2000),dxfattribs={'rotation':90,'layer':'MAPPED'})
    ref.add_attrib('BORE','SP 65',(1100,2000),dxfattribs={'height':100})
    ref.add_attrib('HIDE','150',(1100,2500),dxfattribs={'height':100,'flags':1})
    path=tmp_path/'nested.dxf';doc.saveas(path);before=path.read_bytes()
    old=read_native_texts(path)
    assert not any(t.text=='25A' for t in old)
    rows=read_native_texts(path,extended=True)
    assert {t.text for t in rows}=={'25A','SP 65'}
    t=next(t for t in rows if t.text=='25A')
    assert (t.x,t.y)==pytest.approx((1000,2300))
    assert t.rotation==pytest.approx(90) and t.layer=='MAPPED'
    assert t.native_anchor_mm and t.entity_handle
    assert path.read_bytes()==before


def test_explicit_leader_overrides_text_orientation_and_keeps_original_label_position(tmp_path):
    doc=ezdxf.new();ms=doc.modelspace()
    t=ms.add_text('65A',dxfattribs={'insert':(8000,8000),'height':100,'rotation':0})
    ms.add_leader([(0,2500),(8000,8000)],dxfattribs={'annotation_handle':t.dxf.handle})
    path=tmp_path/'leader.dxf';doc.saveas(path)
    texts=enrich_diameter_annotations(path,[(8000,8000,65)],extended=True)
    assert texts[0].leader_xy_mm==(0,2500)
    points=[(0,0),(0,5000)];edges=[(0,1)];flow=build_flow_tree(points,edges,[[1]],[0])
    ctx=define_diameters(points,edges,flow,texts)
    row=ctx.decisions[0,1]
    assert row['text_mm']==65 and row['text_xy_mm']==[8000,8000]
    assert row['matching_point_mm']==[0,2500] and '지시선' in row['distance_basis']
    assert row['text_entity_handle']==t.dxf.handle
    # Shared F intentionally keeps its previous insertion-point ownership.
    assert not infer_diameter_annotations(points,edges,flow,texts).decisions[0,1].get('text_mm')


def test_conflicting_leader_targets_and_ambiguous_native_source_are_held(tmp_path):
    doc=ezdxf.new();ms=doc.modelspace()
    t=ms.add_text('65A',dxfattribs={'insert':(100,2500),'height':100,'rotation':90})
    for endpoint in [(0,2500),(2000,2500)]:
        ms.add_leader([endpoint,(100,2500)],dxfattribs={'annotation_handle':t.dxf.handle})
    path=tmp_path/'conflict.dxf';doc.saveas(path)
    texts=enrich_diameter_annotations(path,[(100,2500,65)],extended=True)
    assert texts[0].leader_ambiguous
    points=[(0,0),(0,5000)];edges=[(0,1)];flow=build_flow_tree(points,edges,[[1]],[0])
    row=define_diameters(points,edges,flow,texts).decisions[0,1]
    assert row['block_export'] and not row.get('text_mm')
    unmatched=enrich_diameter_annotations(path,[(500,2000,40)],extended=True)
    assert unmatched[0].source_ambiguous
    assert not define_diameters(points,edges,flow,unmatched).decisions[0,1].get('text_mm')


def test_native_center_does_not_overwrite_display_anchor_or_original_rotation():
    points=[(0,0),(0,4000),(0,8000)];edges=[(0,1),(1,2)]
    flow=build_flow_tree(points,edges,[[1],[2]],[0])
    label=DiameterAnnotation('center',100,4000,25,270,200,'mapped','25','H',native_anchor_mm=(100,3600))
    ctx=define_diameters(points,edges,flow,[label],heads={1:'up',2:'up'})
    assert ctx.decisions[0,1]['text_mm']==25
    assert ctx.decisions[0,1]['text_xy_mm']==[100,4000]
    assert ctx.decisions[0,1]['text_rotation_deg']==270
    assert not ctx.decisions[1,2].get('text_mm')


def test_coincident_native_entities_are_not_arbitrarily_chosen(tmp_path):
    doc=ezdxf.new();ms=doc.modelspace()
    for raw in ('65A','SP 65'):
        ms.add_text(raw,dxfattribs={'insert':(100,2000),'height':100,'rotation':90})
    path=tmp_path/'coincident.dxf';doc.saveas(path)
    row=enrich_diameter_annotations(path,[(100,2000,65)],extended=True)[0]
    assert row.source_ambiguous and not row.raw_text


@pytest.mark.parametrize('h_policy',[False,True])
def test_metadata_failure_stops_h_but_retains_f_legacy_fallback(tmp_path,monkeypatch,h_policy):
    from types import SimpleNamespace
    from routes.module_f import bore_context
    path=tmp_path/'source.dxf';ezdxf.new().saveas(path)
    points=[(0,0),(0,4000)];edges={(0,1)}
    board=SimpleNamespace(pts=points,edges=edges,hnodes=[[1]],disk_kinds=['up'],valves=[],network_mode='tree')
    flow=build_flow_tree(points,edges,[[1]],[0])
    sess={'dxf':str(path),'design_settings':{'diameter_policy':'drawing_first_v1' if h_policy else 'legacy'}}
    def fail(*args,**kwargs): raise ValueError('test metadata failure')
    monkeypatch.setattr(bore_context,'enrich_diameter_annotations',fail)
    if h_policy:
        with pytest.raises(ValueError,match='자동 대입을 중단'):
            bore_context.build_bore_context(sess,board,flow,[(100,2000,40)])
    else:
        result=bore_context.build_bore_context(sess,board,flow,[(100,2000,40)])
        assert result.summary['metadata_status']=='read_failed'


def test_source_snapshot_round_trip_keeps_new_native_anchors():
    from dataclasses import asdict
    from src.pipenet_converter.graph.drawing_reference import annotation_from_snapshot
    original=DiameterAnnotation('native',100,200,40,90,100,'mapped','40A','D',
                                native_anchor_mm=(150,230),leader_xy_mm=(0,300))
    wire=json.loads(json.dumps(asdict(original)))
    restored=annotation_from_snapshot(wire)
    assert asdict(restored)==wire
    assert restored.raw_text=='40A' and restored.leader_xy_mm==[0,300]


def test_h_grammar_uses_only_selected_world_texts(monkeypatch):
    from types import SimpleNamespace
    from routes.module_f.common import _boot
    from routes.module_f.api_design import _dia_texts
    _boot()
    world=SimpleNamespace(texts=[('mapped',7,100,200,50,'SP 65'),('mapped',7,300,400,50,'150 x 100')])
    sess={'pick':SimpleNamespace(world=world),'design_settings':{'diameter_policy':'drawing_first_v1'}}
    assert _dia_texts(sess)==[(100.,200.,65)]

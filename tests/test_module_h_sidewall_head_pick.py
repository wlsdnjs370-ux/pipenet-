"""Mirrored sidewall symbols: candidates, manual hits and replay agree."""
from copy import deepcopy

import ezdxf
import pytest

from routes.module_f.common import _boot
from src.pipenet_converter.dxf.symbol_geometry import two_arc_circle

_boot()
from services.cad_import.pipeline.stage1 import read_dxf, explode, DEFAULT_KNOBS
from services.cad_import.pick.board import Board
from services.cad_import.pick.session import PickSession
from routes.module_f.adopt import adopt_heads, select_heads


def make_source(tmp_path, mirrored=True, sx=150., sy=150.):
    """Use the original symbol's dimensions, with generic mapped layer names."""
    doc = ezdxf.new()
    doc.layers.new('TEST-HEAD', dxfattribs={'color': 81})
    block = doc.blocks.new('sidewall-test')
    vertices = [(-1.1697893217206, 0), (1.1697893217206, 0), (0, -1.855721718166023)]
    for a, b in zip(vertices, (vertices[1], vertices[2], vertices[0])):
        block.add_lwpolyline([a, b])
    block.add_lwpolyline([(-.209461244686671, -.6131949975800496, 1),
                          (.209461244686671, -.6131949975800496, 1)],
                         format='xyb', close=True)
    doc.modelspace().add_blockref('sidewall-test', (-169864.227508406 if mirrored else 169864.227508406,
                                                 2992796.888741465), dxfattribs={
        'layer':'TEST-HEAD', 'xscale':sx, 'yscale':sy, 'rotation':90,
        'extrusion':(0,0,-1 if mirrored else 1)})
    path = tmp_path/'symbol.dxf'; doc.saveas(path)
    return path


def board(path):
    w, _ = explode(*read_dxf(path))
    b = Board(w, dict(DEFAULT_KNOBS)); b.mat_done = True
    return b


@pytest.mark.parametrize('mirrored', [False, True])
def test_candidates_align_with_display_and_adopt_without_lowering_threshold(tmp_path, mirrored):
    import remote30_prototype as A
    path = make_source(tmp_path, mirrored)
    b = board(path)
    bundle = A.parse_dxf_bundle(path)
    insert = next(e for e in bundle.entities if e['t']=='I')
    assert insert['p'] == pytest.approx((169864.227508406, 2992796.888741465))
    heads = A.detect_heads(bundle.entities, {'TEST-HEAD':'HEAD'})
    assert len(heads) == 1 and heads[0].confidence >= .75
    assert heads[0].pos == pytest.approx(b.w.circles[0][2:4])
    ps = PickSession(b.w, 'test', b.kn, b); ps.armed=True; ps.mode='헤드'
    cands = [dict(x=h.pos[0], y=h.pos[1], conf=h.confidence) for h in heads]
    result = adopt_heads(ps, select_heads(cands, conf_min=.75))
    assert result['applied']==1 and not result['skipped']
    assert len(b.highlight_geom()['head_circles'])==1
    assert adopt_heads(ps, select_heads(cands, conf_min=.75))['already']==1
    assert len(b.highlight_geom()['head_circles'])==1


def test_rim_triangle_interior_and_disk_are_one_selection_and_undo(tmp_path):
    b = board(make_source(tmp_path))
    points = [(169864.227508406,2992796.888741465), # triangle base / insertion
              (169650.,2992796.888741465),        # inside triangle, outside disk
              b.w.circles[0][2:4]]                # filled disk centre
    signature = None
    for x,y in points:
        result=b.apply_click('헤드', x,y,max_d=.5)
        assert result['동작']=='추가'
        assert len(b.highlight_geom()['head_circles'])==1
        assert 'r' in b.heads[0] and 'tri_side' not in b.heads[0]
        if signature is None: signature=deepcopy(b.heads)
        assert b.heads==signature
        b.undo()
        assert b.heads==[]


def test_plain_triangle_and_circle_interior_but_no_rectangle_guess():
    from services.cad_import.pipeline.stage1 import World
    w=World(); verts=[(0.,0.),(200.,0.),(100.,173.20508)]
    w.segs=[('HEAD',1,a,b) for a,b in zip(verts, (verts[1],verts[2],verts[0]))]
    w.circles=[('HEAD',1,1000.,0.,30.)]
    b=Board(w,dict(DEFAULT_KNOBS));b.mat_done=True
    assert b.apply_click('헤드',100.,70.,max_d=1)['동작']=='추가'
    assert 'tri_side' in b.heads[0]
    assert b.apply_click('헤드',1000.,0.,max_d=1)['동작']=='추가'
    assert b.apply_click('헤드',300.,300.,max_d=1) is None


def test_name_only_unknown_block_stays_below_threshold(tmp_path):
    import remote30_prototype as A
    doc=ezdxf.new();doc.layers.new('TEST-HEAD');b=doc.blocks.new('unknown')
    b.add_line((0,0),(100,0));doc.modelspace().add_blockref('unknown',(100,200),dxfattribs={'layer':'TEST-HEAD'})
    path=tmp_path/'unknown.dxf';doc.saveas(path)
    heads=A.detect_heads(A.parse_dxf_bundle(path).entities,{'TEST-HEAD':'HEAD'})
    assert len(heads)==1 and heads[0].confidence==.7


@pytest.mark.parametrize('bulges,closed', [([0,0],True),([1,-1],True),([.5,.5],True),([1,1],False)])
def test_only_true_two_semicircle_polyline_becomes_circle(bulges, closed):
    assert two_arc_circle([(0,0),(20,0)],bulges,closed) is None


def test_nonuniform_insert_is_not_promoted_to_circle(tmp_path):
    import remote30_prototype as A
    bundle=A.parse_dxf_bundle(make_source(tmp_path,sx=100.,sy=150.))
    assert not [e for e in bundle.entities if e['t']=='C']


def test_nested_insert_with_basepoint_rotation_and_extrusion(tmp_path):
    import remote30_prototype as A
    doc=ezdxf.new();doc.layers.new('TEST-HEAD')
    child=doc.blocks.new('child',base_point=(10,20));child.add_circle((10,20),1)
    parent=doc.blocks.new('parent')
    inner=parent.add_blockref('child',(-30,40),dxfattribs={'extrusion':(0,0,-1),'rotation':90,'xscale':2,'yscale':2})
    outer=doc.modelspace().add_blockref('parent',(-1000,2000),dxfattribs={'layer':'TEST-HEAD','extrusion':(0,0,-1),'rotation':90})
    expected=outer.matrix44().transform(inner.matrix44().transform((10,20,0)))
    path=tmp_path/'nested.dxf';doc.saveas(path)
    circle=next(e for e in A.parse_dxf_bundle(path).entities if e['t']=='C')
    assert circle['c']==pytest.approx((expected.x,expected.y))


def test_nested_equal_length_axes_with_shear_are_not_circle(tmp_path):
    import remote30_prototype as A
    doc=ezdxf.new();doc.layers.new('TEST-HEAD')
    body=doc.blocks.new('body');body.add_lwpolyline([(0,0,1),(20,0,1)],format='xyb',close=True)
    parent=doc.blocks.new('parent');parent.add_blockref('body',(0,0),dxfattribs={'rotation':45})
    doc.modelspace().add_blockref('parent',(100,200),dxfattribs={'layer':'TEST-HEAD','xscale':2,'yscale':1})
    path=tmp_path/'shear.dxf';doc.saveas(path)
    assert not [e for e in A.parse_dxf_bundle(path).entities if e['t']=='C']


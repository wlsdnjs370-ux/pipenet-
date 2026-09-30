"""Mirrored CAD arcs must stay beside their pipes; plan labels remain visible."""
from copy import deepcopy
from pathlib import Path

import ezdxf
import pytest

from routes.module_f.common import _boot
from routes.module_f.world import _world_payload
from src.pipenet_converter.dxf.work_region import crop_world

_boot()
from services.cad_import.pipeline.stage1 import read_dxf, explode


def load(path):
    return explode(*read_dxf(path))[0]


@pytest.fixture
def source(tmp_path):
    doc = ezdxf.new()
    doc.layers.new('PIPE', dxfattribs={'color': 80})
    doc.layers.new('DN', dxfattribs={'color': 2})
    block = doc.blocks.new('MIRROR-C')
    block.add_arc((-100, 20), 10, 315, 225,
                  dxfattribs={'extrusion': (0, 0, -1), 'layer': 'PIPE'})
    block.add_circle((-100, 50), 5, dxfattribs={'extrusion': (0, 0, -1)})
    doc.modelspace().add_blockref('MIRROR-C', (1000, 2000),
                                dxfattribs={'rotation': 90, 'xscale': 2, 'yscale': 2})
    doc.modelspace().add_arc((-100, 20), 10, 315, 225,
                            dxfattribs={'extrusion': (0, 0, -1), 'layer': 'PIPE'})
    doc.modelspace().add_line((90, 20), (110, 20), dxfattribs={'layer': 'PIPE'})
    doc.modelspace().add_text('50', dxfattribs={'insert': (-100, 35), 'height': 5,
                                              'layer': 'DN', 'extrusion': (0, 0, -1)})
    doc.modelspace().add_mtext(r'{\C2;SP 50}', dxfattribs={'insert': (120, 35),
                                                       'char_height': 5, 'layer': 'DN'})
    path = tmp_path/'mirrored.dxf'
    doc.saveas(path)
    return path


def test_negative_normal_arcs_match_ezdxf_wcs(source):
    world = load(source)
    assert world.arcs[0][2:] == pytest.approx((960, 2200, 20))
    assert world.circles[0][2:] == pytest.approx((900, 2200, 10))
    assert world.arcs[1][2:] == pytest.approx((100, 20, 10))
    # Mirrored start/end swap preserves the 270 degree opening, not the minor arc.
    assert world.arc_ang[1] == pytest.approx((315, 270))
    reference = ezdxf.readfile(source).modelspace().query('ARC')[0]
    assert world.arcs[1][2:4] == pytest.approx(tuple(reference.ocs().to_wcs(reference.dxf.center))[:2])
    assert next(t for t in world.texts if t[-1] == '50')[2:4] == (100, 35)


def test_plan_display_labels_native_palette_and_crop_no_mutation(source):
    world = load(source)
    before = deepcopy(world.__dict__)
    old = _world_payload(world)
    assert 'texts' not in old
    got = _world_payload(world, source_display=True)
    assert {t['text'] for t in got['texts']} == {'50', 'SP 50'}
    assert all(t['rotation'] == 0 for t in got['texts'])
    assert next(b for b in got['bundles'] if b['layer'] == 'PIPE')['css'] == '#3fff00'
    assert world.__dict__ == before
    cropped, _ = crop_world(world, {'type': 'polygon', 'points': [[80,0],[115,0],[115,60],[80,60]]})
    result = _world_payload(cropped, source_display=True)
    assert [t['text'] for t in result['texts']] == ['50']
    assert result['counts']['arcs'] == 1


def test_negative_insert_and_bulge_resolve_before_parent_transform(tmp_path):
    doc=ezdxf.new()
    block=doc.blocks.new('C')
    block.add_line((0,0),(20,0))
    doc.modelspace().add_blockref('C',(-100,50),dxfattribs={'extrusion':(0,0,-1)})
    doc.modelspace().add_lwpolyline([(-100,50,1),(-120,50,0)],format='xyb',
                                   dxfattribs={'extrusion':(0,0,-1)})
    path=tmp_path/'insert.dxf';doc.saveas(path)
    world=load(path)
    assert world.segs[0][2:]==((100,50),(80,50))
    assert world.arcs[0][2:]==pytest.approx((110,50,10))
    assert world.arc_ang[0][1]==pytest.approx(180)


def test_wcs_lines_are_not_mirrored_twice_and_tilt_is_explicit():
    world,_=explode({},[{'t':'LINE','10':100,'20':50,'11':120,'21':50,'230':-1}],{})
    assert world.segs[0][2:]==((100,50),(120,50))
    with pytest.raises(ValueError,match='기울어진 OCS'):
        explode({},[{'t':'ARC','10':1,'20':2,'40':3,'50':0,'51':90,'210':1,'230':1}],{})


def test_h_plan_upload_requests_source_display(tmp_path, monkeypatch):
    from flask import Flask
    from routes.module_f import api_open, api_slot, jobs
    seen=[]
    monkeypatch.setattr(api_slot, '_open_job', lambda *a, **k: seen.append(k) or (lambda:None))
    monkeypatch.setattr(api_slot, '_run_job', lambda *a:None)
    app=Flask(__name__)
    api_slot.register(app, _save_upload=lambda *a,**k:tmp_path/'plan.dxf')
    sess=jobs._new_session()
    try:
        for enabled in (False, True):
            res=app.test_client().post('/api/module-f/slot/open',data={
                'sid':sess['id'],'kind':'plan','h_access':'1' if enabled else '0'})
            assert res.status_code==200, res.json
            assert seen[-1]['h_display']==enabled
    finally:
        jobs._SESSIONS.pop(sess['id'],None)

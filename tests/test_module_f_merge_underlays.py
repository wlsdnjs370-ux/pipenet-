"""Slot-specific reference drawings align without changing hydraulic tables."""
from copy import deepcopy

import pytest

from test_module_f_network_editor import session, design
from test_module_f_merge import _riser
from routes.module_f.api_merge import rebuild_merged
from routes.module_f.merge import bake_combined_iso
from routes.module_f.slots import _slot_switch
from routes.module_f.underlays import reference_layers, _upright_reference


def drawings(sess):
    sess['design']['got']['origin_mm']=[1000,1000]
    def world(edges):
        points=[point for edge in edges for point in (edge[:2],edge[2:])]
        xs,ys=zip(*points)
        return dict(bounds=dict(minx=min(xs),miny=min(ys),maxx=max(xs),maxy=max(ys)),
                    bundles=[dict(segs=[v for edge in edges for v in edge],circles=[],arcs=[])])
    t=sess['design']['tables']
    sess['world']=world([[t.nodes[0]['x'],t.nodes[0]['y'],t.nodes[2]['x'],t.nodes[2]['y']]])
    r=_riser()
    # A bent raw path deliberately differs from the uniformly laid out riser.
    r['nodes'][1]['x']=800
    mr=dict(nodes=[dict(label='m1',x=4000,y=9000,elevation=0,io_node='Input'),
                   dict(label='m2',x=7000,y=9000,elevation=0),
                   dict(label='m3',x=7000,y=6000,elevation=0)],
            pipes=[dict(label='m1',**{'in':'m1','out':'m2'},length=3,dia=100),
                   dict(label='m2',**{'in':'m2','out':'m3'},length=3,dia=100)],
            conn_node_label='m3',conn_xy=[7000,6000],
            plan_edges=[[4000,9000,7000,9000],[7000,9000,7000,6000]])
    sess['slots']={'system':dict(key='system-source',riser=r,world=world([[0,0,800,1000],[800,1000,0,3000]])),
                   'machineroom':dict(key='machine-source',machineroom=mr,world=world(mr['plan_edges']))}
    sess['supply_mode']='lsp_gravity'
    rebuild_merged(sess,persist_overrides=False)


def transform(matrix,xy):
    a,b,c,d,e,f=matrix;x,y=xy
    return a*x+c*y+e,b*x+d*y+f


@pytest.mark.parametrize('iso',[False,True])
def test_three_layers_align_at_authoritative_anchors_without_editing_tables(session,iso):
    drawings(session)
    got=session['merged']
    before=deepcopy(got['combined'])
    refs={r['kind']:r for r in reference_layers(session,got,iso=iso,geometry=True)}
    assert all(r['available'] and r['groups'] for r in refs.values())
    assert refs['plan']['bounds']==session['world']['bounds']
    for kind in ('system','machineroom'):
        assert refs[kind]['bounds']==session['slots'][kind]['world']['bounds']
    nodes=bake_combined_iso(got)[0] if iso else got['combined'].nodes
    at={str(n['label']):(n['x'],n['y']) for n in nodes}
    assert transform(refs['plan']['matrix'],(90100,37154.0934468307))==pytest.approx(at['10'],abs=1)
    assert transform(refs['system']['matrix'],(0,3000))==pytest.approx(at['10'],abs=1e-6)
    assert transform(refs['system']['matrix'],(0,0))==pytest.approx(at['1'],abs=1e-6)
    assert transform(refs['machineroom']['matrix'],(7000,6000))==pytest.approx(at['1'],abs=1e-6)
    for label,point in [('m1',(4000,9000)),('m2',(7000,9000))]:
        assert transform(refs['machineroom']['matrix'],point)==pytest.approx(at[label],abs=1.1)
    assert vars(before)==vars(got['combined'])


def test_inactive_plan_and_other_slot_sources_are_read_independently(session):
    drawings(session)
    _slot_switch(session,'machineroom')
    layers=reference_layers(session,session['merged'],iso=True,geometry=True)
    assert [r['available'] for r in layers]==[True,True,True]
    assert [r['source'] for r in layers]==['editor-fixture','system-source','machine-source']
    assert layers[0]['groups'][0]['segs']!=layers[1]['groups'][0]['segs']


def test_preview_metadata_is_small_and_missing_source_is_explicit(session):
    drawings(session)
    layers=reference_layers(session,session['merged'],iso=True)
    assert all('groups' not in r for r in layers)
    session['slots']['system']['world']=None
    layers=reference_layers(session,session['merged'],iso=True,geometry=True)
    assert not layers[1]['available'] and layers[1]['reason']
    assert layers[0]['available'] and layers[2]['available']


@pytest.mark.parametrize('iso',[False,True])
@pytest.mark.parametrize('mode',['hsp_pump','lsp_gravity','lsp_1stage','llsp_2stage'])
def test_offset_system_endpoints_stay_upright_and_match_connection_heights(session,iso,mode):
    drawings(session)
    baseline={r['kind']:r['matrix'] for r in reference_layers(session,session['merged'],iso=iso)}
    source=session['slots']['system']['riser']
    # The old two-point rotation tilts vertical lines when anchor X differs.
    source['nodes'][0]['x']=450
    source['nodes'][-1]['x']=1700
    session['supply_mode']=mode
    rebuild_merged(session,persist_overrides=False)
    got=session['merged']
    before=deepcopy(vars(got['combined']))
    raw_before=deepcopy(source)
    refs={r['kind']:r for r in reference_layers(session,got,iso=iso,geometry=True)}
    matrix=refs['system']['matrix']
    a,b,c,d,_,_=matrix
    assert b==c==0
    assert a==d and a!=0  # no skew, collapse or aspect-ratio distortion
    nodes=bake_combined_iso(got)[0] if iso else got['combined'].nodes
    at={str(n['label']):(n['x'],n['y']) for n in nodes}
    assert transform(matrix,(1700,3000))==pytest.approx(at['10'])
    assert transform(matrix,(450,0))[1]==pytest.approx(at['1'][1])
    # Preserve the source's horizontal offset instead of pretending both
    # schematic endpoints can coincide with one vertical line.
    assert transform(matrix,(450,0))[0]-at['1'][0]==pytest.approx(a*(450-1700))
    if mode=='lsp_gravity':
        assert refs['plan']['matrix']==baseline['plan']
        assert refs['machineroom']['matrix']==baseline['machineroom']
    assert vars(got['combined'])==before
    assert source==raw_before


def test_upright_reference_degenerate_and_level_anchors():
    assert _upright_reference((1,2),(1,2),(0,0),(0,10)) is None
    assert _upright_reference((1,2),(4,5),(0,0),(0,0)) is None
    matrix=_upright_reference((1,2),(11,2),(100,200),(100,220))
    assert matrix==[2,0,0,2,98,196]
    assert transform(matrix,(1,2))==(100,200)

"""Overlaid bowl/brackets must not turn a centre tick into a fourth pipe."""
from copy import deepcopy
import math

import pytest

from test_module_f_loop_fittings_repair import arc, convert
from services.cad_import.design.preserved import review_fittings
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.graph.network import FlowNetwork
from src.pipenet_converter.graph.review_conversion import build_review_network


def site(rotation=0):
    """Two independent adjacent symbols; only the right one has a centre tick."""
    # Main 0–1–2; branch 1–3 feeds the closed ring 3–4–5–6–3.
    # 7/8 are dead tick ends; 9 is its original, coincident raw centre.
    points = [(-1000,0), (0,0), (1000,0), (0,1000), (1000,1000),
              (1000,2000), (0,2000), (0,-68.57), (0,68.57), (0,0),
              (-281,0), (-281,1000), (-281,-1000), (1000,1500)]
    edges = {(0,10),(1,10),(1,2),(1,3),(3,4),(4,13),(5,13),(5,6),(3,6),
             (1,7),(1,8),(10,11),(10,12)}
    drawn = {(7,9),(8,9)}
    arcs = [arc(0,0), arc(.04,.01,135,90), arc(.04,.01,315,90),
            arc(-281,0,135,90), arc(-281,0,315,90)]
    angle = math.radians(rotation)
    def rot(x,y):
        return (x*math.cos(angle)-y*math.sin(angle), x*math.sin(angle)+y*math.cos(angle))
    points = [rot(*p) for p in points]
    for a in arcs:
        a['cx'], a['cy'] = rot(a['cx'],a['cy'])
        a['sa'] = (a['sa']+rotation)%360
    return points, edges, drawn, arcs


@pytest.mark.parametrize('mode',['loop','grid'])
@pytest.mark.parametrize('rotation',[0,90,180,270])
@pytest.mark.parametrize('rise',[.5,.8])
def test_centre_tick_does_not_create_a_false_closed_port(mode,rotation,rise):
    points, edges, drawn, arcs = site(rotation)
    before = deepcopy((points,edges,drawn,arcs))
    g = convert(points,edges,head=13,mode=mode,arcs=arcs,drawn_edges=drawn,branch_rise_m=rise)
    assert g.arc_junctions['N1']=={'kind':'E','top':'V_N1','top_ports':2}
    assert not g.arc_report['review']
    assert g.nodes['N1'].z_m==0 and g.nodes['V_N1'].z_m==rise
    assert all(g.nodes[f'N{n}'].z_m==rise for n in (3,4,5,6))
    assert len(g.arc_report['symbol_strokes'])==1
    evidence=g.arc_report['symbol_strokes'][0]
    assert evidence['ignored_port_edges']==[[1,7],[1,8]]
    assert evidence['source_preserved']
    assert g.cycle_rank==1
    assert (points,edges,drawn,arcs)==before
    net={'nodes_meta_runtime':{n:dict(coords=[r.x_mm/1000,r.y_mm/1000,r.z_m])
                              for n,r in g.nodes.items()},
         'pipe_data':{pid:dict(start=p.start,end=p.end) for pid,p in g.pipes.items()},
         'arc_report':g.arc_report}
    fits=review_fittings(net,{p:(65,'review_default') for p in g.pipes},g.arc_ports)
    instances=[f for row in fits['per_pipe'].values() for f in row['instances']]
    assert [f['type'] for f in instances if f['node']=='N1']==['tee']
    assert [f['type'] for f in instances if f['node']=='V_N1']==['elbow']
    assert not fits['unresolved_kind_items']
    g.validate()


@pytest.mark.parametrize('problem',['no_original','no_overlay','protected_end','protected_site',
                                  'asymmetric','real_branch','different_centres','extra_raw_branch'])
def test_uncertain_or_user_protected_ports_are_not_silently_discarded(problem):
    points,edges,drawn,arcs=site()
    protected=()
    if problem=='no_original': drawn=set()
    if problem=='no_overlay': arcs=arcs[:1]
    if problem=='protected_end': protected=(7,)
    if problem=='protected_site': protected=(1,)
    if problem=='asymmetric': points[7]=(0,-55)
    if problem=='real_branch':
        points.append((0,-1000));edges.add((7,len(points)-1))
    if problem=='different_centres':
        arcs[1]['cx']=20;arcs[2]['cx']=20
    if problem=='extra_raw_branch': drawn.add((9,2))
    g=convert(points,edges,head=13,arcs=arcs,drawn_edges=drawn,protected_nodes=protected)
    assert 'N1' not in g.arc_junctions
    assert any(r['node']=='N1' for r in g.arc_report['review'])
    assert not g.arc_report['symbol_strokes']


def test_height_conflict_still_preserves_the_source_without_a_guessed_rise():
    points,edges,drawn,arcs=site()
    edges.add((2,4))  # Real return bypasses the claimed vertical transition.
    g=convert(points,edges,head=13,arcs=arcs,drawn_edges=drawn)
    assert 'N1' not in g.arc_junctions
    assert any(r['node']=='N1' and '높이 미확정' in r['reason'] for r in g.arc_report['review'])
    assert not g.arc_report['symbol_strokes']
    assert g.cycle_rank==2


@pytest.mark.parametrize('head', [7,11])
def test_a_selected_terminal_is_protected_and_adjacent_symbols_remain_independent(head):
    points,edges,drawn,arcs=site()
    reference=build_flow_tree(points,edges,[{13},{head}],[0])
    g=build_review_network(FlowNetwork.build(reference,'loop'),points,[0,1],
        {0:'상향식',1:'상향식'},datum_m=0,upright_m=.5,pendant_rise_m=.3,
        pendant_drop_m=.5,combo_rise_m=.3,combo_up_m=.2,combo_drop_m=.5,
        arcs=arcs,drawn_edges=drawn)
    assert sum(n.role=='head' for n in g.nodes.values())==2
    if head==7:
        assert 'N1' not in g.arc_junctions
        assert 'N7' in g.nodes
        assert not g.arc_report['symbol_strokes']
    else:
        assert g.arc_junctions['N10']['kind']=='T'
        assert g.arc_junctions['N1']['kind']=='E'
        assert [r['node'] for r in g.arc_report['symbol_strokes']]==['N1']

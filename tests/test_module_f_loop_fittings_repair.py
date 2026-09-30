"""Original connection evidence, cyclic elevations and stable fitting anchors."""
from copy import deepcopy

import networkx as nx
import pytest

from routes.module_f.common import _boot
_boot()
from services.cad_import.design.preserved import review_fittings
from services.cad_import.design.tables import build_design_tables, bfs_order
from routes.module_f.fitting_inspection import build_inspection
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.graph.network import FlowNetwork
from src.pipenet_converter.graph.review_conversion import build_review_network
from src.pipenet_converter.graph.junction_cleanup import duplicate_symbol_joins
from test_module_f_flow_integration import flow_client


def convert(points, edges, head=4, arcs=(), mode='loop', **kw):
    ref=build_flow_tree(points,edges,[{head}],[0])
    net=FlowNetwork.build(ref,mode)
    return build_review_network(net,points,[0],{0:'상향식'},datum_m=0,upright_m=.5,
        pendant_rise_m=.3,pendant_drop_m=.5,combo_rise_m=.3,combo_up_m=.2,
        combo_drop_m=.5,arcs=arcs,**kw)


def arc(x,y,start=135,sweep=270):
    return dict(cx=x,cy=y,r=180,sa=start,sweep=sweep)


@pytest.mark.parametrize('mode',['loop','grid'])
def test_two_supply_returns_raise_once_without_cutting_cycle(mode):
    points=[(-1000,0),(0,0),(2000,0),(0,1000),(1000,1000),(2000,1000),(3000,0)]
    edges={(0,1),(1,2),(1,3),(3,4),(4,5),(2,5),(2,6)}
    g=convert(points,edges,arcs=[arc(0,0),arc(2000,0)],mode=mode)
    assert g.cycle_rank==1
    assert len(g.arc_junctions)==2
    assert {g.nodes[f'N{i}'].z_m for i in (3,4,5)}=={.5}
    assert next(n for n in g.nodes.values() if n.role=='head').z_m==1
    assert sum(p.length_m for p in g.pipes.values())==pytest.approx(8.5)
    assert all(p.length_m==.5 for p in g.pipes.values() if p.id.startswith('V'))
    g.validate()


def test_loop_height_conflict_preserves_source_with_explicit_review():
    points=[(-1000,0),(0,0),(2000,0),(0,1000),(1000,1000),(2000,1000),(3000,0)]
    edges={(0,1),(1,2),(1,3),(3,4),(4,5),(2,5),(2,6)}
    g=convert(points,edges,arcs=[arc(0,0)])
    assert g.cycle_rank==1 and not g.arc_junctions
    assert any('높이 미확정' in r['reason'] for r in g.arc_report['review'])
    assert all(n.z_m==0 for n in g.nodes.values() if n.role!='head')


def test_no_arc_no_change_and_tiny_real_loop_is_preserved():
    points=[(0,0),(20,0),(20,20),(0,20)]
    edges={(0,1),(1,2),(2,3),(0,3)}
    assert not duplicate_symbol_joins(points,edges,edges,[])
    got=convert(points,edges,head=2)
    assert got.cycle_rank==1 and len(got.pipes)==5


def test_source_proven_duplicate_contacts_only():
    # Millimetre-for-millimetre local pattern observed around former node 69.
    points=[(-20,0),(0.043,0),(0,0),(0,68.56),(0,-68.58),(-400,0),(280,0)]
    drawn={(0,1),(2,4),(0,5)}
    edges=drawn|{(1,2),(2,3),(0,3),(0,4),(1,6)}
    before=deepcopy(edges)
    out=duplicate_symbol_joins(points,edges,drawn,[arc(0,0)])
    assert set(out)=={(0,3),(0,4)}
    assert edges==before
    graph=nx.Graph(edges-set(out))
    assert nx.is_connected(graph) and graph.number_of_edges()-graph.number_of_nodes()+1==0
    assert not duplicate_symbol_joins(points,edges,drawn,[arc(0,0)],protected_edges={(0,3),(0,4)})
    assert not duplicate_symbol_joins(points,edges,edges,[arc(0,0)])
    assert not duplicate_symbol_joins(points,edges,drawn,[])


@pytest.mark.parametrize('reverse',[False,True])
def test_fitting_stays_at_physical_node_after_pipe_reordering(reverse):
    nodes={n:dict(coords=list(xy)+[0],type_id='pump' if n=='R' else 'base',elevation_m=0)
           for n,xy in [('R',(-1,0)),('A',(0,0)),('B',(1,0)),('C',(1,1))]}
    pipes={pid:dict(start=b if reverse else a,end=a if reverse else b,length_m=1)
           for pid,a,b in [('RA','R','A'),('BC','B','C'),('AB','A','B')]}
    net=dict(nodes_meta_runtime=nodes,pipe_data=pipes)
    bores={pid:(65,'review_default') for pid in pipes}
    fits=review_fittings(net,bores)
    assert [i for r in fits['per_pipe'].values() for i in r['instances']][0]['node']=='B'
    tbl=build_design_tables(net,{}, {}, [], bores=bores,fittings=fits)
    order,*_=bfs_order(net,'R')
    keys={'nid':{str(i):n for i,n in enumerate(order,1)}}
    got=build_inspection(tbl,{'kfp':net},keys,tbl.nodes)['fittings']
    assert len(got)==1 and keys['nid'][got[0]['node']]=='B'
    assert got[0]['flow_label']=='손실 귀속 배관 · 유향 미확정'
    assert got[0]['eq_m']==1.8


def test_opposite_ends_of_same_pipe_keep_two_fittings():
    nodes={n:dict(coords=list(xy)+[0],type_id='pump' if n=='R' else 'base')
           for n,xy in [('R',(0,0)),('A',(1,0)),('B',(1,1)),('C',(2,1))]}
    pipes={pid:dict(start=a,end=b,length_m=1) for pid,a,b in [('Z1','R','A'),('A1','A','B'),('Z2','B','C')]}
    net=dict(nodes_meta_runtime=nodes,pipe_data=pipes)
    fits=review_fittings(net,{p:(65,'x') for p in pipes})
    assert [i['node'] for i in fits['per_pipe']['A1']['instances']]==['A','B']
    assert fits['per_pipe']['A1']['equivalent_length']==pytest.approx(3.6)


def test_straight_pass_has_no_tee_loss_even_with_two_inactive_sides():
    net={'nodes_meta_runtime':{n:{'coords':[x,0,0]} for n,x in [('A',0),('B',1),('C',2)]},
         'pipe_data':{'a':{'start':'A','end':'B'},'b':{'start':'B','end':'C'}}}
    out=review_fittings(net,{'a':(65,'x'),'b':(65,'x')},{'B':[[0,1,0],[0,-1,0]]})
    assert out['counts']=={} and not out['unresolved_kind_items']


def test_jog_pair_on_loop_has_four_elbows_and_keeps_return():
    points=[(0,0),(1000,0),(2000,0),(3000,0),(4000,0),(4000,-1000),(0,-1000)]
    edges={(0,1),(1,2),(2,3),(3,4),(4,5),(5,6),(0,6)}
    g=convert(points,edges,arcs=[arc(1000,0,90,180),arc(3000,0,270,180)])
    assert g.cycle_rank==1 and len(g.arc_report['jogs'])==1
    assert g.nodes['N2'].z_m==.5 and g.nodes['N5'].z_m==0
    assert len([p for p in g.pipes if p.startswith('V')])==2


@pytest.mark.parametrize('mode', ['loop', 'grid'])
@pytest.mark.parametrize('rise_m', [.5, .8])
def test_single_elbow_symbol_raises_entire_loop_without_a_jog_partner(mode, rise_m):
    """One feed arc below a loop is a height transition, not a flat elbow."""
    points = [(-1000,0), (0,0), (0,1000), (1000,1000), (1000,2000), (0,2000)]
    edges = {(0,1), (1,2), (2,3), (3,4), (4,5), (2,5)}
    before = deepcopy((points, edges))
    got = convert(points, edges, arcs=[arc(0,0)], mode=mode, branch_rise_m=rise_m)
    assert got.cycle_rank == 1
    assert got.arc_report['junctions'] == [dict(node='N1', top='V_N1', kind='elbow_rise')]
    assert not got.arc_report['review'] and not got.arc_report['jogs']
    low, high = got.nodes['N1'], got.nodes['V_N1']
    assert (low.x_mm, low.y_mm) == (high.x_mm, high.y_mm) == (0,0)
    assert low.z_m == 0 and high.z_m == rise_m
    assert {got.nodes[f'N{i}'].z_m for i in (2,3,4,5)} == {rise_m}
    vertical = got.pipes['V1']
    assert (vertical.start, vertical.end, vertical.length_m) == ('N1', 'V_N1', rise_m)
    assert sum(p.length_m for p in got.pipes.values()) == pytest.approx(6.5 + rise_m)
    assert got.arc_ports['N1'] == [[-1.,0.,0.], [0.,0.,1.]]
    assert got.arc_ports['V_N1'] == [[0.,1.,0.], [0.,0.,-1.]]
    assert (points, edges) == before
    got.validate()


def test_single_arc_with_same_level_return_does_not_force_a_false_rise():
    """A closed horizontal bypass contradicts a single rise: preserve/review."""
    points = [(-1000,0), (0,0), (0,1000), (-1000,1000)]
    edges = {(0,1), (1,2), (2,3), (0,3)}
    got = convert(points, edges, head=2, arcs=[arc(0,0)])
    assert got.cycle_rank == 1 and not got.arc_junctions
    assert not got.arc_ports
    assert got.arc_report['review'][0]['node'] == 'N1'
    assert '높이 미확정' in got.arc_report['review'][0]['reason']
    assert {n.z_m for n in got.nodes.values() if n.role != 'head'} == {0}
    assert {p.source_edge for p in got.pipes.values() if p.source_edge} == edges
    got.validate()


def test_single_arc_elevation_is_port_based_not_supply_traversal_based():
    points = [(0,1000), (0,0), (-1000,0), (-1000,-1000)]
    got = convert(points, {(0,1),(1,2),(2,3)}, head=3, arcs=[arc(0,0)])
    assert got.nodes['V_N1'].z_m == 0  # open side is the higher side
    assert got.nodes['N1'].z_m == -.5
    assert got.nodes['N2'].z_m == -.5
    assert got.pipes['V1'].length_m == .5
    got.validate()


def test_real_selected_b1f_roundtrip(flow_client,tmp_path):
    """Opt-in local source audit matching the live 13-head screenshot selection."""
    import os,json
    from pathlib import Path
    if os.environ.get('MODULE_F_FITTING_AUDIT')!='1':
        pytest.skip('실제 도면 대조는 MODULE_F_FITTING_AUDIT=1 로 실행')
    from scripts.audit_loop_fittings import live_board
    from services.cad_import.design.flow import network_for_board,flow_for_board
    from services.cad_import.edit.session import EditSession
    from routes.module_f.preserved_design import build_review_design
    from routes.module_f.api_design import _settings,emit_design_files
    from services.cad_import.design.emit import _pc_models
    board,heads=live_board()
    assert len(heads)==13
    original=deepcopy((board.pts,board.edges))
    tree=flow_for_board(board)
    tree_edges=set(tree.edges)
    from unittest.mock import patch
    from services.cad_import.design.preserved import expand_preserved
    with patch('src.pipenet_converter.graph.review_elevation.overlapping_arc_stroke',return_value=None):
        before=expand_preserved(board,heads)
    repaired=expand_preserved(board,heads)
    assert len(repaired['kfp']['nodes_meta_runtime'])==len(before['kfp']['nodes_meta_runtime'])+1
    assert len(repaired['kfp']['pipe_data'])==len(before['kfp']['pipe_data'])+1
    assert sum(p['length_m'] for p in repaired['kfp']['pipe_data'].values())-sum(
        p['length_m'] for p in before['kfp']['pipe_data'].values())==pytest.approx(.5)
    assert len(before['kfp']['arc_report']['review'])==1
    assert not repaired['kfp']['arc_report']['review']
    c,sess=flow_client
    sess['edit']=EditSession(board,key=board.key,out_dir=str(tmp_path))
    destination=Path(os.environ.get('MODULE_F_FITTING_OUTPUT',str(tmp_path)))
    reports=[]
    for mode in ('loop','grid'):
        board.network_mode=mode
        network=network_for_board(board)
        sess['worst']={'heads':heads,'edges':network.selected_edges(heads),'source_tag':'Z1',
                      'flow_revision':network.revision,'network_mode':mode,'loads':{},'physical_loads':{}}
        cfg=_settings(sess,{'review_default_mm':65,'iso':True})
        result=build_review_design(sess,sess['edit'],cfg,'Z1')
        assert result['ok'],result
        d=sess['design'];tbl=d['tables'];got=d['got']
        assert got['cycle_rank']==4
        assert len(got['kfp']['arc_report']['junctions'])==11
        # Former visible node 82 is a standalone two-port elevation symbol.
        # Audit by source identity; display labels change after new nodes.
        nodes = got['kfp']['nodes_meta_runtime']
        # Cache regeneration changes raw node IDs; match physical CAD sites.
        import math
        def at(x,y):
            source=min(got['node_ref'].values(),key=lambda i:math.dist(board.pts[i][:2],(x,y)))
            assert math.dist(board.pts[source][:2],(x,y))<.1
            return f'N{source}'
        single_sites = {at(753913.618,-144993.200),at(756713.618,-144993.200),
                        at(755663.618,-145666.256)}
        assert {r['node'] for r in got['kfp']['arc_report']['junctions']
                if r['kind']=='elbow_rise'} == single_sites
        assert not got['kfp']['arc_report']['review']
        overlapped=at(758744.743,-153051.534)
        assert got['kfp']['arc_junctions'][overlapped]['kind']=='E'
        assert [r['node'] for r in got['kfp']['arc_report']['symbol_strokes']]==[overlapped]
        labels = {nid:lab for lab,nid in d['keys']['nid'].items()}
        vertical_labels = []
        top=f'V_{overlapped}'
        assert nodes[overlapped]['coords'][:2]==nodes[top]['coords'][:2]
        assert nodes[top]['elevation_m']-nodes[overlapped]['elevation_m']==pytest.approx(.5)
        for end,kind in ((overlapped,'tee'),(top,'elbow')):
            assert [f['type'] for f in tbl.fittings if f['node']==labels[end]]==[kind]
        vpid=next(pid for pid,p in got['kfp']['pipe_data'].items()
                  if {p['start'],p['end']}=={overlapped,top})
        vertical_labels.append(tbl.pipe_labels[vpid])
        for nid in sorted(single_sites):
            top = f'V_{nid}'
            assert nodes[nid]['coords'][:2] == nodes[top]['coords'][:2]
            assert nodes[top]['elevation_m']-nodes[nid]['elevation_m'] == pytest.approx(.5)
            pid = next(pid for pid,p in got['kfp']['pipe_data'].items()
                       if {p['start'],p['end']} == {nid,top})
            vertical_labels.append(tbl.pipe_labels[pid])
            assert got['kfp']['pipe_data'][pid]['length_m'] == .5
            for end in (nid,top):
                fittings = [f for f in tbl.fittings if f['node']==labels[end]]
                assert len(fittings)==1 and fittings[0]['type']=='elbow'
        assert all('node' in f for f in tbl.fittings)
        # Every successful classification is anchored to the two physical ends
        # of its loss-owning pipe, even when table BFS reverses that pipe.
        pmap={p['label']:p for p in tbl.pipes}
        assert all(f['node'] in (pmap[f['pipe']]['in'],pmap[f['pipe']]['out']) for f in tbl.fittings)
        preview=c.get('/api/module-f/design/preview',query_string={'sid':sess['id']})
        assert preview.status_code==200,preview.json
        cards=preview.json['view']['inspection']['fittings']
        realcards=[r for r in cards if r['kind'] in ('elbow','tee','elbow-45')]
        assert all(r['node'] for r in realcards)
        assert any('입체 접속' in r['geometry_source'] for r in realcards)
        shown = {n['label']:n for n in preview.json['view']['nodes']}
        assert shown[labels[overlapped]]['x']==pytest.approx(shown[labels[f'V_{overlapped}']]['x'])
        assert shown[labels[overlapped]]['y']!=pytest.approx(shown[labels[f'V_{overlapped}']]['y'])
        for end,kind in ((overlapped,'tee'),(f'V_{overlapped}','elbow')):
            card=next(r for r in realcards if r['node']==labels[end])
            assert card['kind']==kind
            assert len(card['arms'])==(3 if kind=='tee' else 2)
        for nid in single_sites:
            top = f'V_{nid}'
            assert shown[labels[nid]]['x'] == pytest.approx(shown[labels[top]]['x'])
            assert shown[labels[nid]]['y'] != pytest.approx(shown[labels[top]]['y'])
            for end in (nid,top):
                card = next(r for r in realcards if r['node']==labels[end])
                assert card['kind']=='elbow' and card['eq_m']==1.8
                assert '입체 접속' in card['geometry_source']
        path,error=emit_design_files(sess,destination/mode)
        assert not error,error
        _pc_models()
        from pipenet_converter.sdf_parser import parse_sdf
        restored=parse_sdf(path)
        g=nx.MultiGraph((p.from_node,p.to_node) for p in restored.pipes.values())
        assert nx.is_connected(g) and g.number_of_edges()-g.number_of_nodes()+1==4
        assert len(restored.nozzles)==13
        assert sum(p.length_m for p in restored.pipes.values())==pytest.approx(sum(p['length'] for p in tbl.pipes))
        assert all(abs(p.rise_m)<=p.length_m+1e-6 for p in restored.pipes.values())
        for label in vertical_labels:
            pipe = restored.pipes[label]
            assert pipe.length_m == .5 and abs(pipe.rise_m) == .5
        assert all('flow_direction' in f for f in tbl.fittings)
        reports.append({'mode':mode,'nodes':len(tbl.nodes),'pipes':len(tbl.pipes),'heads':len(tbl.nozzles),
            'cycle_rank':4,'arc_report':got['kfp']['arc_report'],'unresolved':tbl.unresolved,
            'length_m':sum(p['length'] for p in tbl.pipes),'sdf':str(path),'hydraulics_solved':False,
            'overlaid_symbol':{'source_node':overlapped,'lower_label':labels[overlapped],
                'upper_label':labels[f'V_{overlapped}'],'rise_m':.5,
                'fittings':['tee','elbow'],'raw_added_nodes':1,'raw_added_pipes':1,
                'added_length_m':.5},
            'standalone_rises':[{'source_node':nid,'lower_label':labels[nid],
                'upper_label':labels[f'V_{nid}'], 'rise_m':.5,
                'lower_elevation_m':nodes[nid]['elevation_m'],
                'upper_elevation_m':nodes[f'V_{nid}']['elevation_m']} for nid in sorted(single_sites)]})
    assert flow_for_board(board).edges==tree_edges
    assert (board.pts,board.edges)==original
    # Full-network output must not be blocked by one conflicting arc symbol.
    from routes.module_f.preserved_design import emit_review_outputs
    all_out=emit_review_outputs(sess,{'full_kfp':True,'worst_kfp':False,'worst_sdf':False},destination/'all')
    assert all_out['ok'],all_out
    full=sess['kfp']
    full_graph=nx.MultiGraph((p['start'],p['end']) for p in full['pipe_data'].values())
    assert nx.is_connected(full_graph)
    assert full_graph.number_of_edges()-full_graph.number_of_nodes()+1==network.report()['cycle_rank']
    assert any('높이 미확정' in r['reason'] for r in full.get('arc_report',{}).get('review',()))
    reports.append({'mode':'full_kfp','cycle_rank':network.report()['cycle_rank'],
                    'nodes':len(full['nodes_meta_runtime']),'pipes':len(full['pipe_data']),
                    'output':sess['kfp_path'],'hydraulics_solved':False})
    destination.mkdir(exist_ok=True,parents=True)
    (destination/'verification.json').write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')

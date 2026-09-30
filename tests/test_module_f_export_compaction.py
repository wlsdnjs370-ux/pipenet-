"""Serial export reduction preserves physical topology, loss and source data."""
from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

import networkx as nx
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]/'core'))
from core.remote30_full_network import CombinedTables
from src.pipenet_converter.graph.export_compaction import compact_export


def chain():
    return CombinedTables(
        nodes=[dict(label=str(i),x=i*1000,y=0,elevation=0,io_node='Input' if i==0 else 'No')
               for i in range(6)],
        pipes=[dict(label=f'P{i}',**{'in':str(i),'out':str(i+1)},type='KSD 3507',
                    dia=65,inner_mm=67.9,c=120,length=1,elev=0,eq_len=0,status='Normal')
               for i in range(5)],
        nozzles=[dict(label='H',**{'in':'5','out':'@/H'},lib='SP-HEAD',status='1',
                      flow_lmin=80,flow_m3s=80/60000)])


def graph(tables):
    g = nx.MultiGraph()
    g.add_nodes_from(str(n['label']) for n in tables.nodes)
    g.add_edges_from((str(p['in']),str(p['out'])) for p in tables.pipes)
    return g


def test_straight_equal_pipes_only_on_copy_and_reversible_audit():
    tables=chain();before=deepcopy(tables)
    compact,audit=compact_export(tables)
    assert tables==before
    assert [n['label'] for n in compact.nodes]==['0','4','5']
    assert [(p['label'],p['in'],p['out'],p['length']) for p in compact.pipes]==[
        ('P0','0','4',4),('P4','4','5',1)]
    assert audit.nodes_before==6 and audit.nodes_after==3
    assert [p['label'] for p in audit.source_pipes['P0']]==['P0','P1','P2','P3']
    assert compact_export(compact)[0]==compact
    assert sum(p['length'] for p in compact.pipes)==5
    assert compact.nozzles==tables.nozzles


@pytest.mark.parametrize('field,value',[
    ('dia',80),('inner_mm',68),('type','KSD 3562'),('c',130),('status','Closed'),
    ('roughness_mm',.2),('unknown_physical_control',1),('role','branch')])
def test_diameter_material_status_c_and_other_physics_boundaries_survive(field,value):
    tables=chain();tables.pipes[1][field]=value
    compact,_=compact_export(tables)
    assert {'1','2'} <= {n['label'] for n in compact.nodes}


@pytest.mark.parametrize('feature',['head','branch','valve','pressure','review','bend','vertical','reversed','display_corner'])
def test_meaningful_nodes_survive(feature):
    tables=chain();views=[]
    if feature=='head':
        tables.nozzles.append(dict(label='extra',**{'in':'2','out':'@/extra'}))
    elif feature=='branch':
        tables.nodes.append(dict(label='b',x=2000,y=1000,elevation=0))
        tables.pipes.append(dict(tables.pipes[0],label='branch',**{'in':'2','out':'b'}))
    elif feature=='valve':
        tables.valves.append(dict(label='v',**{'in':'2','out':'3'}))
    elif feature=='pressure':
        tables.nodes[2]['pressure_pa']=101325
    elif feature=='review':
        tables.unresolved={'kind_items':[dict(pipe_label='P2',node_label='2',n=1)]}
    elif feature=='bend':
        tables.nodes[2]['y']=1000
    elif feature=='vertical':
        tables.nodes[2]['elevation']=.5
        tables.pipes[1]['elev']=.5;tables.pipes[2]['elev']=-.5
    elif feature=='reversed':
        tables.pipes[2].update({'in':'3','out':'2'})
    else:
        view=deepcopy(tables.nodes);view[2]['y']=500
        views=[view]
    compact,_=compact_export(tables,display_views=views)
    assert '2' in {n['label'] for n in compact.nodes}


def test_endpoint_loss_moves_host_but_not_quantity_or_location():
    tables=chain()
    tables.fittings=[dict(pipe='P1',node='1',type='tee_branch',count=2)]
    tables.equipment=[dict(label='EF1',pipe='P3',editor_node='4',editor_library='ELBOW_90_STD',
                           eq_len=1.8,rel_pos=1,**{'in':'3','out':'4'})]
    tables.pipes[1]['eq_len']=7.4
    compact,audit=compact_export(tables)
    assert {n['label'] for n in compact.nodes}=={'0','1','4','5'}
    assert compact.fittings==[dict(pipe='P1',node='1',type='tee_branch',count=2)]
    eq=compact.equipment[0]
    assert eq['pipe']=='P1' and eq['editor_node']=='4' and eq['rel_pos']==1 and eq['eq_len']==1.8
    assert eq['in']=='1' and eq['out']=='4'
    assert sum(p['eq_len'] for p in compact.pipes)==7.4
    assert len(audit.source_pipes['P1'])==3


def test_unknown_loss_location_blocks_the_whole_host_span():
    tables=chain();tables.equipment=[dict(pipe='P1',eq_len=5,rel_pos=.5)]
    compact,_=compact_export(tables)
    assert {'1','2'} <= {n['label'] for n in compact.nodes}
    assert compact.equipment==tables.equipment


@pytest.mark.parametrize('width',[2,3])
def test_subdivided_loop_and_grid_keep_cycles_and_hydraulic_solution(width):
    from src.pipenet_converter.hydraulics.models import Network,Node,Pipe,Nozzle,Size
    from src.pipenet_converter.hydraulics.solver import solve
    base=nx.grid_2d_graph(width,width)
    tables=CombinedTables()
    for x,y in base:
        tables.nodes.append(dict(label=f'{x}:{y}',x=x*10000,y=y*10000,elevation=0,
                                 io_node='Input' if (x,y)==(0,0) else 'No'))
    for i,(a,b) in enumerate(base.edges):
        al,bl=f'{a[0]}:{a[1]}',f'{b[0]}:{b[1]}'
        tables.nodes.append(dict(label=f'm{i}',x=(a[0]+b[0])*5000,y=(a[1]+b[1])*5000,elevation=0))
        for j,(start,end) in enumerate(((al,f'm{i}'),(f'm{i}',bl))):
            tables.pipes.append(dict(label=f'P{i}-{j}',**{'in':start,'out':end},length=5,
                                     dia=65,type='KSD 3507',c=120,status='Normal',elev=0))
    tables.nodes.append(dict(label='h',x=(width-1)*10000,y=(width-1)*10000,elevation=.5))
    tables.pipes.append(dict(tables.pipes[0],label='stem',length=.5,elev=.5,
                            **{'in':f'{width-1}:{width-1}','out':'h'}))
    tables.nozzles=[dict(label='H',**{'in':'h','out':'@/H'})]
    compact,audit=compact_export(tables)
    assert audit.nodes_after < audit.nodes_before
    for t in (tables,compact):
        g=graph(t)
        assert nx.is_connected(g) and g.number_of_edges()-g.number_of_nodes()+1==(width-1)**2
    def solve_tables(t):
        net=Network(tuple(Node(n['label'],n['elevation']) for n in t.nodes),
            tuple(Pipe(p['label'],p['in'],p['out'],p['length'],p['c'],(Size(65,67.9),)) for p in t.pipes),
            (Nozzle('H','h',80,80,1,12),),'0:0')
        return solve(net,tuple(0 for p in net.pipes),3)
    before,after=solve_tables(tables),solve_tables(compact)
    assert before.converged and after.converged
    assert after.nozzle_flow_lpm==pytest.approx(before.nozzle_flow_lpm,rel=1e-7)
    assert after.pressure_bar==pytest.approx({k:before.pressure_bar[k] for k in after.pressure_bar},rel=1e-7)


def test_known_conflict_is_not_hidden_and_zero_length_not_collapsed():
    tables=chain();tables.pipes[1]['bore_provenance']={'block_export':True}
    tables.pipes[2]['length']=0
    compact,_=compact_export(tables)
    assert len(compact.nodes)==6


def test_merged_files_share_compact_topology_and_audit(tmp_path):
    from routes.module_f.emit import emit_merged, ISO_SPREAD
    from routes.module_f.merge import bake_combined_iso
    tables=chain();before=deepcopy(tables)
    got=dict(combined=tables,parts={'plan':[n['label'] for n in tables.nodes]},system_layout='physical_xy')
    paths=emit_merged(tables,tmp_path,iso_nodes=bake_combined_iso(got)[0],
                      display_reference_labels=got['parts']['plan'],iso_spread=ISO_SPREAD)
    audit=json.loads(Path(paths['compaction']).read_text(encoding='utf8'))
    assert audit['nodes_after']==3
    for key in ('sdf','sdf_iso'):
        root=ET.parse(paths[key]).getroot()
        assert len(list(root.iter('Pipe')))==2
        assert len(list(root.iter('Nozzle')))==1
        assert sum(float(p.get('length')) for p in root.iter('Pipe'))==5
    assert not any('실패' in w for w in paths['warnings']),paths['warnings']
    assert tables==before


def test_proposal_export_reduces_nodes_and_uses_plan_reference_spread(tmp_path):
    from routes.module_f.hydraulic_sizing import prepare, fingerprint
    from routes.module_f.sizing_export import emit_proposal
    from routes.module_f.merge import bake_combined_iso
    from src.pipenet_converter.hydraulics.sizing import propose
    from src.pipenet_converter.render.export_style import EXPORT_SYMBOL_SPREAD
    tables=chain()
    # Internal diameter is resolved from the shipped library by prepare().
    for p in tables.pipes:
        p.pop('inner_mm')
    got=dict(combined=tables,parts={'plan':[n['label'] for n in tables.nodes]},system_layout='physical_xy')
    net,cfg,_=prepare(tables,{'mode':'pump_only'})
    report=propose(net,cfg)
    report.update(fingerprint=fingerprint(tables),model=asdict(net))
    untouched=deepcopy(report)
    files=emit_proposal(got,report,tmp_path)
    exported=json.loads(Path(files['report']).read_text(encoding='utf8'))
    assert exported['export_compaction']['nodes_after']==3
    assert report==untouched
    root=ET.parse(files['sdf_iso']).getroot()
    assert len(list(root.iter('Pipe')))==2
    positions={n.get('label'):(float(n.find('Position').get('x')),float(n.find('Position').get('y')))
               for n in root.iter('Node')}
    iso={n['label']:n for n in bake_combined_iso(got)[0]}
    assert positions['4'][0]-positions['0'][0]==pytest.approx(
        (iso['4']['x']-iso['0']['x'])*3000/5000*EXPORT_SYMBOL_SPREAD,abs=.01)
    assert fingerprint(tables)==report['fingerprint']


def test_plan_export_compacts_and_spreads_without_changing_session(tmp_path,monkeypatch):
    from routes.module_f import api_design
    from routes.module_f.common import _boot
    from src.pipenet_converter.render.export_style import EXPORT_SYMBOL_SPREAD
    _boot()
    tables=chain();before=deepcopy(tables)
    sess=dict(id='compact-plan',key='plan',design={'tables':tables})
    monkeypatch.setattr(api_design,'_design_stale',lambda s:None)
    path,error=api_design.emit_design_files(sess,tmp_path,dict(api_design._DEFAULT_SETTINGS,iso=True))
    assert error is None,error
    root=ET.parse(path).getroot()
    assert len(list(root.iter('Pipe')))==2
    audit=json.loads(path.with_suffix('.compaction.json').read_text(encoding='utf8'))
    assert audit['nodes_after']==3
    assert tables==before
    positions={n.get('label'):float(n.find('Position').get('x')) for n in root.iter('Node')}
    assert positions['4']-positions['0']==pytest.approx(
        4000*3000/5000*EXPORT_SYMBOL_SPREAD*3**.5/2,abs=.01)

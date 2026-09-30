"""Pure model validation; no Flask server or CAD application is required."""
from types import SimpleNamespace

import pytest

from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.graph.network import FlowNetwork
from src.pipenet_converter.graph.review_conversion import build_review_network
from src.pipenet_converter.graph.preserved_properties import apply_review_properties, sync_review_kfp


@pytest.mark.parametrize('separation_mm',[0,0.009839])
def test_zero_alias_and_tiny_real_pipe_preserve_cycle(separation_mm):
    points=[(0,0),(separation_mm,0),(1000,0),(1000,1000),(0,1000)]
    edges={(0,1),(1,2),(2,3),(3,4),(0,4)}
    reference=build_flow_tree(points,edges,[{2}],[0])
    source=FlowNetwork.build(reference,'loop')
    got=build_review_network(source,points,[0],{0:'상향식'},datum_m=2,
        upright_m=.5,pendant_rise_m=.3,pendant_drop_m=.5,
        combo_rise_m=.3,combo_up_m=.2,combo_drop_m=.5)
    assert got.cycle_rank==1
    assert all(p.length_m>0 for p in got.pipes.values())
    assert sum(p.length_m for p in got.pipes.values())==pytest.approx(
        sum(reference.lengths_mm.values())/1000+.5)
    assert len(got.pipes)==(5 if separation_mm==0 else 6)


def test_property_edits_keep_xy_and_sync_calculation_values():
    rows=SimpleNamespace(nodes=[{'label':'1','elevation':0},{'label':'2','elevation':0}],
        pipes=[{'label':'P1','in':'1','out':'2','length':2,'dia':65,'c':120,'type':'KSD 3507'}],
        pipe_labels={'source':'P1'})
    net={'nodes_meta_runtime':{'a':{'coords':[3,4,0]},'b':{'coords':[4,4,0]}},
         'pipe_data':{'source':{'start':'a','end':'b'}}}
    keys={'node':{'2':(7,8)},'pipe':{'P1':(1,2,3,4)}}
    apply_review_properties(rows,keys,[{'field':'elevation','key':[7,8],'new':1},
        {'field':'dia','key':[1,2,3,4],'new':80}])
    sync_review_kfp(rows,net,{'1':'a','2':'b'})
    assert net['nodes_meta_runtime']['b']['coords']==[4,4,1]
    assert net['pipe_data']['source']['nominal_mm']==80
    assert rows.pipes[0]['elev']==1
    assert rows.pipes[0]['bore_provenance']['source']=='user'
    with pytest.raises(ValueError,match='표고차'):
        apply_review_properties(rows,keys,[{'field':'elevation','key':[7,8],'new':3}])

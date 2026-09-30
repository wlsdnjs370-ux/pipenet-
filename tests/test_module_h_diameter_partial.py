"""Safety contracts for source-backed, partial repeat transfer (no live files)."""
from copy import deepcopy
import math

import pytest

from src.pipenet_converter.graph.diameter_definition import define_diameters, resolve_pipe_definitions
from src.pipenet_converter.graph.diameter_inference import DiameterAnnotation as Text
from src.pipenet_converter.graph.flow import build_flow_tree, edge_key


def branches(count=2):
    points = [(i*6000, 0) for i in range(count+2)]
    edges = [(i, i+1) for i in range(count+1)]
    heads, paths = {}, []
    for j in range(count):
        path, previous = [], j+1
        for i in range(5):
            n = len(points); points.append(((j+1)*6000, (i+1)*4000))
            path.append(edge_key(previous, n)); edges.append(path[-1]); heads[n] = 'up'; previous = n
        paths.append(path)
    return points, edges, heads, paths


def text(branch, slot, diameter):
    return Text(f'T{branch}-{slot}', (branch+1)*6000+100, (slot+.5)*4000,
                diameter, 90, 100, 'mapped-diameter', str(diameter), f'H{branch}-{slot}')


def run(fixture, labels, barriers=()):
    points, edges, heads, _ = fixture
    before = deepcopy((points, edges, heads, labels))
    flow = build_flow_tree(points, edges, [[n] for n in heads], [0])
    result = define_diameters(points, edges, flow, labels, heads=heads, barriers=barriers)
    assert before == (points, edges, heads, labels)
    return result


def test_partial_donor_copies_four_slots_not_zero_and_does_not_fill_missing_slot():
    fixture = branches()
    labels = [text(0, i, value) for i, value in enumerate([25, 25, None, 40, 40]) if value]
    result = run(fixture, labels)
    for path in fixture[3]:
        assert [result.decisions[e].get('text_mm') for e in path] == [25, 25, None, 40, 40]
    assert result.summary['repeated_edges'] == 4
    for i in (0, 1, 3, 4):
        row = result.decisions[fixture[3][1][i]]
        assert row['annotation_id'] == f'T0-{i}'
        assert row['text_entity_handle'] == f'H0-{i}'
        assert row['representative_slot'] == i
        assert row['owner_edge'] == list(fixture[3][0][i])


def test_one_conflicting_slot_does_not_block_other_slots_and_direct_target_wins():
    fixture = branches(3)
    labels = [text(j, i, value) for j in (0, 2) for i, value in enumerate([25,25,32,40,40])]
    labels = [t for t in labels if t.identity != 'T2-2'] + [text(2,2,50), text(1,4,65)]
    result = run(fixture, labels)
    rows = [result.decisions[e] for e in fixture[3][1]]
    assert [r.get('text_mm') for r in rows] == [25,25,None,40,65]
    assert rows[2]['block_export'] and rows[2]['candidate_mm'] == [32,50]
    assert rows[4]['definition_source'] == 'drawing_direct'
    assert rows[4]['annotation_id'] == 'T1-4'


def test_complementary_partial_sources_are_original_not_recursive_copies():
    fixture = branches(3)
    labels = [text(0,0,25),text(0,1,25),text(2,3,40),text(2,4,40)]
    result = run(fixture, labels)
    for path in fixture[3]:
        for i, edge in enumerate(path):
            row = result.decisions[edge]
            assert row.get('text_mm') == [25,25,None,40,40][i]
            if i != 2:
                assert row['annotation_id'] == f'T{0 if i<2 else 2}-{i}'


def test_wrong_head_kind_or_valve_does_not_copy():
    fixture = branches()
    labels = [text(0,0,25)]
    fixture[2][fixture[3][1][-1][1]] = 'down'
    assert not run(fixture,labels).decisions[fixture[3][1][0]].get('text_mm')
    fixture[2][fixture[3][1][-1][1]] = 'up'
    assert not run(fixture,labels,barriers=[fixture[3][1][1][1]]).decisions[fixture[3][1][0]].get('text_mm')


def test_disconnected_similar_branches_are_not_donors():
    fixture = branches()
    fixture[1].remove((1,2))
    result = run(fixture,[text(0,0,25)])
    assert not result.decisions[fixture[3][1][0]].get('text_mm')


def test_equal_length_different_bend_is_not_equivalent():
    fixture = branches()
    points, edges, heads, paths = fixture
    a,b = paths[1][0]
    # Same 4000mm length, but a right-angle dogleg in target's first slot.
    points[b] = (14000,2000)
    mid = len(points); points.append((14000,0))
    edges.remove((a,b)); edges.extend([(a,mid),(mid,b)])
    result = run(fixture,[text(0,0,25)])
    assert not result.decisions[edge_key(a,mid)].get('text_mm')


def test_node_renumbering_and_global_rotation_preserve_sizes_and_sources():
    fixture = branches()
    labels = [text(0,0,25),text(0,1,25),text(0,3,40),text(0,4,40)]
    base = run(fixture, labels)
    points, edges, heads, paths = fixture
    order = list(reversed(range(len(points)))); ids = {old:new for new,old in enumerate(order)}
    angle = math.radians(37)
    def rotate(p): return (p[0]*math.cos(angle)-p[1]*math.sin(angle),p[0]*math.sin(angle)+p[1]*math.cos(angle))
    moved = ([rotate(points[n]) for n in order],[(ids[a],ids[b]) for a,b in edges],
             {ids[n]:k for n,k in heads.items()},[])
    moved_text = [Text(t.identity,*rotate((t.x,t.y)),t.nominal_mm,127,t.height_mm,t.layer,t.raw_text,t.entity_handle) for t in labels]
    flow = build_flow_tree(moved[0],moved[1],[[n] for n in moved[2]],[ids[0]])
    got = define_diameters(*moved[:2],flow,moved_text,heads=moved[2])
    for e in edges:
        old,new = base.decisions[edge_key(*e)],got.decisions[edge_key(*(ids[n] for n in e))]
        assert (old.get('text_mm'),old.get('annotation_id')) == (new.get('text_mm'),new.get('annotation_id'))


def test_unknown_tree_does_not_substitute_a_rule_or_default_diameter():
    fixture = branches()
    result = run(fixture,[])
    edge = fixture[3][0][0]
    values,evidence,_ = resolve_pipe_definitions({'pipe_data':{'P':{}}},{'P':edge},result,{'P':15},tree=True,rule=lambda n:65)
    assert values['P'] == (0,'unresolved')
    assert evidence['P']['block_export'] and evidence['P']['rule_mm']==65


def test_two_main_loop_branch_matches_physical_direction_not_node_number():
    fixture = branches()
    points, edges, heads, paths = fixture
    # Connect both terminal ends to an upper main with ports on both sides.
    end1, end2 = paths[0][-1][1], paths[1][-1][1]
    left = len(points); points.extend([(0,20000),(18000,20000)])
    edges.extend([(left,end1),(end1,end2),(end2,left+1)])
    labels = [text(0,i,d) for i,d in enumerate([25,32,40,50,65])]
    baseline = run(fixture, labels)
    assert [baseline.decisions[e].get('text_mm') for e in paths[1]] == [25,32,40,50,65]
    # Reverse only the target branch IDs so its discovered traversal reverses.
    order = list(range(len(points)))
    chain = [paths[1][0][0]] + [e[1] for e in paths[1]]
    for a,b in zip(chain, reversed(chain)): order[a] = b
    ids = {old:new for new,old in enumerate(order)}
    new_points = [points[n] for n in order]
    new_edges = [(ids[a],ids[b]) for a,b in edges]
    new_heads = {ids[n]:kind for n,kind in heads.items()}
    flow = build_flow_tree(new_points,new_edges,[[n] for n in new_heads],[ids[0]])
    result = define_diameters(new_points,new_edges,flow,labels,heads=new_heads)
    for e in paths[1]:
        row = result.decisions[edge_key(*(ids[n] for n in e))]
        assert (row.get('text_mm'),row.get('annotation_id')) == (
            baseline.decisions[e].get('text_mm'),baseline.decisions[e].get('annotation_id'))


def test_length_tolerance_does_not_enable_recursive_donation():
    fixture = branches(3)
    points, _, _, paths = fixture
    # A matches B and B matches C, but A does not match C.
    for j,scale in enumerate((1,1.1,1.25)):
        for e in paths[j]:
            n=e[1];x,y=points[n];points[n]=(x,y*scale)
    result = run(fixture,[text(0,0,25)])
    assert result.decisions[paths[1][0]].get('text_mm') == 25
    assert not result.decisions[paths[2][0]].get('text_mm')


def test_target_fragment_conflict_does_not_overwrite_direct_label():
    from src.pipenet_converter.graph.diameter_definition import DefinitionConfig
    from src.pipenet_converter.graph.diameter_repetition import transfer_repeated_slots
    fixture = branches()
    points, edges, heads, paths = fixture
    a,b=paths[1][0];mid=len(points);points.append((12000,2000))
    edges.remove((a,b));edges.extend([(a,mid),(mid,b)])
    target_path=[edge_key(a,mid),edge_key(mid,b)]+paths[1][1:]
    nodes1=[paths[0][0][0]]+[e[1] for e in paths[0]]
    nodes2=[a,mid,b]+[e[1] for e in paths[1][1:]]
    rows={edge_key(*e):{} for e in edges}
    rows[paths[0][0]]=dict(text_mm=25,annotation_id='original',text_xy_mm=[6100,2000],definition_source='drawing_direct')
    rows[edge_key(a,mid)]=dict(text_mm=65,annotation_id='target',text_xy_mm=[12100,1000],definition_source='drawing_direct')
    transfer_repeated_slots(points,set(rows),[(nodes1,paths[0]),(nodes2,target_path)],heads,{0},rows,DefinitionConfig(),'drawing_first_v1')
    assert rows[edge_key(a,mid)]['text_mm']==65
    assert rows[edge_key(mid,b)]['block_export'] and not rows[edge_key(mid,b)].get('text_mm')


def test_multiple_same_size_sources_choose_physical_nearest_not_node_id():
    points=[(0,0),(4000,0),(8000,0),(12000,0),(16000,0)]
    edges=[(0,1),(1,2),(2,3),(3,4)]
    labels=[Text('left',2000,100,65,0),Text('right',14000,100,65,0)]
    flow=build_flow_tree(points,edges,[[4]],[0])
    base=define_diameters(points,edges,flow,labels)
    assert base.decisions[1,2]['annotation_id']=='left'
    assert base.decisions[2,3]['annotation_id']=='right'
    points2=list(reversed(points));edges2=[(4-a,4-b) for a,b in edges]
    other_flow=build_flow_tree(points2,edges2,[[0]],[4])
    other=define_diameters(points2,edges2,other_flow,labels)
    for a,b in edges:
        assert base.decisions[a,b]['annotation_id']==other.decisions[edge_key(4-a,4-b)]['annotation_id']

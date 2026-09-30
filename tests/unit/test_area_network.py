"""Area-scoped loop extraction must not import unrelated exterior cycles."""
import copy

import pytest

from src.pipenet_converter.graph.area_network import select_area_network
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.graph.network import FlowNetwork
from src.pipenet_converter.graph.review_conversion import build_review_network


def sample(mode="loop"):
    # Source 0, an outside ring 0-1-2-3, local ring 4-5-6-7, heads 8/9.
    points = [(0, 0), (1000, 0), (1000, 1000), (0, 1000),
              (2000, 0), (3000, 0), (3000, 1000), (2000, 1000),
              (3500, 0), (3500, 1000)]
    edges = {(0, 1), (1, 2), (2, 3), (0, 3), (1, 4),
             (4, 5), (5, 6), (6, 7), (4, 7), (5, 8), (6, 9)}
    ref = build_flow_tree(points, edges, [{8}, {9}], [0])
    return points, FlowNetwork.build(ref, mode)


@pytest.mark.parametrize("mode", ["loop", "grid"])
def test_local_loop_survives_and_only_real_supply_connector_remains(mode):
    points, network = sample(mode)
    original = copy.deepcopy(network)
    scope = select_area_network(network, points, [0, 1], [[1900, -100, 3600, 1100]])
    assert scope.edges == network.edges - {(1, 2), (2, 3), (0, 3)}
    assert {(4, 5), (5, 6), (6, 7), (4, 7)} <= scope.local_edges
    assert scope.supply_edges == {(0, 1)}
    assert scope.boundary_edges == {(1, 4)}
    assert scope.report(network)["cycle_rank"] == 1
    assert network == original
    model = build_review_network(network, points, [0, 1], {0: '상향식', 1: '상향식'},
        datum_m=3, upright_m=.5, pendant_rise_m=.3, pendant_drop_m=.5,
        combo_rise_m=.3, combo_up_m=.2, combo_drop_m=.5, source_edges=tuple(scope.edges))
    assert model.cycle_rank == 1
    assert {p.source_edge for p in model.pipes.values() if p.source_edge is not None} == scope.edges
    assert sum(n.role == 'head' for n in model.nodes.values()) == 2
    assert sum(p.length_m for p in model.pipes.values()) == pytest.approx(8.0)
    emitted = network.extraction(points, selected_heads=[0, 1], source_edges=scope.edges)
    assert emitted["scope"] == "region_and_supply" and emitted["cycle_rank"] == 1
    assert {(p['node_a'], p['node_b']) for p in emitted['pipes']} == scope.edges
    # Default non-area F scenario retains its pre-existing all-path policy.
    assert network.selected_edges([0, 1]) == network.edges


def test_disjoint_regions_zero_aliases_and_source_inside():
    points = [(0, 0), (1000, 0), (2000, 0), (2000, 0), (3000, 0), (4000, 0)]
    edges = {(n, n+1) for n in range(5)}
    flow = build_flow_tree(points, edges, [{3}, {5}], [0])
    network = FlowNetwork.build(flow, 'grid')
    scope = select_area_network(network, points, [0, 1],
        [[-100, -100, 100, 100], [1900, -100, 2100, 100], [3900, -100, 4100, 100]])
    assert scope.edges == edges and (2, 3) in scope.local_edges
    assert scope.report(network)['cycle_rank'] == 0


def test_polygon_cutout_is_not_a_box_and_crossings_do_not_connect():
    # The long outside branch 5-6 lies in the polygon's concave cut-out.
    points = [(0, 0), (1000, 0), (1000, 3000), (0, 3000),
              (0, 4000), (2000, 2000), (3000, 2000)]
    edges = {(0, 1), (1, 2), (2, 3), (0, 3), (3, 4), (1, 5), (5, 6)}
    ref = build_flow_tree(points, edges, [{4}, {6}], [0])
    net = FlowNetwork.build(ref, 'loop')
    polygon = {'type': 'polygon', 'points': [[-100, -100], [4000, -100], [4000, 500],
                                            [1500, 500], [1500, 4100], [-100, 4100]]}
    scoped = select_area_network(net, points, [0], [polygon])
    assert (5, 6) not in scoped.edges
    assert scoped.edges <= edges  # no joining of coincident geometric crossings


def test_invalid_scope_and_missing_regions_fail_explicitly():
    points, net = sample()
    with pytest.raises(ValueError, match='영역'):
        select_area_network(net, points, [0], [])
    with pytest.raises(ValueError, match='원본'):
        net.resolve_selection([0], source_edges={(0, 8)})
    with pytest.raises(ValueError, match='연결'):
        net.resolve_selection([0], source_edges={(5, 8)})


def test_necessary_external_route_is_kept_for_isolated_head_region():
    points, net = sample()
    scoped = select_area_network(net, points, [1], [[3490, 990, 3510, 1010]])
    assert (6, 9) in scoped.edges
    assert not {(0, 3), (2, 3)} & scoped.edges
    assert scoped.report(net)['cycle_rank'] == 0


def test_enclosed_freehand_and_input_order_are_stable():
    points, net = sample()
    polygon = {'type':'polygon', 'fill_rule':'enclosed', 'points':[
        [1900,-100],[3600,-100],[3600,1100],[1900,1100],[1900,-100],
        [1800,-200],[1900,-100]]}  # retraced pen tail adds no area
    expected = select_area_network(net, points, [0,1], [[1900,-100,3600,1100]])
    actual = select_area_network(net, points, [1,0], [polygon])
    assert actual == expected


def test_boundary_point_contact_does_not_keep_an_exterior_cycle():
    points, net = sample()
    scoped = select_area_network(net, points, [0,1], [[1000,-100,3600,1100]])
    assert (0,3) not in scoped.local_edges  # outside
    assert (2,3) not in scoped.local_edges  # only one endpoint touches boundary
    assert (0,1) in scoped.supply_edges    # contact is kept only for required supply

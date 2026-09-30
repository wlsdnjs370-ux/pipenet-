"""Displaced arc fittings: explicit contacts, not global distance relaxation."""
from __future__ import annotations

from collections import defaultdict
from copy import deepcopy
import math

import networkx as nx
import pytest

from routes.module_f.common import _boot
from src.pipenet_converter.graph.arc_contacts import offset_arc_contact

_boot()
from services.cad_import.pipeline.flow import thru_arms, join_all, ho_from_spots
from services.cad_import.convert.main_walk import ho_to_kfp_units, sit_arcs
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.graph.network import FlowNetwork
from src.pipenet_converter.graph.review_conversion import build_review_network


def fixture(offset=35.3718922006, angle=0):
    """Measured source motif translated/rotated without project layer rules."""
    theta = math.radians(angle)
    def point(x, y):
        return (160634.227508053 + x * math.cos(theta) - y * math.sin(theta),
                3007451.480906981 + x * math.sin(theta) + y * math.cos(theta))
    pts = [point(-offset, -1000), point(-offset, 1000), point(0, 0),
           point(186.705069071, 0), point(186.705069071, -900)]
    edges = {(0, 1), (2, 3), (3, 4)}
    spot = dict(k='호', cx=pts[2][0], cy=pts[2][1], r=180, sa=45+angle, sweep=270)
    return pts, edges, spot


def contact(pts, edges, spot):
    adj = defaultdict(set)
    for a, b in edges:
        adj[a].add(b); adj[b].add(a)
    return offset_arc_contact(pts, list(edges), adj, spot, [2, 3])


@pytest.mark.parametrize('angle', [0, 90, 180, 270, 37])
def test_only_evidenced_tee_split_at_real_trunk(angle):
    pts, edges, spot = fixture(angle=angle)
    before = deepcopy((pts, edges, spot))
    proof = contact(pts, edges, spot)
    assert proof is not None and proof.offset_mm == pytest.approx(35.3718922006)
    evidence = {}
    p, e, a, splits = thru_arms(pts, edges, [spot], [[2, 3]], evidence=evidence)
    assert splits == 1 and len(p) == len(pts) + 1
    assert (0, 1) not in e and {tuple(sorted(v)) for v in e} >= {(0, 5), (1, 5)}
    final, joins, _ = join_all(p, e, [spot], a)
    assert joins == [('호', 2, 5)] and nx.is_tree(nx.Graph(final))
    assert nx.shortest_path(nx.Graph(final), 0, 4) == [0, 5, 2, 3, 4]
    assert (pts, edges, spot) == before  # original CAD is untouched
    assert evidence[0]['connection_xy'] == pytest.approx(p[5])
    assert sum(math.dist(p[u], p[v]) for u, v in final) == pytest.approx(
        sum(math.dist(pts[u], pts[v]) for u, v in edges) + proof.offset_mm)
    _, twice, _, count = thru_arms(p, final, [spot], a)
    assert count == 0 and twice == frozenset(final)


@pytest.mark.parametrize('offset', [0, 20, 25, 45.1, 60, -35.4])
def test_existing_tolerance_and_large_or_wrong_side_offsets_not_overridden(offset):
    assert contact(*fixture(offset)) is None


@pytest.mark.parametrize('change', [dict(k='헤드'), dict(k='원'), dict(k='획'),
    dict(sa=None), dict(sweep=180), dict(sweep=90), dict(sweep=float('nan')),
    dict(r=80), dict(sa=135)])
def test_nonmatching_symbols_do_not_authorize_a_contact(change):
    p, e, s = fixture(); s.update(change)
    assert contact(p, e, s) is None


def test_no_centre_end_short_stroke_or_parallel_competing_trunk():
    p, e, s = fixture()
    wrong = p[:]; wrong[2] = (p[2][0], p[2][1] + 10)
    assert contact(wrong, e, s) is None
    short = p[:]; short[3] = (p[2][0] + 20, p[2][1])
    assert contact(short, e, s) is None
    double = p + [(p[0][0]+5, p[0][1]), (p[1][0]+5, p[1][1])]
    assert contact(double, e | {(5, 6)}, s) is None
    # Even an already-covered trunk disqualifies the additional offset trunk.
    double = p + [(p[2][0], p[0][1]), (p[2][0], p[1][1])]
    assert contact(double, e | {(5, 6)}, s) is None
    short_trunk = p[:]; short_trunk[1] = (p[1][0], p[2][1]+100)
    assert contact(short_trunk, e, s) is None


def test_genuine_loop_is_retained_without_extra_local_cycle():
    p, e, s = fixture()
    # A real return path outside the symbol. Only the missing local tee closes it.
    e.add((0, 4))
    p, e, a, _ = thru_arms(p, e, [s], [[2, 3]])
    e, joins, _ = join_all(p, e, [s], a)
    assert len(joins) == 1 and len(nx.cycle_basis(nx.Graph(e))) == 1


@pytest.mark.parametrize('mode', ['tree', 'loop', 'grid'])
def test_flow_and_fitting_seat_are_at_restored_tee_not_straight_tip(mode):
    p, e, s = fixture()
    evidence = {}
    p, e, a, _ = thru_arms(p, e, [s], [[2, 3]], evidence=evidence)
    e, _, _ = join_all(p, e, [s], a)
    s.update(evidence[0])
    arcs = ho_from_spots([s])
    assert (arcs[0]['cx'], arcs[0]['cy']) == p[2]
    assert arcs[0]['connection_xy'] == list(p[5])
    # Both conversion paths must seat the arc at the physical three-port node.
    ho = ho_to_kfp_units(arcs, 0, 0)
    xy = {str(n): (x/1000+1, y/1000+1) for n, (x,y) in enumerate(p)}
    assert set(sit_arcs(xy, ho, .18)) == {'5'}
    ref = build_flow_tree(p, e, [{4}], [0])
    assert ref.path(4) == [0, 5, 2, 3, 4]
    if mode == 'tree':
        from test_module_f_arc_symbols import _node, _pipe
        from services.cad_import.convert.engine import convert_to_kfp
        nodes = {n: _node(n, *v, 'head' if n == '4' else 'base') for n, v in xy.items()}
        pipes = {f'P{i}': _pipe(str(a), str(b), math.dist(p[a], p[b])/1000)
                 for i, (a, b) in enumerate(sorted(e))}
        out = convert_to_kfp(dict(kfp=dict(nodes_meta_runtime=nodes, pipe_data=pipes,
            node_counter={'N':0}, pipe_id_counter=0), ho=ho,
            sources=[{'xy': xy['0']}], node_head_kinds={'4':'상향식'}))
        assert out['ok'], out['blockers']
        got = out['kfp']
        assert '5' in got['arc_junctions'] and '2' not in got['arc_junctions']
        assert got['nodes_meta_runtime']['2']['coords'][2] == .5
        assert got['nodes_meta_runtime']['0']['coords'][2] == 0
        return
    net = FlowNetwork.build(ref, mode)
    g = build_review_network(net, p, [0], {0:'상향식'}, datum_m=0,
        upright_m=.5, pendant_rise_m=.3, pendant_drop_m=.5,
        combo_rise_m=.3, combo_up_m=.2, combo_drop_m=.5, arcs=arcs)
    assert g.arc_report['junctions'] == [dict(node='N5', top='V_N5', kind='junction')]
    assert not g.arc_report['review']
    assert g.nodes['N2'].z_m == .5 and g.nodes['N0'].z_m == 0
    assert not any(r['node']=='N2' for r in g.arc_report['junctions'])
    g.validate()


def test_cache_tracks_new_contact_rule():
    from services.cad_import.pipeline.disp_cache import _disp_cache_inputs
    stamp = _disp_cache_inputs('__offset_contact_test__')
    assert any(row['path'].endswith('arc_contacts.py') for row in stamp['code'])


def test_real_planar_adapter_retains_tee_identity_after_grid_snapping():
    from services.cad_import.convert.engine import ensure_planar, _prepare_ho, _walk_main_kfp
    p, e, s = fixture()
    evidence = {}
    p, e, a, _ = thru_arms(p, e, [s], [[2, 3]], evidence=evidence)
    e, _, _ = join_all(p, e, [s], a)
    s.update(evidence[0])
    arcs = ho_from_spots([s])
    payload = dict(key='__offset_only__', pts=p, edges=e,
        hcov=[(*p[n], 50) for n in (1, 4)], ups=[],
        head_kinds=[dict(c=list(p[n]), r=50, kind='하향식') for n in (1, 4)],
        sources=[{'xy': list(p[0])}], ho=arcs)
    planar = ensure_planar(payload)
    assert 'kfp' in planar, planar
    converted_arcs, src = _prepare_ho(planar, planar['kfp'])
    arc = converted_arcs[0]
    tee = arc['connection_node']
    assert tee in planar['kfp']['nodes_meta_runtime']
    walked, reason = _walk_main_kfp(planar['kfp'], converted_arcs, src)
    assert walked is not None, reason
    assert any(n == tee for n, other in walked['branch_first'])
    assert 'connection_node' not in arcs[0]  # caller's source metadata untouched


def test_removed_bound_contact_cannot_reattach_to_nearby_node():
    arc = dict(cx=0, cy=0, r=.18, sa=45, sweep=270,
               connection_evidence='arc_open_branch_unique_through',
               connection_xy=[-.035, 0], connection_node='removed')
    assert not sit_arcs({'other': (0, 0)}, [arc], .18)
    arc['connection_node'] = None
    assert not sit_arcs({'other': (0, 0)}, [arc], .18)

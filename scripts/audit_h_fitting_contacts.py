"""Compare MF20 fitting connections without changing saved user work."""
from __future__ import annotations

from pathlib import Path
from unittest.mock import patch
import json
import math

import networkx as nx

from routes.module_f.common import _boot


def main() -> None:
    """Replay the original picked drawing and publish an inspectable comparison."""
    _boot()
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline, ho_from_spots
    from services.cad_import.edit.io import _board_from_data
    from services.cad_import.edit.session import EditSession
    from services.cad_import.convert.engine import ensure_planar, _prepare_ho, _walk_main_kfp
    from services.cad_import.design.cycle_source import cycle_source
    from src.pipenet_converter.graph.flow import build_flow_tree
    from src.pipenet_converter.graph.network import FlowNetwork
    from src.pipenet_converter.graph.review_conversion import build_review_network

    spec = next(Path('cad_project_editor_g/docs/import/0단계_새찍기').glob('MF20*찍은스펙.json'))
    key = spec.name.removesuffix('_찍은스펙.json')
    st = stage1_body(key)
    with patch('src.pipenet_converter.graph.arc_contacts.offset_arc_contact', return_value=None):
        before = pipeline(st)
    after = pipeline(st)
    assert before['hcov'] == after['hcov']
    assert before['head_kinds'] == after['head_kinds']

    def board(result):
        return _board_from_data('__audit_only__', dict(result, ho=ho_from_spots(result['spots'])))

    old, new = board(before), board(after)
    old_graph = nx.Graph(old.edges)
    new_graph = nx.Graph(new.edges)
    contacts = [s for s in after['spots'] if s.get('connection_xy')]
    assert len(contacts) == 2
    records = []
    for s in contacts:
        centre = (s['cx'], s['cy'])
        # The unchanged trunk point above the arc, directly read from the source.
        old_branch = min(old_graph, key=lambda n: math.dist(old.pts[n], centre))
        new_branch = min(new_graph, key=lambda n: math.dist(new.pts[n], centre))
        joint = min(new_graph, key=lambda n: math.dist(new.pts[n], s['connection_xy']))
        main = next(n for n in new_graph[joint] if abs(new.pts[n][1]-centre[1]) > 180)
        old_main = min(old_graph, key=lambda n: math.dist(old.pts[n], new.pts[main]))
        was_connected = nx.has_path(old_graph, old_branch, old_main)
        # The remote return already reaches the same component: component
        # count alone cannot prove that THIS tee or loop is complete.
        assert not any(math.dist(old.pts[n], s['connection_xy']) < 1e-3
                       for n in old_graph[old_branch])
        assert nx.has_path(new_graph, new_branch, main)
        assert new_graph.has_edge(joint, new_branch)
        assert new_graph.degree(joint) == 3 and new_graph.degree(new_branch) == 2
        comp = nx.node_connected_component(new_graph, main)
        new_heads = [i for i, ns in enumerate(new.hnodes) if set(ns) & comp]
        # Exercise the real planar adapter too (it may compact/snap nodes).
        new.set_source_nodes([main])
        planar = ensure_planar(EditSession(new).convert_payload())
        assert planar.get('kfp') is not None, planar.get('_planar_error')
        ho, source_xy = _prepare_ho(planar, planar['kfp'])
        walked, reason = _walk_main_kfp(planar['kfp'], ho, source_xy)
        assert walked is not None, reason
        from services.cad_import.convert.main_walk import xf_mm_to_m, sit_arcs
        expected = xf_mm_to_m(*s['connection_xy'], *planar['origin_mm'])
        positions = {n: tuple(row['coords'][:2]) for n, row in planar['kfp']['nodes_meta_runtime'].items()}
        seated = sit_arcs(positions, ho, max(a['r'] for a in ho))
        tee = min(positions, key=lambda n: math.dist(positions[n], expected))
        assert tee in seated
        assert any(n == tee for n, other in walked['branch_first'])
        modes = {'tree': {'restored_joint': tee, 'open_branch_detected': True}}
        for mode in ('loop', 'grid'):
            edges = cycle_source(new).edges
            flow = build_flow_tree(new.pts, edges, new.hnodes, [main])
            net = FlowNetwork.build(flow, mode)
            model = build_review_network(net, new.pts, new_heads,
                {i: '상향식' for i in new_heads}, datum_m=0, upright_m=.5,
                pendant_rise_m=.3, pendant_drop_m=.5, combo_rise_m=.3,
                combo_up_m=.2, combo_drop_m=.5, arcs=new.ho)
            # Height defaults here test fitting port propagation only; they are
            # not a hydraulic approval of this dwelling's head specifications.
            assert any(r['node'] == f'N{joint}' for r in model.arc_report['junctions'])
            assert model.nodes[f'N{new_branch}'].z_m == .5
            model.validate()
            modes[mode] = dict(cycle_rank=model.cycle_rank,
                               restored_joint=f'N{joint}', connected_heads=len(new_heads))
        records.append(dict(symbol_xy=centre, trunk_xy=s['connection_xy'],
            gap_mm=s['connection_offset_mm'], before_local_contact=False, after_local_contact=True,
            remote_return_already_connected=was_connected,
            joint_ports=new_graph.degree(joint), branch_tip_ports=new_graph.degree(new_branch),
            conversion=modes))
    report = dict(source=key, head_count=len(new.disks), head_kinds_unchanged=True,
        components_before=nx.number_connected_components(old_graph),
        components_after=nx.number_connected_components(new_graph),
        cycles_before=len(nx.cycle_basis(old_graph)), cycles_after=len(nx.cycle_basis(new_graph)),
        contacts=records)
    assert report['cycles_after'] == report['cycles_before'] + 2
    out = Path('outputs/h_fitting_contact_review'); out.mkdir(parents=True, exist_ok=True)
    (out/'report.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf8')
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()

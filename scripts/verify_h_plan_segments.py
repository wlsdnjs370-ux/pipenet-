"""Verify an offline H snapshot and emit a separate, explicitly draft SDF."""
from __future__ import annotations

import argparse
from collections import Counter
from copy import deepcopy
from dataclasses import asdict
import json
import math
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT / 'core')]

from core.remote30_full_network import CombinedTables
from src.pipenet_converter.graph.continuous_runs import compact_continuous


def main() -> None:
    """Prove source partition, topology, attachments and explicit SDF values."""
    parser = argparse.ArgumentParser()
    parser.add_argument('directory', type=Path)
    args = parser.parse_args()
    snapshot = json.loads((args.directory / 'snapshot.json').read_text(encoding='utf-8'))
    data = snapshot['tables']
    before = CombinedTables(**{k: data[k] for k in ('nodes', 'pipes', 'nozzles', 'fittings', 'equipment')})
    untouched = deepcopy(before)
    after, audit = compact_continuous(before, allow_unassigned=True)
    assert before == untouched
    assert Counter(p['label'] for rows in audit.source_pipes.values() for p in rows) == Counter(p['label'] for p in before.pipes)
    assert after.nozzles == before.nozzles
    assert Counter((r.get('node'), r['type'], r['count']) for r in after.fittings) == Counter((r.get('node'), r['type'], r['count']) for r in before.fittings)
    old_nodes = {n['label']: n for n in before.nodes}
    assert all(n == old_nodes[n['label']] for n in after.nodes)
    import networkx as nx
    def topology(tables):
        graph = nx.MultiGraph((p['in'], p['out']) for p in tables.pipes)
        count = nx.number_connected_components(graph)
        return count, len(graph.edges) - len(graph.nodes) + count
    assert topology(before) == topology(after)
    for p in after.pipes:
        rows = audit.source_pipes[p['label']]
        assert {r['dia'] for r in rows} == {p['dia']}
        assert math.isclose(p['length'], math.fsum(r['length'] for r in rows), abs_tol=1e-8)
        if not p['dia']:
            assert p['bore_provenance']['block_export']
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.design.emit import emit_design_sdf
    from routes.module_f.api_design import _view_opts, _valve_label
    from src.pipenet_converter.render.export_style import EXPORT_SYMBOL_SPREAD
    options = _view_opts(snapshot['settings'])
    options['canvas_units'] *= EXPORT_SYMBOL_SPREAD
    options['head_stub_ratio'] /= EXPORT_SYMBOL_SPREAD
    sdf = emit_design_sdf(after, args.directory / 'plan_compacted_review.sdf',
        project_title='H continuous-pipe verification — NOT VALIDATED',
        iso_ref_label=_valve_label(after), plan_review=True, **options)
    root = ET.parse(sdf).getroot()
    pipes = list(root.iter('Pipe'))
    assert {p.get('label') for p in pipes} == {p['label'] for p in after.pipes}
    assert len(pipes) == len(after.pipes)
    assert len(list(root.iter('Nozzle'))) == len(after.nozzles)
    by_label = {p['label']: p for p in after.pipes}
    for p in pipes:
        row = by_label[p.get('label')]
        assert (p.get('input'), p.get('output')) == (row['in'], row['out'])
        assert p.get('length') == format(float(row['length']), '.6g')
    report = dict(nodes_before=len(before.nodes), nodes_after=len(after.nodes),
        pipes_before=len(before.pipes), pipes_after=len(after.pipes),
        referenced=sum(bool(p['dia']) for p in after.pipes),
        unreferenced=sum(not p['dia'] for p in after.pipes),
        heads=len(after.nozzles), fittings=len(after.fittings),
        total_length_m=audit.total_length_m, components=topology(after)[0], cycles=topology(after)[1],
        source_coverage=len(before.pipes), sdf=str(sdf.resolve()),
        notice='Offline snapshot; original state unchanged. Unknown bores remain blocked. PIPENET GUI not recalculated.')
    for name, obj in [('verification.json', report), ('compaction.json', asdict(audit)), ('canonical_tables.json', asdict(after))]:
        (args.directory / name).write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()

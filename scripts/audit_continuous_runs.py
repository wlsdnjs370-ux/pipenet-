"""Offline proof of canonical contraction on the saved B1F calculation snapshot."""
from __future__ import annotations

from dataclasses import asdict
import json
from pathlib import Path
import sys

ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'core'),str(ROOT/'scripts')]

import networkx as nx
from audit_export_compaction import snapshot_tables
from core.network_editor import Network,apply_edit
from routes.module_f import network_edit as ne
from routes.module_f.hydraulic_sizing import prepare
from src.pipenet_converter.hydraulics.solver import solve


def main() -> None:
    """Compare topology, every retained elevation, losses and solved head flows."""
    path=ROOT/'outputs/module_f_fitting_resolution_20260927/live_before.json'
    snapshot=json.loads(path.read_text(encoding='utf-8'))
    before=Network.from_tables(snapshot_tables(snapshot),protected={'1','10'})
    original=ne.fingerprint(before)
    after,selection=apply_edit(before,dict(op='compact_runs',version=1),ne.catalog())
    assert ne.fingerprint(before)==original
    graphs=[]
    solutions=[]
    lengths=[]
    equivalent=[]
    for graph in (before,after):
        tables=graph.to_tables()
        g=nx.MultiGraph()
        g.add_nodes_from(graph.nodes)
        g.add_edges_from((p.a,p.b) for p in graph.pipes.values())
        graphs.append(dict(nodes=len(g),pipes=g.number_of_edges(),heads=len(tables.nozzles),
                           components=nx.number_connected_components(g),
                           cycles=g.number_of_edges()-len(g)+nx.number_connected_components(g)))
        net,_,_=prepare(tables,{'mode':'pump_only'})
        result=solve(net,tuple(p.initial for p in net.pipes),8)
        assert result.converged
        solutions.append(result)
        lengths.append(sum(p.length_m for p in net.pipes))
        equivalent.append(sum(p.sizes[p.initial].equivalent_m for p in net.pipes))
    assert graphs[0]['cycles']==graphs[1]['cycles']
    assert graphs[0]['components']==graphs[1]['components']
    assert graphs[0]['heads']==graphs[1]['heads']
    assert abs(lengths[0]-lengths[1])<1e-8 and abs(equivalent[0]-equivalent[1])<1e-8
    assert all(n.xyz==before.nodes[k].xyz for k,n in after.nodes.items())
    flow=max(abs(solutions[0].nozzle_flow_lpm[k]-v) for k,v in solutions[1].nozzle_flow_lpm.items())
    pressure=max(abs(solutions[0].pressure_bar[k]-v) for k,v in solutions[1].pressure_bar.items())
    assert flow<.001 and pressure<.0001,(flow,pressure)
    output=ROOT/'outputs/module_f_continuous_runs_20260928'
    output.mkdir(parents=True,exist_ok=True)
    report=dict(source=str(path),scope='Offline historical snapshot, not the current browser session',
                before=graphs[0],after=graphs[1],length_m=lengths[1],equivalent_m=equivalent[1],
                max_nozzle_flow_delta_lpm=flow,max_retained_pressure_delta_bar=pressure,
                fitting_rows_before=len(before.tables.fittings),fitting_rows_after=len(after.tables.fittings),
                equipment_rows_before=len(before.tables.equipment),equipment_rows_after=len(after.tables.equipment),
                compaction=selection['compaction'])
    (output/'verification.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    (output/'canonical_tables.json').write_text(json.dumps(asdict(after.tables),ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in report.items() if k!='compaction'},ensure_ascii=False,indent=2))


if __name__=='__main__':
    main()

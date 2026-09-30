"""Validate compact SDFs against an existing read-only Module F snapshot."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT/'core')]

import networkx as nx
from core.remote30_full_network import CombinedTables
from routes.module_f.hydraulic_sizing import fingerprint, prepare
from routes.module_f.sizing_export import emit_proposal
from routes.module_f.emit import emit_merged, ISO_SPREAD
from routes.module_f.merge import bake_combined_iso, bake_combined_plan
from src.pipenet_converter.hydraulics.models import Network, Node, Nozzle, Pipe, Size
from src.pipenet_converter.hydraulics.solver import solve


def snapshot_tables(snapshot: dict) -> CombinedTables:
    """Rebuild exact editor rows/inspection loss rows from the saved API response."""
    editor = snapshot['merge_editor']
    view = snapshot['merge_preview']['view']
    inputs = {n['label'] for n in view['nodes'] if n.get('input')}
    issues = [dict(pipe_label=r['pipe'],node_label=r['node'],n=r['count'],reason=r['note'])
              for r in view['inspection']['fittings'] if r['kind']=='unresolved']
    return CombinedTables(
        nodes=[dict(label=n['label'],x=n['xyz'][0]*1000,y=n['xyz'][1]*1000,
                    elevation=n['xyz'][2],io_node='Input' if n['label'] in inputs else 'No')
               for n in editor['nodes']],pipes=deepcopy(editor['pipes']),
        fittings=[dict(pipe=r['pipe'],node=r['node'],type=r['kind'],count=r['count'],
                       flow_direction='solver_reference' if r['flow_label'].startswith('손실 귀속') else '',
                       loss_status=r.get('loss_status')) for r in view['inspection']['fittings']
                  if r['origin']=='부속 입력표'],equipment=deepcopy(editor['equipment']),
        nozzles=[dict(label=h['label'],**{'in':h['node'],'out':f"@/{h['label']}"},
                      lib=h['lib'],status='1' if h['active'] else '0',flow_lmin=80,flow_m3s=80/60000)
                 for h in snapshot['sizing_state']['basis']['nozzles']],
        unresolved=dict(kind_items=issues),meta=[('부속 판정 불가',str(sum(r['n'] for r in issues)))])


def main() -> None:
    """Export offline copies, compare cycle rank, loss and a full hydraulic solve."""
    previous = ROOT/'outputs/module_f_fitting_resolution_20260927'
    snapshot = json.loads((previous/'live_before.json').read_text(encoding='utf-8'))
    tables = snapshot_tables(snapshot)
    before = fingerprint(tables)
    output = ROOT/'outputs/module_f_export_compaction_20260927'
    output.mkdir(parents=True,exist_ok=True)
    view = snapshot['merge_preview']['view']
    parts = {part:[n['label'] for n in view['nodes'] if n.get('part')==part]
             for part in ('plan','system','machineroom')}
    got = dict(combined=tables,parts=parts,system_layout='physical_xy',pump_junction='1',mode='hsp_pump')
    normal = emit_merged(tables,output/'normal',iso_nodes=bake_combined_iso(got)[0],
        plan_nodes=bake_combined_plan(got)[0],display_reference_labels=parts['plan'],
        iso_spread=ISO_SPREAD,keep_nodes=('10','1'))
    report = json.loads((previous/'verification.json').read_text(encoding='utf-8'))['report']
    net,_,_ = prepare(tables,{'mode':'joint'})
    report.update(fingerprint=before,model=asdict(net))
    files = emit_proposal(got,report,output/'sizing')
    exported = json.loads(Path(files['report']).read_text(encoding='utf-8'))
    audit = exported['export_compaction']
    nozzle_defs = {h['label']:h for h in report['nozzles']}
    source_defs = {h.label:h for h in net.nozzles}
    proposed = {p['label']:p for p in report['pipes']}
    roots = [ET.parse(files[key]).getroot() for key in ('sdf','sdf_iso')]
    for root in roots:
        physical = {p.get(k) for p in root.iter('Pipe') for k in ('input','output')}
        nodes = tuple(Node(n.get('label'),float(n.get('elevation'))) for n in root.iter('Node')
                      if n.get('label') in physical)
        pipes = []
        for pipe in root.iter('Pipe'):
            label = pipe.get('label')
            old = audit['source_pipes'][label]
            target = proposed[old[0]['label']]
            eq = sum(float(e.get('equivalent-length')) for e in pipe.iter('Equipment'))
            assert abs(eq-sum(proposed[p['label']]['equivalent_m'] for p in old)) < 1e-6
            pipes.append(Pipe(label,pipe.get('input'),pipe.get('output'),float(pipe.get('length')),
                float(pipe.get('roughness-or-c')),(Size(target['after_dn'],target['inner_mm'],eq),)))
        nozzles = tuple(Nozzle(h.get('label'),h.get('input'),source_defs[h.get('label')].k_lpm_sqrt_bar,
            nozzle_defs[h.get('label')]['min_flow_lpm'],nozzle_defs[h.get('label')]['min_pressure_bar'],
            nozzle_defs[h.get('label')]['max_pressure_bar']) for h in root.iter('Nozzle'))
        compact = Network(nodes,tuple(pipes),nozzles,report['source'])
        boundary = root.find(f'.//Node[@label="{report["source"]}"]/Calculation-spec')
        written_pressure = (float(boundary.get('pressure'))-101325)/100000
        solved = solve(compact,tuple(0 for _ in pipes),written_pressure)
        assert solved.converged
        original_graph = nx.MultiGraph((p['in'],p['out']) for p in tables.pipes)
        compact_graph = nx.MultiGraph((p.a,p.b) for p in pipes)
        rank = lambda g: g.number_of_edges()-g.number_of_nodes()+nx.number_connected_components(g)
        assert nx.is_connected(compact_graph) and rank(compact_graph)==rank(original_graph)
        assert len(nozzles)==len(net.nozzles)
        flow_delta = max(abs(v-report['solution']['nozzle_flow_lpm'][k]) for k,v in solved.nozzle_flow_lpm.items())
        pressure_delta = max(abs(v-report['solution']['pressure_bar'][k]) for k,v in solved.pressure_bar.items())
        assert flow_delta < .001 and pressure_delta < .0001, (flow_delta,pressure_delta)
    assert fingerprint(tables)==before
    result = dict(normal=json.loads(Path(normal['compaction']).read_text(encoding='utf8')),
        sizing=audit,cycles=rank(compact_graph),heads=len(net.nozzles),
        max_nozzle_flow_change_lpm=flow_delta,max_retained_node_pressure_change_bar=pressure_delta,
        source_unchanged=True,files=files,notice='Offline snapshot verification; PIPENET GUI not recalculated')
    (output/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k not in ('normal','sizing')},ensure_ascii=False))


if __name__=='__main__':
    main()

"""Check an existing read-only live snapshot; never edit the user's session."""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT/'core')]

from core.remote30_full_network import CombinedTables
from routes.module_f.hydraulic_sizing import prepare, fingerprint
from src.pipenet_converter.graph.fitting_review import fitting_review
from src.pipenet_converter.hydraulics.sizing import propose


def main() -> None:
    """Verify current endpoint edits, loss transfer and a complete hydraulic solve.

    Pipe/node values come from the live editor API; native fitting identities
    come from its inspector. This audit does not reconstruct pump/control links.
    """
    destination = ROOT/'outputs/module_f_fitting_resolution_20260927'
    snapshot_path = destination/'live_before.json'
    snapshot = json.loads(snapshot_path.read_text(encoding='utf-8'))
    summary = snapshot['merge_state']['summary']
    assert summary['pumps'] == summary['valves'] == 0, 'Audit requires link-free supply boundary'
    editor = snapshot['merge_editor']
    view = snapshot['merge_preview']['view']
    basis = snapshot['sizing_state']['basis']
    inputs = {n['label'] for n in view['nodes'] if n.get('input')}
    fittings = [dict(pipe=r['pipe'], node=r['node'], type=r['kind'], count=r['count'],
                     flow_direction='solver_reference' if r['flow_label'].startswith('손실 귀속') else '',
                     loss_status=r.get('loss_status'))
                for r in view['inspection']['fittings'] if r['origin']=='부속 입력표']
    issues = [dict(pipe_label=r['pipe'],node_label=r['node'],n=r['count'],reason=r['note'])
              for r in view['inspection']['fittings'] if r['kind']=='unresolved']
    tables = CombinedTables(
        nodes=[dict(label=n['label'],x=n['xyz'][0]*1000,y=n['xyz'][1]*1000,
                    elevation=n['xyz'][2],io_node='Input' if n['label'] in inputs else 'No')
               for n in editor['nodes']],
        pipes=deepcopy(editor['pipes']), fittings=fittings, equipment=deepcopy(editor['equipment']),
        nozzles=[dict(label=h['label'],**{'in':h['node']},lib=h['lib'],status='1' if h['active'] else '0')
                 for h in basis['nozzles']],
        unresolved=dict(kind_items=issues),meta=[('부속 판정 불가',str(sum(r['n'] for r in issues)))])
    assert (len(tables.nodes),len(tables.pipes),len(tables.fittings),len(tables.nozzles)) == (
        summary['nodes'],summary['pipes'],summary['fittings'],summary['nozzles'])
    assert {p['label'] for p in tables.pipes} == {p['label'] for p in basis['pipes']}
    assert inputs == {basis['source']}
    before = fingerprint(tables)
    review = fitting_review(tables)
    assert len(issues) == len(review.resolved) == 1 and review.count == 0, review
    net, cfg, warnings = prepare(tables,{'mode':'joint'})
    report = propose(net,cfg,progress=print)
    print('RESULT',report['status'],report['pump'], 'VIOLATIONS', report['violations'])
    assert fingerprint(tables) == before
    result = dict(scope='current integrated API snapshot; default review constraints, not final design',
        snapshot=str(snapshot_path),review=asdict(review),equipment=tables.equipment,
        nodes=len(tables.nodes),pipes=len(tables.pipes),heads=len(tables.nozzles),
        warnings=warnings,report=report)
    (destination/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf-8')
    assert report['solution']['converged'],report['violations']


if __name__ == '__main__':
    main()

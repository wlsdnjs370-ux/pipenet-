"""Read-only saved-drawing audit plus a reproducible large-grid benchmark."""
from __future__ import annotations

from dataclasses import asdict
import json
from pathlib import Path
import sys
import time

ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT),str(ROOT/'core')]
from src.pipenet_converter.hydraulics.models import Network, Node, Pipe, Size, Nozzle, Settings
from src.pipenet_converter.hydraulics.sizing import propose


def grid_case(width=25):
    """625 nodes, 1200 physical pipes, 576 independent cycles, 25 emitters."""
    nodes=tuple(Node(f'{x}:{y}',0) for x in range(width) for y in range(width))
    sizes=tuple(Size(d,inner) for d,inner in [(25,27.6),(32,35.7),(40,41.6),(50,52.9),(65,67.9),(80,80.7),(100,105.3)])
    pipes=[]
    for x in range(width):
        for y in range(width):
            for dx,dy in [(1,0),(0,1)]:
                if x+dx<width and y+dy<width:
                    pipes.append(Pipe(str(len(pipes)+1),f'{x}:{y}',f'{x+dx}:{y+dy}',3,120,sizes))
    nozzles=tuple(Nozzle(str(y),f'{width-1}:{y}',80,80,1,12) for y in range(width))
    net=Network(nodes,tuple(pipes),nozzles,'0:0')
    started=time.perf_counter()
    report=propose(net,Settings(mode='joint'),progress=print)
    report['seconds']=time.perf_counter()-started
    report['cycle_rank']=len(pipes)-len(nodes)+1
    print('GRID',report['status'],report['seconds'],report['cycle_rank'],report['pump'])
    assert report['feasible']
    assert report['cycle_rank']==(width-1)**2
    return report


def b1f_case():
    """Reconstruct current saved selection without changing the user's board."""
    from scripts.audit_loop_fittings import live_board
    from services.cad_import.design.preserved import expand_preserved,review_fittings
    from services.cad_import.design.tables import build_design_tables
    from routes.module_f.hydraulic_sizing import prepare
    board,heads=live_board()
    got=expand_preserved(board,heads)
    bores={p:(65,'review_default') for p in got['kfp']['pipe_data']}
    fits=review_fittings(got['kfp'],bores,got['physical_ports'])
    tbl=build_design_tables(got['kfp'],{},got['edge_ref'],[],bores=bores,fittings=fits)
    # No imaginary riser is added to make a full building case. This is the
    # saved plan component, checked through the exact same hydraulic adapter.
    info=dict(heads=len(heads),nodes=len(tbl.nodes),pipes=len(tbl.pipes),cycle_rank=got['cycle_rank'],
              unresolved=getattr(tbl,'unresolved',{}),arc_review=got['kfp'].get('arc_report',{}).get('review',[]))
    try:
        net,cfg,warnings=prepare(tbl,{})
        info.update(result=propose(net,cfg,progress=print),warnings=warnings)
    except ValueError as exc:
        info['blocked_reason']=str(exc)
    print('B1F',json.dumps({k:v for k,v in info.items() if k not in {'unresolved','result','arc_review'}},ensure_ascii=False))
    return info


if __name__=='__main__':
    destination=ROOT/'outputs'/'module_f_hydraulic_sizing_20260927'
    destination.mkdir(exist_ok=True,parents=True)
    result={'grid':grid_case()}
    if '--b1f' in sys.argv:
        result['b1f']=b1f_case()
    (destination/'verification.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,allow_nan=False),encoding='utf-8')

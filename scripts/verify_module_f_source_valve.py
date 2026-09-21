"""Compare AV branch generation on the saved Daemyeong plan, without saving edits."""
from __future__ import annotations

import json
from pathlib import Path
import sys
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT), str(ROOT/'core')]


def main() -> None:
    """Run identical selected-head expansion with old/new boundary behavior."""
    import pandas  # Load native data libraries before the optional Qt library reader.
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.flow import flow_for_board
    from services.cad_import.design.worst import worst_k_heads
    from services.cad_import.design.restrict import expand_worst, tree_loads
    from services.cad_import.design.tables import build_design_tables

    sys.stdout.reconfigure(encoding='utf-8')
    key = '1. 입력도면 대명동 단위세대 평면도'
    es = EditSession.open(key)
    b = es.board
    payload = es.convert_payload()
    flow = flow_for_board(b)
    worst = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30,
                         head_xy=b.disks, flow_tree=flow)
    with patch('src.pipenet_converter.graph.boundaries.separate_valve_picks',
               lambda valves, sources: (list(range(len(valves))), [])):
        old = expand_worst(payload,b,worst,key=key)
    new = expand_worst(payload,b,worst,key=key)
    for name, got in [('before',old), ('after',new)]:
        assert got.get('ok'), (name,got.get('error'))
    def stats(got):
        kfp=got['kfp']; loads=tree_loads(kfp)
        return dict(nodes=len(kfp['nodes_meta_runtime']),pipes=len(kfp['pipe_data']),
                    heads=sum(n.get('type_id')=='head' for n in kfp['nodes_meta_runtime'].values()),
                    blind_pipes=[dict(id=p,**kfp['pipe_data'][p]) for p,v in loads.items() if v==0])
    report = dict(before=stats(old),after=stats(new))
    assert report['before']['heads']==report['after']['heads']==30
    assert report['before']['nodes']-report['after']['nodes']==2
    assert report['before']['pipes']-report['after']['pipes']==2
    assert not report['after']['blind_pipes']
    assert sorted(p['length_m'] for p in report['before']['blind_pipes'])==[.5,2.5]
    # Build the same equipment table and verify the AV did not disappear with
    # its duplicate drawing branch. The source node is the explicit AV pick.
    net=new['kfp']
    from services.cad_import.design.anchor import require_anchor
    root=require_anchor(net['nodes_meta_runtime'],what='공통 절점 검증')
    tables=build_design_tables(net,new['worst'] if 'worst' in new else worst,
                              new['edge_ref'], [], board_pts=b.pts,
                              tree_loads=new['tree_loads'], valve_nodes=[root],
                              phys=new['phys'], origin_mm=new['origin_mm'])
    report['alarm_valve_equipment']=tables.equipment
    assert any(e['desc']=='A/V' for e in tables.equipment)
    dest=ROOT/'data/merge_boundary_review_20260920/plan_comparison.json'
    dest.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False))


if __name__=='__main__':
    main()

"""Read original MF20 symbols; audit parsing/adoption without a live session."""
from __future__ import annotations

import json
import math
import tempfile
import time
from pathlib import Path

import ezdxf

from routes.module_f.common import _boot


def main() -> None:
    """Keep source/saved work untouched and write an inspectable review report."""
    _boot()
    from services.cad_import.pipeline.stage1 import read_dxf, explode, DEFAULT_KNOBS
    from services.cad_import.pick.board import Board
    from services.cad_import.pick.session import PickSession
    from routes.module_f.adopt import adopt_heads, select_heads
    import remote30_prototype as A

    start=time.perf_counter()
    source=next(Path('data/uploads').glob('MF20-001*.dxf'))
    layers,entities,blocks=read_dxf(source)
    name='측벽(드라이펜던트)'  # source-specific audit, never a recognition rule
    inserts=[e for e in entities if e.get('2')==name]
    visible=[e for e in inserts if e.get('67')!=1 and not layers.get(e.get('8'),[7,False])[1]]
    world,_=explode(layers,entities,blocks)
    symbol_world,_=explode(layers,visible,blocks)
    # Copy only these original primitive records into a tiny DXF. This avoids
    # loading the 516 MB architectural HATCH data into ezdxf for the audit.
    doc=ezdxf.new()
    for layer in {e.get('8','0') for e in visible}:
        if layer not in doc.layers:
            doc.layers.new(layer,dxfattribs={'color':layers.get(layer,[7])[0]})
    block=doc.blocks.new(name)
    for e in blocks[name]:
        assert e['t']=='LWPOLYLINE', e
        pts=[(*p,e['bul'][i]) for i,p in enumerate(e['pts'])]
        block.add_lwpolyline(pts,format='xyb',close=bool(e.get('70',0)&1),
                            dxfattribs={'layer':e.get('8','0')})
    for e in visible:
        doc.modelspace().add_blockref(name,(e['10'],e['20']),dxfattribs={
            'layer':e.get('8','0'),'rotation':e.get('50',0),
            'xscale':e.get('41',1),'yscale':e.get('42',1),
            'extrusion':(e.get('210',0),e.get('220',0),e.get('230',1))})
    with tempfile.TemporaryDirectory(prefix='h-sidewall-audit-') as scratch:
        path=Path(scratch)/'source_symbols.dxf';doc.saveas(path)
        bundle=A.parse_dxf_bundle(path)
        cats={ly['name']:ly['auto_category'] for ly in bundle.layers}
        heads=A.detect_heads(bundle.entities,cats)
    b=Board(world,dict(DEFAULT_KNOBS));b.mat_done=True
    saved=next(Path('cad_project_editor_g/docs/import/0단계_새찍기').glob('MF20*찍은스펙.json'))
    spec=json.loads(saved.read_text(encoding='utf8'))
    b.mat=[tuple(p) for p in spec['material_picks']]
    ps=PickSession(world,'audit-only',b.kn,b);ps.mode='헤드';ps.armed=True
    cands=[dict(x=h.pos[0],y=h.pos[1],conf=h.confidence,kind=h.kind) for h in heads]
    first=adopt_heads(ps,select_heads(cands,conf_min=.75))
    count1=len(b.highlight_geom()['head_circles'])
    again=adopt_heads(ps,select_heads(cands,conf_min=.75))
    count2=len(b.highlight_geom()['head_circles'])
    assert not first['skipped'] and not again['skipped']
    eligible=[c for c in symbol_world.circles if cats.get(c[0])=='HEAD']
    assert count1==count2==len(eligible)==len(heads)
    comparison=[]
    for h in heads:
        d=min(math.dist(h.pos,c[2:4]) for c in symbol_world.circles)
        assert d<1e-5 and h.confidence>=.75
        comparison.append(dict(x=h.pos[0],y=h.pos[1],confidence=h.confidence,
                               error_mm=d,evidence=h.kind))
    # Check source insertion, triangle interior and disk centre in the real
    # (not isolated) world: all three must resolve to the same head signature.
    target=(169864.227508406,2992796.888741465)
    signatures=[]
    for x,y in (target,(169650.,target[1]),(169772.248258769,target[1])):
        test=Board(world,dict(DEFAULT_KNOBS));test.mat_done=True;test.mat=b.mat[:]
        rep=test.apply_click('헤드',x,y,max_d=.5)
        assert rep is not None
        signatures.append(test.heads)
    assert signatures[0]==signatures[1]==signatures[2]
    report=dict(source=str(source),source_bytes=source.stat().st_size,
        insert_count=len(inserts),visible_count=len(visible),
        candidate_count=len(heads),selected_count=count1,repeated_selected_count=count2,
        non_head_layer_symbols=len(symbol_world.circles)-len(eligible),
        comparisons=comparison,manual_same_signature=True,
        elapsed_seconds=round(time.perf_counter()-start,2))
    out=Path('outputs/h_sidewall_head_review');out.mkdir(exist_ok=True,parents=True)
    (out/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__=='__main__':
    main()

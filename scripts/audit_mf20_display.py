"""Read-only MF20 source audit; write only a separate display review artifact."""
from pathlib import Path
from collections import Counter
import json
import time

from routes.module_f.common import _boot
from routes.module_f.world import _world_payload


def main() -> None:
    """Compare raw mirrored entities and corrected world without opening a session."""
    _boot()
    from services.cad_import.pipeline.stage1 import read_dxf, explode
    source=next(Path('data/uploads').glob('MF20-001*.dxf'))
    started=time.perf_counter()
    layers, entities, blocks=read_dxf(source)
    mirrored=Counter(e['t'] for e in [*entities,*(e for group in blocks.values() for e in group)]
                     if e.get('230',1)<0)
    world,_=explode(layers,entities,blocks)
    payload=_world_payload(world,source_display=True)
    region=(195000,3018000,207000,3031000)
    target=[a for a in world.arcs if region[0]<a[2]<region[2] and region[1]<a[3]<region[3]]
    report=dict(source=str(source),bytes=source.stat().st_size,mirrored=dict(mirrored),
                counts=payload['counts'],target_arcs=len(target),
                target_texts=[t for t in payload['texts'] if region[0]<t['x']<region[2]
                              and region[1]<t['y']<region[3]],
                elapsed=round(time.perf_counter()-started,2))
    output=Path('outputs/mf20_display_review');output.mkdir(exist_ok=True,parents=True)
    (output/'world.json').write_text(json.dumps(payload,ensure_ascii=False),encoding='utf-8')
    (output/'report.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__=='__main__':
    main()

"""Read actual floor/elevation annotations and DXF dimensions, not drawing spacing."""
from __future__ import annotations

import json
from pathlib import Path
import re
import sys

import ezdxf

ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    """Collect inspectable drawing evidence without changing the source DXF."""
    sys.stdout.reconfigure(encoding='utf-8')
    folder=ROOT/'routes'/'제출용[최종]'
    result={}
    for path in sorted(folder.glob('*대명동*.dxf')):
        doc=ezdxf.readfile(path)
        text, dimensions=[],[]
        def walk(entities, depth=0):
            for entity in entities:
                if entity.dxftype()=='INSERT' and depth<16:
                    yield from walk(entity.virtual_entities(),depth+1)
                else:
                    yield entity
        for entity in walk(doc.modelspace()):
            if entity.dxftype() in ('TEXT','MTEXT','ATTRIB'):
                content=entity.plain_text() if entity.dxftype()=='MTEXT' else entity.dxf.text
                text.append(dict(text=content,xy=list(entity.dxf.insert)[:2],layer=entity.dxf.layer))
            elif entity.dxftype()=='DIMENSION':
                dimensions.append(dict(text=entity.dxf.get('text'),
                                       measurement=str(entity.get_measurement()),
                                       layer=entity.dxf.layer))
        result[path.name]=dict(text=text,dimensions=dimensions)
        print(path.name,'texts',len(text),'dimensions',len(dimensions))
        selected=[r for r in text if re.search(r'(?:층고|층\s*고|[ESF]\.?L\.?\s*[:=+\-]|[ESF]L[+\-]|표고|LEVEL|HEIGHT|2[,.]900|3[,.]000)',r['text'],re.I)]
        print(json.dumps(selected[:100],ensure_ascii=False))
    out=ROOT/'data/merge_boundary_review_20260920/floor_evidence.json'
    out.write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')


if __name__=='__main__':
    main()

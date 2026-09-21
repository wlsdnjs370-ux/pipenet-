"""Read-only inventory of candidate head blocks and native DXF legend text."""
from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path

import ezdxf


def inventory(source: Path) -> dict:
    """Record inspectable geometry and text, without guessing a block's meaning."""
    doc = ezdxf.readfile(source)
    msp = doc.modelspace()
    texts = []
    for e in msp.query("TEXT MTEXT"):
        value = e.plain_text()
        if any(word in value.lower() for word in ("헤드", "sprinkler", "상향", "하향", "측벽", "head")):
            texts.append({"text": value, "layer": e.dxf.layer, "position": list(e.dxf.insert)})
    blocks = []
    instances = Counter(e.dxf.name for e in msp.query("INSERT"))
    for block in doc.blocks:
        ents = list(block)
        if not 0 < len(ents) <= 40 or not any(e.dxftype() in ("CIRCLE", "SOLID") for e in ents):
            continue
        geometry = []
        for e in ents:
            row = {"type": e.dxftype(), "layer": e.dxf.layer, "handle": e.dxf.handle}
            if e.dxftype() in ("CIRCLE", "ARC"):
                row.update(center=list(e.dxf.center), radius=e.dxf.radius)
                if e.dxftype() == "ARC":
                    row.update(start=e.dxf.start_angle, end=e.dxf.end_angle)
            elif e.dxftype() == "LINE":
                row.update(start=list(e.dxf.start), end=list(e.dxf.end))
            elif e.dxftype() == "SOLID":
                row["vertices"] = [list(p) for p in e.vertices()]
            elif e.dxftype() == "LWPOLYLINE":
                row.update(points=[list(p) for p in e.get_points()], closed=e.closed)
            elif e.dxftype() == "HATCH":
                row.update(solid_fill=e.dxf.solid_fill, pattern=e.dxf.pattern_name)
            geometry.append(row)
        blocks.append({"name": block.name, "modelspace_instances": instances[block.name], "geometry": geometry})
    return {"source": str(source), "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "entity_count": len(msp), "texts": texts, "blocks": blocks,
            "circle_layers": dict(Counter(e.dxf.layer for e in msp.query("CIRCLE")))}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("sources", nargs="+", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    rows = [inventory(path) for path in args.sources]
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    for row in rows:
        print(Path(row["source"]).name, row["entity_count"], "entities;", len(row["texts"]),
              "legend texts;", len(row["blocks"]), "small circle/solid blocks")

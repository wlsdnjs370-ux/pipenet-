"""Extract supplied PPTX reference text/images locally, without OCR or editing it."""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET
from zipfile import ZipFile

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


def extract(source: Path, output: Path) -> dict:
    """Preserve the slide 26–36 images and their exact text/crop provenance."""
    output.mkdir(parents=True, exist_ok=True)
    manifest = {"source": source.name, "sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
                "slides": []}
    with ZipFile(source) as archive:
        for number in range(26, 37):
            slide = ET.fromstring(archive.read(f"ppt/slides/slide{number}.xml"))
            rels = ET.fromstring(archive.read(f"ppt/slides/_rels/slide{number}.xml.rels"))
            images = {r.get("Id"): "ppt/media/" + PurePosixPath(r.get("Target", "")).name
                      for r in rels if r.get("Type", "").endswith("/image")}
            row = {"slide": number, "text": [t.text for t in slide.findall(".//a:t", NS)],
                   "pictures": [], "native_shapes": []}
            for shape in slide.findall(".//p:sp", NS):
                geometry = shape.find("p:spPr/a:prstGeom", NS)
                transform = shape.find("p:spPr/a:xfrm", NS)
                identity = shape.find("p:nvSpPr/p:cNvPr", NS)
                if geometry is not None:
                    row["native_shapes"].append({"id": identity.get("id") if identity is not None else None,
                        "shape": geometry.get("prst"),
                        "transform": ET.tostring(transform, encoding="unicode") if transform is not None else None})
            for pic in slide.findall(".//p:pic", NS):
                blip = pic.find("p:blipFill/a:blip", NS)
                if blip is None:
                    continue
                name = images.get(blip.get("{" + NS["r"] + "}embed"))
                if not name:
                    continue
                (output / PurePosixPath(name).name).write_bytes(archive.read(name))
                crop = pic.find("p:blipFill/a:srcRect", NS)
                xfrm = pic.find("p:spPr/a:xfrm", NS)
                row["pictures"].append({"media": PurePosixPath(name).name,
                    "crop_percent_1000": dict(crop.attrib) if crop is not None else {},
                    "transform": ET.tostring(xfrm, encoding="unicode") if xfrm is not None else None})
            manifest["slides"].append(row)
    (output / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    result = extract(args.source, args.output)
    print("Extracted", len(result["slides"]), "reference slides into", args.output)

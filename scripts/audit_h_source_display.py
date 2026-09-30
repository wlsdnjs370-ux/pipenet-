"""Compare H native display to unchanged extraction; never touch a live session."""
from __future__ import annotations

import argparse
import hashlib
import json
import logging
from pathlib import Path
import time

from routes.module_f.common import _boot
from routes.module_f import sub_fastread as sf
from routes.module_f.subdrawing import layer_colors
from routes.module_h_subdrawing import display_payload


def main() -> None:
    """Write review data and timings for a supplied DXF, without editing it."""
    parser = argparse.ArgumentParser(__doc__)
    parser.add_argument("drawing", type=Path)
    parser.add_argument("--output", type=Path, default=Path("outputs/h_source_display_review"))
    args = parser.parse_args()
    logging.getLogger("ezdxf").setLevel(logging.ERROR)
    _boot()
    before = hashlib.sha256(args.drawing.read_bytes()).hexdigest()
    t = time.perf_counter()
    old = sf.read_view(args.drawing)
    old_s = time.perf_counter()-t
    t = time.perf_counter()
    new = sf.read_view(args.drawing, source_display=True)
    new_s = time.perf_counter()-t
    old_fields = {k:v for k,v in new["parsed"].items() if k != "source_display"}
    assert old_fields == old["parsed"], "Calculation parser outputs changed"
    t = time.perf_counter()
    world = display_payload(new["entities"], layer_colors(new["parsed"]), args.drawing,
                            source=new["parsed"]["source_display"])
    payload_s = time.perf_counter()-t
    t = time.perf_counter()
    warm = sf.read_view(args.drawing, source_display=True)
    warm_s = time.perf_counter()-t
    assert warm["parsed"] == new["parsed"]
    assert hashlib.sha256(args.drawing.read_bytes()).hexdigest() == before
    report = dict(source=str(args.drawing.resolve()), source_sha256=before,
                  source_unchanged=True, extraction_identical=True,
                  entities=len(new["entities"]), counts=world["counts"],
                  hidden_bundles=sum(b["source_hidden"] for b in world["bundles"]),
                  bundles=len(world["bundles"]), native_bounds=world["bounds"],
                  saved_cad_view=world.get("source_view_bounds"),
                  full_bounds=world["full_bounds"], unsupported=world["source_unsupported"],
                  timings=dict(legacy_geometry_only_s=round(old_s,3),
                               shared_read_s=round(new_s,3), first_mode=new["how"],
                               payload_s=round(payload_s,3), cache_s=round(warm_s,3),
                               warm_mode=warm["how"]))
    args.output.mkdir(parents=True, exist_ok=True)
    (args.output/"world.json").write_text(json.dumps(world,ensure_ascii=False),encoding="utf-8")
    (args.output/"report.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
    print(json.dumps(report,ensure_ascii=False,indent=2))


if __name__ == "__main__":
    main()

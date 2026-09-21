# -*- coding: utf-8 -*-
"""G9 이전 방출기로 같은 도면을 뽑아 위상 수치를 찍는다. 되돌린 상태에서만 유효."""
import json, os, sys
from pathlib import Path
ROOT = Path.cwd()
for p in (str(ROOT), str(ROOT.parent)):
    if p not in sys.path:
        sys.path.insert(0, p)
KEY = "B1F 현장조사 소화설비 평면도"
from services.cad_import.design.bore import extract_dia_text_points
from services.cad_import.design.emit import emit_design_sdf
from services.cad_import.design.restrict import select_and_expand
from services.cad_import.design.tables import build_design_tables
from services.cad_import.edit.session import EditSession
from services.cad_import.pipeline import handoff, stage1 as s1
es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
payload = es.convert_payload()
srcs = payload.get("sources") or ()
sel = srcs[0].get("tag") if len(srcs) > 1 else None
got = select_and_expand(payload, es.board, k=30, selected_source=sel)
spec = ROOT / "docs" / "import" / "0단계_새찍기" / f"{KEY}_찍은스펙.json"
src = json.loads(spec.read_text(encoding="utf-8")).get("source_dxf")
world = handoff.load_world(KEY, src, s1.World) if src else None
texts = extract_dia_text_points(world.texts) if world else []
tbl = build_design_tables(got["kfp"], got["worst"], got["edge_ref"], texts,
                          board_pts=es.board.pts)
out = Path("tests/_out/g13_preG9.sdf")
emit_design_sdf(tbl, out)
import xml.etree.ElementTree as ET
r = ET.parse(str(out)).getroot()
lens = [float(p.get("length") or 0) for p in r.iter("Pipe")]
res = {"nodes": len(list(r.iter("Node"))), "pipes": len(list(r.iter("Pipe"))),
       "nozzles": len(list(r.iter("Nozzle"))), "length_sum": round(sum(lens), 3)}
Path("tests/_out/_g13_preG9.json").write_text(json.dumps(res, ensure_ascii=False),
                                              encoding="utf-8")
print("PRE-G9", json.dumps(res, ensure_ascii=False))

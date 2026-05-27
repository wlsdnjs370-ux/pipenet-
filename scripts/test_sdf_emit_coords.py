"""emit_sdf 좌표 정규화 + template 잔재 정리 검증.

기존 prototype_runs/ 의 XLSX 를 읽어 PipeTables 를 재구성하고 emit_sdf 로 새 SDF
를 만든 뒤, 좌표 범위와 Graphics/Libraries/Titles 정리 결과를 검증.
"""

from __future__ import annotations

import sys
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from remote30_prototype import PipeTables, emit_sdf

# 토이 데이터 — 좌표를 실제 DXF 범위(수십만 단위)로 둬서 normalize 검증
tables = PipeTables()
tables.nodes = [
    {"label": "1",  "x": 245000, "y": -240000, "elevation": 2.8, "io_node": "Input"},
    {"label": "2",  "x": 260000, "y": -240000, "elevation": 2.8, "io_node": "No"},
    {"label": "3",  "x": 260000, "y": -225000, "elevation": 2.8, "io_node": "No"},
    {"label": "4",  "x": 275000, "y": -225000, "elevation": 2.8, "io_node": "No"},
    {"label": "5",  "x": 275000, "y": -240000, "elevation": 2.8, "io_node": "No"},
]
tables.pipes = [
    {"label": "P1", "in": "1", "out": "2", "type": "K", "dia": 100, "length": 15, "elev": 0, "c": 120, "status": "normal", "group": "main"},
    {"label": "P2", "in": "2", "out": "3", "type": "K", "dia":  80, "length": 15, "elev": 0, "c": 120, "status": "normal", "group": "main"},
    {"label": "P3", "in": "3", "out": "4", "type": "K", "dia":  65, "length": 15, "elev": 0, "c": 120, "status": "normal", "group": "branch"},
    {"label": "P4", "in": "4", "out": "5", "type": "K", "dia":  50, "length": 15, "elev": 0, "c": 120, "status": "normal", "group": "branch"},
]
tables.nozzles = [
    {"label": "Nz1", "in": "5", "out": "@/N", "status": 1, "lib": "SP-HEAD",
     "flow_m3s": 0.001333, "flow_lmin": 80},
]

OUT = Path(__file__).resolve().parent.parent / "data" / "prototype_runs" / "test_sdf_emit" / "out.sdf"
OUT.parent.mkdir(parents=True, exist_ok=True)
emit_sdf(tables, OUT, project_title="emit-test-coords")
print(f"[written] {OUT}  ({OUT.stat().st_size} bytes)")

# ── 검증
tree = ET.parse(OUT)
root = tree.getroot()

# 좌표 범위
xs, ys = [], []
for n in root.iter("Node"):
    pos = n.find("Position")
    xs.append(float(pos.get("x"))); ys.append(float(pos.get("y")))
print(f"\n[Node coords] n={len(xs)}")
print(f"  x: {min(xs):.2f} .. {max(xs):.2f}  (range {max(xs)-min(xs):.2f})")
print(f"  y: {min(ys):.2f} .. {max(ys):.2f}  (range {max(ys)-min(ys):.2f})")

# Pipe bore/length
for p in root.iter("Pipe"):
    print(f"  Pipe {p.get('label')}: bore={p.get('bore')}, length={p.get('length')}, in={p.get('input')}, out={p.get('output')}")

# Graphics > Text-element 잔재 개수
te_count = sum(1 for _ in root.iter("Text-element"))
ul_count = sum(1 for _ in root.iter("User-lib"))
title_count = len(root.findall(".//Network-spray/Title"))
nd_count = len(root.findall(".//Network-spray/Network-description"))

print(f"\n[Template cleanup]")
print(f"  Text-element: {te_count}  (expect 0)")
print(f"  User-lib: {ul_count}  (expect 0)")
print(f"  Network-spray > Title: {title_count}  (expect 1)")
print(f"  Network-spray > Network-description: {nd_count}  (expect 0)")

# Graphics 메타 유지 확인
go = root.find(".//Grid-options")
ls = root.find(".//Link-schemes")
ns = root.find(".//Node-schemes")
print(f"\n[Graphics meta preserved]")
print(f"  Grid-options: {'YES' if go is not None else 'MISSING'}  grid={go.get('grid') if go is not None else '-'}")
print(f"  Link-schemes: {'YES' if ls is not None else 'MISSING'}")
print(f"  Node-schemes: {'YES' if ns is not None else 'MISSING'}")

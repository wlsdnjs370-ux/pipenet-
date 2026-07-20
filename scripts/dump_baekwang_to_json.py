"""배관망.dxf 를 inspect endpoint 와 동일한 JSON 으로 변환 후 표준출력에 요약 인쇄."""

from __future__ import annotations

import io
import json
import math
import sys
from collections import Counter
from pathlib import Path

import ezdxf

# Windows console UTF-8
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

sys.path.insert(0, r"C:\Users\admin\PycharmProjects\JupyterProject")
from sprinkler_remote30_extractor import Remote30Settings, layer_match

DXF = Path(r"C:\Users\admin\PycharmProjects\JupyterProject\대명동201동 단위세대_layer정리_배관망.dxf")
OUT = Path(r"C:\Users\admin\PycharmProjects\JupyterProject\data\gnn_outputs\배관망_inspect.json")
OUT.parent.mkdir(parents=True, exist_ok=True)

s = Remote30Settings()
doc = ezdxf.readfile(str(DXF))
msp = doc.modelspace()

doc_layer_info: dict[str, dict] = {}
hidden_layers: set[str] = set()
for ly in doc.layers:
    try:
        color = int(ly.dxf.color)
    except Exception:
        color = 7
    name = str(ly.dxf.name)
    is_off = bool(ly.is_off())
    is_frozen = bool(ly.is_frozen())
    doc_layer_info[name] = {
        "is_off": is_off,
        "is_frozen": is_frozen,
        "is_locked": bool(ly.is_locked()),
        "color": color,
    }
    if is_off or is_frozen or color < 0:
        hidden_layers.add(name)

entities: list[dict] = []
dropped_types: dict[str, int] = {}
bbox = [float("inf"), float("inf"), float("-inf"), float("-inf")]
MAX_INSERT_DEPTH = 10


def _upd(x: float, y: float) -> None:
    if x < bbox[0]:
        bbox[0] = x
    if y < bbox[1]:
        bbox[1] = y
    if x > bbox[2]:
        bbox[2] = x
    if y > bbox[3]:
        bbox[3] = y


def _render(e, layer_override=None, depth=0):
    etype = e.dxftype()
    own = getattr(e.dxf, "layer", "")
    if layer_override is not None and own in ("0", ""):
        layer = layer_override
    else:
        layer = own or (layer_override or "")
    if layer in hidden_layers:
        return
    if int(getattr(e.dxf, "invisible", 0) or 0) == 1:
        return
    try:
        if etype == "LINE":
            x1, y1, x2, y2 = float(e.dxf.start.x), float(e.dxf.start.y), float(e.dxf.end.x), float(e.dxf.end.y)
            entities.append({"t": "L", "l": layer, "p": [x1, y1, x2, y2]})
            _upd(x1, y1)
            _upd(x2, y2)
        elif etype == "ARC":
            cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
            r = float(e.dxf.radius)
            sa = float(e.dxf.start_angle)
            ea = float(e.dxf.end_angle)
            entities.append({"t": "A", "l": layer, "c": [cx, cy], "r": r, "a": [sa, ea]})
            _upd(cx - r, cy - r)
            _upd(cx + r, cy + r)
        elif etype == "CIRCLE":
            cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
            r = float(e.dxf.radius)
            entities.append({"t": "C", "l": layer, "c": [cx, cy], "r": r})
            _upd(cx - r, cy - r)
            _upd(cx + r, cy + r)
        elif etype == "LWPOLYLINE":
            pts = [[float(p[0]), float(p[1])] for p in e.get_points()]
            if pts:
                for x, y in pts:
                    _upd(x, y)
                entities.append({"t": "PL", "l": layer, "p": pts})
        elif etype == "POLYLINE":
            pts = [[float(v.dxf.location.x), float(v.dxf.location.y)] for v in e.vertices]
            if pts:
                for x, y in pts:
                    _upd(x, y)
                entities.append({"t": "PL", "l": layer, "p": pts})
        elif etype == "INSERT":
            x = float(e.dxf.insert.x)
            y = float(e.dxf.insert.y)
            if depth == 0:
                entities.append({"t": "I", "l": layer, "p": [x, y], "n": str(e.dxf.name)})
            _upd(x, y)
            if depth >= MAX_INSERT_DEPTH:
                dropped_types["INSERT(too deep)"] = dropped_types.get("INSERT(too deep)", 0) + 1
            else:
                try:
                    virtuals = list(e.virtual_entities())
                except Exception:
                    virtuals = []
                for v in virtuals:
                    _render(v, layer_override=layer, depth=depth + 1)
        elif etype == "TEXT":
            x = float(e.dxf.insert.x)
            y = float(e.dxf.insert.y)
            raw = str(e.dxf.text)[:60]
            entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
            _upd(x, y)
        elif etype in ("MTEXT", "ATTRIB", "ATTDEF"):
            x = float(e.dxf.insert.x)
            y = float(e.dxf.insert.y)
            raw = str(getattr(e, "text", "") or getattr(e.dxf, "text", ""))[:60]
            if raw:
                entities.append({"t": "T", "l": layer, "p": [x, y], "v": raw})
            _upd(x, y)
        elif etype == "SPLINE":
            try:
                pts = [[float(p[0]), float(p[1])] for p in e.flattening(1.0)]
            except Exception:
                pts = []
            if pts:
                for px, py in pts:
                    _upd(px, py)
                entities.append({"t": "PL", "l": layer, "p": pts})
        elif etype == "ELLIPSE":
            try:
                pts = [[float(p[0]), float(p[1])] for p in e.flattening(0.5)]
            except Exception:
                pts = []
            if pts:
                for px, py in pts:
                    _upd(px, py)
                entities.append({"t": "PL", "l": layer, "p": pts})
        elif etype == "HATCH":
            paths_out = []
            for path in e.paths:
                pts = []
                for v in getattr(path, "vertices", []) or []:
                    try:
                        pts.append([float(v[0]), float(v[1])])
                    except Exception:
                        continue
                if not pts:
                    for edge in getattr(path, "edges", []) or []:
                        et = type(edge).__name__
                        try:
                            if et == "LineEdge":
                                pts.append([float(edge.start[0]), float(edge.start[1])])
                                pts.append([float(edge.end[0]), float(edge.end[1])])
                            elif et == "ArcEdge":
                                cx = float(edge.center[0])
                                cy = float(edge.center[1])
                                rr = float(edge.radius)
                                sa = float(edge.start_angle)
                                ea2 = float(edge.end_angle)
                                if ea2 < sa:
                                    ea2 += 360.0
                                for k in range(9):
                                    ang = math.radians(sa + (ea2 - sa) * k / 8)
                                    pts.append([cx + rr * math.cos(ang), cy + rr * math.sin(ang)])
                        except Exception:
                            continue
                if pts:
                    paths_out.append(pts)
                    for x, y in pts:
                        _upd(x, y)
            if paths_out:
                biggest = max(paths_out, key=len)
                entities.append({"t": "H", "l": layer, "p": biggest})
            else:
                dropped_types["HATCH(no-geom)"] = dropped_types.get("HATCH(no-geom)", 0) + 1
        elif etype in ("SOLID", "3DFACE", "TRACE"):
            verts = []
            for attr in ("vtx0", "vtx1", "vtx2", "vtx3"):
                try:
                    v = getattr(e.dxf, attr)
                    verts.append([float(v.x), float(v.y)])
                except AttributeError:
                    break
            if len(verts) >= 2 and verts[-1] == verts[-2]:
                verts.pop()
            if len(verts) >= 3:
                for x, y in verts:
                    _upd(x, y)
                entities.append({"t": "S", "l": layer, "p": verts})
        elif etype == "DIMENSION":
            try:
                for v in e.virtual_entities():
                    _render(v, layer_override=layer)
            except Exception:
                pass
        else:
            dropped_types[etype] = dropped_types.get(etype, 0) + 1
    except Exception:
        dropped_types[etype] = dropped_types.get(etype, 0) + 1


for e in msp:
    _render(e)

layer_counts = Counter(en["l"] for en in entities)
layer_type_counts: dict[str, Counter] = {}
for en in entities:
    layer_type_counts.setdefault(en["l"], Counter())[en["t"]] += 1
layer_list = []
for name in sorted(layer_counts.keys()):
    # 콘텐츠 우선(HEAD/PIPE/TEXT) → ARCH 는 최후. _layer_category(대조 서버.py)와 동일 순서.
    if layer_match(name, s.exclude_layer_keywords):
        cat = "EXCLUDE"
    elif layer_match(name, s.head_layer_keywords):
        cat = "HEAD"
    elif layer_match(name, s.pipe_layer_keywords):
        cat = "PIPE"
    elif layer_match(name, s.text_layer_keywords):
        cat = "TEXT"
    elif layer_match(name, s.arch_layer_keywords):
        cat = "ARCH"
    else:
        cat = "OTHER"
    info = doc_layer_info.get(name, {})
    color = int(info.get("color", 7))
    visible = (not info.get("is_off", False)) and (not info.get("is_frozen", False)) and (color >= 0)
    layer_list.append({
        "name": name,
        "count": layer_counts[name],
        "types": dict(layer_type_counts[name]),
        "auto_category": cat,
        "is_off": bool(info.get("is_off", False)),
        "is_frozen": bool(info.get("is_frozen", False)),
        "color": color,
        "visible": visible,
    })

if bbox[0] == float("inf"):
    bbox = [0, 0, 1, 1]
result = {
    "ok": True,
    "dxf_filename": DXF.name,
    "bbox": {"x_min": bbox[0], "y_min": bbox[1], "x_max": bbox[2], "y_max": bbox[3]},
    "layers": layer_list,
    "entities": entities,
    "counts": {"total_entities": len(entities), "layers": len(layer_counts)},
    "dropped_types": dropped_types,
}

OUT.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")

print("=== 변환 완료 ===")
print(f"  source : {DXF.name}")
print(f"  output : {OUT}")
print(f"  파일크기: {OUT.stat().st_size / 1024:.1f} KB")
print(f"  entity 수: {len(entities)}")
print(f"  layer 수 : {len(layer_counts)}")
print(f"  bbox    : x={bbox[0]:.2f}~{bbox[2]:.2f}  y={bbox[1]:.2f}~{bbox[3]:.2f}")
print(f"  dropped : {dropped_types if dropped_types else '{}'}")
print()
print("--- 레이어 목록 ---")
for ly in sorted(layer_list, key=lambda x: -x["count"]):
    types_str = ", ".join(f"{k}={v}" for k, v in sorted(ly["types"].items()))
    print(f"  count={ly['count']:5d}  cat={ly['auto_category']:7s}  color={ly['color']:4d}  {types_str:30s}  |{ly['name']}|")
print()
print("--- entity 타입별 카운트 ---")
type_cnt = Counter(en["t"] for en in entities)
for t, c in type_cnt.most_common():
    print(f"  {t}: {c}")
print()
print("--- entity 샘플 (처음 8개) ---")
for en in entities[:8]:
    print(f"  {json.dumps(en, ensure_ascii=False)}")
print("...")
print("--- entity 샘플 (마지막 5개) ---")
for en in entities[-5:]:
    print(f"  {json.dumps(en, ensure_ascii=False)}")

"""Faithful modelspace display geometry, isolated from hydraulic graph inputs.

Coordinates stay in source DXF units (millimetres in this application). Native
virtual entities apply nested INSERT bases, OCS, rotation and unequal scales.
Curves that are not XY circles are tessellated for display only, at 0.25 mm.
Unsupported entities are counted, never silently used as connection evidence.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass, field
import math
from typing import Any, Iterable

from ezdxf import path as dxf_path
from ezdxf.protocols import SupportsVirtualEntities, virtual_entities

from .dxf_text import display_text


@dataclass
class SourcePrimitive:
    """One source-coloured display primitive; no calculated pipe identity."""

    kind: str
    layer: str
    color: int | str
    xy: list[float]
    hidden: bool = False


@dataclass
class SourceDisplay:
    """Serializable source snapshot cached together with the extraction data."""

    primitives: list[SourcePrimitive] = field(default_factory=list)
    texts: list[dict] = field(default_factory=list)
    unsupported: dict[str, int] = field(default_factory=dict)
    saved_view: dict | None = None


def saved_top_view(document: Any) -> dict | None:
    """Restore an explicit, untwisted modelspace top view; never guess a crop.

    DXF VPORT center is in display coordinates relative to the view target, not
    an absolute WCS location. Perspective, twisted and tiled views are deferred.
    """
    ports = document.viewports.get_config("*Active")
    if len(ports) != 1 or document.header.get("$TILEMODE", 1) != 1:
        return None
    v = ports[0].dxf
    direction, target, center = v.direction, v.target, v.center
    if direction.z <= 0 or math.hypot(direction.x, direction.y) > direction.z * 1e-9:
        return None
    if abs(v.view_twist % 360) > 1e-9 or v.view_mode & 1:
        return None
    # An untouched default viewport is not an authored camera.
    if max(abs(center.x), abs(center.y), abs(target.x), abs(target.y)) < 1e-9:
        return None
    cx, cy = center.x + target.x, center.y + target.y
    height, width = v.height, v.height * v.aspect_ratio
    if min(width, height) <= 0 or not all(math.isfinite(x) for x in (cx, cy, width, height)):
        return None
    return dict(minx=cx-width/2, maxx=cx+width/2, miny=cy-height/2, maxy=cy+height/2)


def source_display(document: Any, *, curve_tolerance_mm: float = 0.25) -> dict:
    """Build a display snapshot from an already loaded document; never reread it."""
    if not math.isfinite(curve_tolerance_mm) or curve_tolerance_mm <= 0:
        raise ValueError("곡선 표시 오차는 양수여야 합니다.")
    result = SourceDisplay()
    skipped: Counter[str] = Counter()

    def emit(kind: str, layer: str, color: int | str, hidden: bool, xy: Iterable) -> None:
        values = [float(v) for v in xy]
        if all(math.isfinite(v) for v in values):
            result.primitives.append(SourcePrimitive(kind, layer, color, values, hidden))
        else:
            skipped["NON_FINITE"] += 1

    def poly(points: Iterable, layer: str, color: int | str, hidden: bool) -> None:
        pts = list(points)
        for a, b in zip(pts, pts[1:]):
            if a != b:
                emit("segs", layer, color, hidden, (a.x, a.y, b.x, b.y))

    def visit(entities: Iterable, layer: str = "0", color: int | str = 7,
              hidden: bool = False, depth: int = 0) -> None:
        if depth > 12:
            raise ValueError("도면 블록 중첩이 12단계를 초과합니다.")
        for ent in entities:
            kind = ent.dxftype()
            own = str(ent.dxf.get("layer", "0"))
            own = layer if own == "0" else own
            record = document.layers.get(own) if own in document.layers else None
            off = hidden or bool(record and (record.is_off() or record.is_frozen()))
            attrs = ent.dxf.all_existing_dxf_attribs()
            if attrs.get("invisible", 0):
                continue
            aci = int(attrs.get("color", 256))
            ink = color if aci == 0 else abs(record.color) if aci == 256 and record else 7 if aci == 256 else aci
            if "true_color" in attrs:
                ink = f"#{int(attrs['true_color']):06x}"
            try:
                if kind == "INSERT":
                    def unsupported(entity: Any, reason: str) -> None:
                        skipped[f"INSERT_{entity.dxftype()}"] += 1
                    for insert in ent.multi_insert() if ent.mcount > 1 else (ent,):
                        visit(insert.attribs, own, ink, off, depth + 1)
                        visit(insert.virtual_entities(skipped_entity_callback=unsupported),
                              own, ink, off, depth + 1)
                elif kind == "LINE":
                    a, b = ent.dxf.start, ent.dxf.end
                    emit("segs", own, ink, off, (a.x, a.y, b.x, b.y))
                elif kind in ("ARC", "CIRCLE"):
                    normal = ent.dxf.extrusion
                    if abs(normal.x) + abs(normal.y) > 1e-10:
                        poly(dxf_path.make_path(ent).flattening(curve_tolerance_mm), own, ink, off)
                        continue
                    c = ent.ocs().to_wcs(ent.dxf.center)
                    r = float(ent.dxf.radius)
                    if kind == "CIRCLE":
                        emit("circles", own, ink, off, (c.x, c.y, r))
                    else:
                        a = math.radians(ent.dxf.start_angle)
                        u = ent.ocs().to_wcs((math.cos(a), math.sin(a), 0))
                        start = math.degrees(math.atan2(u.y, u.x)) % 360
                        sweep = (ent.dxf.end_angle - ent.dxf.start_angle) % 360
                        if normal.z < 0:
                            sweep = -sweep
                        emit("arcs", own, ink, off, (c.x, c.y, r, start, sweep))
                elif kind in ("LWPOLYLINE", "POLYLINE"):
                    # virtual_entities preserves BOTH bulges and the closing edge.
                    visit(ent.virtual_entities(), own, ink, off, depth + 1)
                elif kind in ("SPLINE", "ELLIPSE"):
                    poly(ent.flattening(curve_tolerance_mm), own, ink, off)
                elif kind == "HATCH":
                    for boundary in dxf_path.from_hatch(ent):
                        poly(boundary.flattening(curve_tolerance_mm), own, ink, off)
                elif kind in ("SOLID", "TRACE", "3DFACE"):
                    pts = ent.wcs_vertices(close=True)
                    poly(pts, own, ink, off)
                elif kind in ("TEXT", "MTEXT", "ATTRIB", "ATTDEF"):
                    text = display_text(ent, own, ink, off)
                    if text is not None:
                        result.texts.append(text)
                elif isinstance(ent, SupportsVirtualEntities):
                    visit(virtual_entities(ent), own, ink, off, depth + 1)
                elif kind not in ("POINT", "VIEWPORT"):
                    skipped[kind] += 1
            except (ValueError, TypeError, AttributeError, NotImplementedError, ZeroDivisionError):
                skipped[f"{kind}_ERROR"] += 1

    visit(document.modelspace())
    result.unsupported = dict(skipped)
    result.saved_view = saved_top_view(document)
    return asdict(result)

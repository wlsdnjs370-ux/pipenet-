"""Display-only native DXF labels; never inputs to network construction."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable
import math


@dataclass(frozen=True)
class DisplayText:
    """A plain-text label in CAD millimetres with its original orientation."""

    x: float
    y: float
    height: float
    text: str
    rotation: float
    layer: str
    color: int | str
    align: str = "left"
    vertical: str = "alphabetic"
    hidden: bool = False
    width_factor: float = 1.0


def display_text(ent: Any, layer: str, color: int | str, hidden: bool = False) -> dict | None:
    """Serialize one already transformed native label, without visiting blocks."""
    kind = ent.dxftype()
    if kind in ("ATTRIB", "ATTDEF") and (ent.is_invisible or (kind == "ATTDEF" and not ent.is_const)):
        return None
    text = ent.plain_text()
    if not text.strip():
        return None
    align, vertical = "left", "alphabetic"
    if kind == "MTEXT":
        p = ent.dxf.insert
        height, rotation = float(ent.dxf.char_height), ent.get_rotation()
        attachment = int(ent.dxf.get("attachment_point", 1)) - 1
        align = ("left", "center", "right")[attachment % 3]
        vertical = ("top", "middle", "bottom")[attachment // 3]
    else:
        placement, p, _ = ent.get_placement()
        p = ent.ocs().to_wcs(p)
        name = placement.name
        align = "center" if "CENTER" in name or name == "MIDDLE" else "right" if "RIGHT" in name else "left"
        vertical = "top" if name.startswith("TOP") else "middle" if name.startswith("MIDDLE") else "bottom" if name.startswith("BOTTOM") else "alphabetic"
        height = float(ent.dxf.height)
        angle = math.radians(ent.dxf.get("rotation", 0))
        direction = ent.ocs().to_wcs((math.cos(angle), math.sin(angle), 0))
        rotation = math.degrees(math.atan2(direction.y, direction.x))
    if height > 0 and all(math.isfinite(v) for v in (p.x, p.y, height, rotation)):
        width = float(ent.dxf.get("width", 1)) if kind != "MTEXT" else 1.0
        return asdict(DisplayText(p.x, p.y, height, text, rotation,
                                  layer, color or 7, align, vertical, hidden, width))
    return None


def read_display_texts(path: str | Path, *, document: Any = None) -> list[dict]:
    """Read visible text/attributes, including nested and dimension blocks.

    Off/frozen layers remain available for the user's layer visibility controls.
    Invisible attribute definitions are not displayed. XREF contents require a
    bound DXF, as in the existing geometry reader. No OCR or hydraulic inference.
    """
    import ezdxf
    from ezdxf.protocols import SupportsVirtualEntities, virtual_entities

    doc = document if document is not None else ezdxf.readfile(path)
    result: list[dict] = []

    def visit(entities: Iterable, layer: str = "0", color: int = 7, depth: int = 0,
              hidden: bool = False) -> None:
        if depth > 12:
            raise ValueError("도면 문자 블록이 12단계를 초과합니다. 블록을 정리한 DXF를 사용하세요.")
        for ent in entities:
            own_layer = str(ent.dxf.get("layer", "0"))
            own_layer = layer if own_layer == "0" else own_layer
            own_color = int(ent.dxf.get("color", 256))
            if own_color == 0:
                own_color = color
            elif own_color == 256:
                own_color = abs(doc.layers.get(own_layer).color) if own_layer in doc.layers else 7
            kind = ent.dxftype()
            layer_record = doc.layers.get(own_layer) if own_layer in doc.layers else None
            own_hidden = hidden or bool(layer_record and
                                        (layer_record.is_off() or layer_record.is_frozen()))
            if int(ent.dxf.get("invisible", 0)):
                continue
            if kind in ("ATTRIB", "ATTDEF") and ent.is_invisible:
                continue
            if kind == "INSERT":
                for insert in ent.multi_insert() if ent.mcount > 1 else (ent,):
                    # Attached attributes are already in the parent's coordinate system.
                    visit(insert.attribs, own_layer, own_color, depth + 1, own_hidden)
                    visit(insert.virtual_entities(), own_layer, own_color, depth + 1, own_hidden)
            elif kind in ("TEXT", "MTEXT", "ATTRIB", "ATTDEF"):
                value = display_text(ent, own_layer, own_color, own_hidden)
                if value is not None:
                    result.append(value)
            elif isinstance(ent, SupportsVirtualEntities):
                visit(virtual_entities(ent), own_layer, own_color, depth + 1, own_hidden)

    visit(doc.modelspace())
    return result

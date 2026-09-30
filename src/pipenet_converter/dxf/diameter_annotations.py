"""Enrich already-mapped diameter texts with native DXF rotation; no OCR.

The caller's selected/mapped text list is authoritative. This reader cannot add
unrelated layers, quantities, or new annotation candidates to that list.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
import math
from pathlib import Path
import re
from collections.abc import Sequence

from ..graph.diameter_inference import DiameterAnnotation


@dataclass(frozen=True)
class NativeText:
    """A decoded native text, in world millimetres and world rotation degrees."""
    x: float
    y: float
    height: float
    text: str
    rotation: float
    layer: str
    entity_handle: str = ""
    native_anchor_mm: tuple[float, float] | None = None
    leader_xy_mm: tuple[float, float] | None = None
    leader_ambiguous: bool = False


def read_native_texts(path: str | Path, *, extended: bool = False) -> list[NativeText]:
    """Decode text/INSERT only; preserve ezdxf's native OCS and block transforms."""
    from .annotation_document import read_annotation_document
    result = []
    doc = read_annotation_document(path, extended=extended)
    kinds = {'TEXT','MTEXT','ATTRIB','ATTDEF'} if extended else {'TEXT','MTEXT'}
    leaders: dict[str, set[tuple[float,float]]] = {}
    # Do not expand thousands of car/head/architectural symbols without text.
    text_blocks = {b.name for b in doc.blocks if any(e.dxftype() in kinds or
        (extended and e.dxftype()=='INSERT' and e.attribs) for e in b)}
    references = {b.name:{e.dxf.name for e in b if e.dxftype()=='INSERT'} for b in doc.blocks}
    while True:
        parents = {name for name,children in references.items() if children & text_blocks}
        if parents <= text_blocks:
            break
        text_blocks.update(parents)

    def visit(entities, inherited_layer: str = "0", depth: int = 0, prefix: str = '') -> None:
        if extended and depth > 12:
            raise ValueError('관경 문자 블록이 12단계를 초과합니다. DXF 블록을 정리하세요.')
        for ent in entities:
            layer = str(ent.dxf.get('layer','0'))
            if layer == '0':
                layer = inherited_layer
            if layer in doc.layers and (doc.layers.get(layer).is_off() or doc.layers.get(layer).is_frozen()):
                continue
            kind = ent.dxftype()
            origin = ent
            while origin.source_of_copy is not None:
                origin = origin.source_of_copy
            handle = str(origin.dxf.get('handle','') or '')
            identity = prefix + handle
            if extended and (int(ent.dxf.get('invisible',0)) or (kind in ('ATTRIB','ATTDEF') and ent.is_invisible)):
                continue
            if kind == 'INSERT' and (extended or depth < 2) and (ent.dxf.name in text_blocks or (extended and ent.attribs)):
                for insert in ent.multi_insert() if ent.mcount > 1 else (ent,):
                    # Insertion path plus transformed location disambiguates identical
                    # source handles reused in multiple INSERT/MINSERT instances.
                    child_prefix = f'{identity}@{insert.dxf.insert.x:g},{insert.dxf.insert.y:g}/'
                    if extended:
                        visit(insert.attribs,layer,depth+1,child_prefix)
                    visit(insert.virtual_entities(),layer,depth+1,child_prefix)
            elif extended and kind == 'LEADER':
                target = str(ent.dxf.get('annotation_handle','') or '')
                if target and target != '0' and ent.vertices and ent.dxf.get('has_arrowhead',1):
                    point = ent.vertices[0]
                    leaders.setdefault(prefix+target,set()).add((float(point[0]),float(point[1])))
            elif kind in kinds:
                if kind == 'ATTDEF' and not ent.is_const:
                    continue
                if kind in ('TEXT','ATTRIB','ATTDEF'):
                    point = ent.ocs().to_wcs(ent.dxf.insert)
                    angle = math.radians(ent.dxf.get('rotation',0))
                    direction = ent.ocs().to_wcs((math.cos(angle),math.sin(angle),0))
                    rotation = math.degrees(math.atan2(direction.y,direction.x)) % 360
                    height = float(ent.dxf.height)
                    text = ent.plain_text()
                else:
                    point = ent.dxf.insert
                    rotation = ent.get_rotation() % 360
                    height = float(ent.dxf.char_height)
                    text = ent.plain_text()
                anchor = None
                if extended:
                    from ezdxf.bbox import extents
                    bounds = extents([ent],fast=True)
                    if bounds.has_data:
                        anchor = (bounds.center.x,bounds.center.y)
                result.append(NativeText(point.x,point.y,height,text,rotation,layer,
                                         identity if extended else str(ent.dxf.get('handle', '') or ''),anchor))
    visit(doc.modelspace())
    if extended:
        result = [replace(t, leader_xy_mm=next(iter(leaders[t.entity_handle])) if len(leaders.get(t.entity_handle,()))==1 else None,
                          leader_ambiguous=len(leaders.get(t.entity_handle,()))>1) for t in result]
    return result


def enrich_diameter_annotations(path: str | Path, points: Sequence[Sequence[float]], *,
                                native_texts: Sequence[NativeText] | None = None,
                                extended: bool = False) -> list[DiameterAnnotation]:
    """Match mapped points to metadata; unmatched or ambiguous rotation stays None."""
    result = [DiameterAnnotation(str(i),float(p[0]),float(p[1]),int(p[2]),source_ambiguous=extended) for i,p in enumerate(points)]
    if not result:
        return result
    bins: dict[tuple[int,int],list[int]] = {}
    for i,t in enumerate(result):
        bins.setdefault((math.floor(t.x),math.floor(t.y)),[]).append(i)
    found: dict[int,list[DiameterAnnotation]] = {}
    for text in native_texts if native_texts is not None else read_native_texts(path,extended=extended):
        nums = {int(n) for n in re.findall(r'(?<!\d)\d{2,3}(?!\d)',text.text)}
        for dx in (-1,0,1):
            for dy in (-1,0,1):
                for i in bins.get((math.floor(text.x)+dx,math.floor(text.y)+dy),[]):
                    old=result[i]
                    if old.nominal_mm in nums and math.hypot(old.x-text.x,old.y-text.y)<=1.0:
                        found.setdefault(i,[]).append(replace(old,rotation_deg=text.rotation,height_mm=text.height,
                            layer=text.layer,raw_text=text.text,entity_handle=text.entity_handle,
                            native_anchor_mm=text.native_anchor_mm if extended else None,
                            leader_xy_mm=text.leader_xy_mm if extended else None,
                            leader_ambiguous=text.leader_ambiguous if extended else False,source_ambiguous=False))
    for i,options in found.items():
        # Coincident differently-oriented text entities are not interchangeable.
        if len({round(o.rotation_deg,3) for o in options}) == 1:
            result[i] = options[0]
            if len({(o.raw_text,o.layer,o.entity_handle) for o in options}) > 1:
                # Rotation can agree while the original entity is ambiguous.
                result[i] = replace(result[i],raw_text='',entity_handle='',layer='',
                                    native_anchor_mm=None,leader_xy_mm=None,source_ambiguous=extended)
    return result

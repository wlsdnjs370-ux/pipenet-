"""Map selected, localized DXF geometry to the company symbol classifier.

No entities are added to the pipe graph. The existing candidate extraction stays
authoritative for position and scope. This adapter supplies source-geometry and
actual-fill evidence, then returns a classification and audit reasons.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict, replace
import math
import re
from pathlib import Path
from typing import Any

from .head_symbols import Arc, Detection, HeadCandidate, Segment, SymbolProfile, detect_head
from ..graph.regions import zone_contains


def _projection(p, a, b):
    dx, dy = b[0]-a[0], b[1]-a[1]
    length = math.hypot(dx, dy)
    if not length:
        return a, (1., 0.)
    t = max(0., min(1., ((p[0]-a[0])*dx+(p[1]-a[1])*dy)/(length*length)))
    return (a[0]+t*dx, a[1]+t*dy), (dx/length, dy/length)


class FillEvidence:
    """Read selected-layer HATCH/SOLID/wide-polyline evidence from the source DXF."""

    def __init__(self, source: str | Path | None, layers: set[str]):
        self.complete = False
        self.regions: list[tuple[str, list]] = []
        self.uncertain_layers: set[str] = set()
        if not source or not Path(source).is_file():
            return
        import ezdxf
        from ezdxf.path import from_hatch
        doc = ezdxf.readfile(source)

        def walk(entities, parent_layer=None, depth=0):
            for e in entities:
                layer = parent_layer if e.dxf.layer == "0" and parent_layer else e.dxf.layer
                if e.dxftype() == "INSERT":
                    if depth >= 16 or abs(abs(e.dxf.xscale)-abs(e.dxf.yscale)) > 1e-6:
                        self.uncertain_layers.add(layer)
                        continue
                    try:
                        yield from walk(e.virtual_entities(), layer, depth+1)
                    except (ValueError, TypeError, AttributeError):
                        self.uncertain_layers.add(layer)
                elif layer in layers:
                    yield e, layer

        for e, layer in walk(doc.modelspace()):
            kind = e.dxftype()
            if kind == "HATCH":
                if not e.dxf.solid_fill:
                    self.uncertain_layers.add(layer)
                    continue
                try:
                    loops = [[[float(v.x), float(v.y)] for v in p.flattening(.1, segments=16)]
                             for p in from_hatch(e)]
                    self.regions.append((layer, [p for p in loops if len(p) >= 3]))
                except (ValueError, TypeError, AttributeError):
                    self.uncertain_layers.add(layer)
            elif kind in ("SOLID", "TRACE"):
                self.regions.append((layer, [[[float(v.x),float(v.y)] for v in e.vertices()]]))
            elif kind == "LWPOLYLINE" and (e.dxf.const_width or any(p[2] or p[3] for p in e.get_points())):
                # Do not call a donut empty after discarding its actual width.
                self.uncertain_layers.add(layer)
        self.complete = True

    def body_fill(self, center: tuple, radius: float, layers: set[str]) -> str:
        """Require near-full body coverage; partial/unsupported fill remains unknown."""
        if not self.complete or self.uncertain_layers & layers:
            return "unknown"
        samples = [center] + [(center[0]+radius*.96*math.cos(a*math.pi/8),
                              center[1]+radius*.96*math.sin(a*math.pi/8)) for a in range(16)]
        covered = []
        for point in samples:
            covered.append(any(sum(zone_contains({"points": loop}, point) for loop in loops) % 2
                               for layer, loops in self.regions if layer in layers))
        if all(covered):
            return "solid"
        return "unknown" if any(covered) else "empty"


class WorldSymbolAdapter:
    """Spatially indexed adapter for the existing expanded CAD-world representation."""

    def __init__(self, world: Any, spec: dict, profile: SymbolProfile, *, source: str | None = None):
        self.world, self.spec, self.profile = world, spec, profile
        self.material = {tuple(p) for p in spec.get("material_picks", ())}
        self.selected = {tuple(h["bundle"]) for h in spec.get("heads", ())}
        self.selected.update(tuple(p) for h in spec.get("heads", ()) for p in h.get("mark_bundles", ()))
        self.allowed = self.selected | self.material
        self.grid: dict = defaultdict(set)
        self.rows = []
        self.long = []
        for row in world.segs:
            if not self._selected(row[:2], self.allowed):
                continue
            index = len(self.rows)
            self.rows.append(row)
            a, b = row[2:4]
            x0,x1 = sorted((int(a[0]//2000),int(b[0]//2000)))
            y0,y1 = sorted((int(a[1]//2000),int(b[1]//2000)))
            if (x1-x0+1)*(y1-y0+1) > 1024:
                self.long.append(index)
                continue
            for x in range(x0,x1+1):
                for y in range(y0,y1+1):
                    self.grid[x,y].add(index)
        self.fills = FillEvidence(source or getattr(world, "_source_path", None), {p[0] for p in self.allowed})

    @staticmethod
    def _selected(bundle, allowed):
        return any(bundle[0] == p[0] and (p[1] is None or bundle[1] == p[1]) for p in allowed)

    def classify(self, head: dict, *, explicit_orientation: str | None = None) -> Detection:
        """Classify one already-selected body; leave ambiguous alternatives for review."""
        cx, cy = head["c"]
        radius = float(head.get("head_r") or head.get("tri_side", 0)/math.sqrt(3))
        identity = f"{cx:.3f},{cy:.3f},{radius:.3f}"
        base = dict(candidate_id=identity, profile_id=self.profile.profile_id)
        if radius <= 0:
            return Detection(**base, status="review", reasons=("missing_body",))
        bundle = tuple(head["bundle"])
        mark_bundles = {bundle} | self.material
        for pick in self.spec.get("heads", ()):
            if tuple(pick["bundle"]) == bundle:
                mark_bundles.update(tuple(p) for p in pick.get("mark_bundles", ()))
        indices = set(self.long)
        reach = radius*6
        for x in range(int((cx-reach)//2000),int((cx+reach)//2000)+1):
            for y in range(int((cy-reach)//2000),int((cy+reach)//2000)+1):
                indices.update(self.grid.get((x,y), ()))
        near = [self.rows[i] for i in sorted(indices)
                if math.dist((cx,cy), _projection((cx,cy), *self.rows[i][2:4])[0]) <= reach
                and self._selected(self.rows[i][:2], mark_bundles)]
        tri_edges = {frozenset(map(tuple,e)) for e in head.get("tri_segs", ())}
        vertices = tuple(dict.fromkeys(tuple(p) for e in head.get("tri_segs", ()) for p in e))
        marks = tuple(Segment(tuple(a),tuple(b),ly) for ly,col,a,b in near
                      if max(math.dist((cx,cy),a), math.dist((cx,cy),b)) <= reach
                      and frozenset((tuple(a),tuple(b))) not in tri_edges)
        arcs = []
        complete = True
        for i,(ly,col,x,y,r) in enumerate(self.world.arcs):
            if math.dist((cx,cy),(x,y)) > reach or not self._selected((ly,col), mark_bundles):
                continue
            angle = self.world.arc_ang[i] if i < len(self.world.arc_ang) else None
            if angle is None:
                complete = False
                continue
            arcs.append(Arc((x,y),r,angle[0],angle[0]+angle[1],ly))
        for ly,col,x,y,r in self.world.circles:
            if (math.dist((cx,cy),(x,y))+r < radius*1.2 and r < radius*.9
                    and self._selected((ly,col),mark_bundles)):
                arcs.append(Arc((x,y),r,0,180,ly))
        layers = {p[0] for p in mark_bundles}
        fill = self.fills.body_fill((cx,cy),radius,{bundle[0]})
        anchors = []
        # Inline: actual pipe reaches the centre or the circle rim.
        for ly,col,a,b in near:
            if not self._selected((ly,col), self.material):
                continue
            q, axis = _projection((cx,cy),a,b)
            if math.dist(q,(cx,cy)) <= radius*1.25:
                # Only a line collinear with the centre may bridge a symbol gap.
                cross = abs((q[0]-cx)*axis[1]-(q[1]-cy)*axis[0])
                if cross <= radius*.2:
                    anchors.append(((cx,cy),axis))
        for arc in arcs:
            if 1.8*radius <= math.dist(arc.center,(cx,cy)) <= 5*radius:
                d = (arc.center[0]-cx,arc.center[1]-cy)
                anchors.append((arc.center,(-d[1],d[0])))
        if vertices:
            for mark in marks:
                for p in (mark.start,mark.end):
                    if 1.8*radius <= math.dist(p,(cx,cy)) <= 5*radius:
                        d=(p[0]-cx,p[1]-cy)
                        anchors.append((p,(-d[1],d[0])))
        results = []
        for anchor, axis in anchors:
            length=math.hypot(*axis)
            axis=(axis[0]/length,axis[1]/length)
            verified = False
            for ly,col,a,b in near:
                if not self._selected((ly,col),self.material):
                    continue
                q,d = _projection(anchor,a,b)
                if (abs(d[0]*axis[0]+d[1]*axis[1])>.94 and math.dist(q,anchor)<=radius*1.25
                        and abs((q[0]-anchor[0])*axis[1]-(q[1]-anchor[1])*axis[0])<=radius*.2):
                    verified=True
                    break
            temperature = re.search(r"(?<!\d)(\d{2,3})\s*(?:℃|°\s*[cC]|도)", bundle[0])
            candidate=HeadCandidate(identity,bundle[0],"plan","triangle" if vertices else "circle",
                (cx,cy),radius,anchor,axis,marks,tuple(arcs),vertices,fill,complete,verified,
                explicit_orientation=explicit_orientation,
                explicit_temperature_c=int(temperature.group(1)) if temperature else None)
            results.append(detect_head(candidate,self.profile,selected_head_layers=frozenset(layers)))
        matched = {r.rule_id:r for r in results if r.status == "matched"}
        if len(matched)==1:
            return next(iter(matched.values()))
        if len(matched)>1:
            return Detection(**base,status="review",reasons=("ambiguous_attachment",))
        return next((r for r in results if r.rule_id), Detection(**base,status="review",
                    reasons=("unknown_body_fill" if fill=="unknown" else "incomplete_company_symbol",)))


def apply_company_detection(record: dict, detection: Detection) -> dict:
    """Preserve evidence; only confident supported orientations replace legacy inference."""
    result = dict(record)
    result["company_symbol"] = asdict(detection)
    if detection.status == "matched":
        result["kind_before_company"] = result["kind"]
        result["kind"] = {"upright":"상향식","pendent":"하향식","upright_pendent":"상하향식"}.get(detection.orientation,"미지정")
        if detection.orientation == "sidewall":
            result["company_symbol"]["calculation_review"] = "측벽식 연결 높이/방향을 확인해야 합니다. 상·하향식으로 자동 변환하지 않습니다."
    return result

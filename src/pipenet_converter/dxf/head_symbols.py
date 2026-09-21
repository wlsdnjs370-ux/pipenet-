"""Classify localized CAD head symbols using an explicit company profile.

This is a geometric classifier, not a whole-drawing scanner or ML model. A DXF
adapter must first identify a body, isolate its marks, resolve block transforms,
inspect fill entities, and verify the pipe attachment. No graph is changed here.
Coordinates are millimetres; all matching distances are divided by body radius.
"""
from __future__ import annotations

from dataclasses import dataclass
import json
import math
from pathlib import Path
from typing import Literal

Point = tuple[float, float]
View = Literal["plan", "section"]


@dataclass(frozen=True)
class Segment:
    """A local symbol line, not a hydraulic pipe edge."""

    start: Point
    end: Point
    layer: str


@dataclass(frozen=True)
class Arc:
    """A local symbol arc; start/end are CAD angles in degrees."""

    center: Point
    radius: float
    start_degrees: float
    end_degrees: float
    layer: str


@dataclass(frozen=True)
class HeadCandidate:
    """One localized body and connection evidence from an upstream DXF adapter.

    `fill` describes actual body-area fill, not stroke colour. `empty` is valid
    only after HATCH/SOLID/wide-polyline inspection. Unknown or incomplete input
    deliberately cannot produce a confirmed result. For circles, vertices is
    empty; for triangles it contains all three boundary vertices. Body boundary
    segments are not duplicated in `marks`. Section up must be supplied explicitly.
    """

    candidate_id: str
    layer: str
    view: View
    body: Literal["circle", "triangle"]
    center_mm: Point
    radius_mm: float
    pipe_anchor_mm: Point
    pipe_axis: Point
    marks: tuple[Segment, ...] = ()
    arcs: tuple[Arc, ...] = ()
    vertices_mm: tuple[Point, ...] = ()
    fill: Literal["empty", "solid", "unknown"] = "unknown"
    geometry_complete: bool = False
    pipe_connection_verified: bool = False
    section_up: Point | None = None
    explicit_orientation: str | None = None
    explicit_temperature_c: int | None = None


@dataclass(frozen=True)
class Rule:
    """One versioned semantic rule and its source-page provenance."""

    id: str
    view: str
    body: str
    decoration: str
    attachment: str
    orientation: str
    label: str
    head_count: int
    temperature_c: int | None
    source_slides: tuple[int, ...]


@dataclass(frozen=True)
class SymbolProfile:
    """Explicitly loaded rules; no company or layer is enabled globally."""

    profile_id: str
    source_sha256: str
    rules: tuple[Rule, ...]
    position_tolerance: float
    parallel_cosine: float
    offset_min: float
    offset_max: float
    arc_min: float
    arc_max: float


@dataclass(frozen=True)
class Detection:
    """Inspectable suggestion only; never changes nodes, pipes or manual data."""

    candidate_id: str
    profile_id: str
    status: Literal["matched", "review", "excluded"]
    rule_id: str | None = None
    orientation: str | None = None
    temperature_c: int | None = None
    head_count: int | None = None
    source_slides: tuple[int, ...] = ()
    evidence: tuple[str, ...] = ()
    reasons: tuple[str, ...] = ()


def load_profile(path: str | Path) -> SymbolProfile:
    """Load a declarative profile; fail explicitly for unsupported schema/rules."""
    raw = json.loads(Path(path).read_text(encoding="utf-8"))
    if raw.get("schema_version") != 1:
        raise ValueError("Unsupported head-symbol profile schema")
    rules = tuple(Rule(**{key: (tuple(row[key]) if key == "source_slides" else row[key])
                          for key in Rule.__dataclass_fields__}) for row in raw["rules"])
    if not rules or len({r.id for r in rules}) != len(rules):
        raise ValueError("Head-symbol rules must have unique IDs")
    signatures = {(r.view, r.body, r.decoration, r.attachment) for r in rules}
    if len(signatures) != len(rules):
        raise ValueError("Ambiguous head-symbol rules")
    for r in rules:
        if (r.view not in ("plan", "section") or r.body not in ("circle", "triangle")
                or r.decoration not in ("empty", "cross", "solid")
                or r.attachment not in ("inline", "offset_arch", "offset_ticks", "stem_up", "stem_down")
                or r.orientation not in ("upright", "pendent", "upright_pendent", "sidewall")
                or r.head_count not in (1, 2) or not r.source_slides):
            raise ValueError(f"Unsupported head-symbol rule: {r.id}")
    t = raw["tolerances"]
    result = SymbolProfile(raw["profile_id"], raw["source"]["sha256"], rules,
                           float(t["position_radius_ratio"]), float(t["parallel_cosine"]),
                           float(t["offset_min_radius_ratio"]), float(t["offset_max_radius_ratio"]),
                           float(t["arc_min_sweep_degrees"]), float(t["arc_max_sweep_degrees"]))
    values = (result.position_tolerance, result.parallel_cosine, result.offset_min,
              result.offset_max, result.arc_min, result.arc_max)
    if (not all(math.isfinite(v) for v in values)
            or not 0 < result.position_tolerance < 0.5
            or not 0.8 <= result.parallel_cosine < 1
            or not 1 < result.offset_min < result.offset_max
            or not 0 < result.arc_min <= 180 <= result.arc_max < 360):
        raise ValueError("Invalid head-symbol tolerances")
    return result


def _sub(a: Point, b: Point) -> Point:
    return a[0] - b[0], a[1] - b[1]


def _dot(a: Point, b: Point) -> float:
    return a[0] * b[0] + a[1] * b[1]


def _norm(a: Point) -> float:
    return math.hypot(*a)


def _unit(a: Point) -> Point:
    length = _norm(a)
    if length == 0 or not math.isfinite(length):
        raise ValueError("A direction vector must be finite and nonzero")
    return a[0] / length, a[1] / length


def _near_line(p: Point, a: Point, b: Point, tol: float) -> bool:
    d = _sub(b, a)
    if _norm(d) == 0:
        return _norm(_sub(p, a)) <= tol
    t = max(0.0, min(1.0, _dot(_sub(p, a), d) / _dot(d, d)))
    return _norm(_sub(p, (a[0] + t * d[0], a[1] + t * d[1]))) <= tol


def _decoration(lines: tuple[tuple[Point, Point], ...], fill: str, tol: float) -> str:
    """Detect an X from two diameters, including diameters split at the centre."""
    interior = [(a, b) for a, b in lines
                if _norm(a) < 1 + tol and _norm(b) < 1 + tol
                and _norm(((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)) < 1 - tol]
    if fill == "solid":
        return "solid" if not interior else "unknown"
    if not interior:
        return "empty"
    directions: list[Point] = []
    for a, b in interior:
        if not _near_line((0, 0), a, b, tol):
            return "unknown"
        for p in (a, b):
            if 1 - tol <= _norm(p) <= 1 + tol:
                d = _unit(p)
                if not any(_dot(d, other) > 0.98 for other in directions):
                    directions.append(d)
            elif _norm(p) > tol:
                return "unknown"
    if len(directions) != 4:
        return "unknown"
    for d in directions:
        dots = sorted(_dot(d, other) for other in directions if other != d)
        if dots[0] > -0.94 or max(abs(dots[1]), abs(dots[2])) > tol:
            return "unknown"
    # Do not mistake an axial pipe crossing (+) for this company's diagonal X.
    if any(not 0.4 < abs(d[0]) < 0.9 for d in directions):
        return "unknown"
    return "cross"


def _connector(lines: tuple[tuple[Point, Point], ...], anchor: Point,
               tol: float, cosine: float) -> tuple[bool, bool]:
    """Require a stem, neck bar, and two flanking bars around the pipe anchor."""
    neck = _unit(anchor)
    stem = any(abs(_dot(_unit(_sub(b, a)), neck)) >= cosine
               and _near_line((neck[0], neck[1]), a, b, tol)
               and _near_line(anchor, a, b, tol) for a, b in lines if _norm(_sub(b, a)) > tol)
    flank_signs: set[int] = set()
    neck_bar = False
    for a, b in lines:
        d = _sub(b, a)
        if _norm(d) <= tol:
            continue
        middle = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
        if (abs(_unit(d)[1]) >= cosine and abs(middle[1] - anchor[1]) <= tol
                and 1.2 <= abs(middle[0] - anchor[0]) <= 2.5
                and _norm(d) >= 0.7):
            flank_signs.add(1 if middle[0] > anchor[0] else -1)
        if (abs(_dot(_unit(d), neck)) <= 1 - cosine
                and 1 + tol <= _dot(middle, neck) <= _norm(anchor) - tol
                and _near_line((_dot(middle, neck) * neck[0], _dot(middle, neck) * neck[1]), a, b, tol)
                and _norm(d) >= 0.7):
            neck_bar = True
    return stem, neck_bar and flank_signs == {-1, 1}


def detect_head(candidate: HeadCandidate, profile: SymbolProfile, *,
                selected_head_layers: frozenset[str]) -> Detection:
    """Classify one candidate without scanning other layers or mutating geometry.

    Returns `review` instead of guessing when input is incomplete, contradicts
    explicit metadata, or lacks the full symbol/connection evidence. A match is
    a suggestion, not authorization to create hydraulic nodes or set flow/K.
    """
    c = candidate
    base = {"candidate_id": c.candidate_id, "profile_id": profile.profile_id}
    if c.layer not in selected_head_layers:
        return Detection(**base, status="excluded", reasons=("layer_not_selected",))
    if c.view not in ("plan", "section") or c.body not in ("circle", "triangle"):
        raise ValueError("Unsupported candidate body or drawing view")
    points = (c.center_mm, c.pipe_anchor_mm, c.pipe_axis, *c.vertices_mm,
              *(p for s in c.marks for p in (s.start, s.end)), *(a.center for a in c.arcs))
    scalars = (c.radius_mm, *(v for p in points for v in p),
               *(v for a in c.arcs for v in (a.radius, a.start_degrees, a.end_degrees)))
    if not all(math.isfinite(v) for v in scalars) or c.radius_mm <= 0 or any(a.radius <= 0 for a in c.arcs):
        raise ValueError("Candidate geometry must be finite with positive radii")
    axis = _unit(c.pipe_axis)
    reasons = tuple(reason for ok, reason in (
        (c.geometry_complete, "incomplete_geometry"),
        (c.pipe_connection_verified, "unverified_pipe_connection"),
        (c.fill in ("empty", "solid"), "unknown_body_fill")) if not ok)
    if reasons:
        return Detection(**base, status="review", reasons=reasons)
    normal = (-axis[1], axis[0])

    def local(p: Point) -> Point:
        d = _sub(p, c.center_mm)
        return _dot(d, axis) / c.radius_mm, _dot(d, normal) / c.radius_mm

    lines = tuple((local(s.start), local(s.end)) for s in c.marks if s.layer in selected_head_layers)
    arcs = tuple(a for a in c.arcs if a.layer in selected_head_layers)
    tol = profile.position_tolerance
    anchor = local(c.pipe_anchor_mm)
    decoration = _decoration(lines, c.fill, tol)
    attachment = "unknown"
    if c.body == "circle" and any(_norm(local(a.center)) + a.radius / c.radius_mm < 1 + tol for a in arcs):
        # Nested circles/arcs are different glyphs, including actual legacy DXFs.
        decoration = "unknown"
    if c.body == "circle" and c.view == "plan" and _norm(anchor) <= tol:
        attachment = "inline"
    elif profile.offset_min <= _norm(anchor) <= profile.offset_max and abs(anchor[0]) <= tol:
        stem, ticks = _connector(lines, anchor, tol, profile.parallel_cosine)
        if c.body == "circle" and c.view == "plan" and stem and ticks:
            # The open side of the semicircle faces the body, at the pipe anchor.
            for a in arcs:
                sweep = (a.end_degrees - a.start_degrees) % 360
                mid_angle = math.radians(a.start_degrees + sweep / 2)
                arc_out = (math.cos(mid_angle), math.sin(mid_angle))
                away = _unit(_sub(c.pipe_anchor_mm, c.center_mm))
                if (_norm(_sub(local(a.center), anchor)) <= tol
                        and abs(a.radius / c.radius_mm - 1) <= tol
                        and profile.arc_min <= sweep <= profile.arc_max
                        and _dot(arc_out, away) >= profile.parallel_cosine):
                    attachment = "offset_arch"
        elif c.body == "triangle" and stem:
            if len(c.vertices_mm) != 3:
                raise ValueError("A triangle must have three boundary vertices")
            vertices = tuple(local(p) for p in c.vertices_mm)
            ab, ac = _sub(vertices[1], vertices[0]), _sub(vertices[2], vertices[0])
            if abs(ab[0] * ac[1] - ab[1] * ac[0]) < tol or not 1 - tol <= max(map(_norm, vertices)) <= 1 + tol:
                return Detection(**base, status="review", reasons=("invalid_triangle_boundary",))
            if abs(sum(p[0] for p in vertices)) > tol or abs(sum(p[1] for p in vertices)) > tol:
                return Detection(**base, status="review", reasons=("triangle_center_not_centroid",))
            tip = min(vertices, key=lambda p: _dot(p, anchor))
            if _norm(tip) == 0 or _dot(_unit(tip), _unit(anchor)) > -profile.parallel_cosine:
                return Detection(**base, status="review", reasons=("triangle_direction_unclear",))
            if c.view == "plan" and ticks:
                attachment = "offset_ticks"
            elif c.view == "section" and c.section_up is not None:
                up = _unit(c.section_up)
                direction = _unit(_sub(c.center_mm, c.pipe_anchor_mm))
                alignment = _dot(direction, up)
                if abs(alignment) >= profile.parallel_cosine:
                    attachment = "stem_up" if alignment > 0 else "stem_down"
    matches = [r for r in profile.rules if (r.view, r.body, r.decoration, r.attachment)
               == (c.view, c.body, decoration, attachment)]
    evidence = (f"view={c.view}", f"body={c.body}", f"decoration={decoration}", f"attachment={attachment}")
    if len(matches) != 1:
        return Detection(**base, status="review", evidence=evidence, reasons=("no_unique_complete_rule",))
    rule = matches[0]
    conflicts = tuple(name for bad, name in (
        (c.explicit_orientation is not None and c.explicit_orientation != rule.orientation, "orientation_conflict"),
        (c.explicit_temperature_c is not None and rule.temperature_c is not None
         and c.explicit_temperature_c != rule.temperature_c, "temperature_conflict")) if bad)
    if conflicts:
        return Detection(**base, status="review", rule_id=rule.id, evidence=evidence,
                         source_slides=rule.source_slides, reasons=conflicts)
    return Detection(**base, status="matched", rule_id=rule.id, orientation=rule.orientation,
                     temperature_c=rule.temperature_c, head_count=rule.head_count,
                     source_slides=rule.source_slides, evidence=evidence)

"""PPT-derived symbolic cases, independent of Flask and company DXF availability."""
from dataclasses import replace
import json
import math
from pathlib import Path

import pytest

from src.pipenet_converter.dxf.head_symbols import (
    Arc, HeadCandidate, Segment, detect_head, load_profile,
)

ROOT = Path(__file__).resolve().parents[2]
PROFILE_PATH = ROOT / "configs/head_symbols/company_260404.json"
PROFILE = load_profile(PROFILE_PATH)
LAYERS = frozenset({"HEAD"})


def symbol(variant: str) -> HeadCandidate:
    """Draw normalized geometry according to the slide-26 legend, not rule IDs."""
    lines = []
    arcs = ()
    offset = variant in ("pendent", "pair_offset", "hot", "sidewall")
    if "pair" in variant:
        d = 2 ** -0.5
        lines += [((-d, -d), (d, d)), ((-d, d), (d, -d))]
    if offset:
        lines += [((0, 1), (0, 3)), ((-.8, 1.5), (.8, 1.5)),
                  ((-1.8, 2.3), (-1.8, 3.7)), ((1.8, 2.3), (1.8, 3.7))]
        if variant != "sidewall":
            arcs = (Arc((0, 3), 1, 0, 180, "HEAD"),)
    triangle = variant == "sidewall"
    return HeadCandidate(variant, "HEAD", "plan", "triangle" if triangle else "circle",
                         (0, 0), 1, (0, 3) if offset else (0, 0), (1, 0),
                         tuple(Segment(a, b, "HEAD") for a, b in lines), arcs,
                         ((0, -1), (-math.sqrt(3)/2, .5), (math.sqrt(3)/2, .5)) if triangle else (),
                         "solid" if variant == "hot" else "empty", True, True)


def transform(c: HeadCandidate, angle: float, scale: float) -> HeadCandidate:
    cs, sn = math.cos(math.radians(angle)), math.sin(math.radians(angle))

    def direction(p):
        return (p[0] * cs - p[1] * sn, p[0] * sn + p[1] * cs)

    def point(p):
        x, y = direction(p)
        return (x * scale + 263470.3, y * scale - 231456.8)

    return replace(c, center_mm=point(c.center_mm), pipe_anchor_mm=point(c.pipe_anchor_mm),
                   radius_mm=c.radius_mm * scale, pipe_axis=direction(c.pipe_axis),
                   marks=tuple(Segment(point(s.start), point(s.end), s.layer) for s in c.marks),
                   arcs=tuple(Arc(point(a.center), a.radius * scale, a.start_degrees + angle,
                                  a.end_degrees + angle, a.layer) for a in c.arcs),
                   vertices_mm=tuple(point(p) for p in c.vertices_mm),
                   section_up=direction(c.section_up) if c.section_up else None)


@pytest.mark.parametrize("angle", [0, 37, 90, 180, 270])
@pytest.mark.parametrize("scale", [1, 25, 400])
@pytest.mark.parametrize("variant,orientation,temp,count", [
    ("upright", "upright", 72, 1), ("pair", "upright_pendent", 72, 2),
    ("pendent", "pendent", 72, 1), ("pair_offset", "upright_pendent", 72, 2),
    ("hot", "pendent", 103, 1), ("sidewall", "sidewall", 72, 1),
])
def test_all_plan_variants_rotation_scale_translation(variant, orientation, temp, count, angle, scale):
    c = transform(symbol(variant), angle, scale)
    before = repr(c)
    got = detect_head(c, PROFILE, selected_head_layers=LAYERS)
    assert got.status == "matched", got
    assert (got.orientation, got.temperature_c, got.head_count) == (orientation, temp, count)
    assert 26 in got.source_slides
    assert repr(c) == before


@pytest.mark.parametrize("orientation", ["upright", "pendent"])
@pytest.mark.parametrize("angle", [0, 37, 90, 180])
def test_section_triangles_use_explicit_up_not_plan_direction(orientation, angle):
    c = replace(symbol("sidewall"), view="section", fill="solid", section_up=(0, 1),
                marks=(Segment((0, 1), (0, 3), "HEAD"),))
    if orientation == "upright":
        c = replace(c, pipe_anchor_mm=(0, -3), marks=(Segment((0, -1), (0, -3), "HEAD"),),
                    vertices_mm=tuple((x, -y) for x, y in c.vertices_mm))
    got = detect_head(transform(c, angle, 40), PROFILE, selected_head_layers=LAYERS)
    assert got.status == "matched", got
    assert got.orientation == orientation and got.temperature_c is None


def test_section_requires_explicit_up_and_does_not_relabel_plan_sidewall():
    c = replace(symbol("sidewall"), view="section", fill="solid")
    assert detect_head(c, PROFILE, selected_head_layers=LAYERS).status == "review"
    assert detect_head(symbol("sidewall"), PROFILE, selected_head_layers=LAYERS).orientation == "sidewall"


@pytest.mark.parametrize("change,reason", [
    ({"geometry_complete": False}, "incomplete_geometry"),
    ({"pipe_connection_verified": False}, "unverified_pipe_connection"),
    ({"fill": "unknown"}, "unknown_body_fill"),
    ({"explicit_orientation": "pendent"}, "orientation_conflict"),
    ({"explicit_temperature_c": 68}, "temperature_conflict"),
])
def test_uncertainty_and_explicit_metadata_do_not_get_overwritten(change, reason):
    got = detect_head(replace(symbol("upright"), **change), PROFILE, selected_head_layers=LAYERS)
    assert got.status == "review" and reason in got.reasons
    assert got.orientation is None and got.temperature_c is None


def test_unselected_layer_and_extra_layer_marks():
    assert detect_head(symbol("upright"), PROFILE, selected_head_layers=frozenset()).status == "excluded"
    c = replace(symbol("upright"), marks=tuple(replace(s, layer="WALL") for s in symbol("pair").marks))
    assert detect_head(c, PROFILE, selected_head_layers=LAYERS).orientation == "upright"


@pytest.mark.parametrize("variant", ["pendent", "pair_offset", "hot", "sidewall"])
def test_offset_needs_stem_neck_and_flanks(variant):
    c = symbol(variant)
    assert detect_head(replace(c, marks=c.marks[:-1]), PROFILE, selected_head_layers=LAYERS).status == "review"


def test_wrong_arc_direction_missing_arch_and_extra_body_marks():
    c = symbol("pendent")
    for modified in (replace(c, arcs=()), replace(c, arcs=(replace(c.arcs[0], start_degrees=180, end_degrees=360),))):
        assert detect_head(modified, PROFILE, selected_head_layers=LAYERS).status == "review"
    c = replace(symbol("upright"), marks=(Segment((-.5, 0), (.5, 0), "HEAD"),))
    assert detect_head(c, PROFILE, selected_head_layers=LAYERS).status == "review"


def test_solid_requires_offset_not_generic_dot():
    assert detect_head(replace(symbol("upright"), fill="solid"), PROFILE, selected_head_layers=LAYERS).status == "review"


def test_split_x_lines_still_one_pair_but_plus_not_pair():
    c = symbol("pair")
    c = replace(c, marks=tuple(Segment(a, b, "HEAD") for s in c.marks
                               for a, b in ((s.start, (0, 0)), ((0, 0), s.end))))
    assert detect_head(c, PROFILE, selected_head_layers=LAYERS).head_count == 2
    c = replace(c, marks=(Segment((-1, 0), (1, 0), "HEAD"), Segment((0, -1), (0, 1), "HEAD")))
    assert detect_head(c, PROFILE, selected_head_layers=LAYERS).status == "review"


@pytest.mark.parametrize("change", [{"radius_mm": 0}, {"radius_mm": float("nan")},
                                   {"pipe_axis": (0, 0)}, {"center_mm": (float("inf"), 0)}])
def test_invalid_geometry_fails_clearly(change):
    with pytest.raises(ValueError):
        detect_head(replace(symbol("upright"), **change), PROFILE, selected_head_layers=LAYERS)


def test_profile_provenance_all_slides_and_unique_variants():
    raw = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
    assert raw["source"]["slides"] == list(range(26, 37))
    assert len(PROFILE.rules) == 8
    assert set(n for r in PROFILE.rules for n in r.source_slides) == set(range(26, 37))
    assert len(PROFILE.source_sha256) == 64


def test_bad_profile_rejected(tmp_path):
    raw = json.loads(PROFILE_PATH.read_text(encoding="utf-8"))
    raw["rules"].append(raw["rules"][0])
    path = tmp_path / "bad.json"
    path.write_text(json.dumps(raw), encoding="utf-8")
    with pytest.raises(ValueError, match="unique"):
        load_profile(path)


def test_actual_mf401_pa03_inner_bulge_circle_is_not_plain_upright():
    # MF-401 SHA256 b983aaa654d8efa6faf7330667b8f5de6db04b1f5700323448957cf58cd16ce2.
    # pa-03: outer CIRCLE R1; closed two-vertex bulge polyline has inner R0.125.
    # Normalize the actual block-local geometry; do not invent its temperature.
    c = replace(symbol("upright"), candidate_id="MF401:pa-03",
                arcs=(Arc((0, 0), .125, 0, 180, "HEAD"), Arc((0, 0), .125, 180, 360, "HEAD")))
    got = detect_head(c, PROFILE, selected_head_layers=LAYERS)
    assert got.status == "review" and got.temperature_c is None


def test_actual_mf105_dry_nested_circle_x_is_not_plain_pair():
    # MF-105 SHA256 8bf76108ae07db47321c3a9bdcf0203e64fc9987ae047d0e69f5376f08731c38.
    # SP01-03(DRY): outer bulge circle, diagonal X, extra inner CIRCLE R0.6140449.
    # Same selected-layer/context contract as the synthetic tests, but real geometry.
    a = (-.7627494260668755, -.7918404976371676)
    b = (.7928854897618294, .763794420985505)
    center = (.015068031847477, -.0140230383258313)
    radius = math.dist(a, b) / 2
    c = replace(symbol("pair"), candidate_id="MF105:SP01-03(DRY)", center_mm=center,
                pipe_anchor_mm=center, radius_mm=radius,
                marks=(Segment(a, b, "HEAD"), Segment((a[0], b[1]), (b[0], a[1]), "HEAD")),
                arcs=(Arc(center, .6140449356753379, 0, 180, "HEAD"),
                      Arc(center, .6140449356753379, 180, 360, "HEAD")))
    got = detect_head(c, PROFILE, selected_head_layers=LAYERS)
    assert got.status == "review" and got.head_count is None

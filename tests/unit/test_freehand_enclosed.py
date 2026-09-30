"""The pen keeps every enclosed part, not a hull or parity-cancelled selection."""
from copy import deepcopy
import json
from pathlib import Path
import subprocess

import pytest

from src.pipenet_converter.graph.freehand import enclosed_geometry
from src.pipenet_converter.graph.regions import normalize_zones, zone_contains, zones_signature
from src.pipenet_converter.dxf.work_region import WorkRegion, crop_world
from types import SimpleNamespace

SQUARE = [[0,0],[10,0],[10,10],[0,10]]
BOWTIE = [[0,0],[10,10],[0,10],[10,0]]
CONCAVE = [[0,0],[10,0],[10,10],[7,10],[7,3],[3,3],[3,10],[0,10]]
OVERLAP = SQUARE + [[0,0],[5,0],[5,10],[15,10],[15,0],[5,0]]
NESTED = SQUARE + [[0,0],[2,2],[2,8],[8,8],[8,2],[2,2],[0,0]]


def pen(points: list) -> dict:
    """Use the versioned crop policy without changing legacy region semantics."""
    return {'type':'polygon', 'points':points, 'fill_rule':'enclosed'}


@pytest.mark.parametrize(('points', 'expected_area', 'inside', 'outside'), [
    (BOWTIE, 50, [[5,2],[5,8],[5,5]], [[1,5],[9,5]]),
    (SQUARE + SQUARE, 100, [[5,5]], [[11,5]]),
    (SQUARE + [[0,0]] + list(reversed(SQUARE)), 100, [[5,5]], [[11,5]]),
    (SQUARE + [[0,0],[-5,-5]], 100, [[5,5]], [[-2,-2],[-3,1]]),
    (OVERLAP, 150, [[2,5],[7,5],[12,5]], [[16,5]]),
    (NESTED, 100, [[1,1],[5,5],[9,9]], [[11,5]]),
    (CONCAVE, 72, [[1,8],[9,8],[5,2]], [[5,5],[5,9]]),
])
def test_union_area_containment_and_client_server_parity(points, expected_area, inside, outside):
    original = deepcopy(points)
    zone = normalize_zones([pen(points)])[0]
    geometry = enclosed_geometry(tuple(map(tuple, zone['points'])))
    assert geometry.area == pytest.approx(expected_area)
    region = WorkRegion(zone)
    assert all(region.contains(p) and zone_contains(zone, p) for p in inside)
    assert all(not region.contains(p) and not zone_contains(zone, p) for p in outside)
    # Dense probe checks holes, overlap, tails and boundary consistently in JS.
    probes = inside + outside + [[x+.31,y+.27] for x in range(-1,17) for y in range(-1,12)]
    script_path = Path(__file__).resolve().parents[2] / 'static/module_f_regions.js'
    script = f'''require({json.dumps(str(script_path))}); const z={json.dumps(zone)},pts={json.dumps(probes)};
      console.log(JSON.stringify({{area:ModuleFRegions.area(z),inside:pts.map(p=>ModuleFRegions.contains(z,p))}}));'''
    result = json.loads(subprocess.run(['node','-e',script], check=True, text=True, capture_output=True).stdout)
    assert result['area'] == pytest.approx(expected_area)
    assert result['inside'] == [region.contains(p) for p in probes]
    assert normalize_zones(json.loads(json.dumps([zone]))) == [zone]
    assert zones_signature([zone]) != zones_signature([dict(type='polygon',points=SQUARE)])
    assert points == original


def test_internal_crossings_are_not_clipping_boundaries_and_tail_is_not_selected():
    region = WorkRegion(pen(OVERLAP))
    assert region.clip((-5,5),(20,5)) == [((0,5),(15,5))]
    assert region.whole_segment((1,5),(14,5))
    region = WorkRegion(pen(CONCAVE))
    assert region.clip((-5,5),(15,5)) == [((0,5),(3,5)),((7,5),(10,5))]
    assert not region.whole_segment((1,5),(9,5))
    tail = WorkRegion(pen(SQUARE + [[0,0],[-5,-5]]))
    assert tail.clip((-5,-5),(-1,-1)) == []


def test_all_types_inside_kept_and_boundary_symbols_remain_whole():
    source = SimpleNamespace(segs=[('P',3,(-5,5),(15,5))],raw_segs=[('P',3,(-5,5),(15,5))],
        circles=[('H',2,0,5,1),('H',2,9.8,5,1),('H',2,11,5,2)],
        arcs=[('A',4,0,6,2),('A',4,12,6,1)],arc_ang=[(0,90),(90,90)],
        texts=[('T',7,0,5,1,'SP 150'),('T',7,5,5,1,'25'),('T',7,11,5,1,'100')])
    before = deepcopy(source.__dict__)
    result, report = crop_world(source, pen(SQUARE + SQUARE))
    assert result.segs == [('P',3,(0,5),(10,5))] and result.raw_segs == result.segs
    assert result.circles == source.circles[:2] and result.arcs == source.arcs[:1]
    assert result.arc_ang == [(0,90)] and result.texts == source.texts[:2]
    assert report['boundary_symbols_excluded'] == 0
    replay, _ = crop_world(source, json.loads(json.dumps(result._work_regions[0])))
    assert replay.__dict__ == result.__dict__ and source.__dict__ == before


def test_large_cad_origin_does_not_change_enclosed_faces():
    shifted = [[x+800000000,y-800000000] for x,y in BOWTIE]
    geometry = enclosed_geometry(tuple(map(tuple, shifted)))
    assert geometry.area == pytest.approx(50)
    region = WorkRegion(pen(shifted))
    assert region.contains((800000005,-799999998))
    assert not region.contains((800000001,-799999995))


@pytest.mark.parametrize('points', [[[0,0],[1,1],[2,2]], [[0,0]]*3,
                                  [[0,0],[10,0],[10,float('inf')]], [[0,0]]*513])
def test_empty_or_invalid_strokes_are_still_rejected(points):
    with pytest.raises(ValueError):
        WorkRegion(pen(points))

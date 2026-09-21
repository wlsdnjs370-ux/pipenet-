"""Concave/freehand containment and downstream selection contracts."""
import json
import subprocess
from pathlib import Path

import pytest

from src.pipenet_converter.graph.regions import (
    ZoneRegion, normalize_zones, zone_contains, zones_signature,
)

L_SHAPE = {"type":"polygon", "points":[[0,0],[10,0],[10,3],[3,3],[3,10],[0,10]]}
DIAGONAL = {"type":"polygon", "points":[[0,2],[8,10],[10,8],[2,0]]}


@pytest.mark.parametrize("zone,point,inside", [
    (L_SHAPE,(1,8),True),(L_SHAPE,(8,8),False),(L_SHAPE,(3,8),True),
    (L_SHAPE,(8,3),True),(L_SHAPE,(3,3),True),(L_SHAPE,(3.001,3.001),False),
    (DIAGONAL,(5,5),True),(DIAGONAL,(0,10),False),(DIAGONAL,(4,6),True),
    ([0,0,10,10],(0,0),True),([0,0,10,10],(11,2),False),
])
def test_actual_polygon_not_bounding_box(zone,point,inside):
    assert zone_contains(zone,point) is inside
    root = Path(__file__).resolve().parents[2]
    script = f"require({json.dumps(str(root/'static/module_f_regions.js'))}); console.log(ModuleFRegions.contains({json.dumps(zone)},{json.dumps(point)}));"
    result = subprocess.run(["node","-e",script],capture_output=True,text=True,check=True)
    assert (result.stdout.strip() == "true") is inside


def test_multiple_zones_union_and_downstream_region_dilation():
    from routes.module_f.auto import head_region_of
    region = head_region_of([L_SHAPE,[20,20,30,30]])
    assert region.contains((1,8)) and region.contains((25,25))
    assert not region.contains((8,8))
    assert region.dilate(1).contains((3.5,7))
    assert not region.contains((3.5,7))
    assert [3,3] in region.pts


def test_candidate_replacement_cannot_escape_into_polygon_cutout():
    from routes.module_f.selection import zone_confined_pool
    disks = [(1,8,1),(8,8,1),(2,7,1),(25,25,1)]
    result = zone_confined_pool([0,1,2,3],[0],[L_SHAPE,[20,20,30,30]],disks)
    assert result == [0,2]


def test_signature_changes_when_only_inner_corner_moves_and_json_roundtrip():
    changed=json.loads(json.dumps(L_SHAPE))
    changed["points"][3][0] = 2
    assert zones_signature([changed]) != zones_signature([L_SHAPE])
    zones=normalize_zones([L_SHAPE,[10,10,0,0]])
    assert normalize_zones(json.loads(json.dumps(zones))) == zones
    assert zones[0] == L_SHAPE
    assert zones[1] == [0,0,10,10]


@pytest.mark.parametrize("zones", ["invalid",[{}],[[0,0,0,1]],[[0,0,float("nan"),1]],
    [{"type":"polygon","points":[[0,0],[1,1],[2,2]]}],
    [{"type":"polygon","points":[[0,0]]*513}],[[False,0,1,1]]])
def test_bad_input_is_rejected_not_silently_dropped(zones):
    with pytest.raises(ValueError):
        normalize_zones(zones)


def test_duplicate_closing_vertex_and_large_cad_coordinates():
    pts=[[x+800000000,y-800000000] for x,y in L_SHAPE["points"]]
    got=normalize_zones([{"type":"polygon","points":pts+[pts[0]]}])[0]
    assert len(got["points"])==6
    assert zone_contains(got,(800000001,-799999992))
    assert not zone_contains(got,(800000008,-799999992))

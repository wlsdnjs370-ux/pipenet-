"""Source geometry/visibility and single-read caching cannot alter graph inputs."""
from copy import deepcopy
from unittest.mock import patch

import ezdxf
import pytest

from src.pipenet_converter.render.dxf_source import source_display
from routes.module_f.common import _boot
from routes.module_f import sub_fastread as sf
from routes.module_h_subdrawing import display_payload


def rows(doc, kind):
    return [p for p in source_display(doc)["primitives"] if p["kind"] == kind]


@pytest.mark.parametrize("rotation", [0, 90, 180, 270, 37])
def test_nested_base_rotation_and_arc_direction(rotation):
    doc = ezdxf.new()
    inner = doc.blocks.new("inner", base_point=(10, 20))
    inner.add_arc((10, 20), 5, 0, 90)
    outer = doc.blocks.new("outer")
    outer.add_blockref("inner", (40, 50), dxfattribs={"rotation": rotation})
    doc.modelspace().add_blockref("outer", (100, 200), dxfattribs={"xscale": 2, "yscale": 2})
    assert rows(doc, "arcs")[0]["xy"] == pytest.approx([180, 300, 10, rotation, 90])


def test_negative_ocs_centers_and_signed_sweep():
    doc = ezdxf.new()
    doc.modelspace().add_circle((500, 600), 10, dxfattribs={"extrusion": (0, 0, -1)})
    doc.modelspace().add_arc((500, 600), 10, 0, 90, dxfattribs={"extrusion": (0, 0, -1)})
    assert rows(doc, "circles")[0]["xy"] == pytest.approx([-500, 600, 10])
    assert rows(doc, "arcs")[0]["xy"] == pytest.approx([-500, 600, 10, 180, -90])


def test_closed_polyline_and_bulge_are_not_open_chords():
    doc = ezdxf.new()
    doc.modelspace().add_lwpolyline([(0,0),(40,0),(40,40),(0,40)], close=True)
    assert len(rows(doc, "segs")) == 4
    doc.modelspace().add_lwpolyline([(100,0,1),(120,0,0)], format="xyb")
    arcs = rows(doc, "arcs")
    assert len(arcs) == 1 and abs(arcs[0]["xy"][4]) == 180
    assert len(rows(doc, "segs")) == 4


def test_unequal_scale_circle_is_an_ellipse_not_a_wrong_circle():
    doc = ezdxf.new()
    doc.blocks.new("oval").add_circle((0,0),10)
    doc.modelspace().add_blockref("oval",(100,200), dxfattribs={"xscale":2,"yscale":1})
    display = source_display(doc)
    assert not display["unsupported"]
    assert not any(p["kind"] == "circles" for p in display["primitives"])
    points = [p["xy"] for p in display["primitives"]]
    assert min(p[0] for p in points) == pytest.approx(80)
    assert max(p[0] for p in points) == pytest.approx(120)
    assert min(p[1] for p in points) == pytest.approx(190)
    assert max(p[1] for p in points) == pytest.approx(210)


def visibility_doc():
    doc = ezdxf.new()
    doc.layers.new("OFF", dxfattribs={"color":1}).off()
    doc.layers.new("FROZEN", dxfattribs={"color":2}).freeze()
    doc.layers.new("VISIBLE", dxfattribs={"color":3})
    doc.modelspace().add_line((0,0),(100,100),dxfattribs={"layer":"VISIBLE"})
    doc.modelspace().add_line((1e7,0),(1e7,100),dxfattribs={"layer":"OFF"})
    doc.modelspace().add_circle((2e7,0),20,dxfattribs={"layer":"FROZEN"})
    block = doc.blocks.new("hidden_parent")
    block.add_line((0,0),(10,10), dxfattribs={"layer":"VISIBLE"})
    block.add_text("hidden",dxfattribs={"height":10,"layer":"VISIBLE"})
    doc.modelspace().add_blockref("hidden_parent",(3e7,0),dxfattribs={"layer":"OFF"})
    return doc


def test_hidden_layers_and_parent_preserved_but_excluded_from_initial_fit():
    _boot()
    source = source_display(visibility_doc())
    before = deepcopy(source)
    entities = [{"t":"L", "p":[1,2,3,4], "l":"unchanged"}]
    payload = display_payload(entities, {}, "unused", source=source)
    assert source == before and entities == [{"t":"L", "p":[1,2,3,4], "l":"unchanged"}]
    assert payload["bounds"] == dict(minx=0,miny=0,maxx=100,maxy=100)
    assert payload["full_bounds"]["maxx"] > 1e7
    assert sum(b["source_hidden"] for b in payload["bundles"]) == 3
    assert payload["texts"][0]["hidden"]
    assert any(b["layer"] == "VISIBLE" and b["source_hidden"] for b in payload["bundles"])
    assert payload["shown"] == payload["counts"] and not any(payload["dropped"].values())


def test_empty_or_all_hidden_keeps_objects_without_auto_enabling():
    _boot()
    doc = ezdxf.new()
    doc.layers.new("OFF").off()
    doc.modelspace().add_line((100,200),(300,400),dxfattribs={"layer":"OFF"})
    payload = display_payload([], {}, "unused", source=source_display(doc))
    assert payload["source_all_hidden"]
    assert payload["counts"]["segs"] == 1
    assert all(b["source_hidden"] for b in payload["bundles"])


def test_single_read_cache_and_calculation_entities_unchanged(tmp_path, monkeypatch):
    from remote30_prototype import parse_dxf_for_view
    monkeypatch.setattr(sf, "CACHE_DIR", tmp_path/"cache")
    monkeypatch.setattr(sf, "_CACHE_ON", True)
    monkeypatch.setattr(sf, "CACHE_MIN_PARSE_S", 0)
    path = tmp_path/"drawing.dxf"
    visibility_doc().saveas(path)
    baseline = parse_dxf_for_view(path)
    with patch.object(ezdxf,"readfile", wraps=ezdxf.readfile) as read:
        first = sf.read_view(path, source_display=True)
        assert read.call_count == 1
    without_display = {k:v for k,v in first["parsed"].items() if k != "source_display"}
    assert without_display == baseline
    with patch.object(ezdxf,"readfile", side_effect=AssertionError("cache must not reread DXF")):
        again = sf.read_view(path, source_display=True)
    assert again["how"] == "cache" and again["parsed"] == first["parsed"]
    monkeypatch.setattr(sf, "FP_H_VIEW", "new-renderer")
    assert sf.read_view(path, source_display=True)["how"] != "cache"


def test_fast_stripped_source_matches_original(tmp_path, monkeypatch):
    from test_module_f_sub_fastread import _synthetic
    monkeypatch.setattr(sf,"_CACHE_ON",False)
    path = _synthetic(tmp_path/"used-and-unused.dxf")
    got = sf.read_view(path,source_display=True)
    assert got["how"] == "strip"
    assert got["parsed"]["source_display"] == source_display(ezdxf.readfile(path))


def test_block_colour_and_native_text_placement():
    doc = ezdxf.new()
    block = doc.blocks.new("text")
    block.add_text("65",dxfattribs={"height":5,"insert":(10,0),"width":0.7,"color":0})
    doc.modelspace().add_blockref("text",(100,200),dxfattribs={"color":1,"rotation":90,"xscale":2,"yscale":2})
    text = source_display(doc)["texts"][0]
    assert [text["x"],text["y"],text["height"],text["rotation"],text["width_factor"]] == pytest.approx([100,220,10,90,0.7])
    assert text["color"] == 1


def test_saved_cad_view_uses_target_and_does_not_crop():
    from src.pipenet_converter.render.dxf_source import saved_top_view
    doc=ezdxf.new()
    port=doc.viewports.get_config('*Active')[0]
    port.dxf.center=(-800,20)
    port.dxf.target=(1000,200,99)
    port.dxf.height=100
    port.dxf.aspect_ratio=2
    assert saved_top_view(doc)==dict(minx=100,maxx=300,miny=170,maxy=270)
    port.dxf.view_twist=30
    assert saved_top_view(doc) is None
    port.dxf.view_twist=0
    port.dxf.direction=(1,1,1)
    assert saved_top_view(doc) is None

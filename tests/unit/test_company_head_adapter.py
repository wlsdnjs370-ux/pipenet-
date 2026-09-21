"""Actual-fill DXF adapter and enabled/disabled company-rule integration."""
from dataclasses import replace
import math
from pathlib import Path
from types import SimpleNamespace

import ezdxf
import pytest

from src.pipenet_converter.dxf.head_symbol_adapter import WorldSymbolAdapter, apply_company_detection
from src.pipenet_converter.dxf.head_symbols import load_profile
from test_company_head_symbols import symbol, transform

ROOT = Path(__file__).resolve().parents[2]
PROFILE=load_profile(ROOT/"configs/head_symbols/company_260404.json")


def scene(tmp_path, variant, angle=0):
    c=transform(symbol(variant),angle,30)
    doc=ezdxf.new()
    m=doc.modelspace()
    m.add_circle(c.center_mm,c.radius_mm,dxfattribs={"layer":"HEAD"})
    if c.fill=="solid":
        hatch=m.add_hatch(dxfattribs={"layer":"HEAD"})
        hatch.paths.add_edge_path().add_arc(c.center_mm,c.radius_mm,0,360)
    path=tmp_path/"reference.dxf"
    doc.saveas(path)
    cx,cy=c.pipe_anchor_mm
    axis=c.pipe_axis
    pipe=((cx-axis[0]*120,cy-axis[1]*120),(cx+axis[0]*120,cy+axis[1]*120))
    segs=[("PIPE",7,*pipe)] + [(s.layer,7,s.start,s.end) for s in c.marks]
    vertices=c.vertices_mm
    tri=[(a,b) for a,b in zip(vertices,vertices[1:]+vertices[:1])] if vertices else []
    segs += [("HEAD",7,a,b) for a,b in tri]
    world=SimpleNamespace(segs=segs,circles=[] if vertices else [("HEAD",7,*c.center_mm,c.radius_mm)],
        arcs=[(a.layer,7,*a.center,a.radius) for a in c.arcs],
        arc_ang=[(a.start_degrees,(a.end_degrees-a.start_degrees)%360) for a in c.arcs],
        _source_path=str(path))
    head={"c":c.center_mm,"bundle":("HEAD",7)}
    head.update({"tri_side":c.radius_mm*math.sqrt(3),"tri_segs":tri} if vertices else {"head_r":c.radius_mm})
    spec={"material_picks":[["PIPE",7]],"heads":[{"bundle":["HEAD",7],"r":30}],"head_symbol_profile":"company_260404"}
    return world,spec,head


@pytest.mark.parametrize("variant,orientation", [("upright","upright"),("pair","upright_pendent"),
    ("pendent","pendent"),("pair_offset","upright_pendent"),("hot","pendent"),("sidewall","sidewall")])
@pytest.mark.parametrize("angle",[0,37,90])
def test_selected_world_geometry_and_real_dxf_fill(tmp_path,variant,orientation,angle):
    world,spec,head=scene(tmp_path,variant,angle)
    before=repr((world.segs,world.circles,head))
    got=WorldSymbolAdapter(world,spec,PROFILE).classify(head)
    assert got.status=="matched",got
    assert got.orientation==orientation
    assert got.temperature_c==(103 if variant=="hot" else 72)
    assert repr((world.segs,world.circles,head))==before


def test_missing_fill_source_is_review_and_never_assumed_empty(tmp_path):
    world,spec,head=scene(tmp_path,"pendent")
    world._source_path=None
    got=WorldSymbolAdapter(world,spec,PROFILE).classify(head)
    assert got.status=="review" and "unknown_body_fill" in got.reasons


def test_sidewall_retains_semantics_without_automatic_pendent_conversion(tmp_path):
    world,spec,head=scene(tmp_path,"sidewall")
    got=WorldSymbolAdapter(world,spec,PROFILE).classify(head)
    result=apply_company_detection(dict(head,kind="하향식"),got)
    assert result["kind"]=="미지정"
    assert result["company_symbol"]["orientation"]=="sidewall"
    assert result["company_symbol"]["calculation_review"]


def test_explicit_kind_conflict_keeps_manual_or_layer_decision(tmp_path):
    world,spec,head=scene(tmp_path,"pendent")
    got=WorldSymbolAdapter(world,spec,PROFILE).classify(head,explicit_orientation="upright")
    result=apply_company_detection(dict(head,kind="상향식"),got)
    assert result["kind"]=="상향식"
    assert result["company_symbol"]["status"]=="review"


def test_stage11_calls_profile_and_keeps_legacy_when_disabled(tmp_path,monkeypatch):
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import flow,heads
    world,spec,head=scene(tmp_path,"pendent")
    monkeypatch.setattr(heads,"collect_head_clusters",lambda *a:[head])
    monkeypatch.setattr(heads,"split_head_circles",lambda *a:([head],[],{}))
    monkeypatch.setattr(flow,"classify_head_kind",lambda *a,**k:"상향식")
    st={"w":world,"spec":spec,"knobs":{"small_len":1000},"mat_bundles":[("PIPE",7)]}
    got=flow.stage11_classify_heads(st,arm_index={},owned_half_arc_keys={})
    assert got[0]["kind"]=="하향식" and got[0]["company_symbol"]["temperature_c"]==72
    spec.pop("head_symbol_profile")
    legacy=flow.stage11_classify_heads(st,arm_index={},owned_half_arc_keys={})
    assert legacy[0]["kind"]=="상향식" and "company_symbol" not in legacy[0]

# -*- coding: utf-8 -*-
"""[오너 2026-09-22 · 그림 37~39] 계통도 · 기계실 칸 도면 읽기 — 빠르게, 결과는 그대로.

  「계통도 파트에 B1F 도면을 로드할 때, 너무 시간이 오래걸리는데, 로직 좀 최적화를
   시켜주겠어?」

  ⑴ 안 쓰는 블록 정의만 비운다 — 중첩 블록 · 치수 블록 · 지시선 화살표(치수
     스타일이 번호로 가리킴) · 멀티지시선 내용 블록 · 기본 화살표(_ClosedFilled —
     파일 안에 가리키는 글자가 없고 ezdxf 가 이름으로 찾는다)는 남는다.
  ⑵ 비운 사본으로 읽은 결과 = 원본으로 읽은 결과 (합성 도면 · 실제 도면).
     남긴 블록을 하나라도 비우면 결과가 달라진다 — 시험이 헛돌지 않는다.
  ⑶ 모르면 비우지 않는다 — 모르는 도형 종류 · 이진 DXF · 줄끝 섞임 · 블록 짝.
  ⑷ 같은 도면은 기억 — 꺼낸 것 = 새로 읽은 것 · 코드 판번호가 바뀌면 안 쓴다 ·
     깨진 기억은 버리고 새로 읽는다.
  ⑸ 계통도 칸 열기 — 도면(world)이 ★추적보다 먼저 앉고(★칸 «고르는 중»),
     ★추적이 끝나면 표가 채워진다. 두 번째 열기는 표 · ★추적 전부 기억에서.
"""
from __future__ import annotations

import os
from pathlib import Path

import pytest

ezdxf = pytest.importorskip("ezdxf")

from routes.module_f import sub_fastread as sf  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
UNUSED = {"unused0", "unused1", "unused2"}


def _synthetic(path: Path, *, region: bool = False, fmt: str = "asc") -> Path:
    """쓰는 블록 · 안 쓰는 블록이 섞인 작은 도면. `region` — 모르는 종류(REGION)를 끼운다."""
    from ezdxf.math import Vec2
    from ezdxf.render import mleader
    doc = ezdxf.new("R2018", setup=True)
    msp = doc.modelspace()
    nested = doc.blocks.new("NESTED")
    nested.add_circle((0, 0), 50, dxfattribs={"layer": "6-소화-헤드"})
    used = doc.blocks.new("USED")
    used.add_line((0, 0), (100, 0), dxfattribs={"layer": "6-소화-가지관"})
    used.add_blockref("NESTED", (100, 0))
    arrow = doc.blocks.new("MYARROW")
    arrow.add_line((0, 0), (-1, 0.3), dxfattribs={"layer": "ARW"})
    arrow.add_line((0, 0), (-1, -0.3), dxfattribs={"layer": "ARW"})
    mlb = doc.blocks.new("MLBLOCK")
    mlb.add_circle((0, 0), 30, dxfattribs={"layer": "MLB"})
    for k in range(3):                     # 도면 어디에도 안 놓이는 큰 블록
        blk = doc.blocks.new(f"UNUSED{k}")
        for i in range(3000):
            blk.add_line((i, 0), (i, 10), dxfattribs={"layer": "A-가구"})
    msp.add_line((0, 0), (5000, 0), dxfattribs={"layer": "6-소화-SP-메인"})
    msp.add_blockref("USED", (1000, 1000))
    msp.add_linear_dim(base=(0, 500), p1=(0, 0), p2=(5000, 0), dimstyle="EZDXF").render()
    ds = doc.dimstyles.duplicate_entry("Standard", "ARROWY")
    ds.set_arrows(ldrblk="MYARROW")        # 지시선 화살표 = 사용자 블록(번호로 가리킴)
    msp.add_leader([(0, 0), (100, 100), (200, 100)], dimstyle="ARROWY",
                   dxfattribs={"layer": "LDR"})
    ml = msp.add_multileader_block(style="Standard")
    ml.set_content(name="MLBLOCK")         # 내용 블록 · 화살표는 기본(_ClosedFilled)
    ml.add_leader_line(mleader.ConnectionSide.left, [Vec2(-300, 0)])
    ml.build(insert=Vec2(2000, 2000))
    doc.saveas(path, fmt=fmt)
    if region:                             # ezdxf 는 빈 3D 도형을 안 써서 글자로 끼운다
        raw = path.read_bytes()
        i = raw.find(b"\nENTITIES\n") + len(b"\nENTITIES\n")
        path.write_bytes(raw[:i] + b"  0\nREGION\n  5\nFFF1\n100\nAcDbEntity\n  8\n"
                         b"SOLID\n100\nAcDbModelerGeometry\n 70\n1\n" + raw[i:])
    return path


def _view(path):
    from remote30_prototype import parse_dxf_for_view
    return parse_dxf_for_view(str(path), include_hidden_layers=True)


def _read_stripped(tmp_path, raw, plan, name="copy.dxf"):
    p = tmp_path / name
    p.write_bytes(sf.strip_bytes(raw, plan))
    return _view(p)


@pytest.fixture()
def no_memory(monkeypatch):
    monkeypatch.setattr(sf, "_CACHE_ON", False)


def test_plan_drops_only_unused_blocks(tmp_path):
    raw = _synthetic(tmp_path / "s.dxf").read_bytes()
    plan, why = sf.plan_strip(raw)
    assert why == "ok"
    names = {bk[0].decode() for bk in plan["blocks"]}
    assert {bk[0].decode() for bk in plan["drop"]} == UNUSED
    assert {"used", "nested", "myarrow", "mlblock", "_closedfilled"} <= names - UNUSED
    assert any(n.startswith("*d") for n in names - UNUSED)      # 치수 블록도 남는다
    assert plan["drop_bytes"] > 0.5 * len(raw)


def test_stripped_copy_reads_identically_and_the_test_bites(tmp_path):
    src = _synthetic(tmp_path / "s.dxf")
    raw = src.read_bytes()
    plan, _ = sf.plan_strip(raw)
    whole = _view(src)
    assert _read_stripped(tmp_path, raw, plan) == whole
    # 남긴 블록을 하나라도 비우면 결과가 달라진다 — 남긴 까닭이 실제로 있다.
    for victim in ("myarrow", "mlblock", "nested", "_closedfilled"):
        bad = dict(plan)
        bad["drop"] = plan["drop"] + [bk for bk in plan["blocks"]
                                      if bk[0].decode() == victim]
        assert _read_stripped(tmp_path, raw, bad, f"bad_{victim}.dxf") != whole, victim


def test_read_view_uses_the_copy_and_matches(tmp_path, no_memory):
    src = _synthetic(tmp_path / "s.dxf")
    got = sf.read_view(src)
    assert got["how"] == "strip" and got["drop"] == 3
    assert got["parsed"] == _view(src)
    assert got["entities"] is got["parsed"]["entities"]
    assert "비우고 읽었습니다" in sf.describe(got)


def test_unknown_or_odd_files_are_read_as_is(tmp_path, no_memory):
    # 모르는 도형 종류 — 비우지 않고 원본 그대로
    src = _synthetic(tmp_path / "region.dxf", region=True)
    plan, why = sf.plan_strip(src.read_bytes())
    assert plan is None and "REGION" in why
    got = sf.read_view(src)
    assert got["how"] == "plain" and got["parsed"] == _view(src)
    assert "원본 그대로" in sf.describe(got)
    # 이진 DXF
    plan, why = sf.plan_strip(_synthetic(tmp_path / "b.dxf", fmt="bin").read_bytes())
    assert plan is None and why == "이진 DXF"
    raw = _synthetic(tmp_path / "s.dxf").read_bytes()
    # 줄끝이 섞였다(홀로 선 CR)
    plan, why = sf.plan_strip(raw[:200] + b"\r" + raw[200:])
    assert plan is None and why == "줄끝이 섞여 있음"
    # 블록 짝이 안 맞는다(ENDBLK 하나를 지움)
    i = raw.find(b"\nENDBLK")
    plan, why = sf.plan_strip(raw[:i] + b"\nENDBLX" + raw[i + 7:])
    assert plan is None


def test_memory_roundtrip_code_change_and_broken_file(tmp_path, monkeypatch):
    monkeypatch.setattr(sf, "CACHE_DIR", tmp_path / "cache")
    monkeypatch.setattr(sf, "_CACHE_ON", True)
    monkeypatch.setattr(sf, "CACHE_MIN_PARSE_S", 0.0)
    src = _synthetic(tmp_path / "s.dxf")
    first = sf.read_view(src)
    assert first["how"] == "strip"
    again = sf.read_view(src)
    assert again["how"] == "cache" and again["parsed"] == first["parsed"]
    assert "전에 읽어 둔" in sf.describe(again)
    # 같은 내용이면 파일 이름이 달라도 같은 기억이다
    other = tmp_path / "다른이름.dxf"
    other.write_bytes(src.read_bytes())
    assert sf.read_view(other)["how"] == "cache"
    # 코드 판번호가 바뀌면 기억을 쓰지 않는다 — 옛 판 기억은 지운다
    monkeypatch.setattr(sf, "FP_VIEW", "0" * 16)
    fresh = sf.read_view(src)
    assert fresh["how"] == "strip" and fresh["parsed"] == first["parsed"]
    assert len(list((tmp_path / "cache").glob("subview_*"))) == 1
    # 깨진 기억은 버리고 새로 읽는다
    for p in (tmp_path / "cache").glob("subview_*"):
        p.write_bytes(b"broken")
    got = sf.read_view(src)
    assert got["how"] == "strip" and got["parsed"] == first["parsed"]


def test_memory_keeps_at_most_the_limit(tmp_path, monkeypatch):
    monkeypatch.setattr(sf, "CACHE_DIR", tmp_path / "cache")
    monkeypatch.setattr(sf, "_CACHE_ON", True)
    monkeypatch.setattr(sf, "CACHE_KEEP", 2)
    for k in range(4):
        assert sf._cache_store("view", f"{k:032x}", sf.FP_VIEW, ([k], {}))
    assert len(list((tmp_path / "cache").glob("subview_*"))) == 2


def test_trace_memory_roundtrip(tmp_path, monkeypatch):
    from routes.module_f.sub_trace import layer_rows, plan_trace
    monkeypatch.setattr(sf, "CACHE_DIR", tmp_path / "cache")
    monkeypatch.setattr(sf, "_CACHE_ON", True)
    got = sf.read_view(_synthetic(tmp_path / "s.dxf"))
    rows = layer_rows(got["entities"], got["parsed"])
    tr = plan_trace(got["entities"], rows)
    assert not sf.store_trace(got["key"], rows, tr, seconds=0.01)   # 금방 끝난 것은 안 남김
    assert sf.load_trace(got["key"]) is None
    tr["payload"] = {"화면용": 1}                                    # 화면이 붙이는 칸은 안 남김
    assert sf.store_trace(got["key"], rows, tr, seconds=5.0)
    rows2, tr2 = sf.load_trace(got["key"])
    assert rows2 == rows
    assert "payload" not in tr2
    assert tr2 == {k: v for k, v in tr.items() if k != "payload"}


def test_system_open_draws_before_trace_then_fills(tmp_path, monkeypatch):
    from routes.module_f import api_slot, jobs, sub_trace
    monkeypatch.setattr(sf, "CACHE_DIR", tmp_path / "cache")
    monkeypatch.setattr(sf, "_CACHE_ON", True)
    monkeypatch.setattr(sf, "CACHE_MIN_PARSE_S", 0.0)
    monkeypatch.setattr(sf, "CACHE_MIN_TRACE_S", 0.0)
    src = _synthetic(tmp_path / "s.dxf")
    seen = []
    real = sub_trace.plan_trace

    def spy(entities, rows):
        w = box["sess"].get("world")
        seen.append((w is not None,
                     ((w or {}).get("sub_layers") or {}).get("trace", {}).get("mode")))
        return real(entities, rows)

    monkeypatch.setattr(sub_trace, "plan_trace", spy)
    box = {"sess": jobs._new_session(slot="system")}
    api_slot._sub_open_job(box["sess"], src, "system")()
    first = box["sess"]
    # 도면이 ★추적보다 먼저 앉았고, 그동안 ★칸은 «고르는 중» 이었다
    assert seen == [(True, "pending")]
    assert first["world"]["sub_layers"]["trace"] == sub_trace.public_trace(first["sub_trace"])
    assert first["world"]["sub_layers"]["trace"]["mode"] in ("pipe", "auto")
    # 두 번째 — 표 · ★추적 전부 기억에서(★추적을 다시 만들지 않는다)
    box["sess"] = jobs._new_session(slot="system")
    api_slot._sub_open_job(box["sess"], src, "system")()
    second = box["sess"]
    assert len(seen) == 1
    assert second["world"]["sub_layers"] == first["world"]["sub_layers"]
    assert second["entities"] == first["entities"]
    for k in ("mode", "layers", "reason", "entities", "graph"):
        assert second["sub_trace"].get(k) == first["sub_trace"].get(k), k
    # 기계실 칸은 표 · ★추적이 없다 — 도면만 (읽기는 같은 길)
    box["sess"] = jobs._new_session(slot="machineroom")
    api_slot._sub_open_job(box["sess"], src, "machineroom")()
    assert "sub_layers" not in box["sess"]["world"]
    assert box["sess"]["entities"] == first["entities"]


MF004 = ROOT / "data/uploads/MF-004_005 소화설비 계통도.dxf"
B1F = ROOT / "data/uploads/B1F 현장조사 소화설비 평면도.dxf"


@pytest.mark.skipif(not MF004.is_file(), reason="MF-004 계통도가 없는 PC")
def test_real_system_drawing_reads_identically(no_memory):
    got = sf.read_view(MF004)
    assert got["how"] == "strip"
    assert got["parsed"] == _view(MF004)


@pytest.mark.skipif(not (B1F.is_file() and os.environ.get("MODULE_F_BIG_DXF")),
                    reason="B1F(116 MB) 비교는 MODULE_F_BIG_DXF=1 일 때만 (약 40초)")
def test_b1f_reads_identically(no_memory):
    got = sf.read_view(B1F)
    assert got["how"] == "strip" and got["drop"] == 1246 and got["keep"] == 286
    assert got["parsed"] == _view(B1F)

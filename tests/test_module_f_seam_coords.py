# -*- coding: utf-8 -*-
"""[이음매 좌표] 세 도면이 한 점에서 만나는가 — `ModuleF_이음매좌표_수정지시서.md`.

■ 무엇이 문제였나 (대명동 3장 · 펌프 가압 · `scripts/_merge_seam_probe.py`)

  결합은 **라벨**로 선다. 그래서 어느 부위가 통째로 다른 좌표계에 남아도
  연결성분·고아 참조·Input 개수 검사는 **전부 통과한다** — 종전의 결합 뒤
  검사에 좌표가 하나도 없었던 이유이자, 못 잡던 이유다.

  실측으로 잡힌 것은 하나였다. `insert_source_pump` 가 수원 뒤에 만드는
  `{수원}_pd` 절점이 **`parts` 를 세운 뒤에** 생긴다 — 어느 목록에도 없으니
  굽는 자리의 기본값 `of.get(lab, "plan")` 이 받아 평면 식을 태우고, 평면
  식만 표고 × 1,000 lift 를 얹으므로 지하 −103.6 m 짜리 절점이 화면 밖으로
  날아갔다.

      m1_pd 표고 −103.633 m  →  이음매 배관 m1 이 1,480 → 97,929 (66배)

  조치 전 / 후::

      unclassified          1 (m1_pd)  →  0
      미분류 탓의 거짓 seam   1 (m1)     →  0
      seam_check            —          →  anchor 0.0 · pump 5.7e-13
      layout_status         (없음)      →  {machineroom: ok, riser: ok}

■ 여기서 지키는 것

  ⑴ 결합 뒤에 생긴 절점은 **이웃에게 물어** 제자리에 넣는다 (이름 규칙 아님).
  ⑵ 그래도 모르면 «평면» 으로 조용히 덮지 않는다 — 회전만 하고 보고한다.
  ⑶ 좌표 배치가 폴백으로 떨어지면 그 사실이 표·화면·산출에 남는다.
  ⑷ 투영 세 식은 이음매에서 항등적으로 연속이다 — `seam_check` 가 회귀 감지기.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _fx(name):
    path = _ROOT / "tests" / "test_module_f_merge.py"
    spec = importlib.util.spec_from_file_location("_mf_seam_fx", str(path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return getattr(mod, name)


def _machineroom(x0=900000.0, y0=400000.0, n=4):
    labels = [f"m{i + 1}" for i in range(n)]
    nodes = [{"label": lab, "x": x0 + i * 1500.0, "y": y0 + i * 800.0,
              "elevation": 0.0,
              "io_node": "Input" if i == 0 else "No"}
             for i, lab in enumerate(labels)]
    pipes = [{"label": f"mp{i}", "in": labels[i], "out": labels[i + 1],
              "dia": 100, "length": 1.7} for i in range(n - 1)]
    return {"nodes": nodes, "pipes": pipes,
            "conn_node_label": labels[-1],
            "conn_xy": [nodes[-1]["x"], nodes[-1]["y"]],
            "plan_edges": [[x0, y0, x0 + 2000.0, y0]]}


def _merge(*, mode="lsp_gravity", pump=None, machineroom=None):
    from routes.module_f.merge import merge_network
    return merge_network(_fx("_sample")(), riser=_fx("_riser")(),
                         machineroom=machineroom, mode=mode, pump=pump,
                         source_drop_m=3.0 if pump else 0.0)


# ─────────────────────────────── E1 · 결합 뒤에 생긴 절점
def _combined(nodes, pipes):
    class _C:
        pass
    c = _C()
    c.nodes = nodes
    c.pipes = pipes
    return c


def test_E1_펌프_토출_절점이_이웃을_따라간다():
    """★대명동에서 이것 하나가 이음매를 66배로 벌렸다."""
    from routes.module_f.merge import adopt_late_nodes
    c = _combined([{"label": "m1"}, {"label": "m1_pd"}, {"label": "1"}],
                  [{"in": "m1", "out": "m1_pd"},
                   {"in": "m1_pd", "out": "1"}])
    parts = {"plan": [], "system": ["1"], "machineroom": ["m1"]}
    got = adopt_late_nodes(c, parts, {"m1", "1"})
    assert got == [("m1_pd", "machineroom")], got
    assert "m1_pd" in parts["machineroom"]


def test_E1_이름_규칙이_아니라_연결로_정한다():
    """`_pd` 접미사를 파싱하면 엔진이 이름을 바꾸는 날 조용히 틀린다."""
    from routes.module_f.merge import adopt_late_nodes
    c = _combined([{"label": "1"}, {"label": "ZZZ"}],
                  [{"in": "1", "out": "ZZZ"}])
    parts = {"plan": [], "system": ["1"], "machineroom": []}
    assert adopt_late_nodes(c, parts, {"1"}) == [("ZZZ", "system")]


def test_E1_새_절점이_사슬로_이어져도_닿는다():
    """펌프가 여럿이면 새 절점끼리 붙는다 — 한 바퀴로는 안 닿는다."""
    from routes.module_f.merge import adopt_late_nodes
    c = _combined([{"label": "m1"}, {"label": "a"}, {"label": "b"}],
                  [{"in": "m1", "out": "a"}, {"in": "a", "out": "b"}])
    parts = {"plan": [], "system": [], "machineroom": ["m1"]}
    got = dict(adopt_late_nodes(c, parts, {"m1"}))
    assert got == {"a": "machineroom", "b": "machineroom"}, got


def test_E1_이웃이_없으면_남긴다():
    """★모르면 «평면» 으로 덮지 않는다 — 덮는 순간 표고 lift 를 탄다."""
    from routes.module_f.merge import adopt_late_nodes
    c = _combined([{"label": "1"}, {"label": "떠돌이"}], [])
    parts = {"plan": [], "system": ["1"], "machineroom": []}
    assert adopt_late_nodes(c, parts, {"1"}) == []
    assert "떠돌이" not in parts["system"] + parts["plan"]


def test_E1_결합_뒤_미분류가_0이다():
    got = _merge(mode="hsp_pump",
                 pump={"rated_q_lpm": 800.0, "rated_h_m": 60.0, "count": 2},
                 machineroom=_machineroom())
    ck = got["checks"]
    assert ck["unclassified_n"] == 0, ck["unclassified"]
    # 펌프 토출 절점이 실제로 생겼는지도 못박는다 — 안 생기면 시험이 헛돈다.
    labels = {str(n["label"]) for n in got["combined"].nodes}
    assert any(lab.endswith("_pd") for lab in labels), sorted(labels)[:8]


def test_E1_미분류는_표고_lift_를_타지_않는다():
    """★평면 식만 표고에 1,000 배를 얹는다 — 회귀하면 여기서 잡힌다."""
    from routes.module_f.merge import bake_combined_iso
    c = _combined([{"label": "10", "x": 0.0, "y": 0.0, "elevation": 0.0},
                   {"label": "떠돌이", "x": 100.0, "y": 0.0,
                    "elevation": -103.633}],
                  [])
    c.machine_room_plan_edges = []
    nodes, _e = bake_combined_iso({"combined": c,
                                   "parts": {"plan": ["10"], "system": [],
                                             "machineroom": []}})
    lone = next(n for n in nodes if n["label"] == "떠돌이")
    assert abs(lone["y"]) < 100.0, f"표고 lift 를 탔다: y={lone['y']}"


# ─────────────────────────────── E2 · 배치 폴백을 기록한다
def test_E2_정상이면_ok_라고_적는다():
    got = _merge(machineroom=_machineroom())
    st = got["layout_status"]
    assert st["riser"] == "ok" and st["machineroom"] == "ok", st


def test_E2_기계실이_없으면_해당없음이다():
    """«해당없음» 은 탈이 아니다 — 폴백과 갈라 적는다."""
    got = _merge()
    assert got["layout_status"]["machineroom"].startswith("해당없음")
    assert not any("좌표 배치" in s for s in got["steps"]), got["steps"]


def test_E2_폴백이면_단계기록과_표에_남는다():
    """조용히 원좌표에 남기지 않는다 — 산출물만 봐도 알아야 한다."""
    import remote30_full_network as rn
    src = (_ROOT / "core" / "remote30_full_network.py").read_text(
        encoding="utf-8")
    i = src.index('layout_status["riser"] = (')
    assert "DXF 원좌표에 남는다" in src[i:i + 260], src[i:i + 260]
    assert "layout_status" in rn.CombinedTables().__dict__
    j = src.index('meta.append((f"★좌표 배치')
    assert 'startswith(("폴백", "건너뜀"))' in src[max(0, j - 200):j]


# ─────────────────────────────── E3 · 좌표 검사
def test_E3_좌표_항목이_전부_실린다():
    got = _merge(mode="hsp_pump",
                 pump={"rated_q_lpm": 800.0, "rated_h_m": 60.0, "count": 2},
                 machineroom=_machineroom())
    ck = got["checks"]
    for key in ("bbox_span_mm", "bbox_ratio_to_plan", "part_bbox",
                "pump_seam", "unclassified", "layout_status"):
        assert key in ck, key
    assert set(ck["part_bbox"]) == {"plan", "system", "machineroom"}
    assert ck["part_bbox"]["plan"]["ratio_to_plan"] == 1.0


def test_E3_이음매_배관을_표_length_와_나란히_낸다():
    got = _merge(mode="hsp_pump",
                 pump={"rated_q_lpm": 800.0, "rated_h_m": 60.0, "count": 2},
                 machineroom=_machineroom())
    seam = got["checks"]["pump_seam"]
    assert seam and seam["table_m"] > 0
    # 좌표 거리 ÷ 1000 이 표 length 와 어긋나면 한쪽이 다른 좌표계에 있다.
    assert 0.5 < seam["ratio"] < 2.0, seam


def test_E3_판정하지_않는다():
    """★값만 낸다 — 예외로 올리면 돌던 결합이 통째로 죽는다."""
    src = (_ROOT / "routes" / "module_f" / "merge.py").read_text(
        encoding="utf-8")
    i = src.index("def check_combined(")
    seg = src[i:src.index("\ndef combined_summary(")]
    assert "raise" not in seg, "결합 뒤 검사가 예외를 올린다"


def test_E3_굽은_뒤_이음매_두_식이_같은_점이다():
    """지시서 §0 — 두 이음매는 항등적으로 연속이다. 아니면 식이 바뀐 것이다."""
    from routes.module_f.merge import bake_combined_iso
    got = _merge(mode="hsp_pump",
                 pump={"rated_q_lpm": 800.0, "rated_h_m": 60.0, "count": 2},
                 machineroom=_machineroom())
    rep: dict = {}
    bake_combined_iso(got, report=rep)
    sc = rep["seam_check"]
    assert set(sc) == {"anchor", "pump"}, sc
    assert sc["anchor"] < 1e-6 and sc["pump"] < 1e-6, sc


def test_E3_돌려주는_모양은_그대로다():
    """부르는 자리가 여럿이다 — `report` 는 곁가지여야 한다."""
    from routes.module_f.merge import bake_combined_iso
    got = _merge(machineroom=_machineroom())
    nodes, edges = bake_combined_iso(got)
    assert isinstance(nodes, list) and isinstance(edges, list)


# ─────────────────────────────── §5 금지 사항 — 세 식을 건드리지 않는다
def test_투영_세_식은_그대로다():
    """§0 이 이음매 연속성을 증명한 식이다 — 고치면 이음매가 찢어진다."""
    src = (_ROOT / "routes" / "module_f" / "merge.py").read_text(
        encoding="utf-8")
    i = src.index("    for n in nodes:\n        lab = str(n.get(\"label\"))")
    seg = src[i:i + 900]
    for line in ('n["x"] = a_iso[0] + (x - ax)',
                 'n["y"] = a_iso[1] + (y - ay)',
                 'n["x"] = rx + shift[0]',
                 'n["y"] = ry + shift[1]',
                 'n["y"] = ry + (float(n.get("elevation", 0) or 0) - e_ref) * lift'):
        assert line in seg, line


# ─────────────────────────────── 화면 — 값을 그대로 보인다
def _run_js(body, call):
    """★함수 하나를 떼어 **실제로 돌린다**. 구문 검사만으로는 스코프 밖
    이름(`kv` 같은 함수-지역 헬퍼) 참조를 못 잡는다 — 이 저장소가 이미
    ReferenceError 회귀로 값을 치른 자리다.
    """
    import json
    import subprocess
    src = ('function kv(k, v){ return "<div>" + k + "=" + v + "</div>"; }\n'
           + body + "\nconsole.log(JSON.stringify(" + call + "));\n")
    r = subprocess.run(["node", "-e", src], capture_output=True,
                       encoding="utf-8", errors="replace")
    assert r.returncode == 0, r.stderr
    return json.loads(r.stdout)


def _js_fn(name):
    js = (_ROOT / "static" / "module_f.js").read_text(encoding="utf-8")
    i = js.index(f"function {name}(")
    return js[i:js.index("\n  }\n", i) + 4]


def test_화면이_이음매_숫자를_보인다():
    body = _js_fn("mergeSeamLines")
    ck = {"combined": True, "anchor_gap": "0.0 mm · 표고차 0.000 m",
          "pump_seam": {"pipe": "m11", "coord_mm": 1988.6, "table_m": 1.88,
                        "ratio": 1.058},
          "part_bbox": {"plan": {"ratio_to_plan": 1.0},
                        "system": {"ratio_to_plan": 0.393},
                        "machineroom": {"ratio_to_plan": 0.277}},
          "bbox_ratio_to_plan": 1.086,
          "unclassified_n": 0, "unclassified": [],
          "layout_status": {"machineroom": "ok", "riser": "ok"}}
    out = _run_js(body, f"mergeSeamLines({__import__('json').dumps(ck)})")
    # 좌표 1988.6 mm → 1.989 m 와 표 1.88 m 를 **나란히** 보인다.
    for want in ("기준점 10 벌어짐", "m11", "1.989", "1.88", "0.393", "1.086"):
        assert want in out, (want, out)
    # 정상이면 경고를 띄우지 않는다 — 늑대 소년이 되면 아무도 안 본다.
    assert "warn" not in out, out


def test_화면이_미분류와_폴백을_경고한다():
    import json
    body = _js_fn("mergeSeamLines")
    ck = {"combined": True, "part_bbox": {}, "unclassified_n": 1,
          "unclassified": ["m1_pd"],
          "layout_status": {"riser": "폴백:KeyError 'x' — DXF 원좌표에 남는다"}}
    out = _run_js(body, f"mergeSeamLines({json.dumps(ck)})")
    assert "m1_pd" in out and "warn" in out, out
    assert "DXF 원좌표에 남는다" in out, out


def test_화면_판이_안_붙었으면_아무_말도_안_한다():
    body = _js_fn("mergeSeamLines")
    assert _run_js(body, "mergeSeamLines(null)") == ""
    assert _run_js(body, "mergeSeamLines({combined: false})") == ""


def test_자연낙차는_아무것도_안_바뀐다():
    """★펌프가 없으면 새 절점도 없다 — 이번 조치가 닿지 않아야 한다."""
    got = _merge(machineroom=_machineroom())
    assert got["checks"]["unclassified_n"] == 0
    assert not any("결합 뒤 생긴 절점" in s for s in got["steps"]), got["steps"]

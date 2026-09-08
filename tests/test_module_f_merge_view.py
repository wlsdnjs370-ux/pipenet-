# -*- coding: utf-8 -*-
"""[통합] 결합한 배관망을 **화면에서 볼 수 있다**.

■ 사용자 지적

  「통합 결합하고 나면 그 통합된 형태의 배관망(평면도·계통도·기계실 다 합친
  거)을 스크린에 미리 볼 수 있게 조치해줘. 모듈 A 그 통합처럼. 지금 건 뭐가
  나타나질 않고 있어.」

  맞다 — 결합 단계에는 캔버스에 그리는 갈래가 아예 없었다. 숫자(절점 308 ·
  배관 307)만 뜨는데, 그것으로는 세 도면이 제대로 이어졌는지 판단할 길이 없다.

■ 여기서 지키는 것

  ⑴ 결합망을 «어느 도면에서 왔는지» 와 함께 준다(평면도·계통도·기계실).
  ⑵ 두 도면을 잇는 배관은 «이음매» 로 따로 표시한다 — 결합의 핵심이 그 자리다.
  ⑶ 좌표는 **저장되는 그 좌표**(평면)다. 아이소는 보기 전용이라고 화면이 말한다.
  ⑷ 결합 전에는 오류가 아니라 «아직 없다» 로 답한다(붉은 줄을 남기지 않는다).
"""
from __future__ import annotations

import os

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _load(name):
    """옆 시험의 표본을 그대로 쓴다 — 표본이 갈리면 두 시험이 다른 것을 잰다.

    `tests` 는 패키지가 아니라서 `import tests.…` 가 안 된다(실측: 그렇게
    쓰다 ModuleNotFoundError). 파일 경로로 읽는다.
    """
    import importlib.util
    path = os.path.join(_ROOT, "tests", "test_module_f_merge.py")
    spec = importlib.util.spec_from_file_location("_mf_merge_fixtures", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return getattr(mod, name)


def _sample():
    return _load("_sample")()


def _riser(n: int = 4):
    return _load("_riser")(n)


def _merge(**kw):
    from routes.module_f.merge import merge_network
    return merge_network(_sample(), riser=_riser(), mode="lsp_gravity", **kw)


def test_어느_도면에서_왔는지_남는다():
    got = _merge()
    parts = got.get("parts") or {}
    assert set(parts) == {"plan", "system", "machineroom"}
    assert parts["plan"], "평면도 절점이 하나도 없다"
    assert parts["system"], "계통도 절점이 하나도 없다"


def test_두_망이_겹치는_것은_기준점_하나뿐이다():
    """★시험이 이것을 가르쳐 줬다 — 라벨 «10» 은 양쪽에 다 있다.

    특허 S740 이 평면도 라벨을 +9 해서 기준점이 10 이 되게 맞추기 때문이다.
    즉 그 한 점이 두 망의 이음매다. 그 밖에 겹치는 라벨이 있으면 그건 사고다
    (같은 이름이 다른 배관을 가리킨다).
    """
    got = _merge()
    parts = got["parts"]
    both = set(parts["plan"]) & (set(parts["system"])
                                 | set(parts["machineroom"]))
    assert both == {"10"}, both


def test_기계실이_없으면_그_칸은_빈다():
    got = _merge()
    assert got["parts"]["machineroom"] == []


# ─────────────────────────────── 라우트
def _client():
    import importlib
    import sys
    for p in (_ROOT, os.path.join(_ROOT, "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    c = srv.app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


def _sid(c):
    """세션 하나 — `_new_session` 은 dict 를 돌려준다(id 가 아니다)."""
    from routes.module_f.jobs import _new_session
    sess = _new_session()
    return sess["id"], sess


def test_결합_전에는_오류가_아니라_아직_없다다():
    """404 로 답하면 화면에 들어올 때마다 콘솔에 붉은 줄이 남는다."""
    c = _client()
    sid, _sess = _sid(c)
    r = c.get(f"/api/module-f/merge/preview?sid={sid}")
    assert r.status_code == 200, r.status_code
    d = r.get_json()
    assert d["ok"] is True and d["view"] is None
    assert "결합" in (d.get("message") or "")


def test_결합하면_그릴_것을_준다():
    c = _client()
    sid, sess = _sid(c)
    sess["merged"] = _merge()
    r = c.get(f"/api/module-f/merge/preview?sid={sid}")
    d = r.get_json()
    v = d["view"]
    assert v and v["nodes"] and v["pipes"]
    kinds = {n["part"] for n in v["nodes"]}
    assert kinds <= {"plan", "system", "machineroom"}
    assert "plan" in kinds and "system" in kinds
    assert d["counts"]["plan"] > 0 and d["counts"]["system"] > 0


def test_기준점을_이음매로_표시한다():
    """결합이 제대로 됐는지는 결국 그 한 점을 보고 판단한다."""
    c = _client()
    sid, sess = _sid(c)
    sess["merged"] = _merge()
    d = c.get(f"/api/module-f/merge/preview?sid={sid}").get_json()
    assert d["counts"]["anchor"] == ["10"]
    marked = [n for n in d["view"]["nodes"] if n.get("anchor")]
    assert [n["label"] for n in marked] == ["10"]


def test_두_도면을_잇는_배관은_이음매로_표시된다():
    c = _client()
    sid, sess = _sid(c)
    sess["merged"] = _merge()
    v = c.get(f"/api/module-f/merge/preview?sid={sid}").get_json()["view"]
    seam = [p for p in v["pipes"] if p["part"] == "seam"]
    assert seam, "이음매가 하나도 없다 — 두 망이 안 붙었다는 뜻이다"


def test_아이소는_보기_전용이고_기본이_아니다():
    """저장되는 좌표는 평면이다 — 기본이 아이소면 «보이는 것 ≠ 저장되는 것»."""
    c = _client()
    sid, sess = _sid(c)
    sess["merged"] = _merge()
    plain = c.get(f"/api/module-f/merge/preview?sid={sid}").get_json()
    iso = c.get(f"/api/module-f/merge/preview?sid={sid}&iso=1").get_json()
    assert plain["iso"] is False and iso["iso"] is True
    a = {n["label"]: (n["x"], n["y"]) for n in plain["view"]["nodes"]}
    b = {n["label"]: (n["x"], n["y"]) for n in iso["view"]["nodes"]}
    assert a != b, "아이소인데 좌표가 그대로다"


# ─────────────────────────────── 화면
def test_결합_단계가_캔버스에_그린다():
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    assert 'S.stage === "merge") { drawMerged(); }' in js, \
        "결합 단계에 그리는 갈래가 없다"
    i = js.index("  function drawMerged()")
    src = js[i:js.index("\n  }\n", i) + 4]
    for key in ("seam", "MERGE_COLOR", "mr_plan_edges"):
        assert key in src, key


def test_결합하면_곧바로_그린다():
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index('$("mg-build").onclick')
    src = js[i:js.index("\n  };\n", i)]
    assert "loadMergeView()" in src, "결합 뒤에 화면을 안 그린다"


def test_범례가_세_도면을_말한다():
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("  function renderMergeLegend(d)")
    src = js[i:js.index("\n  }\n", i) + 4]
    for word in ("평면도", "계통도", "기계실", "이음매"):
        assert word in src, word
    assert "보기 전용" in src, "아이소가 보기 전용이라는 말이 없다"


# ─────────────────────────────── 그림의 정확성 [2026-09-08 후속]
def _riser_uneven(n: int = 8):
    """길이가 서로 다른 입상관 — 균등 배치면 비례가 깨지는 것이 보이게."""
    labels = ["1"] + [f"n{i}" for i in range(2, n)] + ["10"]
    lens = [4.0, 0.3, 3.5, 0.5, 4.0, 0.2, 2.5][: n - 1]
    nodes = [{"label": lab, "x": 0, "y": i * 1000, "elevation": float(i)}
             for i, lab in enumerate(labels)]
    nodes[0]["io_node"] = "Input"
    pipes = [{"label": f"r{i}", "in": labels[i], "out": labels[i + 1],
              "dia": 100, "length": lens[i]} for i in range(n - 1)]
    return {"nodes": nodes, "pipes": pipes, "av_node_label": "10"}


def _merge_uneven():
    from routes.module_f.merge import merge_network
    return merge_network(_sample(), riser=_riser_uneven(),
                         mode="lsp_gravity")


def test_라이저_막대는_균등_간격이다():
    """★[2026-09-08 · 사용자 반려] 「통합쪽 배관망 디자인은 이전 버전으로.

    이전 디자인이 더 좋아.」 — 한때 막대 안 간격을 표 길이에 비례시켰다.
    실도면(라이저 49구간)에서는 0.017 m 구간이 사실상 사라지고 긴 구간만 남아
    막대가 한쪽으로 뭉쳤다. 균등 간격으로 되돌렸고, 되살아나면 여기서 잡힌다.

    ★길이의 권위는 표(선언 length)다 — 그림이 아니다. 좌표가 선언을 덮지
      못하게 하는 잠금은 `kfp_sdf_converter.parse_sdf` 에 따로 서 있다.
    """
    import math
    got = _merge_uneven()
    c = got["combined"]
    at = {str(n["label"]): (float(n["x"]), float(n["y"])) for n in c.nodes}
    sysset = set(got["parts"]["system"])
    seg = []
    for p in c.pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        if a in sysset and b in sysset and a in at and b in at:
            seg.append(math.dist(at[a], at[b]))
    assert len(seg) >= 3, seg
    assert max(seg) - min(seg) <= 1.0, f"간격이 균등하지 않다: {sorted(seg)}"


def test_라이저는_평면에서_수직_막대다():
    got = _merge_uneven()
    at = {str(n["label"]): float(n["x"]) for n in got["combined"].nodes}
    xs = {at[lab] for lab in got["parts"]["system"] if lab in at}
    assert len(xs) == 1, xs


def test_아이소에서도_라이저가_수직으로_남는다():
    """★사용자 지적: 「계통도도 수직으로 표현되어야 하는데 기울어져 있고」.

    schematic y 가 이미 수직인데 평면 회전을 그대로 먹이면 막대가 사선이
    된다(실측: x 퍼짐 1,732). 부위마다 맞는 투영을 쓴다 — 평면도는 회전+lift,
    라이저는 기준점의 아이소 위치에 평면 오프셋을 그대로 얹는다.
    """
    import math
    c = _client()
    sid, sess = _sid(c)
    sess["merged"] = _merge_uneven()
    v = c.get(f"/api/module-f/merge/preview?sid={sid}&iso=1").get_json()["view"]
    at = {n["label"]: (n["x"], n["y"]) for n in v["nodes"]}
    part = {n["label"]: n["part"] for n in v["nodes"]}
    xs = {round(at[lab][0], 6) for lab in at if part[lab] == "system"}
    assert len(xs) == 1, f"아이소에서 라이저가 기울었다: {sorted(xs)[:4]}"
    # 균등 간격도 아이소에서 그대로다(수직 평행이동은 길이를 안 바꾼다).
    seg = [math.dist(at[p["a"]], at[p["b"]]) for p in v["pipes"]
           if part.get(p["a"]) == "system" and part.get(p["b"]) == "system"]
    assert max(seg) - min(seg) <= 1.0, f"아이소에서 간격이 갈렸다: {sorted(seg)}"


def test_아이소에서_평면_헤드는_표고만큼_선다():
    """평면도 절점은 회전 + 표고 lift — 자는 평면과 같은 1 m = 1000."""
    c = _client()
    sid, sess = _sid(c)
    sess["merged"] = _merge_uneven()
    v = c.get(f"/api/module-f/merge/preview?sid={sid}&iso=1").get_json()["view"]
    at = {n["label"]: n for n in v["nodes"]}
    # 표본 fixture 의 평면 절점은 전부 표고 0 — lift 항이 0 이어야 한다.
    C30, S30 = 0.8660254037844387, 0.5
    plain = _merge_uneven()["combined"]
    at0 = {str(n["label"]): (float(n["x"]), float(n["y"]))
           for n in plain.nodes}
    for lab, n in at.items():
        if n["part"] != "plan" or lab not in at0:
            continue
        x, y = at0[lab]
        assert abs(n["x"] - (x - y) * C30) < 1e-6
        assert abs(n["y"] - (x + y) * S30) < 1e-6


# ─────────────────────────────── 아이소 산출 [2026-09-08 · 사용자]
def test_아이소_굽기는_한_함수뿐이다():
    """★화면과 파일이 각자 셈하면 «보이는 것 ≠ 저장되는 것» 이 된다.

    사용자 요청: 「저번처럼 나오던 아이소매트릭 형태 위상으로 .sdf 파일이
    출력되었으면 좋겠는데, 그거 되게 잘 그려졌어서.」 — 그래서 미리보기와
    산출이 `merge.bake_combined_iso` 하나를 같이 쓴다.
    """
    api = open(os.path.join(_ROOT, "routes", "module_f", "api_merge.py"),
               encoding="utf-8").read()
    assert api.count("bake_combined_iso(") >= 2, "두 자리가 같은 함수를 안 쓴다"
    assert "iso_nodes=iso_nodes" in api, "산출에 아이소 절점을 안 넘긴다"
    # 굽는 식이 라우트 안에 다시 있으면 안 된다.
    i = api.index("def module_f_merge_preview")
    seg = api[i:api.index("\n    @app.", i)]
    assert "0.8660254" not in seg, "라우트가 제 식으로 다시 굽는다"


def test_아이소_절점은_평면과_다른_자리다():
    """★«평면도» 절점으로 본다.

    라이저는 기준점의 아이소 자리에 평면 오프셋을 얹는 규칙이라, 기준점이
    원점에 있는 표본에서는 좌표가 그대로다(수학이 그렇다 — 결함이 아니다).
    굽혔는지 확인할 자리는 회전을 먹는 평면도 쪽이다.
    """
    got = _merge_uneven()
    from routes.module_f.merge import bake_combined_iso
    iso, _edges = bake_combined_iso(got)
    at0 = {str(n["label"]): (float(n["x"]), float(n["y"]))
           for n in got["combined"].nodes}
    plan = set(got["parts"]["plan"])
    moved = [n["label"] for n in iso
             if str(n["label"]) in plan
             and at0.get(str(n["label"])) != (float(n["x"]), float(n["y"]))]
    assert moved, "평면도 절점이 하나도 안 굽었다"


def test_아이소에서도_라이저는_수직이다():
    got = _merge_uneven()
    from routes.module_f.merge import bake_combined_iso
    iso, _edges = bake_combined_iso(got)
    sysset = set(got["parts"]["system"])
    xs = {round(float(n["x"]), 6) for n in iso if str(n["label"]) in sysset}
    assert len(xs) == 1, f"아이소에서 라이저가 기울었다: {sorted(xs)[:4]}"


def test_산출이_아이소_한_벌을_더_낸다():
    """모듈 A 의 `combined_<id>_iso.sdf` 와 같은 규약 — 좌표만 다른 사본."""
    src = open(os.path.join(_ROOT, "routes", "module_f", "emit.py"),
               encoding="utf-8").read()
    assert "iso_nodes" in src and "_iso" in src
    i = src.index("def emit_merged(")
    body = src[i:src.index("\ndef ", i + 10)]
    for key in ("sdf_iso", "kfp_iso", "has_iso"):
        assert key in body, key
    # ★아이소가 실패해도 본 산출(평면)은 버리지 않는다.
    assert "아이소 SDF 생성 실패" in body


def test_결합망이_없으면_굽지_않는다():
    from routes.module_f.merge import bake_combined_iso, merge_network
    got = merge_network(_sample(), mode="lsp_gravity")   # 계통도 없음
    assert bake_combined_iso(got) == ([], [])


# ─────────────────────────────── 그리는 손 [2026-09-08 · 사용자]
def test_결합망은_모듈A_물_팔레트를_쓴다():
    """★사용자: 「통합쪽 배관망 디자인은 이전 버전으로. 이전 디자인이 더 좋아.」

    «이전 버전» 은 예전부터 쓰던 모듈 A 통합 화면이다. 거기서 색은 «어느
    도면» 이 아니라 **물길의 상하류**를 말한다 — 상류(기계실)가 짙고
    하류(헤드)로 갈수록 옅어진다. 값이 갈리면 같은 망이 두 화면에서 다른
    그림이 되므로 모듈 A 의 WATER 를 그대로 쓴다.
    """
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    a = open(os.path.join(_ROOT, "templates", "remote30_prototype.html"),
             encoding="utf-8").read()
    for key in ("#0369a1", "#7dd3fc", "#22d3ee", "#f0f9ff", "#a5f3fc"):
        assert key in js, f"모듈 F 에 물 팔레트 {key} 가 없다"
        assert key in a, f"모듈 A 에 {key} 가 없다 — 전제가 깨졌다"
    i = js.index("const MERGE_COLOR = {")
    seg = js[i:i + 220]
    assert "WATER.deep" in seg and "WATER.spray" in seg, seg
    assert "machineroom: WATER.deep" in seg, "기계실이 최상류(짙은 물색)가 아니다"


def test_상류부터_그려_하류가_위에_남는다():
    """겹칠 때 헤드 쪽이 보여야 한다 — 모듈 A 와 같은 차례."""
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("  function drawMerged()")
    src = js[i:js.index("\n  }\n", i) + 4]
    order = '["machineroom", "system", "plan", "seam"]'
    assert order in src, src[:200]


def test_절점은_흰_외곽에_채움이다():
    """모듈 A 의 `_drawGraphNode` 와 같은 손 — 끝점만 크게·라벨."""
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("  function drawMergeNode(")
    src = js[i:js.index("\n  }\n", i) + 4]
    assert '"#ffffff"' in src and "endpoint_radius" in src
    assert "MERGE_STYLE.label_offset" in src

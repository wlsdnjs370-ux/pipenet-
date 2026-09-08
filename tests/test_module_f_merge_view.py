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

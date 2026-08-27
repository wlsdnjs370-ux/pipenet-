# -*- coding: utf-8 -*-
"""[F-8] 정찰 · 채택 — 「A 는 제안, 확정은 E 의 클릭」 계약.

이 파일이 지키는 것은 지시서 §0.2 의 권위 규칙 하나다:

    A 는 언제 불려도 항상 «제안» 으로만 들어온다.
    board 에 닿는 것은 언제나 E 의 확정 경로(클릭)뿐이다.

그래서 여기 검사는 「값이 맞나」보다 「어느 경로로 들어갔나」를 본다.
"""
from __future__ import annotations

import inspect
import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from routes.module_f import api_open, recon  # noqa: E402
from routes.module_f.slots import SESSION_KEYS  # noqa: E402


def _src(obj) -> str:
    return inspect.getsource(obj)


# ═══════════════════════════════════════════ F-8a. 정찰
# ── 띠 나누기 — 세는 쪽과 고르는 쪽이 같은 자를 쓴다
def test_띠_경계는_기본값_그대로():
    """D-F8-4 — 0.9 기본 채택 · 0.75~0.9 표시만 · 미만 접힘."""
    assert recon.CONF_HIGH == 0.9 and recon.CONF_MID == 0.75


@pytest.mark.parametrize("conf,band", [
    (1.0, recon.BAND_HIGH), (0.95, recon.BAND_HIGH), (0.9, recon.BAND_HIGH),
    (0.89, recon.BAND_MID), (0.75, recon.BAND_MID),
    (0.74, recon.BAND_LOW), (0.0, recon.BAND_LOW),
])
def test_띠_판정이_경계에서_흔들리지_않는다(conf, band):
    assert recon.band_of(conf) == band


def test_띠_이름은_suggest_가_쓰던_그대로():
    """바꾸면 그 응답을 읽던 화면이 조용히 «0개» 를 그린다."""
    assert recon.BAND_HIGH == "높음(≥0.9)"
    assert recon.BAND_MID == "중간(≥0.75)"
    assert recon.BAND_LOW == "낮음"


def test_띠_세기는_후보_수와_합이_같다():
    cands = [{"conf": c} for c in (1.0, 0.9, 0.88, 0.75, 0.7, 0.1)]
    bands = recon.count_bands(cands)
    assert sum(bands.values()) == len(cands)
    assert bands[recon.BAND_HIGH] == 2
    assert bands[recon.BAND_MID] == 2
    assert bands[recon.BAND_LOW] == 2


def test_묶음_수는_화면이_센_것을_읽는다():
    """여기서 다시 세면 분류 규칙이 두 벌이 되어 카드와 찍기가 달라진다."""
    got = recon.bundle_counts({"cats": {"PIPE": 4, "HEAD": 2, "ARCH": 9}})
    assert got == {"PIPE": 4, "HEAD": 2, "ALARM": 0}
    assert recon.bundle_counts(None) == {"PIPE": 0, "HEAD": 0, "ALARM": 0}


# ── 무해성 — 정찰 실패는 열기 실패가 아니다 (F-8a-4)
class _Boom(Exception):
    pass


@pytest.mark.parametrize("exc", [
    _Boom("A 를 못 불렀다"),
    ImportError("No module named 'remote30_prototype'"),
    SystemExit("엔진이 CLI 태생이라 이렇게 죽는 곳이 있다"),
])
def test_정찰이_어떻게_죽든_열기는_산다(monkeypatch, exc):
    def boom(*a, **k):
        raise exc
    monkeypatch.setattr(recon, "run_recon", boom)

    sess: dict = {}
    api_open._recon_into(sess, "x.dxf", {"cats": {}}, 1.0)   # 올라오면 실패다
    assert sess["recon"]["error"], "사유를 안 남겼다"
    assert type(exc).__name__ in sess["recon"]["error"]
    assert "heads" not in sess["recon"], "실패인데 후보가 있다"


def test_정찰이_성공하면_그대로_앉는다(monkeypatch):
    fake = {"bundles": {"PIPE": 3, "HEAD": 1, "ALARM": 0},
            "heads": [{"x": 1.0, "y": 2.0, "conf": 0.95,
                       "kind": "circle", "layer": "SP"}],
            "bands": {recon.BAND_HIGH: 1, recon.BAND_MID: 0, recon.BAND_LOW: 0},
            "elapsed_ms": 12}
    monkeypatch.setattr(recon, "run_recon", lambda *a, **k: fake)

    sess: dict = {}
    api_open._recon_into(sess, "x.dxf", {"cats": {}}, 1.0)
    assert sess["recon"] is fake


# ── 배선 — 평면도만 정찰한다 (D-F8-2)
def test_열기잡은_도면_종류를_keyword_로_받는다():
    """§3 — 공개 함수 시그니처 파괴 금지, keyword 인자 추가만."""
    sig = inspect.signature(api_open._open_job)
    p = sig.parameters["kind"]
    assert p.kind is inspect.Parameter.KEYWORD_ONLY
    assert p.default == "plan"


def test_정찰은_평면도일_때만_이어진다():
    body = _src(api_open._open_job)
    assert 'if str(kind) == "plan":' in body
    assert "_recon_into(sess, dxf, payload, t_open)" in body


def test_계통도_기계실은_찍기판_자체를_안_연다():
    """두 점 경로가 전부라 헤드 검출이 무의미하다 — 애초에 다른 잡이다."""
    from routes.module_f import api_slot
    body = _src(api_slot.register)
    assert '_open_job(sess, dxf, kind=kind) if kind == "plan"' in body
    assert "_sub_open_job(sess, dxf, kind)" in body


def test_새_도면을_열면_앞_도면의_정찰을_지운다():
    from routes.module_f import api_slot
    body = _src(api_slot.register)
    i = body.index('sess["method"] = None')
    assert '"recon", "suggest"' in body[i:i + 700]


# ── 순서 — world 가 정찰보다 먼저 앉는다
def test_도면은_정찰_전에_화면에_앉는다():
    """정찰이 도는 동안에도 도면이 보여야 한다(F-8a 수용기준)."""
    body = _src(api_open._open_job)
    assert body.index('sess["world"] = payload') < body.index("_recon_into")


# ── 인식은 한 벌 — suggest 와 열기가 같은 것을 쓴다
def test_suggest_는_정찰과_같은_함수를_쓴다():
    from routes.module_f import api_pick
    body = _src(api_pick.register)
    assert "from routes.module_f.recon import run_recon" in body
    assert 'run_recon(dxf, world=sess.get("world"), tag="제안")' in body
    # 옛 몸통이 남아 있으면 언젠가 한쪽만 고쳐진다
    assert "A.detect_heads(" not in body


def test_suggest_응답_모양은_종전_그대로():
    from routes.module_f import api_pick
    body = _src(api_pick.register)
    assert '"ok": True, "n": len(cands), "bands": rec["bands"]' in body
    assert '"candidates": cands' in body
    assert 'sess["suggest"] = cands' in body


# ── 세션 키 — 정찰은 도면별이다(슬롯에 갇힌다)
def test_정찰은_세션전역이_아니다():
    """슬롯 규약(H-0): SESSION_KEYS 여집합이 전부 도면별 상태다."""
    assert "recon" not in SESSION_KEYS
    assert "suggest" not in SESSION_KEYS


# ── 조회 라우트
def test_정찰_조회_라우트가_있다():
    body = _src(api_open.register)
    assert '@app.get("/api/module-f/recon")' in body


def test_조회는_기본으로_후보_좌표를_안_싣는다():
    """카드를 그릴 때마다 3천 점을 내려보내면 새로고침이 그만큼 무거워진다."""
    body = _src(api_open.register)
    i = body.index('@app.get("/api/module-f/recon")')
    seg = body[i:i + 900]
    assert 'request.args.get("heads")' in seg


@pytest.mark.parametrize("rec,state", [
    (None, "none"), ({}, "none"),
    ({"error": "ImportError: x"}, "error"),
    ({"bundles": {}, "bands": {}, "heads": [1, 2], "elapsed_ms": 5}, "ok"),
])
def test_조회_요약이_상태를_구분한다(rec, state):
    assert recon.recon_view(rec)["state"] == state


def test_조회_요약은_후보_수만_말한다():
    v = recon.recon_view({"bundles": {"PIPE": 2}, "bands": {},
                          "heads": [1, 2, 3], "elapsed_ms": 7})
    assert v["n"] == 3 and "heads" not in v


# ── 금지 목록 (§3) — 주입 경로가 생기지 않았는지
def test_정찰은_board_에_손대지_않는다():
    body = _src(recon)
    for banned in ("board.mat", "board.heads", "ps.board", ".click("):
        assert banned not in body, f"정찰이 board 경로를 만졌다: {banned}"


# ═══════════════════════════════════════════ F-8b. 채택
from routes.module_f import adopt  # noqa: E402


class _FakePS:
    """E 의 찍기판을 «토글» 성질만 남기고 흉내 낸다.

    실물 board 없이 확인할 수 있는 것은 여기까지다 — 서명 단위 토글을 채택이
    올바로 다루는가. 실도면 동일성은 scripts/_verify_module_f_adopt.py 가 잰다.
    """

    def __init__(self, shapes, *, sig=None):
        self.shapes = dict(shapes)      # (x, y) → 서명
        self.sig = sig or {}            # 켜져 있는 서명
        self.on: set = set()
        self.calls: list = []
        self.board = self
        self.by_bundle = {}
        self.mat_done = True
        self.head_label = "상향"

    # PickSession 쪽 표면
    def select_pipe(self):
        return True

    def complete_pipe(self):
        return True

    def set_slot(self, label):
        return True

    def click(self, x, y, max_d=None):
        self.calls.append((x, y))
        s = self.shapes.get((x, y))
        if s is None:
            return None                 # 찍을 도형이 없다 = 유령
        if s in self.on:
            self.on.discard(s)
            return {"동작": "취소"}
        self.on.add(s)
        return {"동작": "추가"}


def test_같은_서명을_두_번_채택해도_꺼지지_않는다():
    """토글을 모르면 후보가 서로를 꺼서 결과가 0 에 수렴한다."""
    ps = _FakePS({(0.0, 0.0): "A", (1.0, 1.0): "A", (2.0, 2.0): "B"})
    picks = [(0, {"x": 0.0, "y": 0.0, "conf": 1.0}),
             (1, {"x": 1.0, "y": 1.0, "conf": 1.0}),
             (2, {"x": 2.0, "y": 2.0, "conf": 1.0})]
    got = adopt.adopt_heads(ps, picks)

    assert ps.on == {"A", "B"}, "서명이 꺼진 채 끝났다"
    assert got["applied"] == 2 and got["already"] == 1
    assert got["skipped"] == []


def test_되켜기는_클릭_두_번으로_기록된다():
    """화면이 하는 그대로 — 채택은 클릭의 나열이지 별도 경로가 아니다."""
    ps = _FakePS({(0.0, 0.0): "A", (1.0, 1.0): "A"})
    got = adopt.adopt_heads(ps, [(0, {"x": 0.0, "y": 0.0}),
                                 (1, {"x": 1.0, "y": 1.0})])
    assert ps.calls == [(0.0, 0.0), (1.0, 1.0), (1.0, 1.0)]
    assert got["clicked"] == [[0.0, 0.0], [1.0, 1.0], [1.0, 1.0]]


def test_찍을_도형이_없으면_유령으로_남는다():
    """E 의 확정 게이트를 우회하지 않는다 — 실패는 실패로 센다."""
    ps = _FakePS({(0.0, 0.0): "A"})
    got = adopt.adopt_heads(ps, [(0, {"x": 0.0, "y": 0.0, "conf": 0.95}),
                                 (7, {"x": 9.0, "y": 9.0, "conf": 0.8})])
    assert got["applied"] == 1
    assert [s["i"] for s in got["skipped"]] == [7]
    g = got["skipped"][0]
    assert (g["x"], g["y"], g["conf"]) == (9.0, 9.0, 0.8)
    assert g["why"] == adopt.WHY_NO_SHAPE
    assert ps.on == {"A"}, "유령이 board 를 바꿨다"


def test_유령_좌표는_클릭_목록에_안_들어간다():
    """되짚어 재생할 때 없던 클릭이 끼면 동일성이 깨진다."""
    ps = _FakePS({(0.0, 0.0): "A"})
    got = adopt.adopt_heads(ps, [(0, {"x": 0.0, "y": 0.0}),
                                 (1, {"x": 9.0, "y": 9.0})])
    assert got["clicked"] == [[0.0, 0.0]]


def test_거리_상한이_없으면_남의_헤드를_찍는다():
    """`_click_head` 는 max_d=None 이면 거리와 무관하게 최근접을 받는다."""
    assert adopt.ADOPT_MAX_D_MM == 300.0
    sig = inspect.signature(adopt.adopt_heads)
    assert sig.parameters["max_d"].default == adopt.ADOPT_MAX_D_MM


def test_클릭에_상한을_실제로_넘긴다(monkeypatch):
    seen = []

    class _P(_FakePS):
        def click(self, x, y, max_d=None):
            seen.append(max_d)
            return super().click(x, y, max_d=max_d)

    adopt.adopt_heads(_P({(0.0, 0.0): "A"}), [(0, {"x": 0.0, "y": 0.0})],
                      max_d=123.0)
    assert seen == [123.0]


# ── 후보 고르기
def _cands():
    return [{"x": 0, "y": 0, "conf": 0.95}, {"x": 1, "y": 1, "conf": 0.9},
            {"x": 2, "y": 2, "conf": 0.8}, {"x": 3, "y": 3, "conf": 0.5}]


def test_문턱은_경계를_포함한다():
    got = adopt.select_heads(_cands(), conf_min=0.9)
    assert [i for i, _ in got] == [0, 1]


def test_번호로_고르면_문턱은_무시된다():
    got = adopt.select_heads(_cands(), conf_min=0.9, indices=[2, 3])
    assert [i for i, _ in got] == [2, 3]


def test_조건이_없으면_전부():
    assert len(adopt.select_heads(_cands())) == 4
    assert adopt.select_heads([]) == []


def test_걸러도_원래_번호가_그대로다():
    """번호를 다시 매기면 화면이 «몇 번이 유령인가» 를 후보 목록과 못 맞춘다.

    가운데가 빠져 번호에 구멍이 나는 경우로 본다 — 다시 매기면 [0,1] 이 된다.
    """
    cands = [{"conf": 0.95}, {"conf": 0.5}, {"conf": 0.9}, {"conf": 0.8}]
    assert [i for i, _ in adopt.select_heads(cands, conf_min=0.9)] == [0, 2]


# ── 재료 — /pick/auto 와 한 몸통
def test_재료_채택은_묶음_중점을_클릭한다():
    ps = _FakePS({})
    ps.by_bundle = {("SP", 1): [((0.0, 0.0), (10.0, 20.0))]}
    ps.shapes = {(5.0, 10.0): "SP1"}
    got = adopt.adopt_bundles(
        ps, {"bundles": [{"layer": "SP", "color": 1, "cat": "PIPE"}]}, "PIPE")
    assert ps.calls == [(5.0, 10.0)], "중점이 아니다"
    assert got["applied"] == ["SP"] and got["skipped"] == []


def test_선분이_없는_묶음은_건너뛴다():
    ps = _FakePS({})
    got = adopt.adopt_bundles(
        ps, {"bundles": [{"layer": "빈", "color": 7, "cat": "PIPE"}]}, "PIPE")
    assert got["applied"] == [] and got["skipped"] == ["빈"]


def test_pick_auto_가_같은_몸통을_쓴다():
    """두 길이 갈리면 「추천 일괄」과 「채택」이 다른 결과를 낸다."""
    from routes.module_f import api_pick
    body = _src(api_pick.register)
    assert "adopt_bundles(ps, world, want)" in body
    # 옛 몸통이 남아 있으면 안 된다
    assert "for b in targets:" not in body


# ── 라우트 계약
def test_채택_라우트가_있다():
    from routes.module_f import api_pick
    assert '@app.post("/api/module-f/pick/adopt")' in _src(api_pick.register)


def test_채택은_잡으로_돈다():
    """후보 수천 개 클릭은 무겁다 — 요청 스레드에서 돌리면 게이트웨이가 끊는다."""
    from routes.module_f import api_pick
    body = _src(api_pick.register)
    i = body.index('@app.post("/api/module-f/pick/adopt")')
    seg = body[i:i + 5200]
    assert '_run_job(sess, "인식 결과 채택", job)' in seg
    assert "_job_running(sess)" in seg


def test_채택도_재료_0개면_거부한다():
    """board 의 complete_materials 판정 그대로 — 우회로를 만들지 않는다."""
    from routes.module_f import api_pick
    body = _src(api_pick.register)
    i = body.index('@app.post("/api/module-f/pick/adopt")')
    assert "if not ps.complete_pipe():" in body[i:i + 5200]


def test_채택_응답이_계약대로다():
    from routes.module_f import api_pick
    body = _src(api_pick.register)
    i = body.index('@app.post("/api/module-f/pick/adopt")')
    seg = body[i:i + 5200]
    for k in ("mat_applied", "mat_skipped", "head_applied", "head_skipped",
              "skipped_heads"):
        assert f'"{k}"' in seg, f"응답에 {k} 가 없다"
    assert '"state"' in seg, "갱신된 찍기 상태를 안 돌려준다"


def test_정찰이_실패한_도면은_채택을_거절한다():
    from routes.module_f import api_pick
    body = _src(api_pick.register)
    i = body.index('@app.post("/api/module-f/pick/adopt")')
    assert 'rec.get("error")' in body[i:i + 5200]


# ── 금지 목록 (§3) — 채택도 클릭 경로뿐
def test_채택은_board_에_쓰지_않는다():
    body = _src(adopt)
    for banned in ("board.mat =", "board.heads =", "board.mat.append",
                   "board.heads.append", "board.clicks", "write_pick"):
        assert banned not in body, f"채택이 주입 경로를 만들었다: {banned}"


# ═══════════════════════════════════════════ 자동 화면 — 눈에 보이게
def test_검출_헤드는_빨강이다():
    """신뢰도 색(초록/노랑/회색)은 어두운 도면 위에서 티가 안 났다."""
    html = _script()
    assert 'const HEAD_MARK = "#ff3b30";' in html
    i = html.index("function drawAuto(dim)")
    seg = html[i:i + 500]
    assert "ctx.fillStyle = HEAD_MARK;" in seg
    assert "suggestColor" not in seg, "아직 신뢰도 색을 쓴다"


def test_추출_뒤에는_나머지_도면을_내린다():
    """뽑아낸 망이 드러나야 한다 — 다만 지우지는 않는다(어디서 뽑혔나)."""
    html = _script()
    i = html.index("function drawWorld(dim)")
    seg = html[i:i + 400]
    assert "ctx.globalAlpha = 0.16" in seg
    assert "ctx.setLineDash([2, 4])" in seg


def test_내린_뒤에는_원래대로_되돌린다():
    """알파·점선을 안 되돌리면 다음에 그리는 것이 전부 흐려진다."""
    html = _script()
    i = html.index("function drawWorld(dim)")
    seg = html[i:i + 1800]
    assert "if (dim) { ctx.globalAlpha = 1; ctx.setLineDash([]); return; }" in seg


def test_뽑은_배관망을_따로_그린다():
    html = _script()
    assert "function drawAutoNet()" in html
    i = html.index("function drawAutoNet()")
    seg = html[i:i + 900]
    assert "S.autoView" in seg
    assert "n.head" in seg and "n.input" in seg, "말단·급수 절점 표시가 없다"


def test_추출_전에는_안_내린다():
    """돌리기도 전에 도면이 흐려지면 무엇을 찍는지 안 보인다."""
    html = _script()
    i = html.index("const focus = S.stage")
    seg = html[i:i + 300]
    assert "S.autoDone && S.autoView" in seg


def test_추출_뒤_화면이_뽑은_자리로_맞춰진다():
    """흐리게 내리는 것만으로는 안 드러난다 — 도면 971m 대 설계면적 25m."""
    html = _script()
    assert "function autoNetBounds()" in html
    i = html.index("function curBounds()")
    seg = html[i:i + 700]
    assert 'S.stage === "auto" && S.autoDone' in seg, "「화면 맞춤」이 도면 전체로 간다"
    j = html.index("await loadAutoView();")
    assert "autoNetBounds()" in html[j:j + 600]


def test_뽑은_망을_실제로_받아_온다():
    html = _script()
    assert "async function loadAutoView()" in html
    i = html.index("async function loadAutoView()")
    assert "/api/module-f/auto/preview" in html[i:i + 400]


def test_추출_명칭이_바뀌었다():
    html = _script()
    assert ">배관망 추출<" in html
    assert "<b>5</b>최불리 추출" in html
    # 「추리」로 시작하는 활용형이 하나도 없어야 한다 — 단추만 고치고 안내
    # 문구를 두면 화면 안에서 이름이 둘로 갈린다.
    assert "추리" not in html, "옛 이름이 남아 있다"


def test_서버_로그도_같은_이름을_쓴다():
    from routes.module_f import api_auto
    src = _src(api_auto.register)
    assert "[자동] 최불리 추출 —" in src
    assert "추리기" not in src


def test_알람밸브_단계가_눈에_보인다():
    """`.card h2` 는 10px·faint 구분선이라 「① 알람밸브」가 안 보였다 —
    「알람밸브 지정 버튼이 어디 있는지 모르겠다」는 말을 실제로 들었다."""
    html = _script()
    assert 'id="au-s1"' in html
    assert "<b>1</b>알람밸브 (시작 노드)" in html
    i = html.index("  .step-h{")
    seg = html[i:i + 300]
    assert "12.5px" in seg, "제목이 여전히 작다"


def test_단추가_무엇을_찍는지_말한다():
    html = _script()
    i = html.index('id="au-anchor"')
    assert "알람밸브 찍기" in html[i:i + 120], "«도면에서 찍기» 로는 무엇인지 모른다"


def test_손질의_급수시작과_다름을_밝힌다():
    """모듈 E 의 물흐름/급수 시작과 혼동한다는 지적을 화면에서 직접 푼다."""
    html = _script()
    i = html.index('id="au-s1"')
    assert "급수 시작" in html[i:i + 500]


def test_끝낸_단계가_표시된다():
    """순서가 섞여 보인다는 지적 — «무엇을 이미 했나» 를 화면이 말한다."""
    html = _script()
    for sid in ("au-s1", "au-s2", "au-s3", "au-s4", "au-s5"):
        assert f'id="{sid}-mark"' in html
    i = html.index("const mark = (id, on, txt)")
    seg = html[i:i + 900]
    assert 'mark("au-s1"' in seg and 'mark("au-s5"' in seg


def test_범위_지정이_단계로_승격됐다():
    """어느 구역을 뽑을지가 결과를 가른다 — 접이식에 묻어 두면 안 된다."""
    html = _script()
    assert "<b>4</b>범위 지정" in html
    assert 'id="au-zone-draw"' in html
    # 접이식 잔재가 남으면 두 벌이 된다
    assert 'data-fold="au-zone-body"' not in html
    assert 'id="au-zone-body"' not in html


def test_범위는_안_그려도_되는_단계다():
    """필수로 보이면 사람이 «그려야만 되는 줄» 알고 멈춘다."""
    html = _script()
    i = html.index('id="au-s4"')
    seg = html[i:i + 700]
    assert "도면 전체" in seg
    j = html.index('mark("au-s4"')
    assert '"도면 전체"' in html[j:j + 300], "안 그렸을 때 표시가 없다"


# ── [S270 · S310] 배관망 검출 — 최불리를 고르기 «전» 의 단계
def test_배관망_검출이_3단계다():
    """논리 문서 순서: S270 가지치기 → S310 거리 → S315 범위 → S320 최불리."""
    html = _script()
    assert "<b>3</b>배관망 검출" in html
    assert 'id="au-network"' in html
    # 순서가 뒤집히면 «거리를 어디서 재는지» 를 못 보고 결과만 받는다
    assert html.index("<b>3</b>배관망 검출") < html.index("<b>4</b>범위 지정")
    assert html.index("<b>4</b>범위 지정") < html.index("<b>5</b>최불리 추출")


def test_거리_분포를_보여준다():
    """최불리는 «거리를 내림차순으로 자른 것» 이다 — 그 재료를 보여야 한다."""
    html = _script()
    i = html.index("function renderAutoNet()")
    seg = html[i:i + 1200]
    assert "거리 (밸브→헤드)" in seg
    for k in ("near_m", "mid_m", "far_m"):
        assert k in seg, k
    assert "도달 헤드" in seg


def test_S270_가지치기를_화면에서_고른다():
    """A 는 load_mode 를 기본 off 로 둔다 — 논리 문서는 켜는 것을 전제한다."""
    html = _script()
    assert 'id="au-prune"' in html
    i = html.index('id="au-prune"')
    assert "checked" in html[i:i + 60], "기본이 꺼져 있다"
    j = html.index('$("au-network").onclick')
    assert 'prune: $("au-prune").checked' in html[j:j + 400]


def test_망_검출은_서버가_S270을_켠다():
    from routes.module_f import api_auto
    src = _src(api_auto.register)
    i = src.index('@app.post("/api/module-f/auto/network")')
    seg = src[i:i + 2200]
    assert 'body.get("prune", True)' in seg, "기본이 꺼져 있다"
    assert "prune=prune" in seg


def test_아직_안_돌린_것은_오류가_아니다():
    """404 로 답하면 단계에 들어올 때마다 콘솔에 붉은 줄이 남는다 —
    진짜 오류가 그 사이에 묻힌다."""
    from routes.module_f import api_auto
    src = _src(api_auto.register)
    i = src.index('@app.get("/api/module-f/auto/network-view")')
    # 창을 길이로 자르면 다음 라우트까지 넘어간다 — 경계에서 끊는다.
    j = src.index("    @app.", i + 10)
    seg = src[i:j]
    assert '"summary": None, "view": None' in seg
    # 주석에도 «404» 라고 적혀 있다 — 실행되는 줄만 본다.
    code = "\n".join(ln for ln in seg.splitlines()
                     if not ln.lstrip().startswith("#"))
    assert ", 404)" not in code, "아직 안 돌린 것을 오류로 답한다"


def test_망_검출은_도달_헤드_전부로_돈다():
    """S320 앞까지가 이 단계다 — k 를 전부로 줘야 «물 닿는 망 전체» 가 나온다."""
    from routes.module_f import auto
    src = _src(auto.run_network)
    assert "k_all = max(1, len(cand))" in src
    assert "k=k_all" in src
    assert "load_mode=bool(prune)" in src


def test_검출망은_최불리와_다른_색이다():
    """둘이 같은 색이면 «무엇이 뽑힌 것» 인지 구분이 안 된다."""
    html = _script()
    i = html.index("function drawAutoNetwork()")
    seg = html[i:i + 600]
    assert "#60a5fa" in seg               # 검출망 — 파랑
    j = html.index("function drawAutoNet()")
    assert "#22d3ee" in html[j:j + 600]   # 최불리 — 청록


def test_검출망도_새_도면에서_지워진다():
    from routes.module_f import api_slot
    body = _src(api_slot.register)
    i = body.index('sess["method"] = None')
    assert '"auto_net"' in body[i:i + 800]
    html = _script()
    k = html.index("S.recon = null; S.suggest = null;")
    assert "S.autoNet = null" in html[k:k + 400]


def test_영역_무장은_체크박스_상태를_그대로_쓴다():
    """캔버스 드래그 판정이 그 값을 읽는다 — 단추만 만들고 끊으면 안 그려진다."""
    html = _script()
    assert 'id="au-zone-arm"' in html
    i = html.index('$("au-zone-draw").onclick')
    seg = html[i:i + 500]
    assert '$("au-zone-arm").checked = on;' in seg
    assert '$("au-zone-arm").checked' in html[html.index("const armed ="):
                                              html.index("const armed =") + 300]


# ═══════════════════════════════════════════ 되돌리기 — 모든 단계에서
def test_되돌리기가_단계마다_있다():
    """자동에서 Ctrl+Z 가 아무 일도 안 하던 것 — 「못 되돌리잖아」."""
    html = _script()
    i = html.index("async function undoStep()")
    seg = html[i:i + 1600]
    # 엔진이 기록을 들고 있는 단계는 그 단추를 그대로 누른다
    assert '$("pk-undo").click()' in seg and '$("ed-undo").click()' in seg
    # 자동·계통도는 화면이 쌓은 기록에서 되돌린다
    assert 'item.stage === "auto"' in seg
    assert 'item.stage === "sub"' in seg


def test_자동_되돌리기가_서버까지_되돌린다():
    """화면만 되돌리면 서버는 옛 알람밸브·영역을 그대로 들고 있다."""
    html = _script()
    i = html.index("async function undoStep()")
    seg = html[i:i + 1600]
    assert "/api/module-f/auto/anchor" in seg
    assert "/api/module-f/auto/zones" in seg


def test_되돌릴_자리마다_기록을_남긴다():
    html = _script()
    for label in ("알람밸브 찍기", "알람밸브 지우기", "영역 그리기",
                  "마지막 영역 지우기", "찍은 점 지우기"):
        assert f'markUndo("{label}")' in html or f'markUndo(`{label}' in html \
            or f'markUndo("{label}' in html, label


def test_기록은_스냅샷이다():
    """참조를 그대로 담으면 나중 변경이 «과거» 까지 바꿔 버린다."""
    html = _script()
    i = html.index("function snapAuto()")
    seg = html[i:i + 400]
    assert "S.autoAlarm.slice()" in seg
    assert "S.zones.map((z) => z.slice())" in seg


def test_도면이_바뀌면_기록을_버린다():
    """남겨 두면 Ctrl+Z 가 앞 도면의 좌표를 이 도면에 씌운다."""
    html = _script()
    i = html.index("S.recon = null; S.suggest = null;")
    assert "S.undo = [];" in html[i:i + 400]
    j = html.index("S.sub = { picks: [null, null], arm: null, summary: null };")
    assert "S.undo = [];" in html[j:j + 400], "슬롯을 바꿔도 기록이 남는다"


def test_되돌릴_것이_없으면_그렇게_말한다():
    html = _script()
    i = html.index("async function undoStep()")
    assert "되돌릴 것이 없습니다" in html[i:i + 1600]


# ═══════════════════════════════════════════ F-8e. 실측·골든
def test_골든_도구가_있다():
    p = os.path.join(_ROOT, "scripts", "_golden_module_f_kfp.py")
    assert os.path.isfile(p), "전체망 .kfp 비트동일을 잴 도구가 없다"


def test_실측_도구가_있다():
    p = os.path.join(_ROOT, "scripts", "_measure_module_f_lanes.py")
    assert os.path.isfile(p)


def test_실측은_사용자_저장본을_안_건드린다():
    """공유 작업폴더에 쓰면 사용자의 B1F 저장본을 덮는다 — 임시 폴더로 돌린다."""
    for name in ("_measure_module_f_lanes.py", "_golden_module_f_kfp.py"):
        src = open(os.path.join(_ROOT, "scripts", name),
                   encoding="utf-8").read()
        assert "TemporaryDirectory" in src, name
        assert "handoff.import_write_root = lambda: work" in src, name


def test_실측이_사람조작과_클릭을_가른다():
    """「전체 반영」은 단추 한 번이지만 화면이 후보마다 클릭을 태운다."""
    src = open(os.path.join(_ROOT, "scripts", "_measure_module_f_lanes.py"),
               encoding="utf-8").read()
    assert '"human"' in src and '"clicks"' in src
    assert src.count('"human"') >= 3, "세 차선 모두에 사람 조작 수가 없다"


def test_물닿음은_marks_에서_읽는다():
    """손질 heads 길이를 세면 물길 판정을 안 탄 board 전체가 된다."""
    src = open(os.path.join(_ROOT, "scripts", "_measure_module_f_lanes.py"),
               encoding="utf-8").read()
    assert "design/preview" in src
    assert 'marks.get("total")' in src


@pytest.mark.skipif(
    not os.path.isfile(os.path.join(_ROOT, "docs", "module_f_lanes.md")),
    reason="F-8e 실측이 아직 안 끝났다 — 리포트가 생기면 이 검사가 살아난다")
def test_차선_리포트가_있다():
    p = os.path.join(_ROOT, "docs", "module_f_lanes.md")
    txt = open(p, encoding="utf-8").read()
    for lane in ("자동", "혼합", "수동"):
        assert lane in txt, f"{lane} 차선이 표에 없다"
    assert "비트 동일" in txt, "골든 결과가 없다"


def test_채택의_유일한_쓰기는_클릭이다_():
    """읽기(by_bundle)는 /pick/auto 가 이미 쓰던 것 — 쓰기는 click 뿐이다."""
    body = _src(adopt)
    assert "ps.click(" in body
    assert body.count("ps.board.") == 1
    assert "ps.board.by_bundle.get(" in body


# ═══════════════════════════════════════════ F-8c. 세 차선
def _script() -> str:
    path = os.path.join(_ROOT, "templates", "module_f.html")
    html = open(path, encoding="utf-8").read()
    return html


def test_방식_카드에_세_차선이_있다():
    html = _script()
    for bid in ("mth-auto", "mth-mixed", "mth-manual"):
        assert f'id="{bid}"' in html, f"#{bid} 가 없다"


def test_카드에_정찰_수치_자리가_있다():
    html = _script()
    assert 'id="mth-recon"' in html
    assert "renderRecon" in html and "loadRecon" in html


def test_채택_기준을_화면에서_고른다():
    """D-F8-4 — 기본 0.9, 화면에서 조절 가능."""
    html = _script()
    assert 'id="mth-conf"' in html
    assert "const CONF_CHOICES" in html
    i = html.index("const CONF_CHOICES")
    seg = html[i:i + 260]
    assert "[0.9," in seg and "[0.75," in seg
    assert 'sel.value = "0.9"' in html, "기본이 0.9 가 아니다"


def test_맞는_후보가_0개면_혼합을_잠근다():
    """눌러도 아무 일이 안 일어나는 단추는 «고장» 으로 읽힌다.

    흔한 일이다 — A 는 알려진 블록 참조만 0.95 를 주므로 헤드를 레이어에 직접
    그린 도면은 높음 띠가 0 이 된다(실측 LH306 0/42 · B1F 72/3,338).
    """
    html = _script()
    i = html.index("function renderConfHint()")
    seg = html[i:i + 600]
    assert '$("mth-mixed").disabled = n === 0;' in seg
    assert "후보가 없습니다" in seg, "왜 잠겼는지 안 말한다"


def test_정찰이_실패하면_혼합만_잠근다():
    """수동·자동 두 길은 인식과 무관하게 열려 있어야 한다."""
    html = _script()
    i = html.index('r.state === "none" || r.state === "error"')
    seg = html[i:i + 600]
    assert '$("mth-mixed").disabled = true;' in seg
    assert "mth-manual" not in seg and "mth-auto" not in seg
    assert "왜 잠겼는지" or "시작할 수 없습니다" in seg


def test_정찰_수치는_띠_칸으로_세운다():
    """긴 한 줄로 늘어놓으면 좁은 옆판에서 글자가 토막나 안 읽힌다."""
    html = _script()
    assert 'class="bands"' in html
    i = html.index("function renderRecon()")
    seg = html[i:i + 1400]
    assert 'cell("hi", "높음 ≥0.9"' in seg
    assert 'cell("lo", "낮음"' in seg


def test_도면_이름은_한_줄로_자른다():
    """길면 카드를 밀어낸다 — 자르고 전체는 툴팁에 둔다."""
    html = _script()
    assert 'class="fname" id="mth-file"' in html
    assert "text-overflow:ellipsis" in html
    assert '$("mth-file").title = nm;' in html


def test_모듈_표는_문장_꼬리에_안_매달린다():
    """한글이 접히고 나면 테두리 칩이 꼬리처럼 남아 줄이 꼬여 보인다."""
    html = _script()
    i = html.index('<div class="lane">')
    seg = html[i:i + 1400]
    # 표가 설명 «앞» 에 오고, 설명은 제 span 안에 갇힌다
    assert '<p><span class="tag">MODULE A</span><span>' in seg
    assert '<p><span class="tag">A + E</span><span>' in seg
    j = html.index("  .lane > p{")
    css = html[j:j + 500]
    assert "display:flex" in css
    assert "flex:0 0 auto" in html[j:j + 700], "표가 같이 접힌다"


def test_한글이_단어_가운데서_안_잘린다():
    """「배관 묶음」이 「배관 묶」/「음」으로 갈리면 읽는 눈이 매번 걸린다."""
    html = _script()
    i = html.index("  .hint{")
    assert "word-break:keep-all" in html[i:i + 300]
    j = html.index("  .kv{")
    assert "word-break:keep-all" in html[j:j + 300]


def test_라벨은_안_쪼개진다():
    """값이 길 때 라벨이 먼저 쪼개지면 「배관 / 묶음」 처럼 토막난다."""
    html = _script()
    i = html.index("  .kv{")
    seg = html[i:i + 700]
    assert "grid-template-columns:minmax(0,auto) minmax(0,1fr)" in seg
    assert "white-space:nowrap" in seg
    # 라벨이 길어지는 날에도 값 칸을 밀어내지 않고 제가 말줄임돼야 한다.
    assert "text-overflow:ellipsis" in seg


def test_혼합은_채택까지만_하고_멈춘다():
    """D-F8-5 — commit 까지 자동으로 가지 않는다. 확정은 사람이 한다."""
    html = _script()
    i = html.index('$("mth-mixed").onclick')
    seg = html[i:i + 1400]
    assert "/api/module-f/pick/adopt" in seg
    assert "pick/commit" not in seg, "사람 확정을 건너뛴다"


def test_혼합은_수동_흐름을_쓴다():
    """혼합은 찍기·손질을 그대로 밟는다 — 새 흐름을 만들지 않는다."""
    html = _script()
    i = html.index('$("mth-mixed").onclick')
    seg = html[i:i + 1400]
    assert 'method: "manual"' in seg
    assert 'S.method = "manual"' in seg


def test_유령은_점선으로_그린다():
    """추정과 실측을 한 선으로 그리지 않는다 — 저장소 규약."""
    html = _script()
    i = html.index("function drawSuggest()")
    seg = html[i:i + 1200]
    assert "ctx.setLineDash(ghost ? [3, 3] : [])" in seg


def test_유령_위의_클릭은_가로채지_않는다():
    """유령은 «아직 안 찍힌 것» — 사람이 직접 찍으려고 누르는 자리다."""
    html = _script()
    i = html.index("후보 클릭 → 반영 제외/복원")
    seg = html[i:i + 1200]
    assert "if (S.ghosts && S.ghosts.has(i)) continue;" in seg
    assert "if (!S.showLow && Number(c.conf) < 0.75) continue;" in seg


def test_낮은_띠는_접어_둔다():
    """3천 점 위에 또 겹치면 아무것도 안 보인다."""
    html = _script()
    assert 'id="pk-show-low"' in html
    i = html.index("function drawSuggest()")
    assert "!S.showLow" in html[i:i + 1200]


def test_후보_좌표는_찍기로_갈_때만_받는다():
    """카드는 수치만 받는다 — 3천 점을 카드 그릴 때마다 내려보내지 않는다."""
    html = _script()
    i = html.index("async function loadRecon()")
    assert "heads=1" not in html[i:i + 400], "카드가 좌표까지 받는다"
    j = html.index("async function applyAdopt(")
    assert "heads=1" in html[j:j + 500], "채택 뒤에도 좌표를 안 받는다"


def test_후보_지우기가_유령도_같이_걷는다():
    html = _script()
    i = html.index('$("pk-suggest-clear").onclick')
    seg = html[i:i + 600]
    assert "S.ghosts = null;" in seg and "S.adopted = null;" in seg


# ═══════════════════════════════════════════ F-8d. 탈출로
def _auto_src() -> str:
    from routes.module_f import api_auto
    return _src(api_auto.register)


def test_이어받기_라우트가_있다():
    src = _auto_src()
    assert '@app.post("/api/module-f/auto/handoff")' in src
    assert '@app.get("/api/module-f/auto/handoff-hints")' in src


def test_이어받기는_잡_하나로_끝낸다():
    """채택 → 스펙 저장 → 손질 진입. 사람이 세 번 기다리게 하지 않는다."""
    src = _auto_src()
    i = src.index('@app.post("/api/module-f/auto/handoff")')
    seg = src[i:i + 4600]
    assert "adopt_bundles(ps, world" in seg
    assert "adopt_heads(ps, select_heads(cands)" in seg
    assert "ps.commit()" in seg
    assert "EditSession.open(" in seg
    assert src.count('_run_job(sess, "손질로 이어받기", job)') == 1


def test_이어받기는_후보_전체를_준다():
    """자동이 영역으로 좁혔더라도 넓게 준다 — 좁히기는 손질에서 한다."""
    src = _auto_src()
    i = src.index('@app.post("/api/module-f/auto/handoff")')
    seg = src[i:i + 4600]
    assert "select_heads(cands)" in seg
    assert "conf_min" not in seg, "이어받기가 후보를 미리 잘라낸다"


def test_이어받기는_자동_결과를_지우지_않는다():
    """손질 뒤 사람이 다시 최불리를 고르면 그때 덮인다(기존 규약)."""
    src = _auto_src()
    i = src.index('@app.post("/api/module-f/auto/handoff")')
    seg = src[i:i + 4600]
    assert 'sess.pop("auto"' not in seg and 'sess["auto"] = None' not in seg
    assert 'sess.pop("design"' not in seg
    assert 'sess["method"] = "manual"' in seg, "자동 흐름을 안 떠난다"


def test_이어받기는_자동을_돌린_뒤에만():
    src = _auto_src()
    i = src.index('@app.post("/api/module-f/auto/handoff")')
    seg = src[i:i + 4600]
    assert 'if not sess.get("auto"):' in seg
    assert 'ps = sess.get("pick")' in seg
    assert "_job_running(sess)" in seg


def test_알람밸브와_급수시작은_제안_둘로_나뉜다():
    """합칠지는 미결(BLOCKED §5) — 지금은 사람이 각각 반영한다."""
    src = _auto_src()
    i = src.index('sess["handoff"] = {')
    seg = src[i:i + 400]
    assert '"alarm":' in seg and '"source":' in seg


def test_이어받기_제안_반영은_기존_클릭_경로다():
    """D-F8-3 — 여기서도 주입은 없다."""
    html = _script()
    i = html.index("async function applyHint(")
    seg = html[i:i + 800]
    assert "/api/module-f/edit/mode" in seg
    assert "/api/module-f/edit/click" in seg


def test_이어받기_단추가_자동_화면에_있다():
    html = _script()
    assert 'id="au-handoff"' in html
    i = html.index('$("au-to-design").disabled = !S.autoDone;')
    assert '$("au-handoff").disabled = !S.autoDone;' in html[i:i + 300]


def test_제안은_점선으로_그린다():
    """확정된 것(실선)과 한눈에 갈려야 한다."""
    html = _script()
    i = html.index("function drawHandoffHints()")
    seg = html[i:i + 700]
    assert "ctx.setLineDash([4, 4])" in seg


def test_손질_화면에_제안_반영_단추가_있다():
    html = _script()
    for bid in ("ed-handoff-box", "ed-hint-alarm", "ed-hint-source"):
        assert f'id="{bid}"' in html, f"#{bid} 가 없다"


def test_새_도면을_올리면_이어받기_표시가_사라진다():
    html = _script()
    i = html.index("S.zones = []; S.autoAlarm = null;")
    seg = html[i:i + 500]
    assert "S.handoff = null;" in seg and "S.recon = null;" in seg


def test_채택의_유일한_쓰기는_클릭이다():
    """읽기(by_bundle)는 /pick/auto 가 이미 쓰던 것 — 쓰기는 click 뿐이다."""
    body = _src(adopt)
    writes = {"ps.click(", "ps.select_pipe(", "ps.complete_pipe(",
              "ps.set_slot("}
    assert "ps.click(" in body
    # board 를 통해 무언가를 «호출» 하는 곳은 by_bundle 조회뿐
    assert body.count("ps.board.") == 1
    assert "ps.board.by_bundle.get(" in body
    assert writes  # 계약 문서화용

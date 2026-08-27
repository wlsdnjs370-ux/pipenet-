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

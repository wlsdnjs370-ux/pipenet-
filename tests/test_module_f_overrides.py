# -*- coding: utf-8 -*-
"""[F-11] 완결성 — 지배 띠 채택 · 직접 입력 · 수정 생존.

배포 일반화의 정의(지시서 §0.1)는 인식률이 아니라 **완주 가능성**이다:

    정직성  은닉 오류 0 — 틀리거나 불확실한 지점이 전부 화면에 보인다
    완결성  막다른 길 0 — 표시된 모든 한계에 프로그램 안에서 고칠 길이 있다
    수렴성  같은 수정을 두 번 시키지 않는다

이 파일은 그 셋을 지키는 시험이다.
"""
from __future__ import annotations

import os

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _screen() -> str:
    """화면 소스 한 덩이 — 마크업 + CSS + JS."""
    parts = [open(os.path.join(_ROOT, "templates", "module_f.html"),
                  encoding="utf-8").read()]
    for rel in (("static", "module_f.css"), ("static", "module_f.js")):
        p = os.path.join(_ROOT, *rel)
        if os.path.isfile(p):
            parts.append(open(p, encoding="utf-8").read())
    return "\n".join(parts)


# ═══════════════════════════════════════════ 수렴성 — 두 번 시키지 않는다
def test_재채택이_앞의_채택을_안_지운다():
    """[§0.1 수렴성] 같은 수정을 두 번 시키지 않는다.

    ★F-11a 가 드러낸 결함이다. E 의 찍기는 «문양 서명» 토글이라 이미 찍힌
      서명 위의 클릭이 «취소» 가 된다. `adopt_heads` 는 그 응답을 보고 곧바로
      되클릭해 복원했는데(파일 머리말의 규약), `adopt_bundles` 는 안 그래서
      «다시 채택» 이 앞서 찍은 재료를 통째로 꺼 버렸다.

      실측(LH306): 1차 재료 6묶음·헤드 3 → 2차 뒤 재료 0묶음·헤드 0, 그리고
      「재료를 하나도 못 찍었습니다」로 조립까지 막혔다.

    두 함수가 같은 규약을 지키는지 소스로 못 박는다 — 한쪽만 고쳐지면 이
    결함이 그대로 돌아온다.
    """
    import inspect

    from routes.module_f import adopt

    for fn in (adopt.adopt_bundles, adopt.adopt_heads):
        src = inspect.getsource(fn)
        i = src.index('"취소"')
        # 취소를 만나면 «되클릭» 이 바로 뒤따라야 한다.
        assert "click(" in src[i:i + 400], f"{fn.__name__} 가 복원하지 않는다"
    # 재료 쪽도 «이미 반영» 을 따로 센다 — 건너뜀과 섞으면 사실이 흐려진다.
    assert "already" in inspect.getsource(adopt.adopt_bundles)


# ═══════════════════════════════════════════ F-11a. 지배 띠 채택
def _bands(hi, mid, low):
    from routes.module_f.recon import BAND_HIGH, BAND_LOW, BAND_MID
    return {BAND_HIGH: hi, BAND_MID: mid, BAND_LOW: low}


def test_지배_띠가_세_갈래로_갈린다():
    """[D-F11-2] 절대 임계 0.9 는 A 의 신뢰도가 사실상 이진값이라 퇴화한다.

    블록 참조만 0.95 를 받고 나머지 HEAD 도형은 0.70~0.89 에 몰리므로,
    0.9 는 «대부분» 과 «거의 없음» 사이에서 도면마다 뒤집힌다. 그래서 임계를
    도면의 분포가 정하게 한다.
    """
    from routes.module_f.recon import CONF_HIGH, CONF_MID, dominant_band

    # 실측 세 도면 — 주석의 수치를 그대로 시험으로 못 박는다.
    hi = dominant_band(_bands(87, 14, 18))          # 대명동 73%
    assert hi["rule"] == "high" and hi["conf_min"] == CONF_HIGH
    assert hi["n"] == 87

    b1f = dominant_band(_bands(72, 3163, 103))      # B1F 2%
    assert b1f["rule"] == "high_mid" and b1f["conf_min"] == CONF_MID
    assert b1f["n"] == 72 + 3163

    lh = dominant_band(_bands(0, 40, 2))            # LH306 0%
    assert lh["rule"] == "high_mid" and lh["conf_min"] == CONF_MID
    assert lh["n"] == 40

    none = dominant_band(_bands(0, 0, 5))
    assert none["rule"] == "none" and none["conf_min"] is None
    assert none["n"] == 0


def test_지배_띠_경계는_50퍼센트_이상이다():
    """문지방은 상수다 — 경계에서 어느 쪽으로 가는지 못 박아 둔다."""
    from routes.module_f.recon import (CONF_HIGH, CONF_MID,
                                       DOMINANT_BAND_RATIO, dominant_band)

    assert DOMINANT_BAND_RATIO == 0.5
    # 정확히 50% → 높음만(«이상» 이다)
    assert dominant_band(_bands(50, 30, 20))["conf_min"] == CONF_HIGH
    # 한 개 모자라면 중간까지
    assert dominant_band(_bands(49, 31, 20))["conf_min"] == CONF_MID


def test_발동한_규칙을_반드시_말한다():
    """[F-11a] 조용한 규칙 전환은 새 은닉 오류다 — 카드와 배너 둘 다에 적는다."""
    from routes.module_f.recon import dominant_band

    r = dominant_band(_bands(0, 40, 2))
    assert "중간까지 채택" in r["why"] and "40" in r["why"]
    html = _screen()
    # 카드
    assert "r.adopt.why" in html
    # 손질 진입 배너
    assert "a.why" in html


def test_수동_임계가_규칙을_이긴다():
    """규칙은 기본값이지 잠금이 아니다 — 고급에서 고르면 그것이 이긴다."""
    html = _screen()
    i = html.index("function confMin()")
    seg = html[i:i + 700]
    assert "S.confManual" in seg, "수동 지정을 안 본다"
    # 수동 표시는 사람이 «고르는 순간» 서야 한다.
    assert 'S.confManual = true' in html


def test_게이트가_규칙이_고른_임계로_판단한다():
    """[§16 승계] 게이트는 최후 방어로 남되, 판단 기준이 절대 0.9 가 아니다.

    예전에는 LH306(높음 0/42)이 여기서 막혔다. 이제 규칙이 중간까지 채택하므로
    살아서 지나가고, 규칙조차 0 을 내는 도면에서만 문이 닫힌다.
    """
    html = _screen()
    i = html.index("function reconReady()")
    seg = html[i:i + 1400]
    assert "reconPick(confMin())" in seg, "아직 절대 임계로 막는다"
    assert "reconPick(0.9)" not in seg
    # 막을 때도 «왜» 를 규칙의 문장으로 말한다.
    assert "a.why" in seg

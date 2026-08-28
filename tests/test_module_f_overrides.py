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


# ═══════════════════════════════════════════ F-11b. 직접 입력을 «지우는» 길
def test_채운_값에_갇히지_않는다():
    """[§0.1 완결성] 막다른 길 0 — 잘못 채운 값을 지울 길이 화면에 있어야 한다.

    서버는 이미 「빈 배열을 보내면 지운다」를 규약으로 갖고 있었지만, 그 규약을
    부를 단추가 화면에 없으면 사람에게는 없는 기능이다.
    """
    html = _screen()
    assert "직접 입력 지우기" in html, "지우는 단추가 없다"
    assert "async function dropOverride(d)" in html
    i = html.index("async function dropOverride(d)")
    seg = html[i:i + 1200]
    # 지운 뒤 값이 실제로 바뀌므로 재확정까지 이어 준다 — 저장 경로와 같은 규약.
    assert '$("dg-build").click()' in seg, "지우고 재확정으로 안 이어진다"


def test_지울_때_다른_갈래는_안_건드린다():
    """서버 규약: «칸을 안 보내면 그 갈래는 그대로, 빈 배열이면 지운다».

    한 갈래를 지우면서 다른 갈래를 «안 보내는» 것이 그래서 중요하다. 둘 다
    보내면서 한쪽을 빈 배열로 두면, 부속을 지울 때 등가길이까지 날아간다.
    """
    html = _screen()
    i = html.index("async function dropOverride(d)")
    seg = html[i:i + 1200]
    # 몸통을 먼저 만들고 «해당 갈래만» 채운다.
    assert "const body = { sid: S.sid };" in seg
    assert 'if (d.type === "kind") {' in seg
    # 두 갈래를 한꺼번에 싣지 않는다.
    assert "body.kind" in seg and "body.eq_len" in seg
    assert seg.index("body.kind") < seg.index("body.eq_len")
    assert "else" in seg[seg.index("body.kind"):seg.index("body.eq_len")]


def test_표_확정_필요_배지는_한_곳에서만_써진다():
    """[F-11b-2] 같은 칸을 두 곳에서 쓰면 나중 것이 앞 것을 덮는다.

    ★실제로 그랬다. `renderIssues` 가 「표 확정 필요」를 세워 놓고, 바로 뒤의
      `countFilled()` 가 「채운 칸 0」으로 지워 버려 배지가 뜰 수 없었다.
      그래서 `dg-ov-n` 에 쓰는 곳을 한 곳으로 못 박는다.
    """
    html = _screen()
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    writes = js.count('$("dg-ov-n").textContent')
    assert writes == 0, "배지를 직접 쓰는 곳이 남아 있다"
    i = js.index("function countFilled()")
    seg = js[i:i + 900]
    assert 'const el = $("dg-ov-n");' in seg
    assert "S.ovDirty" in seg, "배지가 «아직 안 들어갔다» 를 안 본다"
    assert "표 확정 필요" in seg


def test_배지를_세우는_곳과_내리는_곳이_다_있다():
    """[정직한 진행 표시] 세울 곳만 있고 내릴 곳이 없으면 배지가 영영 남는다.

    반대로 «세울 곳» 이 없으면 배지는 영영 안 뜬다 — 화면에 코드만 있고
    기능은 없는 상태가 된다. 양쪽을 다 못 박는다.
    """
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    # 세운다 — 저장·지움 두 경로 모두 서버의 `needs_rebuild` 를 그대로 쓴다.
    assert js.count("S.ovDirty = !!") == 2, "저장·지움 둘 다 배지를 세워야 한다"
    # 내린다 — «재확정이 성공한» 자리에서만.
    i = js.index('$("dg-build").onclick')
    seg = js[i:i + 1400]
    assert "S.ovDirty = false;" in seg
    assert seg.index("확정 실패") < seg.index("S.ovDirty = false;"), \
        "실패로 빠지는 throw 앞에서 배지를 내리면 거짓말이 된다"


def test_표에서도_직접_입력이_다른_얼굴이다():
    """[F-11b-3] 목록만이 아니라 «표» 에서도 구별돼야 한다.

    엔진의 부속표에는 그런 칸이 없고 이 항목에서 서버는 불변이다. 그래서
    화면이 이미 받아 둔 `unresolved.applied` 를 표에 겹쳐 놓는다 — 새 판정이
    아니라 표시다.
    """
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    assert "function overrideNoteOf(row, which)" in js
    i = js.index("function overrideNoteOf(row, which)")
    seg = js[i:i + 1400]
    assert 'if (which !== "fittings") return null;' in seg, "부속표에만 쓴다"
    assert "직접 입력 — 부속" in seg and "직접 입력 — 등가길이" in seg
    assert "a.note" in seg, "사유가 표에 안 실린다"
    # 채운 자리가 없으면 표를 안 건드린다 — 없는 칸을 만들지 않는다.
    j = js.index("function renderDesignTable()")
    tab = js[j:j + 1400]
    assert "const hasOv = notes.some(Boolean);" in tab
    assert 'hasOv ? "<th>근거</th>" : ""' in tab
    css = open(os.path.join(_ROOT, "static", "module_f.css"),
               encoding="utf-8").read()
    assert ".ovcell" in css and ".ovdel" in css

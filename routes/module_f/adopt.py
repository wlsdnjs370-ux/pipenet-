# -*- coding: utf-8 -*-
"""[F-8b] 채택 — A 의 «발견» 을 E 의 «찍기 스펙» 으로 번역한다.

넘기는 것은 그래프가 아니라 입력이다(지시서 §0.3). A 가 이었다고 보는 망을
E 의 물길 판정에 그대로 넣으면 헤드 물닿음 0 · 노드 2 만 남는다(G BLOCKED B4)
— 두 파이프라인은 「이어져 있다」의 정의가 다르다. 그래서 번역은 두 줄뿐이다:

    어떤 레이어×색이 배관인가   →  재료 묶음 클릭
    헤드가 어디 있나(신뢰도)    →  헤드 클릭

**쓰기 경로는 `PickSession.click` 하나뿐이다**(D-F8-3). board 에 픽을 넣는
코드는 여기 없다. 그래서 클릭 기록·되돌리기·스펙 저장이 사람이 찍은 것과
완전히 같은 상태가 된다 — 채택한 것을 undo 로 한 번에 되돌릴 수 있는 것도
그 덕이다(D-F8-5).

■ 토글을 다루는 법

E 의 찍기는 «문양 서명» 단위 토글이다 — 한 클릭이 같은 문양의 원 전부를
대표하고, **이미 찍힌 서명 위의 클릭은 «취소»** 가 된다. A 의 후보 3,338개를
그냥 순서대로 클릭하면 같은 서명끼리 서로를 꺼서 결과가 0 에 수렴한다.

그렇다고 board 를 들여다보고 «이미 있는 서명» 을 건너뛰면, 그것은 E 의 판정을
바깥에서 흉내 내는 것이라 언젠가 어긋난다. 화면이 하는 그대로 한다 —
**취소로 응답하면 곧바로 되클릭해 복원하고 «이미 반영» 으로 센다.** 두 클릭이
서로를 지우므로 board 상태는 사람이 그 자리를 두 번 누른 것과 같다.
"""
from __future__ import annotations

# 헤드 후보 좌표에서 E 의 도형까지 허용하는 거리.
#
# ★상한이 없으면 안 된다. `Board._click_head` 는 `max_d=None` 이면 «가장 가까운
#   원» 을 거리와 무관하게 받아들이므로, 후보 자리에 헤드가 없을 때 수 미터
#   떨어진 남의 헤드를 찍는다. 300mm 는 헤드 기호 하나가 들어가는 틈
#   (`HEAD_GAP_JOIN_MAX_MM` 계열)과 같은 자다 — 원 거리는 중심이 아니라
#   테두리까지라, 반지름이 이만한 기호까지 닿는다.
ADOPT_MAX_D_MM = 300.0

WHY_NO_SHAPE = "후보 자리에 찍을 도형이 없음"


def adopt_bundles(ps, world, cat: str = "PIPE") -> dict:
    """레이어 사전이 추천한 묶음을 클릭으로 찍는다 — `/pick/auto` 의 그 경로.

    `board.mat` 에 밀어넣지 않고 **그 묶음의 실제 선분 중점** 으로 정상 클릭을
    태운다. 이 함수를 `/pick/auto` 와 나눠 쓰므로 두 길이 갈릴 수 없다.
    """
    want = str(cat or "PIPE").upper()
    targets = [b for b in ((world or {}).get("bundles") or [])
               if b.get("cat") == want]
    applied, skipped = [], []
    for b in targets:
        segs = ps.board.by_bundle.get((b["layer"], b["color"])) or []
        if not segs:
            skipped.append(b["layer"])
            continue
        a, c = segs[0]
        rep = ps.click((a[0] + c[0]) / 2.0, (a[1] + c[1]) / 2.0)
        if rep is None or rep.get("동작") != "추가":
            skipped.append(b["layer"])
        else:
            applied.append(b["layer"])
    return {"applied": applied, "skipped": skipped, "targets": len(targets)}


def select_heads(cands, *, conf_min=None, indices=None) -> list:
    """채택할 후보를 고른다 — 신뢰도 문턱 또는 사람이 고른 번호.

    반환은 `(원래 번호, 후보)` 짝이다. 번호를 들고 다녀야 화면이 «몇 번이
    유령으로 남았나» 를 후보 목록과 맞출 수 있다.
    """
    items = list(enumerate(cands or ()))
    if indices is not None:
        want = {int(i) for i in indices}
        return [(i, c) for i, c in items if i in want]
    if conf_min is None:
        return items
    lo = float(conf_min)
    return [(i, c) for i, c in items if float(c.get("conf") or 0.0) >= lo]


def adopt_heads(ps, picks, *, max_d: float = ADOPT_MAX_D_MM,
                progress=None) -> dict:
    """고른 후보를 헤드 칸에 찍는다. 실패는 «유령» 으로 남긴다.

    A 의 후보 좌표에 E 가 찍을 도형(원·삼각 해치)이 없으면 클릭이 실패한다 —
    그것이 옳다. E 의 「표시가 없으면 추측하지 않는다」 확정 게이트를 우회하는
    별도 주입 경로를 만들지 않는 것이 이 작업의 요점이기 때문이다. 실패한
    후보는 세어 돌려주고, 화면이 점선으로 남겨 사람이 직접 찍거나 무시한다.
    """
    applied = already = 0
    skipped, clicked = [], []
    total = len(picks)
    for n, (i, c) in enumerate(picks, start=1):
        x, y = float(c["x"]), float(c["y"])
        rep = ps.click(x, y, max_d=max_d)
        act = (rep or {}).get("동작")
        if act == "추가":
            applied += 1
            clicked.append([x, y])
        elif act == "취소":
            # 이미 찍힌 서명을 껐다 — 곧바로 되켠다(화면이 하는 그대로).
            ps.click(x, y, max_d=max_d)
            already += 1
            clicked += [[x, y], [x, y]]
        else:
            skipped.append({"i": i, "x": x, "y": y,
                            "conf": c.get("conf"), "why": WHY_NO_SHAPE})
        if progress is not None and (n % 500 == 0 or n == total):
            progress(n, total, applied, already, len(skipped))
    return {"applied": applied, "already": already, "skipped": skipped,
            "clicked": clicked}

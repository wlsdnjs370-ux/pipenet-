# -*- coding: utf-8 -*-
"""[H-0] 도면 슬롯 — 특허 S650(추가도면 회귀)의 상태.

특허는 평면도 한 장으로 끝나지 않는다. S650 에서 처리할 도면이 남으면 S100 으로
회귀하여 **평면도 · 계통도 · 기계실** 에 같은 절차(제1~4국면)를 반복 적용하고,
S700 이 그 셋을 하나의 배관망으로 결합한다. 그런데 F 의 세션은 도면 한 장짜리
평면 dict 였다 — 도면 종류라는 개념 자체가 없었다.

여기서 하는 일은 그 dict 를 **슬롯 세 칸으로 나누는 것 하나**다. 기존 33개
엔드포인트는 한 줄도 고치지 않는다: 세션 dict 는 여전히 평면이고, 그 내용이
«지금 활성인 슬롯» 의 도면 상태일 뿐이다. 슬롯을 바꿀 때 현재 내용을 슬롯
저장소로 걷어내고, 대상 슬롯의 내용을 그 자리에 편다(`_slot_switch`).

그래서 **도면별 키를 열거하지 않는다.** 세션 전역으로 남을 것(`SESSION_KEYS`)만
적고 나머지 전부를 슬롯 상태로 본다. 실측으로 지금 세션 키는 35개이고 그중
도면별이 30개다 — 열거하면 늘어날 때마다 빠뜨린다. 나중에 누가 도면별 키를
하나 더 늘렸을 때, 그것이 슬롯을 넘나들며 새는 쪽보다 슬롯에 갇히는 쪽이
언제나 안전하다.

활성 슬롯의 상태는 저장소에 **없다** — 평면 dict 그 자체가 활성 슬롯이다.
저장소에는 쉬고 있는 슬롯만 들어 있다. 두 곳에 같은 것을 두면 어느 쪽이
최신인지를 매번 판단해야 하고, 그 판단이 틀리는 날 사용자는 손질한 것이
사라진 화면을 본다.
"""
from __future__ import annotations

SLOT_KINDS = ("plan", "system", "machineroom")
SLOT_LABELS = {
    "plan": "평면도",
    "system": "계통도",
    "machineroom": "기계실",
}

# 슬롯이 바뀌어도 그 자리에 남는 것 — 세션 정체와 «한 번에 하나» 인 잡.
# 잡이 세션 전역인 것은 의도다: _HEAVY_LOCK 이 프로세스 하나짜리라
# 슬롯마다 잡을 두어도 어차피 동시에 돌지 못한다.
SESSION_KEYS = frozenset({"id", "created", "touched", "job", "log",
                          "slots", "active"})


def _check_slot_kind(kind) -> str:
    k = str(kind or "").strip()
    if k not in SLOT_KINDS:
        raise ValueError(
            f"그런 도면 종류가 없습니다: {kind!r} "
            f"(쓸 수 있는 것: {', '.join(SLOT_KINDS)})")
    return k


def _slot_blank() -> dict:
    """빈 도면 상태 — `_new_session` 이 만들던 기본값 그대로."""
    return {
        "dxf": None, "key": None, "pick": None, "edit": None,
        "world": None, "kfp": None, "kfp_path": None,
        "water_path": None, "worst": None,
        "sdf_path": None, "slf_path": None,
    }


def _slot_init(sess: dict, active: str = "plan") -> None:
    """세션에 슬롯 저장소를 붙인다. 활성 슬롯은 저장소에 넣지 않는다."""
    kind = _check_slot_kind(active)
    sess["active"] = kind
    sess["slots"] = {k: _slot_blank() for k in SLOT_KINDS if k != kind}


def _slot_active(sess: dict) -> str:
    """활성 슬롯. 슬롯을 모르는 옛 세션도 평면도로 보고 넘어간다."""
    kind = sess.get("active")
    return kind if kind in SLOT_KINDS else "plan"


def _slot_capture(sess: dict) -> dict:
    """지금 평면 dict 에 펼쳐져 있는 도면 상태를 걷어낸다."""
    return {k: v for k, v in sess.items() if k not in SESSION_KEYS}


def _slot_restore(sess: dict, state: dict) -> None:
    """평면 dict 의 도면 상태를 통째로 갈아끼운다.

    지우고 넣는다 — update 만 하면 **이전 슬롯의 키가 남는다.** 계통도에만
    있던 `design_sdf_path` 가 평면도로 따라가면 남의 산출물을 제 것으로
    보고하게 된다.
    """
    for k in [k for k in sess if k not in SESSION_KEYS]:
        del sess[k]
    sess.update(state)


def _slot_switch(sess: dict, kind) -> str:
    """활성 슬롯을 바꾼다 — 현재 것을 저장소로, 대상 것을 평면으로."""
    target = _check_slot_kind(kind)
    active = _slot_active(sess)
    if target == active:
        return active
    store = sess.setdefault("slots", {})
    store[active] = _slot_capture(sess)
    _slot_restore(sess, store.pop(target, None) or _slot_blank())
    sess["active"] = target
    return target


def _slot_progress(state: dict) -> dict:
    """슬롯 하나의 진행 — S650 이 «이 도면은 끝났나» 를 묻는 자리.

    단계 판정은 `/api/module-f/job` 의 stage 규약과 같은 것을 쓴다. 두 곳이
    다르게 답하면 화면이 슬롯 탭과 단계 표시에서 서로 다른 말을 하게 된다.
    """
    if state.get("edit") is not None:
        stage = "edit"
    elif state.get("pick") is not None:
        stage = "pick"
    else:
        stage = ""
    return {
        "opened": bool(state.get("key") or state.get("dxf")),
        "stage": stage,
        "key": state.get("key"),
        # 제4국면(S500)까지 갔는가 — 병합(S700)이 쓸 수 있는 슬롯인지의 근거.
        "designed": bool(state.get("design_sdf_path")),
    }


def _slot_state(sess: dict) -> dict:
    """세 슬롯의 진행 한 장. 활성은 평면 dict 에서, 나머지는 저장소에서 읽는다."""
    active = _slot_active(sess)
    store = sess.get("slots") or {}
    live = _slot_capture(sess)
    items = []
    for kind in SLOT_KINDS:
        state = live if kind == active else (store.get(kind) or _slot_blank())
        items.append({
            "kind": kind,
            "label": SLOT_LABELS[kind],
            "active": kind == active,
            **_slot_progress(state),
        })
    return {"active": active, "slots": items}

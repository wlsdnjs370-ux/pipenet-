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

import re
import threading

# 슬롯 전환은 평면 dict 를 «순회하며 지우고 다시 채운다». waitress 는 멀티스레드라
# 같은 sid 로 전환 요청이 겹치면(더블클릭·재시도) 순회 중 변경으로 죽거나 —
# 더 나쁘게 — 두 슬롯의 상태가 반쯤 섞인 채 남는다. 전환만 직렬화한다.
# 세션 dict 안에 Lock 을 넣으면 안 된다: SESSION_KEYS 밖이라 _slot_capture 가
# 슬롯 상태로 쓸어가 «슬롯마다 다른 락» 이 된다.
_SWITCH_LOCK = threading.Lock()

SLOT_KINDS = ("plan", "system", "machineroom")
SLOT_LABELS = {
    "plan": "평면도",
    "system": "계통도",
    "machineroom": "기계실",
}

# [오너 2026-09-22 · 그림 42~44] 계통도 여러 장 — «＋ 계통도 추가».
#   연결 축은 평면도 — 계통도 1 — 계통도 2 — … — 계통도 n — 기계실 이다.
#   «system» 이 계통도 1 이고(종전 그대로), 더한 칸은 `system2` · `system3` … 이다.
#   더한 칸의 목록은 세션 전역(`system_extra`)에 둔다 — 도면별 상태가 아니다.
#   칸 이름(계통도 1 · 2 …)은 목록 안의 **자리**로 매긴다(지운 칸 뒤는 당겨진다).
SYSTEM_MAX = 10
_SYSTEM_EXTRA = re.compile(r"system([2-9]|[1-9][0-9])")


def is_system_kind(kind) -> bool:
    """계통도 칸인가 — 계통도 1(`system`) 또는 더한 칸(`system2` …)."""
    k = str(kind or "")
    return k == "system" or bool(_SYSTEM_EXTRA.fullmatch(k))


def slot_role(kind) -> str:
    """칸이 하는 일 — 더한 계통도 칸도 계통도(`system`)와 같은 일을 한다."""
    return "system" if is_system_kind(kind) else str(kind or "")


def system_kinds(sess: dict) -> list:
    """계통도 칸들 — 평면도 쪽(계통도 1)부터 기계실 쪽으로."""
    return ["system", *[k for k in (sess.get("system_extra") or ())
                        if _SYSTEM_EXTRA.fullmatch(str(k))]]


def slot_kinds(sess: dict) -> list:
    """이 세션의 칸 전부 — 평면도 · 계통도 1 … n · 기계실."""
    return ["plan", *system_kinds(sess), "machineroom"]


def slot_label(sess: dict, kind) -> str:
    """칸 이름. 계통도가 한 장뿐이면 종전 그대로 «계통도» 다."""
    if not is_system_kind(kind):
        return SLOT_LABELS[kind]
    kinds = system_kinds(sess)
    if len(kinds) == 1:
        return SLOT_LABELS["system"]
    return f"계통도 {kinds.index(kind) + 1}"

# 슬롯이 바뀌어도 그 자리에 남는 것 — 세션 정체와 «한 번에 하나» 인 잡.
# 잡이 세션 전역인 것은 의도다: _HEAVY_LOCK 이 프로세스 하나짜리라
# 슬롯마다 잡을 두어도 어차피 동시에 돌지 못한다.
#
# ★제5국면(S700 통합)의 상태도 여기 있어야 한다. 통합은 «세 슬롯의 결과를
#   모으는» 자리라 어느 한 도면의 것이 아니다(이 파일 머리말도, api_merge 의
#   `_materials` 도 그렇게 읽는다 — 세 슬롯을 훑어 재료를 모은다).
#   그런데 이 목록에 없으면 `_slot_capture` 가 떠나는 슬롯으로 걷어가고
#   `_slot_restore` 가 지운다. 실측: 급수방식을 고르고 슬롯을 한 번 바꾸면
#   급수방식·낙차·펌프 제원·결합 결과가 **통째로 사라졌다**. 계통도를 고치러
#   다녀오는 것만으로 앞서 정한 것이 없어지는 셈이다.
#
#   ★«도면별 키를 열거하지 않는다» 는 이 파일의 원칙은 그대로다. 여기 적는
#     것은 도면별 키가 아니라 **도면에 딸리지 않는다고 이미 정해 둔** 것들
#     뿐이다 — 늘어나면 그때마다 여기 적어야 한다는 뜻이고, 그것이 옳다.
_MERGE_KEYS = frozenset({
    "merge_editor", "editor_merge_iso", "editor_merge_z", "editor_last_scope",
    "supply_mode",      # S710 급수방식 — 세 도면 공통의 결정
    "source_drop_m",    # S730 수원 낙차
    "pump_spec",        # S710 펌프 제원(펌프 가압에서만)
    "merged",           # S740 결합 결과
    "merge_summary",    # 그 요약
    "merge_files",      # S750 산출 파일 목록
    "sizing_report", "sizing_files",  # optional integrated hydraulic proposals
})
SESSION_KEYS = frozenset({"id", "created", "touched", "job", "log", "_state_revision",
                          "slots", "active",
                          # [오너 2026-09-22] 더한 계통도 칸 목록 — 칸의 짜임이지 도면이 아니다.
                          "system_extra"}) | _MERGE_KEYS


def _check_slot_kind(kind, sess: dict | None = None) -> str:
    """칸 이름 검사. 더한 계통도 칸은 그 칸을 가진 세션(`sess`)에서만 통한다."""
    k = str(kind or "").strip()
    known = slot_kinds(sess) if sess is not None else list(SLOT_KINDS)
    if k not in known:
        raise ValueError(
            f"그런 도면 종류가 없습니다: {kind!r} "
            f"(쓸 수 있는 것: {', '.join(known)})")
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
    return kind if kind in slot_kinds(sess) else "plan"


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
    """활성 슬롯을 바꾼다 — 현재 것을 저장소로, 대상 것을 평면으로.

    전환은 직렬화한다(모듈 머리말의 `_SWITCH_LOCK`). 잡 실행 중 409 는 워커와의
    경쟁만 막는다 — 같은 sid 의 전환 요청 «둘» 이 겹치는 것은 여기서 막는다.
    """
    target = _check_slot_kind(kind, sess)
    with _SWITCH_LOCK:
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
    # ★자동 경로도 찍기판(PickSession)으로 도면을 연다 — 그것이 가장 싸다.
    #   그래서 `pick` 이 있다는 것만으로 「찍기 중」 이라고 하면 안 된다.
    #   자동은 제 단계 이름을 갖는다.
    if str(state.get("method") or "") == "auto":
        stage = "auto" if state.get("auto") else "auto-ready"
    elif state.get("edit") is not None:
        stage = "edit"
    elif state.get("pick") is not None:
        stage = "pick"
    else:
        stage = ""
    return {
        "opened": bool(state.get("key") or state.get("dxf")),
        "stage": stage,
        "method": state.get("method"),
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
    systems = system_kinds(sess)
    for kind in slot_kinds(sess):
        state = live if kind == active else (store.get(kind) or _slot_blank())
        items.append({
            "kind": kind,
            "label": slot_label(sess, kind),
            "active": kind == active,
            # [오너 2026-09-22] 더한 계통도 칸 — 화면이 이름·× 단추를 이것으로 그린다.
            "role": slot_role(kind),
            "removable": kind in systems[1:],
            **_slot_progress(state),
        })
    return {"active": active, "slots": items,
            "can_add_system": len(systems) < SYSTEM_MAX}


def _slot_add_system(sess: dict) -> str:
    """[오너 2026-09-22] «＋ 계통도 추가» — 기계실 쪽 끝에 빈 계통도 칸 하나.

    칸 이름표(id)는 비어 있는 가장 작은 번호다. 새 칸은 빈 도면 상태로
    저장소에 들어간다 — 활성으로 바꾸는 것은 `_slot_switch` 의 일이다.
    """
    with _SWITCH_LOCK:
        extra = list(sess.get("system_extra") or ())
        if len(extra) + 1 >= SYSTEM_MAX:
            raise ValueError(f"계통도는 {SYSTEM_MAX}장까지 올릴 수 있습니다.")
        used = set(extra)
        n = 2
        while f"system{n}" in used:
            n += 1
        kind = f"system{n}"
        sess["system_extra"] = extra + [kind]
        sess.setdefault("slots", {})[kind] = _slot_blank()
    return kind


def _slot_remove_system(sess: dict, kind) -> str:
    """[오너 2026-09-22] 더한 계통도 칸을 지운다(× 단추) — 계통도 1 은 못 지운다.

    지우는 칸이 활성이면 바로 앞(평면도 쪽) 계통도 칸으로 먼저 옮긴다.
    그 칸에 올린 도면·뽑은 경로도 함께 사라진다. 돌려주는 값은 활성 칸이다.
    """
    k = str(kind or "").strip()
    kinds = system_kinds(sess)
    if k not in kinds[1:]:
        raise ValueError(f"지울 수 있는 계통도 칸이 아닙니다: {kind!r}")
    if _slot_active(sess) == k:
        _slot_switch(sess, kinds[kinds.index(k) - 1])
    with _SWITCH_LOCK:
        sess["system_extra"] = [x for x in (sess.get("system_extra") or ()) if x != k]
        (sess.get("slots") or {}).pop(k, None)
    return _slot_active(sess)

# -*- coding: utf-8 -*-
"""[H-1] 제5국면 S700 — 평면도·계통도·기계실을 한 배관망으로.

특허 제5국면은 다섯 단계다::

    S710  급수방식 선택        펌프 가압 / 자연낙차 / 1차 감압 / 2차 감압
    S720  급수방식별 입상관 구성
    S730  기계실 배관 전단 접속   수원 위치 이동 · 낙차 부여
    S740  입상관–헤드배관 결합    기준점 번호 10 을 공통 절점으로
    S750  수리계산 입력파일 생성

이 다섯을 **새로 짜지 않는다.** `core/remote30_full_network.py` 에 이미
모듈 레벨 함수로 다 있다 — `build_riser`(S720 · 급수방식 4종이 그대로
`build_riser_*` 네 갈래다) · `prepend_machine_room_to_riser`(S730) ·
`stitch_riser_and_heads`(S740) · `emit_full_sdf`(S750). 모듈 A 의 통합
라우트(`routes/r30_combined.py`)도 같은 것을 부른다.

그래서 이 파일이 하는 일은 **접합** 하나다.

──────────────────────────────────────────────────────────────────────
왜 A 의 통합 라우트를 재사용하지 않는가
──────────────────────────────────────────────────────────────────────
`remote30_combined_build()` 는 590줄이지만 그 대부분이 «A 의 평면 경로» 다 —
`_PROTOTYPE_JOBS` 에서 잡을 꺼내 `detected_heads` 로 최불리를 고르고
`build_input_tables` 로 헤드망 표를 만든다. F 의 평면 쪽은 그 경로가 아니다.
F 는 사람이 손질한 board 위에서 G 의 `select_and_expand` → `build_design_tables`
로 헤드망 표를 만든다(모듈 E 의 판단 철학).

즉 **다른 것은 평면 쪽 뿐이고, S700 원시함수는 이미 공유돼 있다.** 그러니
라우트 본문을 들어올릴 이유가 없다 — `r30_combined.py` 는 손대지 않는다.
(당초 지시서 H-D2 는 그 승격을 계획했으나, 실측으로 원시함수가 전부 모듈
레벨임을 확인해 접합만 하는 쪽으로 바꿨다. 사본은 여전히 만들지 않는다.)

──────────────────────────────────────────────────────────────────────
접합의 핵심 — 기준점 번호
──────────────────────────────────────────────────────────────────────
특허 S550 은 «기준점 번호 = 10», S740 은 «기준점 번호 10 을 공통 절점으로
결합» 이라고 못박는다. A 의 헤드망도 그 규약이다(라벨 {10, 11, 12, …} 에서
10 이 급수원 = AV).

그런데 G 의 표는 BFS 순서로 1 부터 번호를 매긴다 — 급수원이 «1» 이다.
그대로 결합하면 (1) 기준점이 10 이 아니고 (2) G 의 1~9 가 라이저의 1~9 와
정면으로 충돌한다.

**+9 오프셋** 하나로 둘 다 풀린다: G 의 1 → 10(기준점), 2 → 11, … 이것이
정확히 A 의 헤드망 규약이다. 라벨은 노드표에만 있는 것이 아니라 배관·노즐·
부속·기기의 in/out 에도 박혀 있으므로 **전부 같이** 옮긴다 — 한 곳이라도
빠지면 표가 고아 참조를 갖고, PIPENET 은 그것을 조용히 «Unset» 으로 읽는다.
"""
from __future__ import annotations

from dataclasses import dataclass

# 특허 S550 · S740 — 기준점(급수원 = 알람밸브 접속점)의 번호.
ANCHOR_LABEL = "10"
# G 의 BFS 번호(1부터)를 A 규약(10부터)으로 옮기는 오프셋.
LABEL_OFFSET = int(ANCHOR_LABEL) - 1

# ── S710 급수방식 4종 ────────────────────────────────────────────────
# 키는 **엔진의 이름 그대로** 다(`remote30_full_network.ZoneType` 의 값).
# 특허의 네 갈래와 정확히 1:1 이므로 새 이름을 지어 사전을 하나 더 두지
# 않는다 — 이름이 둘이면 어느 쪽이 권위인지를 매번 정해야 한다.
#
#     펌프 가압   → HSP_PUMP      build_riser_hsp_pump
#     자연낙차    → LSP_GRAVITY   build_riser_lsp_gravity
#     1차 감압    → LSP_1STAGE    build_riser_lsp_1stage
#     2차 감압    → LLSP_2STAGE   build_riser_llsp_2stage
#
# 네 빌더 모두 `av_node_label="10"` 을 세운다 — 특허 S740 의 기준점과 같다.
SUPPLY_MODES = {
    "hsp_pump":    "펌프 가압",
    "lsp_gravity": "자연낙차 (고가수조)",
    "lsp_1stage":  "1차 감압",
    "llsp_2stage": "2차 감압",
}
# 펌프 가압이면 수원·기계실이 망 최하부에 놓인다(3D 아이소뷰 Z 방향이 뒤집힌다).
PUMP_MODES = frozenset({"hsp_pump"})


class MergeError(ValueError):
    """접합이 성립하지 않는다 — 임의로 메우지 않고 올린다(S340 원칙 승계)."""


@dataclass
class HeadTables:
    """A 의 `PipeTables` 규약을 만족하는 헤드망 표.

    G 의 `PipeTablesG` 와 필드가 같다(G 가 A 규약을 그대로 따랐다). 그래도
    새로 담는 이유는 **라벨을 옮겨야** 하기 때문이다 — 원본을 제자리에서
    고치면 같은 세션의 표 보기·산출이 함께 흔들린다.
    """

    nodes: list
    pipes: list
    nozzles: list
    fittings: list
    equipment: list
    meta: list


def _shift(label, offset: int = LABEL_OFFSET):
    """라벨 하나를 옮긴다. 숫자가 아니면 그대로 둔다(`?` · `@/3` 등)."""
    s = str(label)
    try:
        return str(int(s) + offset)
    except (TypeError, ValueError):
        return s


def to_head_tables(tbl, *, offset: int = LABEL_OFFSET) -> HeadTables:
    """G 의 설계 표 → A 의 헤드망 표 규약(기준점 10).

    라벨이 박혀 있는 자리를 빠짐없이 옮긴다::

        nodes.label
        pipes.label(이름은 그대로) · pipes.in · pipes.out
        nozzles.in                     (nozzles.out 은 `@/n` 노즐 참조라 불변)
        fittings.in · fittings.out     (fittings.pipe 는 배관 이름이라 불변)
        equipment.in · equipment.out   (equipment.pipe 도 배관 이름)

    ★배관·부속·기기의 `pipe`/`label` 은 **노드 라벨이 아니다.** 같이 옮기면
      배관 이름이 어긋나 부속표가 통째로 고아가 된다(실측으로 한 번 그랬다).
    """
    if tbl is None:
        raise MergeError("설계 표가 없습니다 — 먼저 표를 확정하세요.")

    def sh(v):
        return _shift(v, offset)

    nodes = []
    for r in (getattr(tbl, "nodes", None) or ()):
        row = dict(r)
        row["label"] = sh(row.get("label"))
        nodes.append(row)

    pipes = []
    for r in (getattr(tbl, "pipes", None) or ()):
        row = dict(r)
        row["in"] = sh(row.get("in"))
        row["out"] = sh(row.get("out"))
        pipes.append(row)

    nozzles = []
    for r in (getattr(tbl, "nozzles", None) or ()):
        row = dict(r)
        row["in"] = sh(row.get("in"))
        nozzles.append(row)

    fittings = []
    for r in (getattr(tbl, "fittings", None) or ()):
        row = dict(r)
        row["in"] = sh(row.get("in"))
        row["out"] = sh(row.get("out"))
        fittings.append(row)

    equipment = []
    for r in (getattr(tbl, "equipment", None) or ()):
        row = dict(r)
        row["in"] = sh(row.get("in"))
        row["out"] = sh(row.get("out"))
        equipment.append(row)

    out = HeadTables(nodes=nodes, pipes=pipes, nozzles=nozzles,
                     fittings=fittings, equipment=equipment,
                     meta=list(getattr(tbl, "meta", None) or ()))
    _check_anchor(out)
    return out


def _check_anchor(ht: HeadTables) -> None:
    """기준점이 10 이고 급수원인가 — S740 이 성립하는지 여기서 본다."""
    labels = {str(n.get("label")) for n in ht.nodes}
    if ANCHOR_LABEL not in labels:
        raise MergeError(
            f"헤드망에 기준점 «{ANCHOR_LABEL}» 이 없습니다 — 결합할 절점이 "
            f"없습니다 (특허 S740). 라벨: {sorted(labels)[:8]}…")
    anchor = next(n for n in ht.nodes if str(n.get("label")) == ANCHOR_LABEL)
    if str(anchor.get("io_node")) != "Input":
        raise MergeError(
            f"기준점 «{ANCHOR_LABEL}» 이 급수원(Input)이 아닙니다 — "
            f"G 의 BFS 뿌리와 어긋났습니다 (io_node={anchor.get('io_node')!r}).")
    # 고아 참조 — 옮기다 한 자리를 빠뜨리면 여기서 잡힌다.
    for name, rows, keys in (("배관", ht.pipes, ("in", "out")),
                             ("노즐", ht.nozzles, ("in",)),
                             ("부속", ht.fittings, ("in", "out")),
                             ("기기", ht.equipment, ("in", "out"))):
        for r in rows:
            for k in keys:
                v = str(r.get(k))
                if v not in labels and v not in ("?", "None"):
                    raise MergeError(
                        f"{name}표가 없는 절점을 가리킵니다: {r.get('label') or r.get('pipe')}"
                        f".{k}={v!r}")


def check_supply_mode(mode) -> str:
    """S710 — 사람이 고른 급수방식. 자동 추정하지 않는다(H-D4).

    도면에는 급수방식이 적혀 있지 않다. 관종·상하향과 같은 부류로, 설계 협의에서
    정해지는 값이다 — 여기서 추측하면 라이저 구조가 통째로 달라진다.
    """
    m = str(mode or "").strip().lower()
    if m not in SUPPLY_MODES:
        raise MergeError(
            "급수방식을 먼저 고르세요 (S710) — "
            + " · ".join(f"{k}({v})" for k, v in SUPPLY_MODES.items()))
    return m


def zone_type_of(mode: str):
    """급수방식 이름 → 엔진의 `ZoneType`."""
    from remote30_full_network import ZoneType
    return ZoneType(check_supply_mode(mode))


def build_riser_for(mode: str, ctx):
    """S720 — 고른 급수방식으로 입상관을 만든다.

    `ctx.zone_spec.zone_type` 이 라우터의 분기 키다. 여기서 그것만 갈아끼우고
    나머지 판단은 전부 엔진에 맡긴다.
    """
    from remote30_full_network import build_riser
    zt = zone_type_of(mode)
    spec = getattr(ctx, "zone_spec", None)
    if spec is None:
        raise MergeError("project_context 에 zone_spec 이 없습니다.")
    spec.zone_type = zt
    return build_riser(ctx)

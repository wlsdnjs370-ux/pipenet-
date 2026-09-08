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

from dataclasses import dataclass, field

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
    # [D2] 끝점이 «?»·«None» 인 행 — 예외로 막지 않고 **세어서** 올린다.
    #   표를 만들 때 라벨을 못 찾은 배관 끝점이 `label_of.get(a, "?")` 로
    #   그렇게 남는데(design/tables.py `pipe_row`), 고아 검사가 그 표식을
    #   명시적으로 면제해 왔다 — 끝점 없는 배관이 검사 셋을 다 지나 SDF 까지
    #   갔다. 지금 예외로 승격하면 돌던 실행이 통째로 실패하므로, 건수만
    #   보고하고 승격 여부는 사람이 정한다(지시서 D2).
    dangling: list = field(default_factory=list)


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
    out.dangling = _check_anchor(out)["dangling"]
    return out


def _check_anchor(ht: HeadTables) -> dict:
    """기준점이 10 이고 급수원인가 — S740 이 성립하는지 여기서 본다.

    [D2] 끝점이 «?»·«None» 인 행은 **모아서 돌려준다**(예외 아님). 모르는
    절점을 가리키는 것은 종전대로 즉시 올린다 — 그쪽은 «옮기다 빠뜨린» 자리라
    성격이 다르다.
    """
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
    dangling: list = []
    for name, rows, keys in (("배관", ht.pipes, ("in", "out")),
                             ("노즐", ht.nozzles, ("in",)),
                             ("부속", ht.fittings, ("in", "out")),
                             ("기기", ht.equipment, ("in", "out"))):
        for r in rows:
            for k in keys:
                v = str(r.get(k))
                if v in ("?", "None"):
                    dangling.append((name, str(r.get("label")
                                               or r.get("pipe") or ""), k, v))
                elif v not in labels:
                    raise MergeError(
                        f"{name}표가 없는 절점을 가리킵니다: {r.get('label') or r.get('pipe')}"
                        f".{k}={v!r}")
    return {"dangling": dangling}


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


def riser_tables_from(riser: dict):
    """계통도 추출 결과(dict) → 엔진의 `RiserTables`.

    `extract_system_path` 는 dict 를 돌려주고 S730·S740 은 `RiserTables` 를
    받는다. 좌표는 **손대지 않는다** — 계통도의 실좌표를 수직 막대로 재배치하는
    일은 `stitch_riser_and_heads` 안의 `_layout_riser_as_schematic` 이 이미
    한다(그 함수의 주석이 v1 실좌표 경로를 명시적으로 다룬다). 여기서 미리
    옮기면 그 배치와 이중으로 어긋난다.
    """
    from remote30_full_network import RiserTables
    r = riser or {}
    nodes = list(r.get("nodes") or [])
    pipes = list(r.get("pipes") or [])
    if not nodes or not pipes:
        raise MergeError("계통도 추출 결과가 비어 있습니다 — 입상관을 만들 수 없습니다.")
    av = str(r.get("av_node_label") or "")
    if not av:
        raise MergeError("계통도에 알람밸브 절점이 없습니다 (S740 결합점).")
    return RiserTables(nodes=nodes, pipes=pipes,
                       pumps=list(r.get("pumps") or []),
                       valves=list(r.get("valves") or []),
                       av_node_label=av)


def label_offset_for(method) -> int:
    """평면도가 어느 길로 왔는지에 따라 라벨 오프셋이 갈린다.

    G(수동 · E 경로)의 표는 BFS 로 **1** 부터 매기므로 +9 를 먹여 기준점을 10 으로
    올린다. 모듈 A(자동)의 `build_input_tables` 는 **처음부터 10** 이다
    (`counter = [10]`) — 거기에 또 +9 를 먹이면 기준점이 19 가 되어 S740 결합이
    성립하지 않는다. 두 경로의 표가 같은 자리에 들어가므로 여기서 가른다.
    """
    return 0 if str(method or "").lower() == "auto" else LABEL_OFFSET


def merge_network(head_tbl, *, riser=None, machineroom=None, mode: str,
                  source_drop_m: float = 0.0, pump=None,
                  method: str = "manual",
                  head_orientation: str = "pendent",
                  head_stub_pct: float = 2.5):
    """S720 → S730 → S740 — 세 도면을 한 배관망으로.

    `head_tbl` 은 G 의 설계 표(`PipeTablesG`) 그대로 받는다 — 라벨 옮기기는
    여기서 한다. `riser`·`machineroom` 은 각 슬롯의 추출 결과 dict 이며 **없어도
    된다**: 계통도가 없으면 평면도 단독으로 지나간다(지시서 H-5).

    반환: {"combined", "head_tables", "attached", "steps"} — steps 는 어느
    단계가 실제로 돌았는지다. 화면이 «기계실을 붙였다» 고 말하려면 근거가 있어야
    한다(붙이지 못했는데 붙였다고 하면 그 순간 보고가 거짓이 된다).
    """
    from remote30_full_network import (
        normalize_pipe_bores, prepend_machine_room_to_riser,
        stitch_riser_and_heads)

    mode = check_supply_mode(mode)
    is_pump = mode in PUMP_MODES
    off = label_offset_for(method)
    ht = to_head_tables(head_tbl, offset=off)
    steps: list[str] = [
        "S740 기준점 10 정합"
        + (" (자동 경로 — 이미 10)" if off == 0 else f" (+{off})")]
    # [D2] 끝점 없는 배관 — 0 건이면 줄을 넣지 않는다(없는 것을 말하지 않는다).
    if ht.dangling:
        _kinds = sorted({d[0] for d in ht.dangling})
        steps.append(f"★끝점 없는 배관 {len(ht.dangling)}건 "
                     f"({' · '.join(_kinds)}) — 표에서 라벨을 못 찾은 자리입니다")

    if not riser:
        # 계통도가 없다 — 평면도 단독. 결합할 입상관이 없으므로 여기서 끝난다.
        return {"combined": None, "head_tables": ht, "attached": False,
                "mode": mode, "steps": steps + ["계통도 없음 — 평면도 단독"]}

    rt = riser_tables_from(riser)
    steps.append(f"S720 입상관 ({SUPPLY_MODES[mode]}) · 절점 {len(rt.nodes)}")

    # ★[D1] 기계실 평면이 «붙는 자리» = 라이저의 Input 노드. **prepend 전에**
    #   잡아 둔다.
    #
    #   `prepend_machine_room_to_riser` 는 기계실을 앞에 붙여 돌려주므로 그
    #   뒤의 `rt.nodes[0]` 은 기계실 수원(m1)이다. 종전에는 그것을
    #   `pump_junction_label` 로 넘겼는데, `stitch` 는 그 라벨을
    #   `translated_riser_nodes`(= 기계실 라벨을 이미 **제외한** 목록)에서
    #   찾는다 → `pump_node` 가 **항상 None** → 기계실 노드가 원 DXF 좌표에
    #   방치되고 평면 형상(plan_edges)도 통째로 빈다.
    #   실측(대명동 3장): bbox span 36,150 → 986,199 mm · emit 배율 0.083 →
    #   0.003 · 기계실 12노드 중 11개가 원좌표 그대로 · plan_edges 0.
    #
    #   고르는 규칙을 `prepend_machine_room_to_riser` 와 **같게** 둔다 —
    #   두 곳이 다른 노드를 고르면 평면이 엉뚱한 데 붙는다.
    _riser_input_label = next(
        (str(n.get("label")) for n in rt.nodes
         if str(n.get("io_node", "")).lower() == "input"), None)
    if _riser_input_label is None:
        _riser_input_label = next(
            (str(n.get("label")) for n in rt.nodes
             if str(n.get("label")) == "1"), None)

    mr_labels: list[str] = []
    mr_plan_edges = None
    mr_conn_xy = None
    attached = False
    if machineroom:
        mr_labels = [str(n.get("label")) for n in (machineroom.get("nodes") or ())]
        mr_plan_edges = machineroom.get("plan_edges")
        conn = machineroom.get("conn_xy")
        if conn and len(conn) >= 2:
            try:
                mr_conn_xy = (float(conn[0]), float(conn[1]))
            except (TypeError, ValueError):
                mr_conn_xy = None
        rt, attached = prepend_machine_room_to_riser(
            machineroom, rt, at_bottom=is_pump,
            source_drop_below_lowest_m=float(source_drop_m or 0.0))
        # ★붙었는지를 그대로 전한다. 엔진은 못 붙이면 원본 라이저를 조용히
        #   돌려준다(안전) — 그 조용함을 화면까지 들고 가면 안 된다.
        steps.append(f"S730 기계실 전단 접속 · {'성공' if attached else '미접속'}"
                     + (f" · 수원 낙차 {source_drop_m} m" if is_pump else ""))
        if not attached:
            mr_labels = []

    combined = stitch_riser_and_heads(
        rt, ht,
        machine_room_labels=mr_labels or None,
        pump_junction_label=(_riser_input_label if attached else None),
        machine_room_plan_edges=mr_plan_edges,
        machine_room_at_bottom=is_pump,
        machine_room_conn_xy=mr_conn_xy,
    )

    if is_pump and pump:
        from remote30_full_network import insert_source_pump
        try:
            combined = insert_source_pump(
                combined,
                rated_q_lpm=float(pump.get("rated_q_lpm") or 0.0),
                rated_h_m=float(pump.get("rated_h_m") or 0.0),
                count=int(pump.get("count") or 1))
            steps.append("펌프 삽입 (수원 직후)")
        except (TypeError, ValueError) as exc:
            raise MergeError(f"펌프 제원이 올바르지 않습니다: {exc}") from None

    # 관경 꼬임 정규화 — 상류(입상관)가 하류(가지)보다 얇아지는 것을 편다.
    # 결합 전에는 두 망이 각자 관경을 정했으므로 이음매에서 꼬이기 쉽다.
    try:
        fixed = normalize_pipe_bores(combined.nodes, combined.pipes)
        steps.append(f"관경 정규화 · 고친 배관 {fixed}")
    except Exception as exc:  # noqa: BLE001 — 정규화 실패로 결합을 버리지 않는다
        steps.append(f"관경 정규화 건너뜀 ({type(exc).__name__}: {exc})")

    # ★어느 절점이 «어느 도면에서 왔는지» 를 남긴다. 결합망을 화면에 그릴 때
    #   세 도면을 색으로 갈라 보여야 «통합된 형태» 가 눈에 들어온다. 라벨은
    #   결합 뒤에도 안 바뀌므로(평면도만 +offset) 여기서 한 번 세워 두면 된다.
    mr = set(mr_labels or ())
    riser_labels = [str(n.get("label")) for n in rt.nodes
                    if str(n.get("label")) not in mr]
    out = {"combined": combined, "head_tables": ht, "attached": attached,
           "mode": mode, "steps": steps,
           # 기계실 평면이 라이저에 붙는 그 노드 — 아이소로 굽을 때 기계실
           # 군집을 어디에 다시 맞출지의 기준이다.
           "pump_junction": (_riser_input_label if attached else None),
           "parts": {"plan": [str(n.get("label")) for n in ht.nodes],
                     "system": riser_labels,
                     "machineroom": sorted(mr)}}
    # [D5] 결합 뒤 검사 — 보고만 한다(예외 아님). 이상이 있으면 그 사실을
    #   단계 기록에 남겨 화면이 그대로 읽게 한다.
    out["checks"] = check_combined(out)
    ck = out["checks"]
    if ck.get("components", 1) != 1:
        steps.append(f"★연결성분 {ck['components']}개 {ck['component_sizes']}"
                     f" — 한 망으로 안 붙었습니다")
    if ck.get("dangling_pipes_n"):
        steps.append(f"★노드표에 없는 끝점을 가리키는 배관 "
                     f"{ck['dangling_pipes_n']}건")
    if ck.get("orphan_fittings") or ck.get("orphan_equipment"):
        steps.append(f"★고아 부속 {len(ck['orphan_fittings'])} · 고아 기기 "
                     f"{len(ck['orphan_equipment'])}")
    if len(ck.get("inputs") or ()) != 1:
        steps.append(f"★급수원(Input) 절점이 {len(ck.get('inputs') or ())}개"
                     f" — 정확히 1 이어야 합니다")
    if ck.get("anchor_gap"):
        steps.append(f"S740 두 기준점 {ck['anchor_gap']}")
    return out


def bake_combined_iso(got: dict, *, iso_z_scale: float = 1.0):
    """결합망을 30° 아이소매트릭 좌표로 굽는다 — **화면과 파일이 쓰는 그 한 식**.

    ★부위마다 «맞는» 투영이 다르다. 한 식으로 다 굽으면 깨진다:

      · 평면도 — 평면이니 30° 회전 + 표고 lift. lift 는 평면과 같은 자
        (1 m = 1000 · §T3 로 절점 좌표가 mm)라 설계 화면과 규칙이 같다.
      · 계통도(라이저) — schematic y 가 **이미 수직**이다. 회전을 먹이면
        수직 막대가 사선이 된다(실측 x 퍼짐 1,732 · 사용자 지적 「계통도가
        기울어져 있다」). 기준점의 아이소 자리에 평면 오프셋을 그대로 얹는다.
      · 기계실 — 평면 군집이라 회전하되, 접속점(펌프 junction)이 라이저의
        «새» 자리에 그대로 붙도록 평행이동한다. 안 하면 이음매가 찢어진다.

    돌려주는 것: (절점 사본, 기계실 평면 edge 사본) — 원본은 건드리지 않는다.
    산출(.sdf)과 미리보기가 **같은 함수**를 써야 「보이는 것 = 저장되는 것」이
    성립한다(이 저장소가 설계 화면에서 이미 값을 치른 규칙이다).
    """
    c = (got or {}).get("combined")
    if c is None:
        return [], []
    parts = (got.get("parts") or {})
    of = {}
    for kind in ("system", "machineroom", "plan"):
        for lab in (parts.get(kind) or ()):
            of[str(lab)] = kind

    nodes = [dict(n) for n in (getattr(c, "nodes", None) or ())]
    zs = float(iso_z_scale or 1.0)
    cos30, sin30 = 0.8660254037844387, 0.5

    def _rot(x, y):
        return ((x - y) * cos30, (x + y) * sin30)

    at0 = {str(n.get("label")): (float(n.get("x", 0) or 0),
                                 float(n.get("y", 0) or 0)) for n in nodes}
    ax, ay = at0.get(ANCHOR_LABEL, (0.0, 0.0))
    a_iso = _rot(ax, ay)
    e_ref = next((float(n.get("elevation", 0) or 0) for n in nodes
                  if str(n.get("label")) == ANCHOR_LABEL), 0.0)
    lift = 1000.0 * zs

    pj = got.get("pump_junction")
    pj_xy = at0.get(str(pj)) if pj else None
    shift = (0.0, 0.0)
    if pj_xy is not None:
        new_pj = (a_iso[0] + (pj_xy[0] - ax), a_iso[1] + (pj_xy[1] - ay))
        rot_pj = _rot(*pj_xy)
        shift = (new_pj[0] - rot_pj[0], new_pj[1] - rot_pj[1])

    for n in nodes:
        lab = str(n.get("label"))
        x = float(n.get("x", 0) or 0)
        y = float(n.get("y", 0) or 0)
        kind = of.get(lab, "plan")
        if kind == "system":
            n["x"] = a_iso[0] + (x - ax)
            n["y"] = a_iso[1] + (y - ay)
        elif kind == "machineroom":
            rx, ry = _rot(x, y)
            n["x"] = rx + shift[0]
            n["y"] = ry + shift[1]
        else:
            rx, ry = _rot(x, y)
            n["x"] = rx
            n["y"] = ry + (float(n.get("elevation", 0) or 0) - e_ref) * lift

    edges = []
    for e in (getattr(c, "machine_room_plan_edges", None) or ()):
        r1 = _rot(float(e[0]), float(e[1]))
        r2 = _rot(float(e[2]), float(e[3]))
        edges.append([r1[0] + shift[0], r1[1] + shift[1],
                      r2[0] + shift[0], r2[1] + shift[1]])
    return nodes, edges


def check_combined(got: dict) -> dict:
    """[D5] 결합 **뒤** 검사 — 전부 «보고» 다. 예외로 올리지 않는다.

    종전에는 결합 뒤를 보는 코드가 하나도 없었다(`api_merge` 는 급수방식만
    검사했다). 결합이 성립했는지 판단할 근거를 여기서 만든다 — 판정은 사람이
    한다. 값은 화면 응답에도 그대로 실린다.
    """
    c = (got or {}).get("combined")
    if c is None:
        return {"combined": False}
    nodes = list(getattr(c, "nodes", None) or ())
    pipes = list(getattr(c, "pipes", None) or ())
    labels = {str(n.get("label")) for n in nodes}
    plabels = {str(p.get("label")) for p in pipes}

    adj = {lab: set() for lab in labels}
    dangling = []
    for p in pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        if a in adj and b in adj:
            adj[a].add(b)
            adj[b].add(a)
        else:
            dangling.append({"pipe": str(p.get("label")), "in": a, "out": b})
    seen, sizes = set(), []
    for n in adj:
        if n in seen:
            continue
        stack, size = [n], 0
        seen.add(n)
        while stack:
            u = stack.pop()
            size += 1
            for v in adj[u]:
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
        sizes.append(size)
    sizes.sort(reverse=True)

    xs = [float(n.get("x", 0) or 0) for n in nodes]
    ys = [float(n.get("y", 0) or 0) for n in nodes]
    meta = dict(getattr(c, "meta", None) or ())
    return {
        "combined": True,
        "components": len(sizes),
        "component_sizes": sizes[:6],
        "dangling_pipes": dangling[:20],
        "dangling_pipes_n": len(dangling),
        "orphan_fittings": [str(f.get("pipe"))
                            for f in (getattr(c, "fittings", None) or ())
                            if str(f.get("pipe")) not in plabels][:20],
        "orphan_equipment": [str(e.get("pipe"))
                             for e in (getattr(c, "equipment", None) or ())
                             if e.get("pipe")
                             and str(e.get("pipe")) not in plabels][:20],
        "inputs": [str(n.get("label")) for n in nodes
                   if str(n.get("io_node", "")).lower() == "input"],
        "anchor_gap": meta.get("S740 두 기준점 거리"),
        "renamed": meta.get("배관 라벨 개명"),
        "bbox": ({"minx": min(xs), "maxx": max(xs),
                  "miny": min(ys), "maxy": max(ys),
                  "span_x": max(xs) - min(xs), "span_y": max(ys) - min(ys)}
                 if xs and ys else None),
    }


def combined_summary(got: dict) -> dict:
    """결합 결과 한 장 — 화면·산출이 같은 수치를 말하게 한다."""
    c = got.get("combined")
    if c is None:
        return {"merged": False, "mode": got.get("mode"),
                "steps": got.get("steps") or [],
                "nodes": len(got["head_tables"].nodes),
                "pipes": len(got["head_tables"].pipes),
                "nozzles": len(got["head_tables"].nozzles)}
    return {
        "merged": True,
        "mode": got.get("mode"),
        "attached": bool(got.get("attached")),
        "steps": got.get("steps") or [],
        # [D5] 결합 뒤 검사 — 화면이 볼 수 있게 그대로 통과시킨다.
        "checks": got.get("checks"),
        "nodes": len(getattr(c, "nodes", ()) or ()),
        "pipes": len(getattr(c, "pipes", ()) or ()),
        "nozzles": len(getattr(c, "nozzles", ()) or ()),
        "pumps": len(getattr(c, "pumps", ()) or ()),
        "valves": len(getattr(c, "valves", ()) or ()),
        "fittings": len(getattr(c, "fittings", ()) or ()),
    }


# [정리 2026-08-31] `build_riser_for(mode, ctx)` 를 지웠다 — S720 입상관 생성의
#   얇은 어댑터였는데 **저장소 어디에서도 안 불렸다**(비추적 파일까지 훑었다).
#   짝인 `zone_type_of` 는 시험이 직접 쓰므로 그대로 둔다.
#   되살릴 일이 생기면 커밋 이력에 있다 — 죽은 채로 두면 「쓰이는 코드」로 오해된다.

"""통합 배관망 자체 수리해석 — 과토출 유량으로 유속을 수렴시키는 내경 사이징.

왜 필요한가::

    도면 텍스트 관경 + 규약 최소경으로 SDF 를 내면 PIPENET 수리계산에서 많은
    구간이 유속 상한(가지 6, 그 외 10 m/s)을 넘는다. 기존 사이징
    (remote30_full_network.size_pipes_by_velocity)이 헤드당 80 L/min 고정
    **수요유량**으로 1 회 산정하기 때문이다. 실제 해석은 헤드를 K√P 오리피스로
    풀어, 펌프에 가까운 헤드일수록 80 을 크게 넘겨 토출(과토출)한다. 그 합이
    상류 배관을 지나며 유속을 밀어올린다.

    이 모듈은 같은 원리(Hazen-Williams + K√P)로 망을 직접 풀어 과토출 유량을
    얻고, 그 유량으로 유속 초과 배관을 한 치수 올린 뒤 다시 푸는 것을 더 이상
    바뀌지 않을 때까지 반복한다.

해석 방식 (설계 기준 역산)::

    펌프 양정을 몰라도 풀 수 있게, 수리적으로 가장 먼 헤드의 방수압을 NFTC
    최소치(0.1 MPa)에 고정하고 나머지 헤드 압력을 그 상대치로 얻는다.
    → 헤드유량 가정 → 트리 누적으로 배관유량 → 마찰·표고 손실 → 헤드압력
    → q=K√P 로 헤드유량 갱신 … 을 감쇠 고정점 반복으로 수렴시킨다.

한계 (의도적)::

    * 피팅(<Fittings>)의 등가길이는 모델링하지 않는다. equipment 의 명시적
      등가길이만 더한다. 손실 과소 → 헤드 간 압력차 과소 → 과토출 과소 평가
      방향이라, 사이징 판정에는 safety 여유를 곱해 상쇄한다.
    * 루프(분기영역 루프)가 있으면 트리 누적이 일부 유량을 이중 계상한다.
      과대 산정 = 안전측이라 그대로 둔다.
"""

from __future__ import annotations

import math
from collections.abc import Iterable
from dataclasses import dataclass, field
from functools import lru_cache

from hb_rules import (
    BRANCH_PIPE_V_LIMIT,
    KGFCM2_TO_BAR,
    MAIN_PIPE_V_LIMIT,
    get_inner_diameter_mm,
)
from nftc_rules import head_pressure_min_mpa
from remote30_full_network import PIPE_BORE_LADDER_MM

# 표준형 스프링클러 헤드: q[L/min] = K·√P[bar]. 80 L/min @ 0.1 MPa → K = 80.
NOZZLE_K_FACTOR = 80.0
BAR_PER_M_WATER = 0.0980665
DEFAULT_C_FACTOR = 120.0
DEFAULT_PIPE_MATERIAL = "KSD 3507"

_PI_OVER_4 = 0.7853981633974483
_MPA_PER_BAR = 0.1


def hazen_williams_drop_bar(q_lpm: float, length_m: float, c_factor: float,
                            inner_dia_mm: float) -> float:
    """Hazen-Williams 마찰손실 [bar]. 유량 L/min, 길이 m, 내경 mm.

    식은 pipenet_validator._calc_hw_friction_loss_kgcm2 와 같지만(계수 6.174e4)
    그 결과는 kgf/cm² 다. 이 솔버는 표고손실을 BAR_PER_M_WATER 로 bar 로 환산해
    더하고 헤드 K 계수도 bar 기준이므로, 여기서 bar 로 맞춰야 두 항이 같은
    척도가 된다. 두 구현의 일치는 tests/test_hydraulic_solver.py 가 검산한다.
    """
    if q_lpm <= 0 or length_m <= 0 or inner_dia_mm <= 0 or c_factor <= 0:
        return 0.0
    dp_kgfcm2 = (6.174e4 * (q_lpm ** 1.85) * length_m
                 / ((c_factor ** 1.85) * (inner_dia_mm ** 4.87))) / 0.1
    return dp_kgfcm2 * KGFCM2_TO_BAR


@lru_cache(maxsize=256)
def inner_diameter_mm(nominal_mm: float, material: str = DEFAULT_PIPE_MATERIAL) -> float:
    """호칭경(mm) → KS 실내경(mm). 표에 없으면 호칭값을 그대로 쓴다.

    유속·마찰손실은 실내경 기준이어야 한다. 호칭 25A 의 실내경은 27.5 mm 로
    호칭값보다 크므로, 호칭값을 쓰면 유속이 과대 평가되어 불필요하게 굵어진다.
    """
    try:
        nominal = int(round(float(nominal_mm)))
    except (TypeError, ValueError):
        return 0.0
    if nominal <= 0:
        return 0.0
    return get_inner_diameter_mm(f"{nominal}A", material) or float(nominal)


def unresolved_inner_diameters(bores: Iterable[float], material: str) -> dict[str, int]:
    """실내경을 못 찾아 호칭값으로 때운 구간을 호칭별로 센다.

    때우는 것 자체는 유지한다 — 해석이 멈추면 안 된다. 다만 때웠다는 사실이
    리포트에 드러나야 한다. CPVC2 는 표에 25A~80A 뿐이므로 그 위가 세어졌다면
    "굵은 CPVC 배관" 이 아니라 재질 판단이 틀렸다는 신호다.
    """
    out: dict[str, int] = {}
    for b in bores:
        try:
            nominal = int(round(float(b)))
        except (TypeError, ValueError):
            continue
        if nominal <= 0:
            continue
        if get_inner_diameter_mm(f"{nominal}A", material) is None:
            key = f"{nominal}A"
            out[key] = out.get(key, 0) + 1
    return out


def velocity_mps(q_lpm: float, inner_dia_mm: float) -> float:
    """평균 유속 [m/s] — v = Q/A. 내경 0 이면 inf."""
    if inner_dia_mm <= 0:
        return float("inf")
    area = _PI_OVER_4 * (inner_dia_mm / 1000.0) ** 2
    return (float(q_lpm) / 60000.0) / area


def _next_bore(nominal_mm: float) -> int | None:
    """사다리에서 한 치수 위. 최대치면 None."""
    for d in PIPE_BORE_LADDER_MM:
        if d > nominal_mm + 1e-9:
            return d
    return None


def _snap_up(nominal_mm: float) -> int:
    """사다리 값 이상으로 스냅(올림). 최대치 초과면 최대치."""
    for d in PIPE_BORE_LADDER_MM:
        if d >= nominal_mm - 1e-9:
            return d
    return PIPE_BORE_LADDER_MM[-1]


# ────────────────────────────────────────────────────────────────────────────
# 위상
# ────────────────────────────────────────────────────────────────────────────


@dataclass
class Topology:
    """source 기준으로 방향이 잡힌 배관망. 배열 인덱스는 pipes 리스트와 일치."""

    source: str
    depth: dict[str, int]
    up: list[str]
    down: list[str]
    order: list[int] = field(default_factory=list)      # 하류→상류 (깊은 것부터)
    children: dict[str, list[int]] = field(default_factory=dict)
    head_count: dict[str, int] = field(default_factory=dict)
    elevation: dict[str, float] = field(default_factory=dict)


_UNREACHED = 10 ** 9


def build_topology(nodes: list[dict], pipes: list[dict],
                   nozzles: list[dict] | None = None) -> Topology:
    """Input 노드에서 BFS 하여 각 배관의 상류/하류 끝점과 처리 순서를 정한다."""
    adj: dict[str, list[str]] = {}
    for p in pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        adj.setdefault(a, []).append(b)
        adj.setdefault(b, []).append(a)

    source = next((str(n.get("label")) for n in nodes
                   if str(n.get("io_node", "")).lower() == "input"), None)
    if source is None or source not in adj:
        source = str(pipes[0].get("in")) if pipes else ""

    depth: dict[str, int] = {source: 0} if source else {}
    queue = [source] if source else []
    head = 0
    while head < len(queue):
        cur = queue[head]
        head += 1
        for nb in adj.get(cur, ()):
            if nb not in depth:
                depth[nb] = depth[cur] + 1
                queue.append(nb)

    def _d(lbl: str) -> int:
        return depth.get(lbl, _UNREACHED)

    up: list[str] = []
    down: list[str] = []
    for p in pipes:
        a, b = str(p.get("in")), str(p.get("out"))
        if _d(a) <= _d(b):
            up.append(a); down.append(b)
        else:
            up.append(b); down.append(a)

    children: dict[str, list[int]] = {}
    for i, u in enumerate(up):
        children.setdefault(u, []).append(i)

    head_count: dict[str, int] = {}
    for nz in (nozzles or []):
        lbl = str(nz.get("in") or nz.get("input_node") or nz.get("input") or "")
        if lbl:
            head_count[lbl] = head_count.get(lbl, 0) + 1

    elevation: dict[str, float] = {}
    for n in nodes:
        try:
            elevation[str(n.get("label"))] = float(n.get("elevation") or 0.0)
        except (TypeError, ValueError):
            elevation[str(n.get("label"))] = 0.0

    order = sorted(range(len(pipes)), key=lambda i: _d(down[i]), reverse=True)
    return Topology(source=source, depth=depth, up=up, down=down, order=order,
                    children=children, head_count=head_count, elevation=elevation)


def classify_pipe_roles(topo: Topology, pipes: list[dict]) -> list[str]:
    """배관별 역할 — "branch"(가지배관) 또는 "other"(교차·주·입상).

    NFTC 의 6/10 m/s 는 배관 **역할**로 갈린다. 호칭경(≤50A)으로 대신 가르면
    "관경을 정하려고 관경으로 상한을 정하는" 순환이 생기므로 위상으로 판정한다.

    가지배관 = 하류 부분망이 헤드를 포함하면서 분기가 없는 단일 사슬.
    말단 교차배관(가지 하나만 먹이는 구간)은 가지로 분류되지만, 더 엄격한
    6 m/s 를 적용하는 쪽이라 안전측이다.

    노즐이 하나도 없으면 판정 근거가 없으므로 호칭경 규칙(≤50A → 가지)으로
    물러난다.
    """
    if not topo.head_count:
        return ["branch" if _bore_of(p) <= 50 else "other" for p in pipes]

    is_chain = [False] * len(pipes)
    has_head = [False] * len(pipes)
    for i in topo.order:  # 하류부터 — 자식은 이미 확정
        kids = topo.children.get(topo.down[i], ())
        is_chain[i] = len(kids) <= 1 and all(is_chain[k] for k in kids)
        has_head[i] = (topo.head_count.get(topo.down[i], 0) > 0
                       or any(has_head[k] for k in kids))
    return ["branch" if (is_chain[i] and has_head[i]) else "other"
            for i in range(len(pipes))]


def _bore_of(p: dict) -> float:
    try:
        return float(p.get("dia") or 0.0)
    except (TypeError, ValueError):
        return 0.0


# ────────────────────────────────────────────────────────────────────────────
# 해석
# ────────────────────────────────────────────────────────────────────────────


def accumulate_flows(topo: Topology, head_flow: dict[str, float]) -> list[float]:
    """하류→상류 누적으로 배관별 통과유량 [L/min]. 인덱스는 pipes 와 일치."""
    flows = [0.0] * len(topo.up)
    subtree = dict(head_flow)
    for i in topo.order:
        q = subtree.get(topo.down[i], 0.0)
        flows[i] = q
        subtree[topo.up[i]] = subtree.get(topo.up[i], 0.0) + q
    return flows


def demand_flows(topo: Topology, head_lpm: float = 80.0) -> list[float]:
    """헤드당 설계유량 고정(과토출 미고려) 기준의 배관 유량 — 예비 사이징용."""
    return accumulate_flows(topo, {h: head_lpm * c for h, c in topo.head_count.items()})


@dataclass
class Solution:
    pipe_flow_lpm: list[float]
    node_pressure_bar: dict[str, float]
    head_flow_lpm: dict[str, float]
    source_pressure_bar: float
    iterations: int
    converged: bool


def _pipe_arrays(pipes: list[dict], equiv_length_m: dict[str, float] | None,
                 material: str) -> tuple[list[float], list[float], list[float]]:
    """해석에 쓸 (등가길이 포함 길이, C계수, 실내경) 배열."""
    eq = equiv_length_m or {}
    lengths, c_factors = [], []
    for p in pipes:
        try:
            L = float(p.get("length") or 0.0)
        except (TypeError, ValueError):
            L = 0.0
        lengths.append(max(0.0, L + eq.get(str(p.get("label", "")), 0.0)))
        try:
            c_factors.append(float(p.get("c") or DEFAULT_C_FACTOR))
        except (TypeError, ValueError):
            c_factors.append(DEFAULT_C_FACTOR)
    return lengths, c_factors, [inner_diameter_mm(_bore_of(p), material) for p in pipes]


def solve_network(
    pipes: list[dict],
    topo: Topology,
    *,
    equiv_length_m: dict[str, float] | None = None,
    material: str = DEFAULT_PIPE_MATERIAL,
    k_factor: float = NOZZLE_K_FACTOR,
    min_head_bar: float | None = None,
    damping: float = 0.5,
    tol_lpm: float = 0.2,
    max_iter: int = 80,
) -> Solution:
    """헤드 과토출을 포함한 배관 유량·노드 압력을 감쇠 고정점 반복으로 푼다.

    최원단 헤드의 방수압을 min_head_bar 에 고정하고 나머지를 상대 압력으로
    얻으므로 펌프 양정 입력이 필요 없다.
    """
    if min_head_bar is None:
        min_head_bar = head_pressure_min_mpa() / _MPA_PER_BAR

    n = len(pipes)
    lengths, c_factors, bores = _pipe_arrays(pipes, equiv_length_m, material)

    head_nodes = list(topo.head_count)
    head_flow = {h: k_factor * math.sqrt(min_head_bar) * topo.head_count[h]
                 for h in head_nodes}
    ascending = list(reversed(topo.order))

    pipe_flow = [0.0] * n
    pressure: dict[str, float] = {}
    source_p = min_head_bar
    iterations = 0
    converged = not head_nodes  # 헤드가 없으면 반복할 대상 자체가 없다

    for iterations in range(1, max_iter + 1):
        pipe_flow = accumulate_flows(topo, head_flow)

        rel = {topo.source: 0.0}
        for i in ascending:
            u, d = topo.up[i], topo.down[i]
            if u not in rel:
                continue
            drop = hazen_williams_drop_bar(pipe_flow[i], lengths[i], c_factors[i], bores[i])
            lift = (topo.elevation.get(d, 0.0) - topo.elevation.get(u, 0.0)) * BAR_PER_M_WATER
            rel[d] = rel[u] - drop - lift

        reached = [h for h in head_nodes if h in rel]
        if not reached:
            break
        source_p = min_head_bar - min(rel[h] for h in reached)
        pressure = {lbl: v + source_p for lbl, v in rel.items()}

        delta = 0.0
        for h in reached:
            target = k_factor * math.sqrt(max(0.0, pressure[h])) * topo.head_count[h]
            new = head_flow[h] + damping * (target - head_flow[h])
            # 폭주 방지 — 심하게 얇은 망에서는 (말단압 고정 탓에) 근접 헤드 압력이
            # 한 회차 만에 수십 배로 튀어 반복이 발산한다. 회차당 배율을 묶으면
            # 같은 고정점에 기하급수적으로 접근하되 오버슈트하지 않는다.
            new = min(max(new, head_flow[h] * 0.5), head_flow[h] * 2.0)
            delta = max(delta, abs(new - head_flow[h]))
            head_flow[h] = new
        if delta < tol_lpm:
            # 소스에서 도달 못 한 헤드가 있으면 그 헤드의 유량·압력은 초기 가정
            # 그대로다 — 나머지가 안정됐다고 "수렴"이라 부르면 근거 없는 합격이 된다.
            converged = len(reached) == len(head_nodes)
            break

    return Solution(pipe_flow_lpm=pipe_flow, node_pressure_bar=pressure,
                    head_flow_lpm=head_flow, source_pressure_bar=source_p,
                    iterations=iterations, converged=converged)


def solve_at_source_pressure(
    pipes: list[dict],
    topo: Topology,
    source_bar: float,
    *,
    equiv_length_m: dict[str, float] | None = None,
    material: str = DEFAULT_PIPE_MATERIAL,
    k_factor: float = NOZZLE_K_FACTOR,
    damping: float = 0.4,
    tol_lpm: float = 0.05,
    max_iter: int = 400,
) -> Solution:
    """수원 게이지압을 못 박고 푼다 — K-Fire_Solver/EPANET 이 실제로 푸는 조건.

    solve_network 는 최원단 헤드 방수압을 고정해 필요 수원압을 역산한다(설계기준).
    산출물 KFP/SDF 는 반대로 수원 조건을 싣고 나가므로, 옥상수조처럼 수두가 설계
    필요압을 넘으면 헤드가 과토출하고 같은 관경에서도 유속이 몇 배로 뛴다. 그
    괴리를 재려면 내보낼 경계조건 그대로 한 번 더 풀어야 한다.
    """
    lengths, c_factors, bores = _pipe_arrays(pipes, equiv_length_m, material)

    head_flow = {h: k_factor * math.sqrt(max(0.0, source_bar)) * c
                 for h, c in topo.head_count.items()}
    ascending = list(reversed(topo.order))
    pipe_flow = [0.0] * len(pipes)
    pressure: dict[str, float] = {}
    iterations = 0
    converged = not topo.head_count

    for iterations in range(1, max_iter + 1):
        pipe_flow = accumulate_flows(topo, head_flow)
        pressure = {topo.source: source_bar}
        for i in ascending:
            u, d = topo.up[i], topo.down[i]
            if u not in pressure:
                continue
            drop = hazen_williams_drop_bar(pipe_flow[i], lengths[i], c_factors[i], bores[i])
            lift = (topo.elevation.get(d, 0.0) - topo.elevation.get(u, 0.0)) * BAR_PER_M_WATER
            pressure[d] = pressure[u] - drop - lift
        reached = [h for h in topo.head_count if h in pressure]
        if not reached:
            break
        delta = 0.0
        for h in reached:
            target = k_factor * math.sqrt(max(0.0, pressure[h])) * topo.head_count[h]
            new = head_flow[h] + damping * (target - head_flow[h])
            delta = max(delta, abs(new - head_flow[h]))
            head_flow[h] = new
        if delta < tol_lpm:
            converged = len(reached) == len(topo.head_count)
            break

    return Solution(pipe_flow_lpm=pipe_flow, node_pressure_bar=pressure,
                    head_flow_lpm=head_flow, source_pressure_bar=source_bar,
                    iterations=iterations, converged=converged)


# ────────────────────────────────────────────────────────────────────────────
# 수렴 사이징
# ────────────────────────────────────────────────────────────────────────────


def converge_bores_by_velocity(
    nodes: list[dict],
    pipes: list[dict],
    nozzles: list[dict],
    *,
    equipment: list[dict] | None = None,
    material: str = DEFAULT_PIPE_MATERIAL,
    safety: float = 1.1,
    keep_existing: bool = True,
    max_outer: int = 24,
    export_source_bar: float | None = None,
    head_stub_bore_mm: float | None = None,
    reset_bores: list[float] | None = None,
) -> dict:
    """모든 배관이 유속 상한 이하가 될 때까지 해석↔내경승급을 반복 (in-place).

    한 회차 = 망을 풀어 과토출 유량을 얻고 → 상한 초과 배관을 승급하고 → 상류 ≥
    하류 단조성을 회복. 내경은 오직 커지기만 하고 사다리 최대치(200A)가 상한이라
    반드시 종료한다. 해석이 발산한 회차(극단적으로 얇은 망은 말단압 0.1 MPa 를
    만족시킬 수 없어 소스압이 폭주)에는 유량 크기를 믿지 않고 한 치수만 올린다.

    Args:
        safety: 사이징 판정에만 곱하는 유량 여유. 피팅 등가길이 미모델링분을 상쇄.
        keep_existing: 도면/규약이 준 기존 내경보다 얇게 만들지 않음.
        export_source_bar: 산출물이 실제로 싣고 나가는 수원 게이지압 [bar]. 주면
            사이징 후 그 경계로 한 번 더 풀어 `export` 블록(과토출·유속)을 담는다.
        head_stub_bore_mm: 헤드 접속 신축배관 호칭경. 이 구간은 제조사 규격품이라
            승급 대상이 아니므로 판정만 하고 `규격고정` 사유로 표에 남긴다.
        reset_bores: 주면 사이징 전에 각 배관 내경을 이 값으로 되돌린다. 이미 한 번
            승급된 망을 도면 상태에서 다시 역산할 때 쓴다.

    기록 필드 (배관 dict, 표시·검증용):
        flow_lpm, velocity_mps, v_limit, v_over, pipe_role, inner_dia_mm

    Returns:
        수렴 리포트 dict. `changes` 는 관경이 바뀐 배관과 상한을 못 맞춘 배관의
        전후 대조표, `baseline` 은 승급 전 전 배관 현황 — 승급 사유가 유속인지
        단조성 회복인지 구분해 담는다.
    """
    empty = {"iterations": 0, "converged": True, "changed": 0,
             "max_velocity_before": 0.0, "max_velocity_after": 0.0,
             "violations_before": 0, "violations_after": 0,
             "source_pressure_bar": 0.0, "overdischarge_ratio_max": 1.0,
             "solver_converged": True, "pipes": 0, "head_stub_bore_mm": None,
             "head_stub_violations": 0, "head_stub_max_velocity": 0.0,
             "unresolved_inner_dia": {"material": material, "by_nominal": {}},
             "export": None, "changes": [], "baseline": []}
    if not pipes:
        return empty

    if reset_bores is not None:
        for p, bore in zip(pipes, reset_bores):
            p["dia"] = bore

    topo = build_topology(nodes, pipes, nozzles)
    roles = classify_pipe_roles(topo, pipes)
    limits = [BRANCH_PIPE_V_LIMIT if r == "branch" else MAIN_PIPE_V_LIMIT for r in roles]

    equiv: dict[str, float] = {}
    for eq in (equipment or []):
        lbl = str(eq.get("pipe", ""))
        try:
            equiv[lbl] = equiv.get(lbl, 0.0) + float(eq.get("eq_len") or 0.0)
        except (TypeError, ValueError):
            continue

    def _velocities(flows: list[float]) -> list[float]:
        return [velocity_mps(flows[i], inner_diameter_mm(_bore_of(p), material))
                for i, p in enumerate(pipes)]

    def _summary(vs: list[float]) -> tuple[float, int]:
        mx = max((v for v in vs if v != float("inf")), default=0.0)
        viol = sum(1 for i, v in enumerate(vs) if v > limits[i] + 1e-9)
        return mx, viol

    # 진입 상태(도면 관경 + 규약 최소경) 를 기존 방식 그대로 — 헤드당 설계유량
    # 고정 — 으로 채점한 값이 "before". 과토출 기준으로 재면 아래 예비 사이징이
    # 없는 상태에선 해가 발산해 비교값이 되지 못한다.
    bore_before = [_bore_of(p) for p in pipes]
    demand = demand_flows(topo)
    v_before = _velocities(demand)
    mv_before, vio_before = _summary(v_before)
    # 승급 전 전 배관 현황 — `changes` 는 바뀐 것만 담아 설계 검토용 전체 표가 되지
    # 못한다. 어디가 상한을 넘는지 먼저 보여주고 역산 여부를 판단하게 하려면 도면
    # 그대로의 모든 구간이 필요하다.
    baseline = [{
        "label": str(p.get("label", "")),
        "in": str(p.get("in", "")), "out": str(p.get("out", "")),
        "role": roles[i],
        "bore_mm": round(bore_before[i], 1),
        "inner_dia_mm": (round(inner_diameter_mm(bore_before[i], material), 1)
                         if bore_before[i] > 0 else None),
        "flow_lpm": round(demand[i], 1),
        "velocity_mps": (round(v_before[i], 2)
                         if v_before[i] != float("inf") else None),
        "v_limit": limits[i],
        "v_over": v_before[i] > limits[i] + 1e-9,
    } for i, p in enumerate(pipes)]
    # 승급 사유 — 유속 때문에 굵어진 것과 상류≥하류를 맞추느라 굵어진 것은
    # 설계 검토에서 의미가 달라 분리 기록한다.
    reasons: list[list[str]] = [[] for _ in pipes]

    def _mark(i: int, reason: str) -> None:
        if reason not in reasons[i]:
            reasons[i].append(reason)

    for i, p in enumerate(pipes):
        if keep_existing:
            p["dia"] = int(_snap_up(bore_before[i]) if bore_before[i] > 0
                           else PIPE_BORE_LADDER_MM[0])
        else:
            p["dia"] = int(PIPE_BORE_LADDER_MM[0])
        if _bore_of(p) != bore_before[i]:
            _mark(i, "기준초기화" if not keep_existing else "규격스냅")

    def _bump(flows: list[float], *, single_step: bool = False) -> int:
        """유속 상한을 넘는 배관을 승급하고 단조성 회복. 바뀐 횟수 반환.

        single_step 이면 회차당 한 치수만 올린다 — 해석이 발산했을 때는 유량의
        크기를 믿을 수 없고 "너무 얇다"는 방향만 유효하기 때문.
        """
        changed = 0
        for i, p in enumerate(pipes):
            bore = _bore_of(p)
            while velocity_mps(flows[i] * safety, inner_diameter_mm(bore, material)) > limits[i]:
                nxt = _next_bore(bore)
                if nxt is None:
                    break
                bore = nxt
                changed += 1
                if single_step:
                    break
            if bore != _bore_of(p):
                p["dia"] = int(bore)
                _mark(i, "유속초과")
        # 상류가 하류보다 얇으면 배관망이 성립하지 않는다 — 단조 비증가 회복.
        for i in topo.order:
            kid_max = max((_bore_of(pipes[k]) for k in topo.children.get(topo.down[i], ())),
                          default=0.0)
            if kid_max > _bore_of(pipes[i]):
                pipes[i]["dia"] = int(kid_max)
                _mark(i, "단조성회복")
                changed += 1
        return changed

    # 예비 사이징 — 설계유량으로 먼저 상식적인 굵기를 만든다. 25A 일색 상태에서
    # 바로 해석에 들어가면 말단압 고정 탓에 근접 헤드 유량이 폭주해 전 구간이
    # 최대경으로 튄다.
    changed_total = _bump(demand)

    sol = solve_network(pipes, topo, equiv_length_m=equiv, material=material)
    outer = 0
    for outer in range(1, max_outer + 1):
        changed = _bump(sol.pipe_flow_lpm, single_step=not sol.converged)
        changed_total += changed
        if changed == 0:
            break
        sol = solve_network(pipes, topo, equiv_length_m=equiv, material=material)

    v_after = _velocities(sol.pipe_flow_lpm)
    mv_after, vio_after = _summary(v_after)
    changes = []
    for i, p in enumerate(pipes):
        idia = inner_diameter_mm(_bore_of(p), material)
        v = v_after[i]
        over = v > limits[i] + 1e-9
        p["flow_lpm"] = round(sol.pipe_flow_lpm[i], 1)
        p["velocity_mps"] = round(v, 2) if v != float("inf") else None
        p["v_limit"] = limits[i]
        p["v_over"] = over
        p["pipe_role"] = roles[i]
        p["inner_dia_mm"] = round(idia, 1)
        # 상한을 못 맞춘 배관은 관경이 안 바뀌었어도 표에 남긴다 — 사다리 최대치로도
        # 해소되지 않는 구간은 설계 불가가 아니라 대개 추출 오접합 신호라 판정 보류다.
        if _bore_of(p) == bore_before[i] and not over:
            continue
        if over:
            _mark(i, "한계도달")
        changes.append({
            "label": str(p.get("label", "")),
            "in": str(p.get("in", "")), "out": str(p.get("out", "")),
            "role": roles[i],
            "bore_before": round(bore_before[i], 1),
            "bore_after": int(_bore_of(p)),
            "inner_dia_mm": round(idia, 1),
            "velocity_before": (round(v_before[i], 2)
                                if v_before[i] != float("inf") else None),
            "velocity_after": round(v, 2) if v != float("inf") else None,
            "v_limit": limits[i],
            "v_over": over,
            "flow_lpm": round(sol.pipe_flow_lpm[i], 1),
            "reasons": reasons[i],
        })

    base = NOZZLE_K_FACTOR * math.sqrt(head_pressure_min_mpa() / _MPA_PER_BAR)
    per_head = [q / max(1, topo.head_count.get(h, 1))
                for h, q in sol.head_flow_lpm.items()]

    # 산출물 경계조건 재검증 — 사이징은 최원단 0.1 MPa 를 못 박고 필요 수원압을
    # 역산하지만, KFP/SDF 는 수원 조건을 싣고 나간다. 옥상수조 수두가 필요압을
    # 넘으면 전 헤드가 과토출해 같은 관경에서도 유속이 몇 배가 된다.
    export = None
    export_sol = None
    if export_source_bar is not None:
        export_sol = solve_at_source_pressure(pipes, topo, export_source_bar,
                                              equiv_length_m=equiv, material=material)
        ev = _velocities(export_sol.pipe_flow_lpm)
        emx, eviol = _summary(ev)
        eper = [q / max(1, topo.head_count.get(h, 1))
                for h, q in export_sol.head_flow_lpm.items()]
        export = {
            "source_pressure_bar": round(export_source_bar, 3),
            "required_source_bar": round(sol.source_pressure_bar, 3),
            "surplus_bar": round(export_source_bar - sol.source_pressure_bar, 3),
            "max_velocity": round(emx, 2),
            "violations": eviol,
            "head_flow_min": round(min(eper), 1) if eper else 0.0,
            "head_flow_max": round(max(eper), 1) if eper else 0.0,
            "overdischarge_ratio_max": round(max(eper) / base, 2) if eper else 1.0,
            "solver_converged": export_sol.converged,
        }

    # 헤드 접속 신축배관 — 산출 단계에서 헤드마다 삽입되는 규격품이라 승급 대상이
    # 아니다. 사이징 대상 배관 목록에 없어 표에서 통째로 빠지면 이 구간의 유속
    # 초과가 산출물에서만 드러나므로, 상한을 넘는 것만 판정해 남긴다.
    stub_violations = 0
    stub_max_v = 0.0
    if head_stub_bore_mm:
        stub_inner = inner_diameter_mm(head_stub_bore_mm, material)
        flow_of = (export_sol or sol).head_flow_lpm
        for h, q in sorted(flow_of.items(), key=lambda kv: -kv[1]):
            per = q / max(1, topo.head_count.get(h, 1))
            v = velocity_mps(per, stub_inner)
            stub_max_v = max(stub_max_v, v if v != float("inf") else 0.0)
            if v <= BRANCH_PIPE_V_LIMIT + 1e-9:
                continue
            stub_violations += 1
            changes.append({
                "label": f"FX@{h}", "in": str(h), "out": str(h), "role": "branch",
                "bore_before": round(float(head_stub_bore_mm), 1),
                "bore_after": int(head_stub_bore_mm),
                "inner_dia_mm": round(stub_inner, 1),
                "velocity_before": round(v, 2), "velocity_after": round(v, 2),
                "v_limit": BRANCH_PIPE_V_LIMIT, "v_over": True,
                "flow_lpm": round(per, 1), "reasons": ["규격고정"],
            })

    return {
        "iterations": outer,
        "converged": vio_after == 0,
        "changed": changed_total,
        "max_velocity_before": round(mv_before, 2),
        "max_velocity_after": round(mv_after, 2),
        "violations_before": vio_before,
        "violations_after": vio_after,
        "source_pressure_bar": round(sol.source_pressure_bar, 3),
        "overdischarge_ratio_max": round(max(per_head) / base, 2) if per_head else 1.0,
        "solver_converged": sol.converged,
        "pipes": len(pipes),
        "head_stub_bore_mm": head_stub_bore_mm,
        "head_stub_violations": stub_violations,
        "head_stub_max_velocity": round(stub_max_v, 2),
        "unresolved_inner_dia": {
            "material": material,
            "by_nominal": unresolved_inner_diameters(
                [_bore_of(p) for p in pipes]
                + ([head_stub_bore_mm] if head_stub_bore_mm else []),
                material),
        },
        "export": export,
        "changes": changes,
        "baseline": baseline,
    }

# -*- coding: utf-8 -*-
"""link-prediction 연결복원 — 학습데이터 생성 (정답 SDF 자가지도).

전략
====
DXF 에는 신뢰할 ground truth 가 없다(모든 도면이 끝점 미스냅 작도 → raw 그래프 파탄,
clean_candidate_survey.py 참조). 반면 답안 SDF 는 파이프마다 start/end 노드 ID 가
명시된 **완벽한 위상 정답**이다. 그래서:

  1) 답안 SDF 의 깨끗한 그래프를 읽고
  2) LH306 실측 분포(틈≈1세그먼트)에 맞춰 junction 노드를 'un-snap'(합성손상)해
     CAD 처럼 끊고
  3) 손상 전 위상으로 끊긴 끝단쌍을 자동 라벨링(양성=원래 같은 노드, 음성=근접 타 노드)
  4) **추론 시점에 관측 가능한 피처만** 뽑는다 (손상 전 정보 누수 금지).

피처 (모두 손상 후 관측 가능)
  · gap         : 두 끝단 거리
  · gap_norm    : gap / 중앙 엣지길이
  · align_A/B   : 각 파이프 진행방향이 상대 끝단을 향하는 각(0=정면)
  · align_max/min
  · cos_dirs    : 두 파이프 방향 코사인 (직선연속≈-1, 직교 tee≈0)
  · nominal_match / nominal_absdiff_norm : 호칭경 일치/차이
  · pt_deg_A/B  : 각 끝단 위치의 군집 차수 (loose tip=1)
"""
from __future__ import annotations

import math
import random
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from kfp_sdf_converter import parse_sdf  # noqa: E402

FEATURE_NAMES = [
    "gap", "gap_norm", "align_A", "align_B", "align_max", "align_min",
    "cos_dirs", "nominal_match", "nominal_absdiff_norm",
]


# ── 끝단(endpoint) 표현 ──────────────────────────────────────────────────────
@dataclass
class _EP:
    pipe_id: str
    side: str          # 's' | 'e'
    node_id: str       # 손상 전 원래 노드 (라벨용 — 피처엔 안 씀)
    coord: tuple       # 손상 후 좌표 (x, y)
    anchor: tuple      # 파이프 반대편(또는 인접 waypoint) 좌표 — 방향 산출용
    nominal: float     # 호칭경 mm (0=미상)


def _node_xy(net, nid):
    n = net.nodes.get(nid)
    return (n.x, n.y) if n is not None else (0.0, 0.0)


def _pipe_anchor(net, pipe, side):
    """끝단의 인접 앵커 좌표 — waypoint 있으면 그쪽, 없으면 반대편 노드."""
    wps = pipe.waypoints or []
    if side == "s":
        return tuple(wps[0]) if wps else _node_xy(net, pipe.end)
    return tuple(wps[-1]) if wps else _node_xy(net, pipe.start)


def _median(vals):
    s = sorted(v for v in vals if v > 0)
    return s[len(s) // 2] if s else 1.0


def build_endpoints(net, *, break_prob=0.6, jitter_lo=0.25, jitter_hi=1.2,
                    rng=None) -> tuple[list[_EP], float]:
    """깨끗한 SDF → 손상된 끝단 목록 + 중앙 엣지길이.

    junction(차수≥2) 노드를 break_prob 확률로 un-snap: 인접 파이프 끝단 좌표를
    노드좌표 + jitter(=U(lo,hi)×국소 엣지길이, 임의 방향)로 흩뿌려 끊는다.
    차수1(헤드·수원 등 진짜 단말)은 끊지 않음 — 정상 단말.
    """
    rng = rng or random.Random()

    # 노드별 인접 끝단 + 차수
    incident: dict[str, list[tuple[str, str]]] = defaultdict(list)
    for pid, p in net.pipes.items():
        incident[p.start].append((pid, "s"))
        incident[p.end].append((pid, "e"))

    # 엣지 길이 (좌표 기반) — jitter 스케일·정규화 기준
    edge_len = {}
    for pid, p in net.pipes.items():
        a = _node_xy(net, p.start)
        b = _node_xy(net, p.end)
        edge_len[pid] = math.hypot(b[0] - a[0], b[1] - a[1])
    med_edge = _median(edge_len.values())

    broken = set()
    for nid, eps in incident.items():
        if len(eps) >= 2 and rng.random() < break_prob:
            broken.add(nid)

    eps_out: list[_EP] = []
    for pid, p in net.pipes.items():
        for side, nid in (("s", p.start), ("e", p.end)):
            base = _node_xy(net, nid)
            if nid in broken:
                # 국소 스케일 = 인접 엣지 길이 (없으면 전역 중앙값)
                scale = edge_len.get(pid) or med_edge
                mag = rng.uniform(jitter_lo, jitter_hi) * max(scale, med_edge * 0.1)
                ang = rng.uniform(0, 2 * math.pi)
                coord = (base[0] + mag * math.cos(ang), base[1] + mag * math.sin(ang))
            else:
                coord = base
            eps_out.append(_EP(
                pipe_id=pid, side=side, node_id=nid, coord=coord,
                anchor=_pipe_anchor(net, p, side),
                nominal=float(p.nominal_mm or 0.0),
            ))
    return eps_out, med_edge


# ── 후보쌍 + 피처 ────────────────────────────────────────────────────────────
def _unit(dx, dy):
    n = math.hypot(dx, dy)
    return (dx / n, dy / n) if n > 1e-12 else (0.0, 0.0)


def _angle_deg(u, v):
    c = max(-1.0, min(1.0, u[0] * v[0] + u[1] * v[1]))
    return math.degrees(math.acos(c))


def _point_degrees(eps, eps_tol):
    """끝단 위치의 군집 차수 — 같은 점에 몇 개의 끝단이 모여있나 (loose tip=1)."""
    cell = max(eps_tol, 1e-9)
    grid = defaultdict(list)
    for i, ep in enumerate(eps):
        grid[(int(ep.coord[0] // cell), int(ep.coord[1] // cell))].append(i)
    deg = [1] * len(eps)
    tol_sq = eps_tol * eps_tol
    for i, ep in enumerate(eps):
        kx, ky = int(ep.coord[0] // cell), int(ep.coord[1] // cell)
        cnt = 0
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for j in grid.get((kx + dx, ky + dy), ()):
                    o = eps[j]
                    if (o.coord[0] - ep.coord[0]) ** 2 + (o.coord[1] - ep.coord[1]) ** 2 <= tol_sq:
                        cnt += 1
        deg[i] = cnt  # 자기 자신 포함
    return deg


def candidate_pairs(eps, med_edge, *, search_factor=2.5, eps_tol_factor=0.02):
    """근접 끝단쌍 후보 생성 + 피처/라벨.

    search radius = search_factor × med_edge. 같은 위치에 모인(미손상 junction)
    끝단끼리는 후보에서 제외 — 추론 시 끊긴 끝단만 후보가 되므로.
    """
    R = search_factor * med_edge
    eps_tol = max(eps_tol_factor * med_edge, 1e-9)
    pt_deg = _point_degrees(eps, eps_tol)

    # loose 끝단만: 같은 점에 1개만 있는(=연결 안 된) 끝단
    loose = [i for i in range(len(eps)) if pt_deg[i] <= 1]

    cell = max(R, 1e-9)
    grid = defaultdict(list)
    for i in loose:
        grid[(int(eps[i].coord[0] // cell), int(eps[i].coord[1] // cell))].append(i)

    R_sq = R * R
    rows = []  # (feat_dict, label, i, j)
    seen = set()
    for i in loose:
        ei = eps[i]
        kx, ky = int(ei.coord[0] // cell), int(ei.coord[1] // cell)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for j in grid.get((kx + dx, ky + dy), ()):
                    if j == i:
                        continue
                    key = (min(i, j), max(i, j))
                    if key in seen:
                        continue
                    ej = eps[j]
                    gap2 = (ej.coord[0] - ei.coord[0]) ** 2 + (ej.coord[1] - ei.coord[1]) ** 2
                    if gap2 > R_sq:
                        continue
                    # 같은 파이프 양끝은 자기연결 — 제외
                    if ei.pipe_id == ej.pipe_id:
                        continue
                    seen.add(key)
                    feat = _features(ei, ej, med_edge)
                    label = 1 if ei.node_id == ej.node_id else 0
                    rows.append((feat, label, i, j))
    return rows


def _features(ei: _EP, ej: _EP, med_edge) -> dict:
    gap = math.hypot(ej.coord[0] - ei.coord[0], ej.coord[1] - ei.coord[1])
    di = _unit(ei.coord[0] - ei.anchor[0], ei.coord[1] - ei.anchor[1])  # A 진행방향(밖으로)
    dj = _unit(ej.coord[0] - ej.anchor[0], ej.coord[1] - ej.anchor[1])
    ab = _unit(ej.coord[0] - ei.coord[0], ej.coord[1] - ei.coord[1])
    ba = (-ab[0], -ab[1])
    align_a = _angle_deg(di, ab)   # A 가 B 를 향하는가 (0=정면)
    align_b = _angle_deg(dj, ba)
    cos_dirs = di[0] * dj[0] + di[1] * dj[1]
    nm = 1.0 if (ei.nominal > 0 and ej.nominal > 0 and ei.nominal == ej.nominal) else 0.0
    if ei.nominal > 0 and ej.nominal > 0:
        nd = abs(ei.nominal - ej.nominal) / max(ei.nominal, ej.nominal)
    else:
        nd = 1.0
    return {
        "gap": gap,
        "gap_norm": gap / med_edge if med_edge > 0 else gap,
        "align_A": align_a,
        "align_B": align_b,
        "align_max": max(align_a, align_b),
        "align_min": min(align_a, align_b),
        "cos_dirs": cos_dirs,
        "nominal_match": nm,
        "nominal_absdiff_norm": nd,
    }


def dataset_from_net(net, *, seed=0, **kw) -> list[tuple[dict, int]]:
    rng = random.Random(seed)
    eps, med = build_endpoints(net, rng=rng,
                               break_prob=kw.get("break_prob", 0.6),
                               jitter_lo=kw.get("jitter_lo", 0.25),
                               jitter_hi=kw.get("jitter_hi", 1.2))
    rows = candidate_pairs(eps, med,
                           search_factor=kw.get("search_factor", 2.5))
    return [(f, lab) for (f, lab, _i, _j) in rows]


if __name__ == "__main__":
    # 빠른 sanity: 첫 REMOTE 답안 한 건 손상→후보 통계
    import validate_sdf as vs
    files = vs._remote_answer_files()
    print(f"REMOTE 답안 {len(files)}건")
    f0 = files[0]
    net = parse_sdf(f0)
    xs = max(n.x for n in net.nodes.values()) - min(n.x for n in net.nodes.values())
    ys = max(n.y for n in net.nodes.values()) - min(n.y for n in net.nodes.values())
    print(f"표본: {Path(f0).name}")
    print(f"  노드 {len(net.nodes)} · 파이프 {len(net.pipes)} · 좌표 spread x={xs:.2f} y={ys:.2f}")
    ds = dataset_from_net(net, seed=1)
    pos = sum(1 for _f, l in ds if l == 1)
    print(f"  후보쌍 {len(ds)} (양성 {pos} · 음성 {len(ds) - pos})")

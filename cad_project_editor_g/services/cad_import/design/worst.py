# -*- coding: utf-8 -*-
"""[G1] 최불리 K 선정 — 앵커 방식(지시서 D1).

`routes/module_f/remote30.py` 에서 **로직 변경 없이** 옮겨 왔다. 순수 그래프
함수라 Qt·Flask 의존이 없다. 모듈 F 의 결과와 앵커·헤드 집합·far_m·max_load 가
완전히 일치해야 이식이 성공한 것이다(§G1 수용 기준).

「먼 순서 K개」가 아니라 앵커 방식인 이유는 worst_k_heads 의 docstring 에 있다.
"""
from __future__ import annotations

import heapq
import math
import sys
from pathlib import Path

# NFPC 103 이 요구하는 «가장 불리한 헤드 K개». 기본 30.
REMOTE_K_DEFAULT = 30


def worst_k_heads(pts, edges, hnodes, sources, k=REMOTE_K_DEFAULT,
                   only_heads=None) -> dict:
    """앵커 기반 «최불리 배관망» 추출 — 수리계산의 설계면적 그 자체.

    ─ 왜 «먼 순서 K개» 가 아니라 앵커인가 ────────────────────────────
    NFPC 103 의 기준개수(K)는 «하나의 설계구역 안에서 동시에 방수되는 인접
    K개» 다. 급수원에서 먼 순서로 그냥 K개를 뽑으면 도면 곳곳의 막다른 헤드가
    섞여 뽑힌다 — B1F 실측: 먼 순서 30개는 대각 95.9m 로 흩어졌고, 앵커 방식은
    30.3m 로 한 구역에 뭉쳤다. 흩어진 30개로는 설계면적이 성립하지 않는다.

    ─ 세 단계 ────────────────────────────────────────────────────────
    ① 앵커 = 급수원에서 **배관 거리로** 가장 먼(가장 불리한) 헤드. 여기가
       기준압을 잡는 지점 — 급수원↔앵커 거리가 «최원 유하거리» 다.
    ② 설계면적 = 앵커에서 **배관 거리로** 가까운 K개(유클리드 아님 — 실제 물이
       같은 관을 타고 함께 흐르는 무리라야 한다).
    ③ corridor = 그 K개를 급수원까지 잇는 최단경로의 합집합. 각 간선의
       **담당 헤드 수(load)** 를 함께 낸다 — NFPC 별표1 이 최소 호칭경을 정할
       때 쓰는 바로 그 값이라, 이 최대값이 주배관 관경을 결정한다.

    `only_heads` : 도면이 여러 장일 때 한 장으로 범위를 좁힌다. 앵커도 그
        범위 안에서 고른다(장이 다르면 앵커가 남의 도면으로 튄다).
    """
    adj: dict[int, list[int]] = {}
    for a, b in edges:
        adj.setdefault(a, []).append(b)
        adj.setdefault(b, []).append(a)

    def dijkstra(seeds):
        INF = float("inf")
        dist: dict[int, float] = {}
        prev: dict[int, int] = {}
        pq: list[tuple[float, int]] = []
        for s in seeds:
            if isinstance(s, int) and 0 <= s < len(pts):
                dist[s] = 0.0
                heapq.heappush(pq, (0.0, s))
        while pq:
            d, u = heapq.heappop(pq)
            if d > dist.get(u, INF):
                continue
            for v in adj.get(u, ()):
                nd = d + math.dist(pts[u], pts[v])
                if nd < dist.get(v, INF):
                    dist[v] = nd
                    prev[v] = u
                    heapq.heappush(pq, (nd, v))
        return dist, prev

    # ① 급수원 기점 — 헤드마다 부착 노드·유하거리
    src_dist, prev = dijkstra(list(sources))
    head_node: dict[int, int] = {}
    head_far: dict[int, float] = {}
    for hi, nodes in enumerate(hnodes):
        if only_heads is not None and hi not in only_heads:
            continue
        reach = [n for n in nodes if n in src_dist]
        if not reach:
            continue
        node = min(reach, key=lambda n: src_dist[n])
        head_node[hi] = node
        head_far[hi] = src_dist[node]

    reachable = len(head_far)
    empty = {"heads": [], "anchor": None, "edges": set(), "nodes": set(),
             "loads": {}, "reachable": reachable, "unreachable": 0,
             "far_m": 0.0, "near_m": 0.0, "span_m": 0.0, "total_m": 0.0,
             "max_load": 0}
    if not head_far:
        return empty

    k = max(1, min(int(k), reachable))
    anchor = max(head_far, key=head_far.get)   # 가장 불리한 헤드

    # ② 앵커 기점 — 배관 거리로 가까운 K개 = 설계면적
    an_dist, _ = dijkstra([head_node[anchor]])
    ranked = sorted(head_node,
                    key=lambda hi: an_dist.get(head_node[hi], float("inf")))
    picked = ranked[:k]
    span = an_dist.get(head_node[picked[-1]], 0.0) if picked else 0.0

    # ③ corridor — K개 → 급수원 경로 합집합 + 담당 헤드 수
    loads: dict[tuple[int, int], int] = {}
    keep_nodes: set[int] = set()
    for hi in picked:
        cur = head_node[hi]
        keep_nodes.add(cur)
        while cur in prev:
            nxt = prev[cur]
            key = (min(cur, nxt), max(cur, nxt))
            loads[key] = loads.get(key, 0) + 1
            keep_nodes.add(nxt)
            cur = nxt

    total = sum(math.dist(pts[a], pts[b]) for a, b in loads)
    return {
        "heads": picked,
        "anchor": anchor,
        "dists": {hi: head_far[hi] for hi in picked},
        "edges": set(loads),
        "loads": loads,
        "nodes": keep_nodes,
        "reachable": reachable,
        "unreachable": 0,          # picked 는 전부 도달 헤드 중에서 골랐다
        "far_m": round(head_far[anchor] / 1000.0, 2),   # 앵커 = 최원 유하거리
        "near_m": round(min(head_far[hi] for hi in picked) / 1000.0, 2),
        "span_m": round(span / 1000.0, 2),              # 설계면적 폭(배관거리)
        "total_m": round(total / 1000.0, 2),            # corridor 총연장
        "max_load": max(loads.values(), default=0),     # 주배관 관경 결정값
    }


def sheet_frames(board) -> list[dict]:
    """한 파일에 도면이 여러 장 들어 있는지 — 모듈 A 의 규칙을 그대로 부른다.

    국내 도서는 도면 한 장이 곧 파일 하나가 아니다(A 실측 — 죽전 6장·청라
    포레스트 3장·대구오페라 단위세대 5장). 여러 장을 한 망으로 보면 최불리 30 이
    서로 다른 도면의 헤드를 섞어 뽑아 계산이 성립하지 않는다.

    A 의 `detect_sheet_frames` 는 헤드 좌표(`.pos`)만 본다 — 문턱도 상수가 아니라
    그 도면의 헤드 간격에서 잰다. 그래서 규칙을 베끼지 않고 그대로 호출한다.
    """
    disks = getattr(board, "disks", None) or ()
    if len(disks) < 24:
        return []

    class _Head:  # A 가 보는 것은 .pos 하나뿐이다
        __slots__ = ("pos",)

        def __init__(self, p):
            self.pos = p

    # 모듈 A 는 저장소 루트에서만 import 된다. G 는 제 트리를 cwd 로 도는 별개
    # 프로세스라 루트가 sys.path 에 없다 — 여기서 붙인다(A 는 읽기 전용 참조다).
    # 저장소 밖으로 G 를 떼어 내면 아래 예외 경로가 그대로 받아 장 나누기만 꺼진다.
    _repo = Path(__file__).resolve().parents[4]
    if str(_repo) not in sys.path:
        sys.path.append(str(_repo))
    try:
        from remote30_prototype import detect_sheet_frames
    except Exception as exc:  # noqa: BLE001 — A 가 없어도 손질은 돌아야 한다
        print(f"[G] 도면 장 나누기 건너뜀 — 모듈 A 미탑재: {exc}")
        return []
    try:
        return detect_sheet_frames(
            [_Head((float(d[0]), float(d[1]))) for d in disks])
    except Exception as exc:  # noqa: BLE001
        print(f"[G] 도면 장 나누기 실패: {exc}")
        return []

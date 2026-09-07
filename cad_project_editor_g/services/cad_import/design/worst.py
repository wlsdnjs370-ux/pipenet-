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
                   only_heads=None, source_index: int | None = None,
                   head_xy=None) -> dict:
    """«가장 불리한 헤드 K개» — 수리계산의 설계면적 그 자체.

    ─ 규칙은 두 줄이다 (2026-09-07 · 사용자 확정) ────────────────────
    1순위  사람이 «영역» 을 지정했으면 **그 안의 헤드만** 후보다.
           (영역이 없으면 도면 전체가 후보 — 곧바로 2순위로 간다.)
    2순위  알람밸브(급수원) 노드에서 **배관을 따라 잰 길이**가 긴 순서로
           기준개수 K 개를 추린다. **가장 긴 것부터**.

    ★«배관을 따라» 가 핵심이다 — 직선 변위가 아니다. DXF 배관 레이어로 세운
      그래프 위에서 Dijkstra 로 잰다. 옆 가지관의 헤드는 공간으로 코앞이어도
      주관을 돌아오므로 유하거리가 길고, 그것이 곧 «불리함» 이다.

    ─ 여기까지 온 길 (BLOCKED §31) ──────────────────────────────────
    종전에는 ①앵커(가장 먼 헤드)를 잡고 ②그 둘레를 «채우는» 방식이었다.
    채우는 자를 배관거리 → 직사각형 → 가지관으로 세 번 바꿔 봤지만, 어느
    자든 «가장 불리한 헤드» 를 빠뜨리는 것이 문제였다 — 실측(B1F·K=30)에서
    가장 먼 30개 중 7개(23%)만 뽑히고 347.9 m 짜리 대신 306.9 m 짜리를
    골랐다. 지금 규칙은 그 자를 없앤다: **불리한 순서 그대로 K 개**다.

    ★대가는 적어 둔다. 뽑힌 K 개는 도면 여러 곳에 흩어질 수 있다(B1F 실측
      공간폭 50.4 m). 한 구역으로 모으고 싶으면 사람이 «영역» 을 지정한다 —
      그것이 1순위인 이유다. 프로그램이 대신 «구역» 을 지어내지 않는다.

    ─ 그리고 corridor ───────────────────────────────────────────────
    뽑은 K 개를 급수원까지 잇는 최단경로의 합집합. 각 간선의 **담당 헤드
    수(load)** 를 함께 낸다 — NFPC 별표1 이 최소 호칭경을 정할 때 쓰는 바로
    그 값이라, 이 최대값이 주배관 관경을 결정한다.

    `only_heads` : 1순위 «영역» 이 여기로 들어온다(도면 장 나누기도 같은 자리).
    `head_xy` : 헤드의 제 좌표(board 의 disks). 뽑힌 무리가 **얼마나 넓게
        퍼졌나**(area_*)를 재는 데만 쓴다 — 선정에는 쓰지 않는다.
    `source_index` : [F-1 · D4] 급수원이 여럿일 때 **어느 하나 기준**인지.
        지정하면 `sources[source_index]` 하나만 seed 로 Dijkstra 를 돈다 —
        전체망 `.kfp` 변환의 `source_selection_required` 와 같은 규약이다.
        급수원이 둘이면 앵커·최원 유하거리가 달라지므로, «어느 급수원에서든
        가장 먼 헤드» 는 수리계산 입력이 못 된다(BLOCKED B2 — 이것으로 해소).
        `None` 이면 종전 그대로 전부 seed(하위호환 — 산출 비트 동일).
    """
    if source_index is not None:
        src_list = list(sources)
        if not (0 <= int(source_index) < len(src_list)):
            raise ValueError(
                f"급수원 번호가 범위를 벗어났습니다: {source_index} "
                f"(급수원 {len(src_list)}곳)")
        sources = [src_list[int(source_index)]]

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
    empty = {"heads": [], "worst_head": None, "worst_path": [],
             "worst_path_m": 0.0, "edges": set(), "nodes": set(),
             "loads": {}, "reachable": reachable, "unreachable": 0,
             "far_m": 0.0, "near_m": 0.0, "span_m": 0.0, "total_m": 0.0,
             "area_w_m": 0.0, "area_h_m": 0.0, "area_m2": 0.0,
             "max_load": 0}
    if not head_far:
        return empty

    k = max(1, min(int(k), reachable))
    # 기준 헤드 = 가장 불리한 헤드. 2순위 규칙에서 그것은 곧 «1등» 이다.
    worst_head = min(head_far, key=lambda hi: (-head_far[hi], hi))

    # ② 2순위 — 유하거리가 **긴 순서 그대로** K 개. 가장 긴 것부터.
    #
    # ★같은 입력에 같은 산출이어야 한다. 유하거리가 똑같은 헤드가 여럿일 때
    #   (나란한 가지관에서 흔하다) 순서가 흔들리므로 번호로 못 박는다.
    ranked = sorted(head_node, key=lambda hi: (-head_far[hi], hi))
    picked = ranked[:k]

    # 뽑힌 무리가 얼마나 넓게 퍼졌나 — **선정에는 안 쓰고 보고만 한다.**
    #   흩어지면 그 사실이 수치로 보여야 사람이 «영역» 을 지정할지 판단한다.
    def _xy(hi):
        if head_xy is not None and hi < len(head_xy):
            p = head_xy[hi]
            return float(p[0]), float(p[1])
        return pts[head_node[hi]]

    xs = [_xy(hi)[0] for hi in picked]
    ys = [_xy(hi)[1] for hi in picked]
    box_w = max(xs) - min(xs)
    box_h = max(ys) - min(ys)
    span = max(box_w, box_h)

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

    # ④ 최원 유하거리 «경로» — 급수원 → 앵커. corridor 전체가 아니라 그 한 줄이다.
    #
    # far_m 은 이 경로의 길이인데, 화면에는 corridor 만 굵기로 그려져 있어서
    # «어느 줄이 그 거리인지» 가 안 보였다. 기준압을 잡는 지점이 앵커라면
    # 그 압이 어느 관을 타고 오는지도 같이 보여야 한다 — 관경을 키울지
    # 경로를 줄일지는 그 줄을 봐야 정할 수 있다.
    #
    # prev 는 ① 의 급수원 기점 Dijkstra 가 남긴 최단경로 트리다. 앵커의 부착
    # 노드에서 거슬러 올라가면 그 경로가 그대로 나온다(다시 풀지 않는다).
    worst_path: list[int] = []
    cur = head_node.get(worst_head)
    seen: set[int] = set()
    while cur is not None and cur not in seen:
        worst_path.append(cur)
        seen.add(cur)
        cur = prev.get(cur)
    worst_path.reverse()          # 접속점 → 기준 헤드 방향
    worst_path_m = round(
        sum(math.dist(pts[worst_path[i]], pts[worst_path[i + 1]])
            for i in range(len(worst_path) - 1)) / 1000.0, 2)

    return {
        "heads": picked,
        "worst_head": worst_head,
        # 급수원에서 앵커까지의 절점 열. 화면이 이 줄을 따로 그린다.
        "worst_path": worst_path,
        "worst_path_m": worst_path_m,
        "dists": {hi: head_far[hi] for hi in picked},
        "edges": set(loads),
        "loads": loads,
        "nodes": keep_nodes,
        "reachable": reachable,
        "unreachable": 0,          # picked 는 전부 도달 헤드 중에서 골랐다
        "far_m": round(head_far[worst_head] / 1000.0, 2),   # 기준 헤드까지 = 최원 유하거리
        "near_m": round(min(head_far[hi] for hi in picked) / 1000.0, 2),
        # 뽑힌 K 개가 걸친 범위 — **선정 기준이 아니라 결과의 모양**이다.
        # 흩어져 있으면 이 수가 커진다. 그때 «영역» 을 지정할지는 사람이 본다.
        "span_m": round(span / 1000.0, 2),
        "area_w_m": round(box_w / 1000.0, 2),
        "area_h_m": round(box_h / 1000.0, 2),
        "area_m2": round(box_w * box_h / 1e6, 1),
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

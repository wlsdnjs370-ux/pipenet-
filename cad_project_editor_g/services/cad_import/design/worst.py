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

# 설계면적을 «어떻게 채우나» — 사람이 고른다(BLOCKED §31).
#
#   rect   앵커를 품는 직사각형을 K개가 담길 때까지 넓힌다. 낱개 단위라
#          모양이 가장 조밀하지만, 한 가지관을 반만 담을 수 있다.
#   branch 앵커가 달린 «가지관» 을 통째로 담고, 공간으로 가까운 다음
#          가지관으로 넘어간다. 반쪽 가지관이 안 생긴다.
#
# 둘 다 방호구역은 모른다 — 그 판단은 사람이 «영역 지정» 으로 한다.
DESIGN_AREA_RULES = ("rect", "branch")
DESIGN_AREA_DEFAULT = "rect"


def _pipe_runs(edges) -> dict:
    """그래프를 «런» 으로 자른다 — 분기점 사이의 배관 한 줄이 곧 가지관이다.

    가지관은 이름이 아니라 **위상**으로 안다: 티(분기·차수 3 이상)와 끝(차수 1)
    사이를 잇는, 가운데가 전부 차수 2 인 길. 스프링클러 도면에서 그 한 줄이
    곧 «가지관» 이다.

    반환: 차수 2 인 노드 → 런 번호. 분기점·끝점은 여러 런에 걸치므로 뺀다.
    """
    adj: dict = {}
    for a, b in edges:
        adj.setdefault(a, set()).add(b)
        adj.setdefault(b, set()).add(a)
    run_of: dict = {}
    rid = 0
    for start in adj:
        if len(adj[start]) == 2 or start in run_of:
            continue
        # 분기점·끝점에서 출발해 차수 2 만 밟고 다음 분기점까지 간다.
        for nb in adj[start]:
            cur, prev = nb, start
            chain = []
            while cur not in run_of:
                deg = len(adj.get(cur, ()))
                if deg == 2:
                    chain.append(cur)
                    nxt = next(iter(adj[cur] - {prev}), None)
                    if nxt is None:
                        break
                    cur, prev = nxt, cur
                    continue
                # ★끝점(차수 1)은 **그 줄의 끝**이다 — 런에 넣는다.
                #   빼면 가지관 맨 끝 헤드가 제 홀로 무리가 되는데, 앵커는
                #   거의 언제나 그 자리다(급수원에서 가장 먼 헤드). 실제로
                #   그래서 앵커의 가지관이 «헤드 1개» 로 잡혔다.
                if deg == 1:
                    chain.append(cur)
                break                  # 다음 분기점 — 여기서 줄이 끊긴다
            if chain:
                for n in chain:
                    run_of[n] = rid
                rid += 1
    return run_of


def _fill_by_branch(head_node, anchor, xy, edges, k) -> list:
    """[규칙 «가지관»] 앵커의 가지관을 통째로 담고, 가까운 가지관으로 넘어간다.

    ★«반쪽 가지관» 을 안 만드는 것이 이 규칙의 뜻이다. 직사각형은 상자가 줄을
      가로질러 자르지만, 실무의 설계면적은 가지관 단위로 잡는 일이 많다.

    순서는 이렇다.
      ① 앵커가 달린 가지관의 헤드를 전부 담는다.
      ② 아직 안 담은 가지관 중, **이미 담은 헤드에 공간으로 가장 가까운**
         것을 골라 통째로 담는다. (배관 거리로 고르면 옆 줄을 건너뛴다 —
         그것이 §31 에서 고친 결함이다.)
      ③ K 를 넘기면 그 가지관 안에서 가까운 것부터 잘라 K 를 맞춘다.

    같은 입력에 같은 산출이라야 하므로 동점은 (거리, 번호)로 못 박는다.
    """
    run_of = _pipe_runs(edges)
    # 헤드를 가지관별로 모은다. 분기점에 바로 달린 헤드는 제 무리를 이룬다
    # (티 자리의 헤드 — 드물지만 남의 줄에 끼워 넣으면 그 줄이 늘어난다).
    groups: dict = {}
    for hi, n in head_node.items():
        key = run_of.get(n)
        groups.setdefault(("run", key) if key is not None else ("node", n),
                          []).append(hi)
    home = None
    for gk, members in groups.items():
        if anchor in members:
            home = gk
            break
    picked: list = []
    used: set = set()

    def _near(hi):
        return min((math.dist(xy(hi), xy(p)) for p in picked), default=0.0)

    def _take(gk):
        # ★동점이면 «앵커에 가까운 쪽» 을 먼저. 나란한 가지관에서는 이미 담은
        #   줄과의 거리가 위아래 똑같이 나오는데, 그때 번호순으로 집으면
        #   설계면적이 앵커 반대쪽 끝부터 자란다.
        members = sorted(groups[gk],
                         key=lambda hi: (_near(hi),
                                         math.dist(xy(hi), xy(anchor)), hi))
        for hi in members:
            if len(picked) >= k:
                return
            picked.append(hi)

    if home is not None:
        used.add(home)
        picked.append(anchor)
        for hi in sorted(groups[home], key=lambda h: math.dist(xy(h),
                                                               xy(anchor))):
            if hi != anchor and len(picked) < k:
                picked.append(hi)
    while len(picked) < k:
        best, bd = None, None
        for gk, members in groups.items():
            if gk in used:
                continue
            d = min(math.dist(xy(hi), xy(p)) for hi in members for p in picked)
            # 동점이면 앵커에 가까운 가지관부터 — 설계면적은 앵커를 중심으로
            # 자라야 한다.
            a = min(math.dist(xy(hi), xy(anchor)) for hi in members)
            if bd is None or (d, a, str(gk)) < (bd, best[1], str(best[0])):
                best, bd = (gk, a), d
        if best is None:
            break                      # 더 담을 가지관이 없다
        used.add(best[0])
        _take(best[0])
    return picked


def worst_k_heads(pts, edges, hnodes, sources, k=REMOTE_K_DEFAULT,
                   only_heads=None, source_index: int | None = None,
                   head_xy=None, rule: str = DESIGN_AREA_DEFAULT) -> dict:
    """앵커 기반 «최불리 배관망» 추출 — 수리계산의 설계면적 그 자체.

    ─ 세 단계 ────────────────────────────────────────────────────────
    ① 앵커 = 급수원에서 **배관 거리로** 가장 먼(가장 불리한) 헤드. 여기가
       기준압을 잡는 지점 — 급수원↔앵커 거리가 «최원 유하거리» 다.
    ② 설계면적 = 앵커를 품는 **직사각형** 안의 K개 — 공간으로 가까운 순서.
    ③ corridor = 그 K개를 급수원까지 잇는 최단경로의 합집합. 각 간선의
       **담당 헤드 수(load)** 를 함께 낸다 — NFPC 별표1 이 최소 호칭경을 정할
       때 쓰는 바로 그 값이라, 이 최대값이 주배관 관경을 결정한다.

    ─ ②가 «배관 거리» 에서 «직사각형» 으로 바뀐 이유 (2026-09-07) ────
    종전 ②는 앵커에서 **배관 거리**로 가까운 K개였다. 사용자 지적으로 실측해
    보니 그 자가 설계면적을 두 조각으로 갈라 놓고 있었다
    (`scripts/_probe_design_area_map.py` · B1F · K=30):

        · 앵커 줄에서 8개(0~19.9 m · 2.8 m 간격)를 집고,
          **4.3 m 를 건너뛰어** 24.2~36.2 m 떨어진 곳의 22개를 집었다.
        · 배관 거리는 공간 거리와 딴판이다 — 앵커 둘레 헤드의 배관/직선 비가
          중앙값 **6.2배**, 가장 심한 것은 직선 3.0 m 인데 배관 100.4 m(34배).
          옆 가지관은 공간으로 코앞인데 주관을 돌아가느라 «멀다» 고 읽힌다.
        · 그래서 «가장 먼 30개» 중 7개(23%)만 뽑히고, 347.9 m 짜리를 두고
          306.9 m 짜리를 뽑았다. 사용자가 본 그대로다.

    설계면적은 이름 그대로 **면적**이다(NFPC/NFPA 는 ㎡ 로 규정한다). 불은
    공간으로 번지지 관을 타고 번지지 않으므로, «어느 헤드가 함께 열리나» 는
    공간이 정한다. 그래서 앵커를 품는 직사각형(체비셰프 상자)을 K개가 담길
    때까지 넓힌다 — 사람이 화면에서 «영역 지정» 으로 그리는 그 사각형과 같은
    개념이고, 실제로 사람이 좁게 그리면 그것이 그대로 설계면적이 된다.

    ★한계는 적어 둔다: 이 상자는 **방호구역을 모른다.** 공간으로 붙어 있어도
      다른 계통에서 물을 받는 헤드가 섞일 수 있다(실측: 상자 30개 중 25개가
      앵커보다 80.5 m 앞에서 갈라진 계통). 그 판단은 사람이 «영역 지정» 으로
      한다 — 프로그램이 대신 정하지 않는다(BLOCKED §31).

    `only_heads` : 도면이 여러 장일 때 한 장으로 범위를 좁힌다. 앵커도 그
        범위 안에서 고른다(장이 다르면 앵커가 남의 도면으로 튄다).
    `head_xy` : 헤드의 **제 좌표**(board 의 disks). ②의 직사각형은 이 값으로
        잰다. 안 주면 «부착 노드» 좌표로 대신한다 — 드롭·후렉시블 길이만큼
        어긋나지만 상자 크기에 견주면 작다.
    `rule` : ②를 채우는 방식(`DESIGN_AREA_RULES`). `"rect"` 는 직사각형,
        `"branch"` 는 가지관 통째. **사람이 고른다** — 어느 쪽이 맞는지는
        현장·도면마다 다르고, 프로그램이 정할 문제가 아니다.
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
    worst_head = max(head_far, key=head_far.get)   # 가장 불리한 헤드 = 기준 헤드

    # ② 앵커를 품는 직사각형 = 설계면적. 공간으로 가까운 K개.
    #
    # 상자는 «체비셰프 거리»(max(|dx|,|dy|))로 넓힌다 — 원이 아니라 사각형이라야
    # 사람이 화면에서 그리는 «영역 지정» 과 같은 모양이 된다.
    #
    # ★같은 입력에 같은 산출이어야 한다. 상자 경계에서 여러 헤드가 같은 값이면
    #   순서가 흔들리므로 (상자 → 직선 → 번호) 로 못 박는다.
    def _xy(hi):
        if head_xy is not None and hi < len(head_xy):
            p = head_xy[hi]
            return float(p[0]), float(p[1])
        return pts[head_node[hi]]

    ax, ay = _xy(worst_head)

    def _box(hi):
        x, y = _xy(hi)
        return max(abs(x - ax), abs(y - ay))

    if rule == "branch":
        picked = _fill_by_branch(head_node, worst_head, _xy, edges, k)
    else:
        ranked = sorted(
            head_node,
            key=lambda hi: (_box(hi), math.dist(_xy(hi), (ax, ay)), hi))
        picked = ranked[:k]
    # 설계면적의 «폭» — 상자의 긴 변. 종전에는 배관 거리였다.
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
        "span_m": round(span / 1000.0, 2),              # 설계면적 직사각형의 긴 변
        # 설계면적은 이름 그대로 «면적» 이다 — 규정이 ㎡ 로 말하는 그 값을
        # 산출물이 직접 낸다. 종전에는 어디에도 없었다.
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

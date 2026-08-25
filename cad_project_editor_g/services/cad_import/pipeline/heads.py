# -*- coding: utf-8 -*-
"""시제품 헤드/기호원호 소유 모듈 — 본체 poc6 로직 복사.

실행 경로에서 `_tmp_dxf_extract_poc6` 를 부르지 않는다.
격자 도우미만 stage1 공유 import.
"""
import math
from collections import defaultdict

from services.cad_import.pipeline.stage1 import _grid_put, _grid_near


def head_clusters(w, bundle, knobs, gap=None):
    """정리서 §3: 헤드 묶음의 '작은 도형'을 근접 군집 → 헤드 실체.
    작은 도형 = 선(≤small_len)·원/원호(r≤small_r). 살수반경 큰 원은 자연 제외.
    gap(군집 간격)은 묶음별 관행 값 — 전역 한 값이면 틱 벌어진 심볼(상향식
    625mm)과 이웃 헤드 촘촘한 심볼이 동시에 못 산다(MF 실측: 전역 700 →
    하향식 이웃 62쌍 오합병). 반환: [{"c", "ents", "arcs"}]"""
    lay0, col0 = bundle
    small_len = knobs["small_len"]
    small_r = knobs["small_r"]
    gap = gap if gap is not None else knobs["cluster_gap"]
    items = []   # (ref_pt, kind, data)
    for ly, c, a, b in w.segs:
        if ly != lay0 or (col0 is not None and c != col0):
            continue
        if math.hypot(b[0] - a[0], b[1] - a[1]) <= small_len:
            m = ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2)
            items.append((m, "seg", (a, b)))
    for ly, c, cx, cy, r in w.circles:
        if ly != lay0 or (col0 is not None and c != col0):
            continue
        if r <= small_r:
            items.append(((cx, cy), "circle", r))
    arcs_here = []
    for ly, c, cx, cy, r in w.arcs:
        if ly != lay0 or (col0 is not None and c != col0):
            continue
        if r <= small_r:
            items.append(((cx, cy), "arc", r))
    # 근접 군집(유니온 파인드, 격자 가속)
    n = len(items)
    parent = list(range(n))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    grid = defaultdict(list)
    for i, (p, _k, _d) in enumerate(items):
        _grid_put(grid, gap, p[0], p[1], i)
    for i, (p, _k, _d) in enumerate(items):
        for j in _grid_near(grid, gap, p[0], p[1]):
            if j <= i:
                continue
            q = items[j][0]
            if math.hypot(q[0] - p[0], q[1] - p[1]) <= gap:
                ri, rj = find(i), find(j)
                if ri != rj:
                    parent[rj] = ri
    groups = defaultdict(list)
    for i in range(n):
        groups[find(i)].append(i)
    out = []
    for lst in groups.values():
        xs = [items[i][0][0] for i in lst]
        ys = [items[i][0][1] for i in lst]
        cl = {"c": (sum(xs) / len(xs), sum(ys) / len(ys)),
              "ents": [items[i] for i in lst],
              "arcs": [items[i][0] for i in lst if items[i][1] == "arc"],
              "bundle": (lay0, col0)}
        out.append(cl)
    return out


def head_disks(clusters, small_r=300.0):
    """기호용 헤드 도형 — 군집의 원·원호를 각자 제 중심·제 반지름으로.

    ★군집 평균 중심을 쓰지 않는다 [2026-08-02]. 원과 원호를 평균 내면 둘 다
    아닌 자리가 나온다 — MF3 EPS 실측: 원 (916317,465644) · 원호
    (916781,465644) 인데 군집 평균은 (916539,…) 로 원에서 242mm 어긋났다.
    MF3는 원=헤드 실체, 원호=티 자리 표식이라 두 도형이 서로 다른 틈을
    설명하므로 반드시 갈라 놓아야 한다.
    원·원호가 없는 군집(짧은 획 뭉치)에 가짜 반지름을 씌우지도 않는다.
    """
    out = []
    for cl in clusters:
        for p, k2, d in cl["ents"]:
            if k2 in ("circle", "arc") and 0 < d <= small_r:
                out.append((float(p[0]), float(p[1]), float(d)))
    return out


def head_cover_disks(clusters, small_r=300.0):
    """4단계 헤드걸침용 디스크 — CIRCLE + 측벽 삼각형 등가원.

    헤드가 관 위에 겹쳐 그려져 관이 끊긴 자리를 잇는 판정이라 실제 헤드
    원만 쓴다. 접속 원호는 헤드 실체가 아니고, 군집 평균·가짜 150mm는
    도면에 없는 원을 만들어 낸다(MF3 원 없는 군집 573개에 오발생).
    ★이 함수를 head_disks 와 섞지 마라 — 한 함수를 2단계와 4단계가 같이
    쓰다가 4단계만 고치려다 2단계가 무너졌다(MF3 이음 163→80 ·
    MF 덩어리 379→1638 실측).
    ★헤드 몸통 원 하나만 [2026-08-02]. split_head_circles가 고른 몸통이 있으면
      그것만 쓴다 — MF3 (926341,465644)처럼 헤드 안에 작은 원이 또 그려진
      동심 헤드에서 장식 원(r=82)까지 걸침 판정 재료로 쓰지 않기 위해서다.
    ★측벽 삼각형 [2026-08-19 오너 방안①]: circle/head_r 없는 군집에
      tri_side 가 있으면 (c, tri_side/√3) 을 넣는다. 이미 원이 있으면 기존
      경로 — 중복 원 금지.
    """
    out = []
    for cl in clusters:
        if "head_r" in cl:
            out.append((float(cl["c"][0]), float(cl["c"][1]),
                        float(cl["head_r"])))
            continue
        n0 = len(out)
        for p, k2, d in cl["ents"]:
            if k2 == "circle" and 0 < d <= small_r:
                out.append((float(p[0]), float(p[1]), float(d)))
        if len(out) > n0:
            continue
        side = cl.get("tri_side")
        if side:
            out.append((float(cl["c"][0]), float(cl["c"][1]),
                        float(side) / math.sqrt(3.0)))
    return out


def closed_tris(segs, vert_tol=1.0):
    """세 획이 꼭짓점에서 맞닿아 닫힌 삼각형 — 삼각형 헤드 판별의 손발.

    [2026-08-05 오너 지시 · D8 고침] "삼각형 헤드가 있으니 규칙에 넣고
    로직에 넣도록 하자. 유저가 찍은 것으로 찾도록 하자."(규칙서 §3-4)
    segs = [(a, b), ...]. 꼭짓점 관용 vert_tol 은 A1 겹침 정리와 같은
    1mm — apt 실측 10개 전부 틈 0.000mm 라 지어낼 숫자가 없다.
    등변·자 대조는 여기서 안 한다(기하만) — 부르는 쪽 몫.
    반환: [{"verts": [p×3], "sides": [작은→큰], "segs": [(a,b)×3]}]
    """
    n = len(segs)
    if n < 3:
        return []
    adj = defaultdict(set)
    for i in range(n):
        ai, bi = segs[i]
        for j in range(i + 1, n):
            aj, bj = segs[j]
            if min(math.hypot(ai[0] - aj[0], ai[1] - aj[1]),
                   math.hypot(ai[0] - bj[0], ai[1] - bj[1]),
                   math.hypot(bi[0] - aj[0], bi[1] - aj[1]),
                   math.hypot(bi[0] - bj[0], bi[1] - bj[1])) <= vert_tol:
                adj[i].add(j)
                adj[j].add(i)
    out = []
    for i in sorted(adj):
        for j in sorted(adj[i]):
            if j <= i:
                continue
            for k in sorted(adj[j]):
                if k <= j or k not in adj[i]:
                    continue
                tri = [segs[i], segs[j], segs[k]]
                pts = [t[0] for t in tri] + [t[1] for t in tri]
                used = [False] * 6
                verts, ok = [], True
                for a2 in range(6):
                    if used[a2]:
                        continue
                    best = None
                    for b2 in range(a2 + 1, 6):
                        if used[b2]:
                            continue
                        d = math.hypot(pts[a2][0] - pts[b2][0],
                                       pts[a2][1] - pts[b2][1])
                        if best is None or d < best[0]:
                            best = (d, b2)
                    if best is None or best[0] > vert_tol:
                        ok = False        # 여섯 끝점이 세 쌍으로 안 닫힌다
                        break
                    used[a2] = used[best[1]] = True
                    verts.append(pts[a2])
                if not ok or len(verts) != 3:
                    continue
                out.append({"verts": verts,
                            "sides": sorted(
                                math.hypot(b[0] - a[0], b[1] - a[1])
                                for a, b in tri),
                            "segs": tri})
    return out


def tri_head_of(cl, knobs=None):
    """군집이 「찍은 변 길이」 자를 통과한 닫힌 등변 삼각형이면 헤드로.

    규칙서 §3-4 [2026-08-05 오너 지시 · D8 고침]. 발동 조건 = 그 묶음에
    삼각형 픽(`tri_rulers`)이 있을 때뿐 — 자동 탐지는 없다.
    판별: ① 세 획이 꼭짓점 ≤1mm(A1 관용)로 닫힘 ② 세 변이 서로
    ±max(5mm,10%) 안(등변) ③ 세 변 전부 찍은 변 ±max(5mm,10%) 안.
    자리 = 세 꼭짓점의 무게중심(원의 「중심」에 해당하는 유일한 자리).
    실측(apt): 진짜 10개 변 169.9~170.5 통과 · 범례 2개 변 255 자연 제외.
    반환: 헤드 dict 또는 None. 도구(_tmp_stage0pick)도 이 판별을 그대로
    쓴다 — 판정 한 벌(문지기 검사 ②의 취지).
    """
    kn = dict(head_size_eps=5.0, head_size_rel=0.10, a1_lat=1.0)
    if knobs:
        kn.update(knobs)
    eps = float(kn["head_size_eps"])
    rel = float(kn["head_size_rel"])
    segs = [d for _p, k2, d in cl["ents"] if k2 == "seg"]
    for tri in closed_tris(segs, vert_tol=float(kn["a1_lat"])):
        s_lo, _s_mid, s_hi = tri["sides"]
        if s_hi - s_lo > max(eps, s_hi * rel):
            continue                       # 부등변 — 헤드 아님
        for side, lb in (cl.get("tri_rulers") or ()):
            slack = max(eps, side * rel)
            if all(abs(s - side) <= slack for s in tri["sides"]):
                vx = tri["verts"]
                return dict(cl,
                            c=(sum(p[0] for p in vx) / 3.0,
                               sum(p[1] for p in vx) / 3.0),
                            tri_side=float(side), tri_segs=tri["segs"],
                            label=lb if lb is not None
                            else cl.get("label"))
    return None


def collect_head_clusters(w, spec, knobs):
    """헤드 픽대로 군집을 모으고 각 군집에 「자 목록」을 붙인다 — 그릇 v2.

    픽 = (레이어×색×반지름) [2026-08-04 오너 확정 · D7]. 같은 묶음을 크기만
    달리해 여러 번 찍으면(apt 노랑 42·52.6·60) 군집은 **한 번만** 모으고
    자만 여러 개 붙인다 — 픽마다 다시 모으면 같은 군집이 통째로 중복된다.
    본체(process)·0단계(head_candidates)·대장 프로브가 전부 이 함수를 쓴다
    (판정 한 벌 — 문지기 검사 ②의 취지).
    ★삼각형 픽 [2026-08-05 · D8]: `tri_side` 픽은 원 자가 아니라 삼각형
    자(`tri_rulers`)로 붙는다 — 판별은 tri_head_of.
    """
    by = {}
    for hs in spec.get("heads") or []:
        key = (tuple(hs["bundle"]), hs.get("cluster_gap"))
        by.setdefault(key, []).append(hs)
    clusters = []
    for (bundle, gap), hss in by.items():
        cls = head_clusters(w, bundle, knobs, gap=gap)
        rulers = [(float(hs["r"]), hs.get("label"))
                  for hs in hss if "r" in hs]
        tri_rulers = [(float(hs["tri_side"]), hs.get("label"))
                      for hs in hss if "tri_side" in hs]
        for cl in cls:
            cl["rulers"] = rulers
            if tri_rulers:
                cl["tri_rulers"] = tri_rulers
            cl["label"] = hss[0].get("label")
        clusters += cls
    return clusters


def split_head_circles(clusters, knobs=None):
    """헤드 실체 = 「찍은 반지름」 자를 통과한 동그라미 **각각** [그릇 v2].

    오너: "헤드는 99% 동그라미다. 안에 X자나 작은 원이 또 있어도 겉모양은
    동그라미다. 반원 달린 것은 전부 분기점 기호이니 그대로 살려라."
    (외국 도면의 삼각형 헤드는 규칙으로만 인정 — 표본 도면이 하나도 없어
     3선 닫힘 판별은 실물이 올 때 붙인다. D8)

    ★자 = 유저가 찍은 반지름 [2026-08-04 오너 확정 · D7 고침 2026-08-05].
      구 방식(그 묶음에서 가장 많은 크기로 짐작)은 폐기 — 유저가 이미 준
      답을 두고 추측했고, 크기 섞인 도면에서 소수 크기를 잘랐다(apt 19개 ·
      3F 탕비실 r150 3개 — r125 6개는 헤드가 아니라 접속부 기호였다).
      작도 오차 slack = ±max(5mm, 10%) 유지. 자가 여러 개면(같은 묶음 크기
      다른 픽) 원마다 제일 가까운 자 하나에 댄다.
    ★자 통과 원이 여럿이면 **각각을 헤드로 쪼갠다** [D12 고침 2026-08-05].
      구 방식(가장 큰 하나만)은 7도면에서 46개를 조용히 버렸다(MF2 14 ·
      apt 24 · MF 3 · 3F 3 · BF4 1 · MF4 1). 쪼갤 때 나머지 증거(획·원호·
      탈락 원)는 **가장 가까운 통과 원 하나에만** 붙인다 — 전부 복사하면
      이음 판정(head_disks·rim_dist·헤드원호)이 같은 증거를 두 번 센다.
      통과 원이 하나면 군집을 그대로 쓴다(증거 배치가 구판과 동일).
    ★자리 = 그 원의 중심 — 불변(군집 평균 폐기 이력 §3-2 유지).
    ★동그라미 없는 군집 = 분기 표시(marks) — 불변. 2단계 기호 증거로 쓴다.
      예외 하나 [2026-08-05 오너 지시 · D8 고침 · §3-4]: **삼각형 픽**
      (`tri_side`)이 있는 묶음에서 그 군집이 닫힌 등변 삼각형(찍은 변 자
      ±max(5mm,10%) 통과)이면 헤드다 — tri_head_of 가 판별하고, 자리는
      세 꼭짓점의 무게중심. 자동 탐지 없음(찍은 묶음에서만).
    반환: (heads, marks, info) — 계약 동일. info["r0"]/["slack"] = 묶음의
      첫 자(0단계 씨앗·표시용) · info["rulers"] = 묶음별 자 전체.
      dropped 의 "묶음버림"은 v2에서 소멸(항상 0건) — 회귀 칸은 유지된다.
    """
    kn = dict(head_size_eps=5.0, head_size_rel=0.10)
    if knobs:
        kn.update(knobs)
    eps = float(kn["head_size_eps"])
    rel = float(kn["head_size_rel"])
    r0s, slacks, rulers_by = {}, {}, {}
    for cl in clusters:
        b = cl.get("bundle")
        rs = cl.get("rulers") or ()
        if rs and b not in r0s:
            r0s[b] = rs[0][0]
            slacks[b] = max(eps, rs[0][0] * rel)
            rulers_by[b] = list(rs)
    heads, marks, n_size_drop = [], [], 0
    n_tri = 0                    # 삼각형 헤드 [2026-08-05 D8]
    # ★손실 명세 [2026-08-05 D14] — 버릴 때 이유·자리를 남긴다.
    n_circ_total = 0
    dropped = []                 # [{"at","r","why"}] — 사진·부검용
    for cl in clusters:
        rulers = cl.get("rulers") or ()
        passing, failed, extras = [], [], []
        for e in cl["ents"]:
            p, k2, d = e
            if k2 != "circle" or d <= 0:
                extras.append(e)
                continue
            n_circ_total += 1
            best = None
            for (r, lb) in rulers:
                gap = abs(d - r)
                if gap <= max(eps, r * rel) and \
                        (best is None or gap < best[0]):
                    best = (gap, lb)
            if best is None:
                failed.append(e)
                dropped.append({"at": (float(p[0]), float(p[1])),
                                "r": float(d), "why": "크기자"})
            else:
                passing.append((e, best[1]))
        if not passing:
            # ★삼각형 헤드 [2026-08-05 오너 지시 · D8 고침 · §3-4] —
            #   삼각형 픽이 있는 묶음에서, 원 없는 군집이 닫힌 등변
            #   삼각형(찍은 변 자 통과)이면 헤드다. 아니면 종전대로
            #   분기 표시(marks).
            if cl.get("tri_rulers"):
                tri = tri_head_of(cl, kn)
                if tri is not None:
                    heads.append(tri)
                    n_tri += 1
                    continue
            if failed:
                n_size_drop += 1
            marks.append(cl)
            continue
        if len(passing) == 1:
            (p, _k2, d), lb = passing[0]
            heads.append(dict(cl, c=(float(p[0]), float(p[1])),
                              head_r=float(d),
                              label=lb if lb is not None
                              else cl.get("label")))
            continue
        centers = [e[0] for (e, _lb) in passing]
        buckets = [[] for _ in passing]
        for e in extras + failed:
            x, y = e[0]
            i = min(range(len(passing)),
                    key=lambda k3: (centers[k3][0] - x) ** 2 +
                                   (centers[k3][1] - y) ** 2)
            buckets[i].append(e)
        for i, ((p, _k2, d), lb) in enumerate(passing):
            own = [(p, "circle", d)] + buckets[i]
            heads.append(dict(cl, c=(float(p[0]), float(p[1])),
                              head_r=float(d),
                              label=lb if lb is not None
                              else cl.get("label"),
                              ents=own,
                              arcs=[q for (q, k4, _d4) in own
                                    if k4 == "arc"]))
    return heads, marks, {"r0": r0s, "slack": slacks, "rulers": rulers_by,
                          "n_size_drop": n_size_drop, "n_tri": n_tri,
                          "n_circ": n_circ_total, "dropped": dropped}


def iter_material_symbol_arcs(w, mat_set, head_lays=None, small_r=300.0):
    """배관기호 원호 [2026-08-02 오너 확정].

    재료로 지정한 **레이어**의 ARC(원호)만. 색이 재료 묶음과 달라도 포함한다.
    CIRCLE은 반환하지 않는다(헤드 원과 혼동·3F CIRCLE 이음 0 실측).
    반환: (ly, c, cx, cy, r) 이터레이터.
    """
    mat_layers = {ly for ly, _c in mat_set}
    head_lays = set(head_lays or ())
    for ly, c, cx, cy, r in w.arcs:
        if ly in head_lays or ly not in mat_layers:
            continue
        if 0 < r <= small_r:
            yield ly, c, cx, cy, r


def iter_outside_symbol_arcs(w, mat_set, head_lays=None, small_r=300.0):
    """재료 레이어 **밖**의 작은 원호 — 2단계 잇기 근거로만 쓴다.

    분리안 [2026-08-03 오너 지시]
    ---------------------------
    MF4 부검: 이 도면은 접속 기호를 배관(FIRE1)과 **다른 레이어(FIRE)** 에
    그렸다. 그래서 `iter_material_symbol_arcs` 가 405개 중 0개를 넘겨
    2단계 이음이 25곳으로 무너졌다(정상 377곳).

    그렇다고 레이어 잣대를 통째로 넓히면 안 된다 — **3단계 기호제외 게이트**
    까지 넓어져 MF3 3단계가 68→66으로 2곳 죽는다(오너 실측). 그래서 갈라 쓴다:

      · 2단계 잇기 근거  = 재료 레이어 원호 + **이 함수(재료 밖 원호)**
      · 3단계 기호제외 게이트 = 재료 레이어 원호만 (`iter_material_symbol_arcs`,
        지금 그대로 — 손대지 않는다)

    넓혀도 안전한 이유: 2단계는 "원호 **양옆**에 자유단이 있어야" 잇는다.
    건축 문짝 호·치수선 호는 배관 틈에 앉아 있지 않아 아무것도 못 잇는다
    (MF3 +2,674개·BF4 +2,792개 들이부어도 이음 +0 실측, 2026-08-03).
    반환: (ly, c, cx, cy, r) 이터레이터.
    """
    mat_layers = {ly for ly, _c in mat_set}
    head_lays = set(head_lays or ())
    for ly, c, cx, cy, r in w.arcs:
        if ly in head_lays or ly in mat_layers:
            continue
        if 0 < r <= small_r:
            yield ly, c, cx, cy, r


# ------------------------------------------- 헤드 종류 판정용 읽기 헬퍼
#   [2026-08-08 오너 확정] 두 칸 찍기 + 자동 분류. 군집/split 은 손대지 않는다.

_ARM_CTR = 5.0       # flow_water.ARM_CTR 과 동일 — 순환 import 피함
_HEAD_TOUCH = 50.0   # flow_water.HEAD_TOUCH 과 동일
_ARM_CELL = 500.0    # 후보화 전용 격자 — 최종 거리 판정식에는 쓰지 않는다


def _seg_dist(a, b, px, py):
    """점→선분 최단거리와 선분 파라미터 t∈[0,1]."""
    ax, ay = a[0], a[1]
    bx, by = b[0], b[1]
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    if L2 < 1e-18:
        return math.hypot(ax - px, ay - py), 0.0
    t = ((px - ax) * dx + (py - ay) * dy) / L2
    t = 0.0 if t < 0.0 else (1.0 if t > 1.0 else t)
    return math.hypot(ax + t * dx - px, ay + t * dy - py), t


def _is_head_circle(qx, qy, qr, cx, cy, r):
    """이 원이 헤드 원 자체인가 (같은 자리·같은 크기)."""
    return (math.hypot(qx - cx, qy - cy) <= 1.0
            and abs(qr - r) <= max(5.0, r * 0.10))


def short_marks_in_disk(w, cx, cy, r, small_len):
    """원 안 짧은 문양 — 긴 관통배관·헤드원 제외.

    반환: [{"bundle": (ly, c), "kind": "seg"|"arc"|"circle"}, ...]
    안 찍은 레이어를 분류에 쓰지 말 것 — 찍기 때 후보 안내·자동 동일묶음
    판별용. 분류는 찍힌 mark_bundle 만 mark_bundle_in_disk 로 본다.
    """
    out = []
    r2 = float(r) * float(r)
    for ly, c, a, b in w.segs:
        ln = math.hypot(b[0] - a[0], b[1] - a[1])
        if not (0 < ln <= small_len):
            continue
        mx = (a[0] + b[0]) * 0.5
        my = (a[1] + b[1]) * 0.5
        if (mx - cx) * (mx - cx) + (my - cy) * (my - cy) <= r2:
            out.append({"bundle": (ly, c), "kind": "seg"})
    for ly, c, ax, ay, ar in w.arcs:
        if not (0 < ar <= r):
            continue
        if _is_head_circle(ax, ay, ar, cx, cy, r):
            continue
        if (ax - cx) * (ax - cx) + (ay - cy) * (ay - cy) <= r2:
            out.append({"bundle": (ly, c), "kind": "arc"})
    for ly, c, qx, qy, qr in w.circles:
        if qr <= 0:
            continue
        if _is_head_circle(qx, qy, qr, cx, cy, r):
            continue
        if math.hypot(qx - cx, qy - cy) + qr <= r + 1.0:
            out.append({"bundle": (ly, c), "kind": "circle"})
    return out


def mark_bundles_in_disk(w, cx, cy, r, small_len):
    """원 안 문양 묶음(레이어×색) 목록 — 등장 순·중복 제거."""
    seen, out = set(), []
    for m in short_marks_in_disk(w, cx, cy, r, small_len):
        b = m["bundle"]
        if b in seen:
            continue
        seen.add(b)
        out.append(b)
    return out


def mark_bundle_in_disk(w, mark_bundle, cx, cy, r, small_len):
    """찍힌 mark_bundle 도형이 원 안에 있으면 True.

    안 찍은 레이어는 보지 않는다 — mark_bundle 묶음만 스캔.
    """
    mb = tuple(mark_bundle)
    mly, mc = mb
    r2 = float(r) * float(r)
    for ly, c, a, b in w.segs:
        if ly != mly or c != mc:
            continue
        ln = math.hypot(b[0] - a[0], b[1] - a[1])
        if not (0 < ln <= small_len):
            continue
        mx = (a[0] + b[0]) * 0.5
        my = (a[1] + b[1]) * 0.5
        if (mx - cx) * (mx - cx) + (my - cy) * (my - cy) <= r2:
            return True
    for ly, c, ax, ay, ar in w.arcs:
        if ly != mly or c != mc or not (0 < ar <= r):
            continue
        if _is_head_circle(ax, ay, ar, cx, cy, r):
            continue
        if (ax - cx) * (ax - cx) + (ay - cy) * (ay - cy) <= r2:
            return True
    for ly, c, qx, qy, qr in w.circles:
        if ly != mly or c != mc or qr <= 0:
            continue
        if _is_head_circle(qx, qy, qr, cx, cy, r):
            continue
        if math.hypot(qx - cx, qy - cy) + qr <= r + 1.0:
            return True
    return False


def build_head_arm_index(w, mat_set):
    """재료 선 필터·길이와 양 끝점 격자를 한 번 만든다."""
    cell = _ARM_CELL
    mat_set = frozenset(tuple(b) for b in mat_set)
    rows = []
    grid = defaultdict(list)
    for order, (ly, c, a, b) in enumerate(w.segs):
        if (ly, c) not in mat_set:
            continue
        ln = math.hypot(b[0] - a[0], b[1] - a[1])
        if ln <= 0:
            continue
        rid = len(rows)
        rows.append((order, a, b, ln))
        cells = {(int(a[0] // cell), int(a[1] // cell)),
                 (int(b[0] // cell), int(b[1] // cell))}
        for key in cells:
            grid[key].append(rid)
    return {"cell": cell, "rows": rows, "grid": grid}


def _head_arm_candidates(index, hx, hy, radius):
    """끝점이 판정 반경에 들 수 있는 셀의 선을 원래 순서로 반환."""
    cell = index["cell"]
    ids = set()
    for gx in range(int((hx - radius) // cell),
                    int((hx + radius) // cell) + 1):
        for gy in range(int((hy - radius) // cell),
                        int((hy + radius) // cell) + 1):
            ids.update(index["grid"].get((gx, gy), ()))
    return [index["rows"][i] for i in sorted(ids)]


def head_has_arm(w, mat_set, hx, hy, hr, small_len,
                 ctr=None, touch=None, index=None):
    """짧은 재료 선이 헤드에서 바깥으로 뻗으면 하향식 팔.

    상향식 끊김은 한쪽만 짧게 잘리고 맞은편은 긴 관인 경우가 많다 —
    맞은편은 길이 제한 없이 본다. 마주 보면 팔이 아니다 [2026-08-08].
    """
    ctr = _ARM_CTR if ctr is None else float(ctr)
    touch = _HEAD_TOUCH if touch is None else float(touch)
    mat_set = set(mat_set)
    short_dirs = []
    all_dirs = []
    if index is None:
        rows = ((order, a, b,
                 math.hypot(b[0] - a[0], b[1] - a[1]))
                for order, (ly, c, a, b) in enumerate(w.segs)
                if (ly, c) in mat_set)
    else:
        rows = _head_arm_candidates(
            index, hx, hy, max(ctr, float(hr) + touch))
    for _order, a, b, ln in rows:
        if ln <= 0:
            continue
        for (px, py), (ox, oy) in ((a, b), (b, a)):
            d = math.hypot(px - hx, py - hy)
            if not (d <= ctr or abs(d - hr) <= touch):
                continue
            od = math.hypot(ox - hx, oy - hy)
            if od <= hr + touch:
                continue
            u = ((ox - hx) / od, (oy - hy) / od)
            all_dirs.append(u)
            if ln <= small_len:
                short_dirs.append(u)
    if not short_dirs:
        return False
    for ux, uy in short_dirs:
        for vx, vy in all_dirs:
            if ux * vx + uy * vy < -0.5:
                return False               # 관통 끊김
    return True


def pipe_under_disk(w, mat_set, hx, hy, hr):
    """재료 관이 원 밑(안)을 지나가면 True — 상향식 관위."""
    mat_set = set(mat_set)
    for ly, c, a, b in w.segs:
        if (ly, c) not in mat_set:
            continue
        d, _t = _seg_dist(a, b, hx, hy)
        if d <= hr:
            return True
    return False


# ------------------------------------------- 문양 기하 지문 [2026-08-08]
#   원 클릭 한 번 → 원 안 문양을 짧은 태그로 기억 → 같은 지문 원만 같이 선택.
#   바깥으로 나가는 팔/배관은 중점이 원 밖이면 자연 제외. 중심 통과 짧은
#   표식(끝이 살짝 나가도)은 중점·중심거리로 남긴다(MF4 FIRE1 빗금 실측).

def _qr_bucket(qr):
    """반지름 5mm 버킷."""
    return int(round(float(qr) / 5.0)) * 5


def _n_bucket(n):
    """개수 버킷: 1 / 2 / 3+."""
    n = int(n)
    if n <= 1:
        return 1
    if n == 2:
        return 2
    return 3


def _seg_orient(ang_deg):
    """선 방향 → h|v|d (15° 버킷). ang_deg 는 [0, 180)."""
    a = float(ang_deg) % 180.0
    if min(a, 180.0 - a) <= 15.0:
        return "h"
    if abs(a - 90.0) <= 15.0:
        return "v"
    return "d"


def _angle_diff(a, b):
    """두 선분 방향각 차이 (0~90)."""
    d = abs(float(a) - float(b)) % 180.0
    if d > 90.0:
        d = 180.0 - d
    return d


def mark_fp_key(fp):
    """지문 → 비교·_hkey 용 정규화 튜플."""
    if fp is None:
        return ("empty",)
    if fp == "empty" or fp == ["empty"] or fp == ("empty",):
        return ("empty",)
    tags = []
    for t in fp:
        if t == "empty":
            return ("empty",)
        if isinstance(t, (list, tuple)):
            tags.append(tuple(t))
        else:
            tags.append((t,))
    if not tags:
        return ("empty",)
    return tuple(sorted(tags, key=repr))


# 지문용 공간 격자 — 판정 규칙이 아니라 근처만 보기 위한 성능 장치.
_FP_CELL = 500.0


def _fp_grid_near(grid, cell, cx, cy, rad):
    """원 (cx,cy,rad) 과 겹치는 칸의 후보를 모두 낸다."""
    x0, x1 = cx - rad, cx + rad
    y0, y1 = cy - rad, cy + rad
    for gx in range(int(x0 // cell), int(x1 // cell) + 1):
        for gy in range(int(y0 // cell), int(y1 // cell) + 1):
            yield from grid.get((gx, gy), ())


def _fp_build_spatial(w, small_len):
    """짧은 선·호·원을 중점/중심 격자에 넣는다 — 표 만들 때 1회."""
    cell = _FP_CELL
    small_len = float(small_len)
    sgrid = defaultdict(list)
    for ly, c, a, b in w.segs:
        ln = math.hypot(b[0] - a[0], b[1] - a[1])
        if not (0 < ln <= small_len):
            continue
        mx = (a[0] + b[0]) * 0.5
        my = (a[1] + b[1]) * 0.5
        ang = math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])) % 180.0
        sgrid[(int(mx // cell), int(my // cell))].append(
            (ly, c, a, b, ln, mx, my, ang))
    agrid = defaultdict(list)
    for ly, c, ax, ay, ar in w.arcs:
        if ar <= 0:
            continue
        agrid[(int(ax // cell), int(ay // cell))].append(
            (ly, c, float(ax), float(ay), float(ar)))
    cgrid = defaultdict(list)
    for ly, c, qx, qy, qr in w.circles:
        if float(qr) <= 0:
            continue
        cgrid[(int(qx // cell), int(qy // cell))].append(
            (ly, c, float(qx), float(qy), float(qr)))
    return {"cell": cell, "sgrid": sgrid, "agrid": agrid, "cgrid": cgrid}


def _fp_collect(cx, cy, r, small_len, w=None, spatial=None):
    """원 안 짧은 문양 후보 — 전수(w) 또는 격자(spatial). 결과는 같아야 한다."""
    cx, cy, r = float(cx), float(cy), float(r)
    lim = r + 1.0
    lim2 = lim * lim
    segs = []
    if spatial is not None:
        cell = spatial["cell"]
        for ly, c, a, b, ln, mx, my, ang in _fp_grid_near(
                spatial["sgrid"], cell, cx, cy, lim):
            if (mx - cx) * (mx - cx) + (my - cy) * (my - cy) > lim2:
                continue
            da = math.hypot(a[0] - cx, a[1] - cy)
            db = math.hypot(b[0] - cx, b[1] - cy)
            pd, _t = _seg_dist(a, b, cx, cy)
            if da > lim and db > lim and pd > 0.25 * r:
                continue
            segs.append({"ly": ly, "c": c, "ln": ln, "pd": pd, "ang": ang,
                         "both": da <= lim and db <= lim})
        arcs = []
        for ly, c, ax, ay, ar in _fp_grid_near(
                spatial["agrid"], cell, cx, cy, lim):
            if not (0 < ar <= r):
                continue
            if _is_head_circle(ax, ay, ar, cx, cy, r):
                continue
            if (ax - cx) * (ax - cx) + (ay - cy) * (ay - cy) <= lim2:
                arcs.append({"ly": ly, "c": c, "kind": "arc"})
        circs = []
        for ly, c, qx, qy, qr in _fp_grid_near(
                spatial["cgrid"], cell, cx, cy, lim):
            if _is_head_circle(qx, qy, qr, cx, cy, r):
                continue
            if math.hypot(qx - cx, qy - cy) + qr <= lim:
                circs.append({"ly": ly, "c": c, "qx": qx, "qy": qy, "qr": qr,
                              "kind": "circle"})
        return segs, arcs, circs

    small_len = float(small_len)
    # 짧은 선 — 중점이 원 안(r+1). 양 끝이 원 밖이어도 중심을 가르는
    # 짧은 표식(MF4 FIRE1 빗금: 끝 187·r150)은 중점이 원에 들어 살아남고,
    # 바깥으로 뻗는 팔은 중점이 원 밖이라 빠진다.
    for ly, c, a, b in w.segs:
        ln = math.hypot(b[0] - a[0], b[1] - a[1])
        if not (0 < ln <= small_len):
            continue
        mx = (a[0] + b[0]) * 0.5
        my = (a[1] + b[1]) * 0.5
        if (mx - cx) * (mx - cx) + (my - cy) * (my - cy) > lim2:
            continue
        da = math.hypot(a[0] - cx, a[1] - cy)
        db = math.hypot(b[0] - cx, b[1] - cy)
        pd, _t = _seg_dist(a, b, cx, cy)
        if da > lim and db > lim and pd > 0.25 * r:
            continue
        ang = math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])) % 180.0
        segs.append({"ly": ly, "c": c, "ln": ln, "pd": pd, "ang": ang,
                     "both": da <= lim and db <= lim})
    arcs = []
    for ly, c, ax, ay, ar in w.arcs:
        if not (0 < ar <= r):
            continue
        if _is_head_circle(ax, ay, ar, cx, cy, r):
            continue
        if (ax - cx) * (ax - cx) + (ay - cy) * (ay - cy) <= lim2:
            arcs.append({"ly": ly, "c": c, "kind": "arc"})
    circs = []
    for ly, c, qx, qy, qr in w.circles:
        if qr <= 0:
            continue
        if _is_head_circle(qx, qy, qr, cx, cy, r):
            continue
        if math.hypot(qx - cx, qy - cy) + qr <= lim:
            circs.append({"ly": ly, "c": c, "qx": qx, "qy": qy, "qr": qr,
                          "kind": "circle"})
    return segs, arcs, circs


def _fp_tags(cx, cy, r, segs, arcs, circs):
    """모아 둔 문양 후보 → 지문 태그 리스트."""
    tags = []
    used_seg = set()
    used_circ = set()

    # 1) dot — 중심 근처 작은 원
    for i, ci in enumerate(circs):
        if math.hypot(ci["qx"] - cx, ci["qy"] - cy) > 0.35 * r:
            continue
        if ci["qr"] > 0.5 * r:
            continue
        tags.append(["dot", ci["ly"], ci["c"], _qr_bucket(ci["qr"])])
        used_circ.add(i)

    # 중심 통과 짧은 선 후보
    ctr = [(i, s) for i, s in enumerate(segs) if s["pd"] <= 0.25 * r]
    # 2) x — 같은 묶음 선 2개 교차각 55~90°
    x_done = False
    by_b = defaultdict(list)
    for i, s in ctr:
        by_b[(s["ly"], s["c"])].append((i, s))
    # 묶음 순서를 고정 — 후보가 여럿이면 같은 지문이 나오게(격자/전수 동일).
    for (ly, c), lst in sorted(by_b.items(), key=repr):
        if x_done or len(lst) < 2:
            continue
        found = None
        for ai in range(len(lst)):
            for bi in range(ai + 1, len(lst)):
                if 55.0 <= _angle_diff(lst[ai][1]["ang"],
                                       lst[bi][1]["ang"]) <= 90.0:
                    found = (lst[ai][0], lst[bi][0])
                    break
            if found:
                break
        if found:
            tags.append(["x", ly, c])
            used_seg.add(found[0])
            used_seg.add(found[1])
            x_done = True

    # 3) bar — 중심 통과 짧은 선 1개(묶음별, x 에 안 쓰인 것)
    if not x_done:
        for (ly, c), lst in sorted(by_b.items(), key=repr):
            left = [(i, s) for i, s in lst if i not in used_seg]
            if len(left) != 1:
                continue
            i, s = left[0]
            tags.append(["bar", ly, c, _seg_orient(s["ang"])])
            used_seg.add(i)

    # 4) mk — 나머지 원안 짧은 선·호·원
    bag = defaultdict(int)
    for i, s in enumerate(segs):
        if i in used_seg:
            continue
        bag[(s["ly"], s["c"], "seg")] += 1
    for a in arcs:
        bag[(a["ly"], a["c"], "arc")] += 1
    for i, ci in enumerate(circs):
        if i in used_circ:
            continue
        bag[(ci["ly"], ci["c"], "circle")] += 1
    for (ly, c, kind), n in sorted(bag.items(), key=repr):
        tags.append(["mk", ly, c, kind, _n_bucket(n)])

    if not tags:
        return ["empty"]
    tags.sort(key=repr)
    return tags


def mark_fingerprint(w, cx, cy, r, small_len, spatial=None):
    """헤드 원 안 문양 지문 — JSON 가능 리스트.

    태그(이름 있는 쪽 우선, 쓴 도형은 mk 에서 제외):
      · ["dot", ly, c, qr_bucket]  중심 근처 작은 원
      · ["x", ly, c]               중심 통과 짧은 선 2개 교차 55~90°
      · ["bar", ly, c, "h"|"v"|"d"] 중심 통과 짧은 선 1개
      · ["mk", ly, c, kind, n_bucket] 나머지 원안 짧은 선·호·원
      · ["empty"]                  아무 것도 없음

    ※ 찍기 확장·빈헤드 판정은 `pick_mark_fp` / `circles_with_mark_pick`
      을 쓴다(재료 확정 후 문양 묶음 자동 채택). 이 함수는 전 레이어
      그대로의 옛 지문·분류 폴백용.
    """
    cx, cy, r = float(cx), float(cy), float(r)
    segs, arcs, circs = _fp_collect(
        cx, cy, r, small_len, w=w, spatial=spatial)
    return _fp_tags(cx, cy, r, segs, arcs, circs)


# ------------------------------------------- 재료확정·문양자동채택 [2026-08-09]
# 확정 규칙 (구현 주석 = 정본, 기록 문서만 보지 말 것):
#   1) 배관(재료) 완료 후에만 헤드 찍기. 배관을 **실제로 다시 찍으면**
#      헤드 해제(`1` 키만으로는 해제하지 않음).
#   2) 유저는 잡선 없는 깨끗한 헤드 원을 찍는다(UI 안내).
#   3) 클릭 원 안 · 재료 묶음이 아닌 (레이어×색) = 문양 묶음으로 자동 채택.
#   4) 같은 문양이 후보에 들어 있으면 잡선이 더 있어도 같이 선택(포함 매칭).
#   5) 문양 묶음이 없으면 빈 헤드. 빈 확장에서는 비재료를 전부 무시
#      (건축 원·호·선 포함). 문양 있는 헤드는 문양 픽으로만 고른다.
#   6) 빈 헤드에 «아무 문양이나 포함»은 금지(문양 픽과 칸·묶음으로 분리).


def _fp_filter_bundles(segs, arcs, circs, bundles):
    """원안 후보를 지정 묶음(레이어×색)만 남긴다."""
    keep = {tuple(b) for b in bundles}
    segs2 = [s for s in segs if (s["ly"], s["c"]) in keep]
    arcs2 = [a for a in arcs if (a["ly"], a["c"]) in keep]
    circs2 = [c for c in circs if (c["ly"], c["c"]) in keep]
    return segs2, arcs2, circs2


def _fp_non_mat_bundles(segs, arcs, circs, mat_set):
    """재료가 아닌 묶음 — 클릭 기준 문양 자동 채택 대상."""
    mat = {tuple(b) for b in (mat_set or ())}
    out = set()
    for s in segs:
        b = (s["ly"], s["c"])
        if b not in mat:
            out.add(b)
    for a in arcs:
        b = (a["ly"], a["c"])
        if b not in mat:
            out.add(b)
    for c in circs:
        b = (c["ly"], c["c"])
        if b not in mat:
            out.add(b)
    return out


def mark_fp_on_bundles(w, cx, cy, r, small_len, bundles, spatial=None):
    """지정 묶음 도형만으로 지문. bundles 비면 empty."""
    cx, cy, r = float(cx), float(cy), float(r)
    if not bundles:
        return ["empty"]
    segs, arcs, circs = _fp_collect(
        cx, cy, r, small_len, w=w, spatial=spatial)
    segs, arcs, circs = _fp_filter_bundles(segs, arcs, circs, bundles)
    return _fp_tags(cx, cy, r, segs, arcs, circs)


def disk_is_empty_head(w, cx, cy, r, small_len, mat_set, spatial=None):
    """빈 헤드 후보?

    빈 픽(문양묶음=[]) 에서는 **비재료를 보지 않는다** [2026-08-09].
    건축 원·호·선이 끼어도 빈으로 본다. 상하향 등 문양 있는 헤드는
    문양 픽(다른 칸·다른 묶음)으로 고른다 — 빈 확장에 섞이지 않음.
    """
    # 인자·spatial 은 호출부와 서명 맞추기용. 비재료를 안 보므로 검사 없음.
    return True


def pick_mark_fp(w, cx, cy, r, small_len, mat_set, spatial=None):
    """클릭 원 → (지문, 문양묶음목록).

    깨끗한 원 기준으로 재료 밖 묶음을 문양으로 채택한다.
    문양 없으면 empty · 묶음 [].
    """
    cx, cy, r = float(cx), float(cy), float(r)
    segs, arcs, circs = _fp_collect(
        cx, cy, r, small_len, w=w, spatial=spatial)
    marks = sorted(_fp_non_mat_bundles(segs, arcs, circs, mat_set),
                   key=repr)
    if not marks:
        return ["empty"], []
    segs, arcs, circs = _fp_filter_bundles(segs, arcs, circs, marks)
    return _fp_tags(cx, cy, r, segs, arcs, circs), marks


def mark_fp_contains(want_fp, got_fp):
    """포함 매칭 — 찍은 문양 태그가 후보에 모두 있으면 참.

    empty 끼리만 빈 매칭(아무 문양이나 열지 않음).
    """
    want = mark_fp_key(want_fp)
    got = mark_fp_key(got_fp)
    if want == ("empty",):
        return got == ("empty",)
    if got == ("empty",):
        return False
    return set(want).issubset(set(got))


def circles_with_mark_pick(w, bundle, r, want_fp, mark_bundles, mat_set,
                           small_len, spatial=None, index=None):
    """재료확정 후 헤드 확장 — [(ly,c,cx,cy,r), ...].

    · 문양 묶음 있음: 그 묶음만 본 지문으로 포함 매칭(다른 잡선 레이어 무시)
    · 문양 없음(빈): disk_is_empty_head 인 원만(비재료 짧은 선 무시)
    """
    ly0, col0 = bundle
    r = float(r)
    slack = max(5.0, r * 0.10)
    marks = [tuple(b) for b in (mark_bundles or ())]
    empty_pick = (not marks) or mark_fp_key(want_fp) == ("empty",)
    out = []
    if index is not None:
        if col0 is not None:
            rows = ((ly0, col0, *row[:3])
                    for row in index["by_bundle"].get((ly0, col0), ()))
        else:
            rows = ((ly, c, *row[:3])
                    for (ly, c), values in index["by_bundle"].items()
                    if ly == ly0 for row in values)
    else:
        rows = (row for row in w.circles
                if row[0] == ly0 and (col0 is None or row[1] == col0))
    for ly, c, cx, cy, qr in rows:
        if abs(float(qr) - r) > slack:
            continue
        cx, cy, qr = float(cx), float(cy), float(qr)
        if empty_pick:
            if disk_is_empty_head(w, cx, cy, qr, small_len, mat_set,
                                  spatial=spatial):
                out.append((ly, c, cx, cy, qr))
            continue
        got = mark_fp_on_bundles(
            w, cx, cy, qr, small_len, marks, spatial=spatial)
        if mark_fp_contains(want_fp, got):
            out.append((ly, c, cx, cy, qr))
    return out


def build_fp_index(w, small_len):
    """원마다 문양 지문을 1회만 찍어 표로 만든다 — 옛 완전일치 경로용.

    재료확정·문양자동채택 확장은 `circles_with_mark_pick` + spatial.
    반환: {"by_bundle": {(ly,c): [(cx,cy,qr,fp,fp_key), ...]},
           "n": 원수, "spatial": 격자, "small_len": ...}
    """
    spatial = _fp_build_spatial(w, small_len)
    by_bundle = defaultdict(list)
    n = 0
    for ly, c, cx, cy, qr in w.circles:
        if float(qr) <= 0:
            continue
        cx, cy, qr = float(cx), float(cy), float(qr)
        segs, arcs, circs = _fp_collect(
            cx, cy, qr, small_len, spatial=spatial)
        fp = _fp_tags(cx, cy, qr, segs, arcs, circs)
        key = mark_fp_key(fp)
        by_bundle[(ly, c)].append((cx, cy, qr, fp, key))
        n += 1
    return {"by_bundle": dict(by_bundle), "n": n,
            "spatial": spatial, "small_len": float(small_len)}


def fp_from_index(index, bundle, cx, cy, r, tol=1.0):
    """표에서 (묶음·자리·r) 에 맞는 지문 — 없으면 None."""
    if not index:
        return None
    r = float(r)
    slack = max(5.0, r * 0.10)
    tol2 = float(tol) * float(tol)
    for x, y, qr, fp, _key in index["by_bundle"].get(tuple(bundle), ()):
        if abs(qr - r) > slack:
            continue
        if (x - cx) * (x - cx) + (y - cy) * (y - cy) <= tol2:
            return fp
    return None


def circles_with_fp(w, bundle, r, fp, small_len, index=None,
                    mat_set=None, mark_bundles=None):
    """같은 (묶음×r×지문) 원 목록.

    mat_set 이 주어지면 재료확정·문양자동채택 경로
    (`circles_with_mark_pick`). 없으면 옛 완전일치(index 조회).
    """
    if mat_set is not None:
        spatial = (index or {}).get("spatial") if index else None
        return circles_with_mark_pick(
            w, bundle, r, fp, mark_bundles, mat_set, small_len,
            spatial=spatial, index=index)
    ly0, col0 = bundle
    r = float(r)
    slack = max(5.0, r * 0.10)
    want = mark_fp_key(fp)
    out = []
    if index is not None:
        rows = []
        if col0 is not None:
            rows = [((ly0, col0), index["by_bundle"].get((ly0, col0), ()))]
        else:
            rows = [((ly, c), lst)
                    for (ly, c), lst in index["by_bundle"].items()
                    if ly == ly0]
        for (ly, c), lst in rows:
            for cx, cy, qr, _fp, key in lst:
                if abs(qr - r) > slack:
                    continue
                if key == want:
                    out.append((ly, c, cx, cy, qr))
        return out
    for ly, c, cx, cy, qr in w.circles:
        if ly != ly0 or (col0 is not None and c != col0):
            continue
        if abs(float(qr) - r) > slack:
            continue
        got = mark_fp_key(mark_fingerprint(w, cx, cy, qr, small_len))
        if got == want:
            out.append((ly, c, float(cx), float(cy), float(qr)))
    return out

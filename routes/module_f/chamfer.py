# -*- coding: utf-8 -*-
"""[챔퍼] 45° 모서리 조각을 지우고 두 직선을 모서리까지 잇는다.

■ 목적 — 평면도와 아이소의 위상이 1:1 · 누락·휨 없이 (2026-09-09 오너)

  등각에서 나올 수 있는 각도는 셋뿐이다::

      평면 가로 → 화면 +30°   ·   평면 세로 → +150°   ·   헤드 스텁 → 수직

  평면의 45° 대각은 그 격자를 벗어나 **휘어** 보인다. 그리고 도면이 모서리를
  45°로 잘라 그리면(챔퍼) 노드정리(SSOT)가 «일직선» 만 병합하므로 그 조각이
  남아, 한 실제 배관이 「직선 – 챔퍼 – 직선」 **세 토막**이 된다.

  토막이 나면 셋이 한꺼번에 무너진다 (대명동 K=30 실측)::

      배관 262 중 0.3 m 미만 조각 140 (53 %) · 그중 45° 대각 91
      조각마다 근처 치수 텍스트를 **제각기** 물어 → 관경 역전 26건
      담당 헤드 1개인데 65A 같은 과대 관경 49건 (근거가 전부 «도면 텍스트»)

■ 왜 «병합» 이 아니라 «모서리 복원» 인가

  「직선 – 챔퍼 – 직선」을 **한 직선으로 접으면** 그 배관은 대각선이 된다 —
  양 끝이 어긋나 있기 때문이다. 등각에서 여전히 휜다. 목적을 못 이룬다.

  그래서 챔퍼를 **지우고** 두 직선을 교점까지 **연장**한다. 가로는 +30°,
  세로는 +150° 로 격자 위에만 남는다.

      pi ─가로─ i ─45°─ j ─세로─ pj      →      pi ─가로─ x ─세로─ pj
                                                        (x = 두 직선의 교점)

  길이는 모서리를 도는 만큼 **늘어난다**(실측 중앙값 +58.6 mm · 합 +2,899 mm).
  줄지 않으므로 마찰손실이 과소 산정될 일이 없다 — 보수측이다.

■ ★건드리지 않는 것

  · 양옆이 **직교가 아닌** 챔퍼 (실측 75곳) — 그 옆은 실제 대각 주행이다.
    접기 지시서 §6 이 「가지관의 2~7 m 45° 는 실제 주행이라 펴면 망가진다」고
    못박은 그것이다. 챔퍼와 주행은 **길이로 갈린다.**
  · 양 끝 차수가 2 가 아닌 조각 (실측 218곳) — 분기가 걸려 있어 지우면 그
    가지가 떨어진다.
  · 급수원·알람밸브가 가리키는 점 — 그 둘은 **인덱스로** 살아서, 없애면
    가리키는 자리가 통째로 어긋난다.
"""
from __future__ import annotations

import math

# 챔퍼로 볼 최대 길이(mm). ★이 자가 «챔퍼» 와 «실제 대각 주행» 을 가른다 —
#   대명동 실측: 챔퍼는 71·141 mm, 주행은 2~7 m 다. 사이가 비어 있어 300 mm
#   어디에 그어도 같은 답이 나온다.
MAX_MM = 300.0
# 각도 판정 여유(도).
TOL = 3.0


def _ang(a, b):
    """0~180. 가로 0 · 세로 90."""
    return math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])) % 180.0


def _meet(a1, a2, b1, b2):
    """두 직선의 교점. 평행이면 None."""
    d1 = (a2[0] - a1[0], a2[1] - a1[1])
    d2 = (b2[0] - b1[0], b2[1] - b1[1])
    den = d1[0] * d2[1] - d1[1] * d2[0]
    if abs(den) < 1e-9:
        return None
    t = ((b1[0] - a1[0]) * d2[1] - (b1[1] - a1[1]) * d2[0]) / den
    return (a1[0] + t * d1[0], a1[1] + t * d1[1])


def collapse_zero_edges(board, *, eps=0.001):
    """★같은 자리에 있는 두 점을 하나로 합친다 — 좌표는 그대로, 연결은 정리된다.

    실측(대명동 손질판)에서 이런 자리가 나왔다::

        880(275401,-234671) 차수 3 – 1227 · 챔퍼 70 mm
            880–1227    69.6 mm · 135.0°
            880–2243     0.0 mm ·   0.0°   ← ★길이 0
            880–2455    69.6 mm · 135.0°   ← 880–1227 의 겹친 복제
            1227–2455    0.0 mm ·   0.0°   ← ★길이 0

    인덱스가 한쪽은 880·1227, 다른 쪽은 2243·2455 다 — **두 묶음에서 온 같은
    자리**다(도면이 배관을 두 줄로 그린 것이 겹쳐 들어왔다).

    이것이 사슬의 시작이다::

        같은 좌표에 점이 둘 → 길이 0 간선 → 차수가 3·4 로 부푼다
          → 챔퍼를 못 편다(실측 218곳이 막힌 이유)
          → 표에 길이 0 배관이 실린다(실측 48개)
          → 배관망이 조각나고 관경이 튄다

    ★**좌표는 안 바뀌지만 «연결» 은 바뀔 수 있다.** 처음에 「형상은 안 바뀐다」고
      적었다가 실측으로 틀렸음을 확인했다 — 조건 없이 650개를 합쳤더니 표 전체
      총연장 +9.4 % · 평면 X 교차 +4 였다. 겹쳐 그린 두 줄에서 **한쪽 끝만**
      같은 자리면, 합치는 순간 그 점이 Y 자로 갈라져 «없던 분기» 가 생긴다.

      그래서 지금은 **합쳐도 분기가 늘지 않는 짝만** 합친다(합친 점의 이웃
      수가 원래 둘 중 많은 쪽을 안 넘을 것). 좌표가 안 바뀌는 것과 망이 안
      바뀌는 것은 다른 이야기다 — 못 합친 것은 세어서 돌려준다.

    ★번호는 «작은 쪽» 으로 모은다. 급수원·알람밸브가 인덱스로 살므로, 그것이
      가리키던 점이 사라지면 함께 옮겨 준다(안 그러면 가리키는 자리가 어긋난다).
    """
    pts = list(getattr(board, "pts", None) or ())
    edges = {tuple(sorted(e)) for e in (getattr(board, "edges", None) or ())}
    before = len(edges)

    # 합칠 짝 — union-find 로 사슬(셋 이상 겹침)까지 한 번에 모은다.
    par = {}

    def find(a):
        while par.get(a, a) != a:
            par[a] = par.get(par[a], par[a])
            a = par[a]
        return a

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra == rb:
            return
        lo, hi = (ra, rb) if ra < rb else (rb, ra)
        par[hi] = lo

    nb: dict = {}
    for (i, j) in edges:
        nb.setdefault(i, set()).add(j)
        nb.setdefault(j, set()).add(i)

    n_zero = n_skip = 0
    why: dict = {}
    for (i, j) in edges:
        if math.dist(pts[i], pts[j]) > eps:
            continue
        n_zero += 1
        # ★«합쳐도 없던 연결이 안 생기는» 짝만 합친다.
        #
        #   두 줄이 겹쳐 그려졌을 때 **양 끝이 모두** 같은 자리면 합치는 것이
        #   순수한 중복 제거다. 그런데 **한쪽 끝만** 같으면 합치는 순간 그 점이
        #   Y 자로 갈라져 «없던 분기» 가 생기고, 최불리 경로가 딴 길로 간다.
        #
        #   실측(대명동 K=30)에서 그 대가가 나왔다 — 조건 없이 650개를 합쳤더니
        #   표 총연장 +9.4 % · 평면 X 교차 +4.
        #
        #   자는 «합친 뒤 차수» 다. 두 점을 합쳐 만든 점의 이웃 수가 원래 둘 중
        #   많은 쪽보다 커지면 **분기가 는 것**이다 — 그때만 거른다.
        #
        #       차수 2 + 차수 2  → 합쳐도 2   (단순 중복점 · 안전)
        #       차수 3 + 차수 2  → 합쳐도 3   (Y 자가 이미 있었다 · 안전)
        #       차수 3 + 차수 3  → 4         (★없던 분기가 생긴다 · 거른다)
        ni = nb.get(i, set()) - {j}
        nj = nb.get(j, set()) - {i}
        if len(ni | nj) > max(len(nb.get(i, ())), len(nb.get(j, ()))):
            n_skip += 1
            why["합치면 분기가 는다"] = why.get("합치면 분기가 는다", 0) + 1
            continue
        union(i, j)
    if not par:
        return {"zero_edges": n_zero, "merged": 0, "skipped": n_skip,
                "why": why, "before": before, "after": before}

    rep = {}
    for a in list(par):
        rep[a] = find(a)
    merged = sum(1 for a, r in rep.items() if a != r)

    def R(v):
        return rep.get(v, v)

    out = set()
    for (i, j) in edges:
        a, b = R(i), R(j)
        if a == b:
            continue                    # 길이 0 간선 자체는 사라진다
        out.add(tuple(sorted((a, b))))
    board.edges = frozenset(out)

    # 급수원·알람밸브가 사라진 점을 가리키면 대표 점으로 옮긴다.
    for name in ("sources", "valves"):
        cur = list(getattr(board, name, None) or ())
        if not cur:
            continue
        setattr(board, name, [R(v) if isinstance(v, int) else v
                              for v in cur])
    return {"zero_edges": n_zero, "merged": merged, "skipped": n_skip,
            "why": why, "before": before, "after": len(out)}


def find_chamfers(board, *, max_mm=MAX_MM, tol=TOL):
    """모서리로 복원할 수 있는 챔퍼를 찾는다 — **좌표는 안 건드린다.**

    돌려주는 것: ([(i, j, pi, pj, 교점)], 왜 걸렀는지 셈)
    """
    pts = list(getattr(board, "pts", None) or ())
    edges = {tuple(sorted(e)) for e in (getattr(board, "edges", None) or ())}
    nb: dict = {}
    for i, j in edges:
        nb.setdefault(i, []).append(j)
        nb.setdefault(j, []).append(i)
    # 급수원·알람밸브가 가리키는 점은 없애면 안 된다(둘 다 인덱스로 산다).
    pinned = {int(n) for n in (getattr(board, "sources", None) or ())
              if isinstance(n, int)}
    pinned |= {int(n) for n in (getattr(board, "valves", None) or ())
               if isinstance(n, int)}

    def is45(a, b):
        v = _ang(pts[a], pts[b])
        return abs(v - 45) <= tol or abs(v - 135) <= tol

    def is_ortho(a, b):
        v = _ang(pts[a], pts[b])
        return v <= tol or v >= 180 - tol or abs(v - 90) <= tol

    out, why = [], {}

    def skip(k):
        why[k] = why.get(k, 0) + 1

    for (i, j) in edges:
        if math.dist(pts[i], pts[j]) >= max_mm:
            continue
        if not is45(i, j):
            continue
        if len(nb.get(i, ())) != 2 or len(nb.get(j, ())) != 2:
            skip("양 끝 차수≠2 (분기가 걸려 있다)")
            continue
        if i in pinned or j in pinned:
            skip("급수원·알람밸브가 가리키는 점")
            continue
        pi = [n for n in nb[i] if n != j][0]
        pj = [n for n in nb[j] if n != i][0]
        if not (is_ortho(pi, i) and is_ortho(j, pj)):
            skip("양옆이 직교가 아님 (실제 대각 주행)")
            continue
        a1, a2 = _ang(pts[pi], pts[i]), _ang(pts[j], pts[pj])
        if abs(a1 - a2) < tol or abs(abs(a1 - a2) - 180) < tol:
            skip("양옆이 서로 평행 (모서리가 아니다)")
            continue
        x = _meet(pts[pi], pts[i], pts[j], pts[pj])
        if x is None:
            skip("교점을 못 구함")
            continue
        out.append((i, j, pi, pj, x))
    return out, why


def restore_corners(board, *, max_mm=MAX_MM, tol=TOL):
    """찾은 챔퍼를 모서리로 편다. board 를 제자리에서 고친다.

    ★`pts` 의 **번호는 그대로** 둔다 — 급수원·알람밸브가 인덱스로 살기 때문이다
      (`flexfold` 에서 같은 이유로 배운 것). 챔퍼의 한 끝 `i` 를 교점으로
      **옮기고**, 다른 끝 `j` 는 간선만 걷어 고아로 남긴다. 안 쓰이는 점은
      `build_planar_graph` 가 무시한다.

          pi ─ i ─ j ─ pj        →        pi ─ i(=모서리) ─ pj

    돌려주는 것: {"corners", "why", "grew_mm", "before", "after"}
    """
    found, why = find_chamfers(board, max_mm=max_mm, tol=tol)
    edges = {tuple(sorted(e)) for e in (getattr(board, "edges", None) or ())}
    before = len(edges)
    pts = list(board.pts)
    grew = 0.0
    done = 0
    used = set()
    for i, j, pi, pj, x in found:
        # 한 조각이 다른 조각과 점을 나눠 쓰면 뒤엣것은 건너뛴다 — 앞의 조작으로
        # 이웃 관계가 이미 바뀌었으므로 지어내지 않는다.
        if {i, j, pi, pj} & used:
            why["앞 조작과 점을 나눠 씀"] = why.get("앞 조작과 점을 나눠 씀", 0) + 1
            continue
        e1 = tuple(sorted((pi, i)))
        e2 = tuple(sorted((i, j)))
        e3 = tuple(sorted((j, pj)))
        if not (e1 in edges and e2 in edges and e3 in edges):
            why["간선이 이미 바뀜"] = why.get("간선이 이미 바뀜", 0) + 1
            continue
        old = (math.dist(pts[pi], pts[i]) + math.dist(pts[i], pts[j])
               + math.dist(pts[j], pts[pj]))
        edges.discard(e2)
        edges.discard(e3)
        pts[i] = (float(x[0]), float(x[1]))
        edges.add(tuple(sorted((i, pj))))
        grew += (math.dist(pts[pi], pts[i]) + math.dist(pts[i], pts[pj])) - old
        used |= {i, j, pi, pj}
        done += 1
    board.pts = pts
    board.edges = frozenset(edges)
    return {"corners": done, "why": why, "grew_mm": round(grew, 1),
            "before": before, "after": len(edges),
            "found": len(found)}

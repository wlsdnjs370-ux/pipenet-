# -*- coding: utf-8 -*-
"""세션 A: 유저 손질용 순수 망 로직. 화면 없음.

노드 합치 여유만 `pipeline.stage1.SNAP`(Graph.node 와 동일)을 SSOT로 쓴다.
"""
import json
import math
import os
import tempfile
import unittest

from services.cad_import.pipeline.stage1 import SNAP  # Graph.node 와 같은 노드 합치 여유(30mm)

USER_AXIS_ANG = 5.0
USER_LAT_TOL = 20.0
USER_COVER_SLACK = 150.0  # 본체 r1_cover_slack 과 동일
# deprecated: 유저가 두 관을 직접 고르므로 틈 길이 거부는 쓰지 않음.
USER_CORNER_GAP = 300.0
USER_EDITS_SUFFIX = "_유저손질.json"


def _project_on_seg(p, a, b):
    dx, dy = b[0] - a[0], b[1] - a[1]
    den = dx * dx + dy * dy
    t = 0.0 if den == 0 else max(0.0, min(1.0,
        ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / den))
    q = (a[0] + t * dx, a[1] + t * dy)
    return q, t


def _pt_seg_d2(p, a, b):
    q, _t = _project_on_seg(p, a, b)
    return (p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2


def pick_seg(segs, x, y, max_d):
    """클릭 반경 안의 가장 가까운 선분을, 없으면 None을 돌려준다."""
    best = min(enumerate(segs), key=lambda z: _pt_seg_d2((x, y), *z[1]),
               default=(None, None))
    return best[1] if best[1] is not None and _pt_seg_d2(
        (x, y), *best[1]) <= max_d * max_d else None


def pick_head(disks, x, y, max_d):
    """클릭 반경 안의 가장 가까운 헤드 (cx,cy,r). 원판·테두리 바깥 max_d까지.

    거리 = max(0, 중심거리 − r). 원 안을 눌러도 고른다.
    """
    best = None  # (d, disk)
    for disk in disks:
        hx, hy, hr = disk[0], disk[1], disk[2]
        d = max(0.0, math.hypot(x - hx, y - hy) - float(hr))
        if d <= max_d and (best is None or d < best[0]):
            best = (d, (float(hx), float(hy), float(hr)))
    return None if best is None else best[1]


def head_bridge(seg, disk, ang_tol=USER_AXIS_ANG, lat_tol=USER_LAT_TOL):
    """선택 관 자유단 → 헤드 «중심» 한 선분. 조건 불만족이면 None.

    왼쪽 가로선이 곧 헤드 팔인 자리(헤드 쪽 짧은 관 없음)용.
    · 자유단 = 헤드 중심에 더 가까운 끝, 관 방향이 그 끝을 넘어 헤드를 향함
    · 횡이탈 ≤ USER_LAT_TOL (배관↔배관 이음과 동일)
    · 이미 head_nodes 기준(중심 ARM_CTR · 테두리 HEAD_TOUCH)으로 붙었으면 None
    · 도착점은 중심 [2026-08-13 오너 — 접속은 정확히 헤드 중심에서.
      kfp 와 같은 상태로 편집을 완성한다]. 여유값은 water SSOT.
    """
    from services.cad_import.pipeline.flow import ARM_CTR, HEAD_TOUCH

    aa = _segments(seg)
    if not aa:
        return None
    a0, a1 = aa[0]
    hx, hy, hr = float(disk[0]), float(disk[1]), float(disk[2])
    if hr <= 0:
        return None
    da = math.hypot(a0[0] - hx, a0[1] - hy)
    db = math.hypot(a1[0] - hx, a1[1] - hy)
    end, other = (a0, a1) if da <= db else (a1, a0)
    dx, dy = end[0] - other[0], end[1] - other[1]
    plen = math.hypot(dx, dy)
    if plen < 1e-9:
        return None
    ux, uy = dx / plen, dy / plen
    wx, wy = hx - end[0], hy - end[1]
    along = wx * ux + wy * uy
    lat = abs(-wx * uy + wy * ux)
    if along <= 0 or lat > lat_tol:
        return None
    # 관 방향과 끝→중심 방향이 USER_AXIS_ANG 안인지(횡이탈만으로 부족할 때)
    d_end = math.hypot(wx, wy)
    if d_end < 1e-9:
        return None
    cos_min = math.cos(math.radians(ang_tol))
    if (wx * ux + wy * uy) / d_end < cos_min:
        return None
    if d_end <= ARM_CTR or abs(d_end - hr) <= HEAD_TOUCH:
        return None
    ctr = (hx, hy)
    if math.hypot(ctr[0] - end[0], ctr[1] - end[1]) < 1e-9:
        return None
    return (tuple(end), ctr)


def _segments(value):
    """선분 하나 ((x,y),(x,y)) 또는 그 선분들의 묶음을 같은 형태로 만든다."""
    if len(value) == 2 and len(value[0]) == 2 and isinstance(value[0][0], (int, float)):
        return [value]
    return list(value)


def colinear_bridges(seg_a, seg_b, segs=None, ang_tol=USER_AXIS_ANG,
                     lat_tol=USER_LAT_TOL):
    """두 선택 관 사이에서만 segs의 빈 구간을 끝↔끝으로 잇는다."""
    aa, bb = _segments(seg_a), _segments(seg_b)
    selected = aa + bb
    if not selected:
        return []
    p, q = selected[0]
    dx, dy = q[0] - p[0], q[1] - p[1]
    length = math.hypot(dx, dy)
    if length == 0:
        return []
    ux, uy = dx / length, dy / length
    cos_min = math.cos(math.radians(ang_tol))

    def project(v):
        return (v[0] - p[0]) * ux + (v[1] - p[1]) * uy

    def lateral(v):
        return abs(-(v[0] - p[0]) * uy + (v[1] - p[1]) * ux)

    def interval(a, b):
        sx, sy = b[0] - a[0], b[1] - a[1]
        slen = math.hypot(sx, sy)
        if (slen == 0 or abs((sx * ux + sy * uy) / slen) < cos_min or
                max(lateral(a), lateral(b)) > lat_tol):
            return None
        ta, tb = project(a), project(b)
        return min(ta, tb), max(ta, tb)

    selected_intervals = [interval(a, b) for a, b in selected]
    if any(x is None for x in selected_intervals):
        return []

    a_span = (min(x[0] for x in selected_intervals[:len(aa)]),
              max(x[1] for x in selected_intervals[:len(aa)]))
    b_span = (min(x[0] for x in selected_intervals[len(aa):]),
              max(x[1] for x in selected_intervals[len(aa):]))
    if a_span[1] <= b_span[0]:
        lo, hi = a_span[1], b_span[0]
    elif b_span[1] <= a_span[0]:
        lo, hi = b_span[1], a_span[0]
    else:
        return []

    candidates = selected if segs is None else list(segs) + selected
    covered = sorted((max(lo, x[0]), min(hi, x[1]))
                     for a, b in candidates if (x := interval(a, b)) is not None
                     and x[0] < hi and x[1] > lo)
    out, cursor = [], lo
    for a, b in covered:
        if a > cursor:
            out.append(((p[0] + cursor * ux, p[1] + cursor * uy),
                        (p[0] + a * ux, p[1] + a * uy)))
        cursor = max(cursor, b)
    if cursor < hi:
        out.append(((p[0] + cursor * ux, p[1] + cursor * uy),
                    (p[0] + hi * ux, p[1] + hi * uy)))
    return out


def colinear_chain_segs(pts, edges, seed_seg, ang_tol=USER_AXIS_ANG,
                        lat_tol=USER_LAT_TOL):
    """seed 와 같은 축·같은 물덩이에서 연속으로 붙은 공선 토막들.

    클릭한 한 토막만으로 T자 수선이 실패하지 않도록 trunk 후보를 모은다.
    다른 물덩이는 끌어오지 않는다. 직선 전체 일괄 이음이 아니다.
    """
    aa = _segments(seed_seg)
    if not aa:
        return []
    seed = (tuple(aa[0][0]), tuple(aa[0][1]))
    a0, a1 = seed
    dx, dy = a1[0] - a0[0], a1[1] - a0[1]
    length = math.hypot(dx, dy)
    if length < 1e-9:
        return [seed]
    ux, uy = dx / length, dy / length
    cos_min = math.cos(math.radians(ang_tol))

    def is_colinear_edge(i, j):
        p, q = pts[i], pts[j]
        sx, sy = q[0] - p[0], q[1] - p[1]
        slen = math.hypot(sx, sy)
        if slen < 1e-9:
            return False
        if abs((sx * ux + sy * uy) / slen) < cos_min:
            return False
        for v in (p, q):
            lat = abs(-(v[0] - a0[0]) * uy + (v[1] - a0[1]) * ux)
            if lat > lat_tol:
                return False
        return True

    seed_e = _edge_matching_seg(pts, edges, seed)
    if seed_e is None:
        return [seed]
    bodies = recompute_bodies(pts, edges)
    body_of = {n: bi for bi, body in enumerate(bodies) for n in body}
    seed_bi = body_of.get(seed_e[0])
    adj = {}
    for i, j in edges:
        adj.setdefault(i, set()).add(j)
        adj.setdefault(j, set()).add(i)
    collected = {tuple(sorted(seed_e))}
    stack = [seed_e[0], seed_e[1]]
    seen = {seed_e[0], seed_e[1]}
    while stack:
        n = stack.pop()
        for m in adj.get(n, ()):
            e = tuple(sorted((n, m)))
            if e in collected:
                continue
            if body_of.get(n) != seed_bi or body_of.get(m) != seed_bi:
                continue
            if not is_colinear_edge(n, m):
                continue
            collected.add(e)
            if m not in seen:
                seen.add(m)
                stack.append(m)
    return [((float(pts[i][0]), float(pts[i][1])),
             (float(pts[j][0]), float(pts[j][1]))) for i, j in collected]


def corner_bridges(seg_a, seg_b, gap_tol=None, ang_tol=USER_AXIS_ANG):
    """대략 직교인 두 관의 가장 가까운 끝↔끝만 한 선분으로 잇는다.

    일직선(colinear)이 아닐 때 쓰는 코너 이음. 사이 추정·연장 금지.
    |방향 내적| ≤ sin(ang_tol) 이면 90±ang 수준으로 본다.
    브리지가 어느 관 축에도 평행하지 않으면(≈45°) 거부 — 직교는 T자 수선.
    길이 0(이미 같은 점)만 제외. gap_tol은 deprecated·무시.
    """
    del gap_tol  # deprecated: 유저 직접 선택이므로 틈 길이로 거부하지 않음
    aa, bb = _segments(seg_a), _segments(seg_b)
    sin_max = math.sin(math.radians(ang_tol))
    cos_min = math.cos(math.radians(ang_tol))
    best = None  # (dist2, (pa, pb), dir_a, dir_b)
    for a0, a1 in aa:
        adx, ady = a1[0] - a0[0], a1[1] - a0[1]
        alen = math.hypot(adx, ady)
        if alen < 1e-9:
            continue
        aux, auy = adx / alen, ady / alen
        for b0, b1 in bb:
            bdx, bdy = b1[0] - b0[0], b1[1] - b0[1]
            blen = math.hypot(bdx, bdy)
            if blen < 1e-9:
                continue
            bux, buy = bdx / blen, bdy / blen
            if abs(aux * bux + auy * buy) > sin_max:
                continue
            for pa in (a0, a1):
                for pb in (b0, b1):
                    d2 = (pa[0] - pb[0]) ** 2 + (pa[1] - pb[1]) ** 2
                    if best is None or d2 < best[0]:
                        best = (d2, (tuple(pa), tuple(pb)), (aux, auy),
                                (bux, buy))
    if best is None or best[0] < 1e-18:
        return []
    (pa, pb), (aux, auy), (bux, buy) = best[1], best[2], best[3]
    gx, gy = pb[0] - pa[0], pb[1] - pa[1]
    glen = math.hypot(gx, gy)
    if glen < 1e-9:
        return []
    gux, guy = gx / glen, gy / glen
    # 축정렬 L만 — 대각선(두 축 어느 쪽과도 비평행) 거부
    if (abs(gux * aux + guy * auy) < cos_min and
            abs(gux * bux + guy * buy) < cos_min):
        return []
    return [(pa, pb)]


def tee_bridges(seg_a, seg_b, ang_tol=USER_AXIS_ANG, chain_a=None,
                chain_b=None):
    """직교 T자 — 한쪽 관 끝 → 다른 관 **몸통** 위 수선발.

    가로가 세로 옆구리에 닿아야 하는 자리 [2026-08-08 오너].
    끝↔끝(L자)은 corner_bridges 몫이므로 수선발이 몸통 안(0<t<1)일 때만.
    chain_a/chain_b: 공선 연속 토막 — stub는 선택 관, trunk는 체인 전부.
    반환: [(stub_end, foot, trunk_seg), ...] 최대 1개.
    """
    aa, bb = _segments(seg_a), _segments(seg_b)
    trunks_a = _segments(chain_a) if chain_a is not None else aa
    trunks_b = _segments(chain_b) if chain_b is not None else bb
    sin_max = math.sin(math.radians(ang_tol))
    best = None  # (dist, stub_end, foot, trunk)

    def consider(stub, trunk):
        nonlocal best
        t0, t1 = trunk
        tx, ty = t1[0] - t0[0], t1[1] - t0[1]
        tlen = math.hypot(tx, ty)
        if tlen < 1e-9:
            return
        tux, tuy = tx / tlen, ty / tlen
        s0, s1 = stub
        sx, sy = s1[0] - s0[0], s1[1] - s0[1]
        slen = math.hypot(sx, sy)
        if slen < 1e-9:
            return
        if abs((sx / slen) * tux + (sy / slen) * tuy) > sin_max:
            return
        for end, other in ((s0, s1), (s1, s0)):
            # 관이 end 쪽으로 이어져 trunk 를 향할 때만
            ux, uy = end[0] - other[0], end[1] - other[1]
            ulen = math.hypot(ux, uy)
            if ulen < 1e-9:
                continue
            ux, uy = ux / ulen, uy / ulen
            wx, wy = end[0] - t0[0], end[1] - t0[1]
            t = wx * tux + wy * tuy
            if not (0.0 < t < tlen):
                continue
            foot = (t0[0] + t * tux, t0[1] + t * tuy)
            fx, fy = foot[0] - end[0], foot[1] - end[1]
            dist = math.hypot(fx, fy)
            if dist < 1e-9:
                continue
            if fx * ux + fy * uy <= 0:
                continue  # 수선이 관 뒤쪽
            if best is None or dist < best[0]:
                best = (dist, tuple(end), foot, (tuple(t0), tuple(t1)))

    for sa in aa:
        for tb in trunks_b:
            consider(sa, tb)
    for sb in bb:
        for ta in trunks_a:
            consider(sb, ta)
    if best is None:
        return []
    return [(best[1], best[2], best[3])]


def _find_node(pts, xy, snap=SNAP):
    """SNAP 안 최근접 노드. 없으면 None. Graph.node 와 같은 눈금."""
    found = _find_nodes(pts, xy, snap=snap)
    if not found:
        return None
    x, y = float(xy[0]), float(xy[1])
    return min(found, key=lambda i: (pts[i][0] - x) ** 2 + (pts[i][1] - y) ** 2)


def _find_nodes(pts, xy, snap=SNAP):
    """SNAP 안 노드 전부 — 같은 자리 중복 노드를 빠뜨리지 않는다."""
    x, y = float(xy[0]), float(xy[1])
    lim = float(snap) * float(snap)
    return [i for i, (px, py) in enumerate(pts)
            if (px - x) ** 2 + (py - y) ** 2 <= lim]


def _node(pts, xy, snap=SNAP):
    """SNAP 안이면 기존 노드, 없으면 추가. Graph.node 와 같은 눈금."""
    found = _find_node(pts, xy, snap=snap)
    if found is not None:
        return found
    pts.append((float(xy[0]), float(xy[1])))
    return len(pts) - 1


def head_cover_ok(bridge, disks, cover_slack=USER_COVER_SLACK):
    """헤드 원이 브리지 축을 자르고 양끝 남는 길이가 cover_slack 이하면 True.

    본체 join_by_head_cover 의 head_explains 와 같은 취지:
    ① 횡이탈 < r  ② 원 테두리 밖 남는 길이 양끝 각각 ≤ cover_slack.
    """
    pa, pb = bridge
    dx, dy = pb[0] - pa[0], pb[1] - pa[1]
    length = math.hypot(dx, dy)
    if length < 1e-9:
        return False
    ux, uy = dx / length, dy / length
    for qx, qy, qr in disks:
        if qr <= 0:
            continue
        wx, wy = qx - pa[0], qy - pa[1]
        along = wx * ux + wy * uy
        lat = abs(-wx * uy + wy * ux)
        if lat >= qr:
            continue
        half = math.sqrt(qr * qr - lat * lat)
        if along - half <= cover_slack and length - (along + half) <= cover_slack:
            return True
    return False


def apply_joins(pts, edges, bridges, snap=SNAP):
    """bridge 좌표쌍을 노드/무방향 간선으로 추가해 (pts, edges)를 돌려준다."""
    pts2, edges2 = list(pts), {tuple(sorted(e)) for e in edges}
    for a, b in bridges:
        edges2.add(tuple(sorted((
            _node(pts2, a, snap=snap), _node(pts2, b, snap=snap)))))
    return pts2, frozenset(edges2)


def _edge_matching_seg(pts, edges, seg, tol=1.0):
    """seg 양 끝과 맞는 현재 간선. 없으면 None."""
    a, b = seg
    na = _find_nodes(pts, a, snap=tol)
    nb = _find_nodes(pts, b, snap=tol)
    edge_set = {tuple(sorted(e)) for e in edges}
    for i in na:
        for j in nb:
            e = tuple(sorted((i, j)))
            if e in edge_set and i != j:
                return e
    return None


def apply_tee_bridge(pts, edges, stub_end, foot, trunk_seg, snap=SNAP):
    """T자 이음: trunk 를 foot 에서 쪼개고 stub_end↔foot 간선을 단다.

    foot 이 trunk 몸통에 새 노드로 안 들어가면 물이 trunk 로 못 넘어간다.
    """
    pts2 = list(pts)
    edges2 = {tuple(sorted(e)) for e in edges}
    trunk_e = _edge_matching_seg(pts2, edges2, trunk_seg)
    i_stub = _node(pts2, stub_end, snap=snap)
    # foot 은 trunk 위 점이라 SNAP 이 옆 가지 노드로 끌려가면 안 된다 —
    # trunk 끝점이면 그 노드, 아니면 좌표 고정 노드.
    if trunk_e is not None:
        ia, ib = trunk_e
        if math.hypot(pts2[ia][0] - foot[0], pts2[ia][1] - foot[1]) <= snap:
            i_foot = ia
        elif math.hypot(pts2[ib][0] - foot[0], pts2[ib][1] - foot[1]) <= snap:
            i_foot = ib
        else:
            i_foot = _node(pts2, foot, snap=0.0)
            if i_foot != ia and i_foot != ib:
                edges2.discard(trunk_e)
                edges2.add(tuple(sorted((ia, i_foot))))
                edges2.add(tuple(sorted((i_foot, ib))))
    else:
        i_foot = _node(pts2, foot, snap=snap)
    if i_stub != i_foot:
        edges2.add(tuple(sorted((i_stub, i_foot))))
    return pts2, frozenset(edges2)


def apply_head_bridge(pts, edges, bridge):
    """관말→헤드: 관 끝은 SNAP, 헤드 쪽(중심)은 좌표 고정(짧은 틈이 SNAP에 접히지 않게)."""
    pts2, edges2 = list(pts), {tuple(sorted(e)) for e in edges}
    a, b = bridge
    ia = _node(pts2, a, snap=SNAP)
    ib = _node(pts2, b, snap=0.0)
    if ia != ib:
        edges2.add(tuple(sorted((ia, ib))))
    return pts2, frozenset(edges2)


# 같은 자리 겹선 판정 — SNAP(30)은 짧은 관의 옆 노드까지 먹어 과삭제된다.
_DELETE_TWIN_TOL = 1.0  # mm


def insert_node_on_pipe(pts, edges, x, y, max_d, snap=SNAP):
    """관 위 수선점에 접속 노드를 넣는다. 끝점 SNAP 안이면 그 노드.

    반환: (pts, edges, node). 반경 밖이면 node is None.
    """
    p = (float(x), float(y))
    best = None
    for i, j in edges:
        a, b = pts[i], pts[j]
        q, _t = _project_on_seg(p, a, b)
        d2 = (p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2
        if best is None or d2 < best[0]:
            best = (d2, i, j, q)
    pts2 = list(pts)
    edges2 = {tuple(sorted(e)) for e in edges}
    if best is None or best[0] > float(max_d) * float(max_d):
        return pts2, frozenset(edges2), None
    _d2, i, j, q = best
    if math.hypot(pts2[i][0] - q[0], pts2[i][1] - q[1]) <= snap:
        return pts2, frozenset(edges2), i
    if math.hypot(pts2[j][0] - q[0], pts2[j][1] - q[1]) <= snap:
        return pts2, frozenset(edges2), j
    nid = _node(pts2, q, snap=0.0)
    edges2.discard(tuple(sorted((i, j))))
    if nid != i:
        edges2.add(tuple(sorted((i, nid))))
    if nid != j:
        edges2.add(tuple(sorted((nid, j))))
    return pts2, frozenset(edges2), nid


def apply_deletes(pts, edges, deletes):
    """원본/유저 이음 모두 좌표 선분 또는 (노드, 노드)로 삭제한다.

    좌표 선분: 끝점마다 «같은 자리»(1mm) 노드를 모아, 그 사이 간선 **전부**를
    지운다 [2026-08-08 MF 우상단 — 중복 노드 쌍이 남아 안 지워지던 원인].
    1mm 안에 없으면 SNAP 최근접 한 쌍으로 한 번 더 찾는다(저장 좌표 여유).
    """
    doomed = set()
    for item in deletes:
        if isinstance(item, dict):
            item = (item["a"], item["b"])
        if len(item) == 2 and all(isinstance(n, int) for n in item):
            doomed.add(tuple(sorted(item)))
            continue
        a, b = item
        na = _find_nodes(pts, a, snap=_DELETE_TWIN_TOL)
        nb = _find_nodes(pts, b, snap=_DELETE_TWIN_TOL)
        hit = set()
        for ia in na:
            for ib in nb:
                if ia != ib:
                    hit.add(tuple(sorted((ia, ib))))
        if not hit:
            ia, ib = _find_node(pts, a), _find_node(pts, b)
            if ia is not None and ib is not None and ia != ib:
                hit.add(tuple(sorted((ia, ib))))
        doomed |= hit
    return frozenset(tuple(sorted(e)) for e in edges if tuple(sorted(e)) not in doomed)


def recompute_bodies(pts, edges):
    """간선에 있는 노드의 연결 성분 목록을 돌려준다."""
    adj = {}
    for a, b in edges:
        adj.setdefault(a, set()).add(b)
        adj.setdefault(b, set()).add(a)
    out, unseen = [], set(adj)
    while unseen:
        todo, body = [unseen.pop()], set()
        while todo:
            n = todo.pop()
            if n in body:
                continue
            body.add(n)
            for m in adj[n]:
                if m not in body:
                    unseen.discard(m)
                    todo.append(m)
        out.append(frozenset(body))
    return out


def wet_heads(reach, head_nodes):
    """도달 노드 집합에 하나라도 닿은 헤드 인덱스 집합."""
    return {i for i, nodes in enumerate(head_nodes) if set(nodes) & set(reach)}


def user_edits_path(key, dwg_dir):
    return os.path.join(dwg_dir, f"{key}{USER_EDITS_SUFFIX}")


# 종류 정규화 로직은 전부 kinds.py 로 모았다 — 여기 있던 본문도 그리로.
# import 경로 호환을 위한 재수출이다(engine·preflight·planar·board·io 가
# 여기서 가져간다).
from services.cad_import.kinds import apply_kind_overrides  # noqa: F401,E402


def _xy_of_pick(rec):
    if isinstance(rec, dict):
        xy = rec.get("xy") or ()
    else:
        xy = rec
    if not xy or len(xy) < 2:
        return None
    return (float(xy[0]), float(xy[1]))


def user_edit_valve_picks(key, dwg_dir):
    """유저손질 JSON 의 밸브 찍기 좌표. 파일 없으면 []."""
    path = user_edits_path(key, dwg_dir)
    if not os.path.exists(path):
        return []
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    out = []
    for rec in raw.get("valve_picks") or []:
        xy = _xy_of_pick(rec)
        if xy is not None:
            out.append({"xy": [xy[0], xy[1]]})
    return out


def apply_user_edits(key, pts, edges, dwg_dir):
    """파일 없으면 (pts, edges, [], []). 있으면 joins/deletes/밸브찍기 적용 후
    (pts, edges, sources_raw, kind_overrides).

    kind_overrides 는 이음에 쓰지 않는다 — 호출측이 head_kinds 에만 덮는다.
    """
    path = user_edits_path(key, dwg_dir)
    if not os.path.exists(path):
        return pts, edges, [], []
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    for rec in raw.get("joins") or []:
        bridges = [(tuple(x[0]), tuple(x[1])) for x in rec.get("bridges") or []]
        if (rec.get("kind") == "헤드" or rec.get("head")) and bridges:
            for br in bridges:
                pts, edges = apply_head_bridge(pts, edges, br)
        elif rec.get("kind") == "T자" and bridges:
            trunks = rec.get("trunks") or []
            for i, br in enumerate(bridges):
                if i < len(trunks):
                    tr = trunks[i]
                    pts, edges = apply_tee_bridge(
                        pts, edges, br[0], br[1],
                        (tuple(tr[0]), tuple(tr[1])))
                else:
                    pts, edges = apply_joins(pts, edges, [br])
        else:
            pts, edges = apply_joins(pts, edges, bridges)
    edges = apply_deletes(pts, edges, list(raw.get("deletes") or []))
    for rec in raw.get("valve_picks") or []:
        xy = _xy_of_pick(rec)
        if xy is None:
            continue
        pts, edges, _nid = insert_node_on_pipe(
            pts, edges, xy[0], xy[1], max_d=SNAP)
    return (pts, edges, list(raw.get("sources") or []),
            list(raw.get("kind_overrides") or []))


class UserNetTest(unittest.TestCase):
    def test_gates(self):
        # 한 번의 양끝 관 픽: 중간 관이 덮지 않는 두 틈만 이음.
        segs = [((0, 0), (10, 0)), ((20, 0), (30, 0)), ((40, 0), (50, 0))]
        bridges = colinear_bridges(segs[0], segs[-1], segs)
        self.assertEqual(2, len(bridges))
        self.assertEqual([((10.0, 0.0), (20.0, 0.0)),
                          ((30.0, 0.0), (40.0, 0.0))], bridges)
        self.assertEqual([], colinear_bridges(((0, 0), (100, 0)),
                                              ((120, 0), (220, math.tan(
                                                  math.radians(6)) * 100))))
        self.assertEqual([], colinear_bridges(((0, 0), (100, 0)),
                                              ((120, 25), (220, 25))))
        pts = [(0, 0), (10, 0), (20, 0), (30, 0)]
        edges = {(0, 1), (1, 2), (2, 3)}
        split = recompute_bodies(pts, apply_deletes(
            pts, edges, [((10, 0), (20, 0))]))
        self.assertEqual(2, len(split))
        # 같은 자리 중복 노드·간선 — 한 번에 전부 [2026-08-08 MF]
        d_pts = [(0.0, 0.0), (10.0, 0.0), (0.0, 0.0), (10.0, 0.0)]
        d_edges = {(0, 1), (2, 3)}
        d_after = apply_deletes(d_pts, d_edges, [((0.0, 0.0), (10.0, 0.0))])
        self.assertEqual(0, len(d_after))
        joined_pts, joined_edges = apply_joins(
            [(0, 0), (10, 0)], {(0, 1)}, [((10, 0), (20, 0))])
        self.assertEqual(1, len(recompute_bodies(joined_pts, joined_edges)))
        # float 미세 어긋남도 SNAP 안이면 기존 노드에 붙어 물덩이가 합쳐진다.
        drift = 1e-9
        a_pts = [(0.0, 0.0), (10.0, 0.0), (20.0 + drift, 0.0), (30.0, 0.0)]
        a_edges = {(0, 1), (2, 3)}
        self.assertEqual(2, len(recompute_bodies(a_pts, a_edges)))
        b_pts, b_edges = apply_joins(
            a_pts, a_edges, [((10.0 + drift, 0.0), (20.0, 0.0))])
        self.assertEqual(4, len(b_pts))  # 새 점 없이 기존 4노드만
        self.assertEqual(1, len(recompute_bodies(b_pts, b_edges)))
        self.assertEqual({0}, wet_heads({0, 1}, [{1}, {2}]))
        self.assertEqual(((10, 0), (20, 0)), pick_seg(
            [((10, 0), (20, 0))], 14, 3, 5))
        i_pts, i_edges, i_n = insert_node_on_pipe(
            [(0.0, 0.0), (100.0, 0.0)], {(0, 1)}, 40.0, 2.0, max_d=10.0)
        self.assertEqual(3, len(i_pts))
        self.assertAlmostEqual(40.0, i_pts[i_n][0], places=6)
        self.assertAlmostEqual(0.0, i_pts[i_n][1], places=6)
        self.assertEqual(2, len(i_edges))
        # 끝점 SNAP 안이면 새 노드 없이 그 노드
        e_pts, e_edges, e_n = insert_node_on_pipe(
            [(0.0, 0.0), (100.0, 0.0)], {(0, 1)}, 5.0, 0.0, max_d=20.0)
        self.assertEqual(0, e_n)
        self.assertEqual(2, len(e_pts))
        self.assertEqual(1, len(e_edges))
        # 헤드걸침: 원이 틈 한가운데를 덮고 양끝 여유 ≤ 150 → True
        self.assertTrue(head_cover_ok(
            ((10.0, 0.0), (20.0, 0.0)), [(15.0, 0.0, 3.0)]))
        # 헤드가 한쪽에만 있고 반대가 멀리 빔 → False
        self.assertFalse(head_cover_ok(
            ((10.0, 0.0), (400.0, 0.0)), [(30.0, 0.0, 25.0)]))
        print("gate: one_pick bridges=2 outer=0 angle6=0 lateral25=0 "
              "bodies=1->2 head_cover=ok/reject")
        # 관말→헤드: HEAD_TOUCH(50) 밖 끊김 → 중심까지 bridge [2026-08-13]
        disk = (100.0, 0.0, 10.0)
        br = head_bridge(((0.0, 0.0), (30.0, 0.0)), disk)
        self.assertIsNotNone(br)
        self.assertEqual((30.0, 0.0), tuple(br[0]))
        self.assertAlmostEqual(100.0, br[1][0], places=6)  # 중심 x=100
        self.assertAlmostEqual(0.0, br[1][1], places=6)
        # 이미 head_nodes 여유(테두리±HEAD_TOUCH) 안 → None
        self.assertIsNone(head_bridge(((0.0, 0.0), (50.0, 0.0)), disk))
        # 헤드가 자유단 뒤쪽 → None
        self.assertIsNone(head_bridge(((20.0, 0.0), (80.0, 0.0)),
                                      (-10.0, 0.0, 10.0)))
        # 횡이탈 과다 → None
        self.assertIsNone(head_bridge(((0.0, 0.0), (30.0, 0.0)),
                                      (100.0, 40.0, 10.0)))
        self.assertEqual((100.0, 0.0, 10.0), pick_head([disk], 100.0, 0.0, 5.0))
        self.assertIsNone(pick_head([disk], 200.0, 0.0, 5.0))

    def test_corner_bridges(self):
        # L자 끊김: 가로 끝 (100,0) ↔ 세로 끝 (150,0)
        bridges = corner_bridges(((0, 0), (100, 0)), ((150, 0), (150, 100)))
        self.assertEqual(1, len(bridges))
        a, b = bridges[0]
        self.assertEqual({(100.0, 0.0), (150.0, 0.0)},
                         {tuple(a), tuple(b)})
        # 평행 → 0
        self.assertEqual([], corner_bridges(((0, 0), (100, 0)),
                                            ((0, 50), (100, 50))))
        # 둔각 120° → 0
        dx = math.cos(math.radians(120.0)) * 100.0
        dy = math.sin(math.radians(120.0)) * 100.0
        self.assertEqual([], corner_bridges(
            ((0, 0), (100, 0)), ((120, 0), (120 + dx, dy))))
        # 멀리 있어도 최근접 끝↔끝이면 bridge 1 (틈 길이 거부 없음)
        far = corner_bridges(((0, 0), (100, 0)), ((500, 0), (500, 100)))
        self.assertEqual(1, len(far))
        fa, fb = far[0]
        self.assertEqual({(100.0, 0.0), (500.0, 0.0)},
                         {tuple(fa), tuple(fb)})
        # 이미 같은 점 → 0
        self.assertEqual([], corner_bridges(
            ((0, 0), (100, 0)), ((100, 0), (100, 100))))
        # T자: 가로 끝 → 세로 몸통 수선 (L자와 다름)
        tee = tee_bridges(((0.0, 50.0), (80.0, 50.0)),
                          ((100.0, 0.0), (100.0, 100.0)))
        self.assertEqual(1, len(tee))
        end, foot, trunk = tee[0]
        self.assertEqual((80.0, 50.0), end)
        self.assertAlmostEqual(100.0, foot[0], places=6)
        self.assertAlmostEqual(50.0, foot[1], places=6)
        pts, edges = apply_tee_bridge(
            [(0.0, 50.0), (80.0, 50.0), (100.0, 0.0), (100.0, 100.0)],
            {(0, 1), (2, 3)}, end, foot, trunk)
        self.assertEqual(1, len(recompute_bodies(pts, edges)))
        self.assertEqual(4, len(edges))  # stub + trunk×2 + T이음
        # L자는 tee 가 아니라 corner (몸통 t 가 끝점)
        self.assertEqual([], tee_bridges(((0, 0), (100, 0)),
                                         ((150, 0), (150, 100))))
        # 대각선 끝↔끝(≈45°) 거부 — 직교는 T자 수선
        self.assertEqual([], corner_bridges(
            ((-50.0, 100.0), (0.0, 100.0)),
            ((100.0, 0.0), (100.0, -50.0))))
        # 쪼갠 세로: 아래 토막만 클릭해도 chain 으로 T자
        split_pts = [(0.0, 50.0), (80.0, 50.0),
                     (100.0, 0.0), (100.0, 40.0), (100.0, 200.0)]
        split_edges = {(0, 1), (2, 3), (3, 4)}
        chain_v = colinear_chain_segs(
            split_pts, split_edges, ((100.0, 0.0), (100.0, 40.0)))
        self.assertEqual(2, len(chain_v))
        tee_split = tee_bridges(
            ((0.0, 50.0), (80.0, 50.0)), ((100.0, 0.0), (100.0, 40.0)),
            chain_b=chain_v)
        self.assertEqual(1, len(tee_split))
        self.assertAlmostEqual(100.0, tee_split[0][1][0], places=6)
        self.assertAlmostEqual(50.0, tee_split[0][1][1], places=6)
        print("gate: corner L=1 parallel=0 obtuse120=0 far=1 samept=0 "
              "tee=1 tee_split=4 tee_not_L=0 diag45=0 chain_tee=1")

    def test_apply_user_edits(self):
        # 새 끝점은 SNAP(30) 밖이어야 별 노드가 된다(Graph.node 와 동일).
        pts, edges = [(0.0, 0.0), (10.0, 0.0)], frozenset([(0, 1)])
        with tempfile.TemporaryDirectory() as td:
            miss = apply_user_edits("__missing__", pts, edges, td)
            self.assertEqual((pts, edges, [], []), miss)
            path = user_edits_path("__t__", td)
            with open(path, "w", encoding="utf-8") as f:
                json.dump({
                    "version": 1,
                    "joins": [{"a": [[0, 0], [10, 0]], "b": [[50, 0], [80, 0]],
                               "bridges": [[[10, 0], [50, 0]]]}],
                    "deletes": [],
                    "sources": [{"tag": "Z1", "xy": [0, 0]}],
                    "kind_overrides": [
                        {"c": [100.0, 0.0], "r": 5.0, "kind": "하향식"}],
                }, f)
            pts2, edges2, srcs, kovs = apply_user_edits("__t__", pts, edges, td)
            self.assertEqual(3, len(pts2))
            self.assertIn((1, 2), {tuple(sorted(e)) for e in edges2})
            self.assertEqual([{"tag": "Z1", "xy": [0, 0]}], srcs)
            self.assertEqual(
                [{"c": [100.0, 0.0], "r": 5.0, "kind": "하향식"}], kovs)
            # override 는 kind 만 — 이음 개수 불변
            hk = apply_kind_overrides(
                [{"c": (100.0, 0.0), "head_r": 5.0, "kind": "상향식"},
                 {"c": (200.0, 0.0), "head_r": 5.0, "kind": "상향식"}],
                kovs)
            self.assertEqual("하향식", hk[0]["kind"])
            self.assertEqual("상향식", hk[1]["kind"])
            # 구 JSON(필드 없음) → 빈 override
            path2 = user_edits_path("__old__", td)
            with open(path2, "w", encoding="utf-8") as f:
                json.dump({"version": 1, "joins": [], "deletes": [],
                           "sources": []}, f)
            _p, _e, _s, kovs2 = apply_user_edits("__old__", pts, edges, td)
            self.assertEqual([], kovs2)
            path3 = user_edits_path("__v__", td)
            with open(path3, "w", encoding="utf-8") as f:
                json.dump({
                    "version": 1, "joins": [], "deletes": [],
                    "sources": [],
                    "valve_picks": [{"xy": [40.0, 0.0]}],
                }, f)
            vpts, vedges, _s, _k = apply_user_edits(
                "__v__", [(0.0, 0.0), (100.0, 0.0)], {(0, 1)}, td)
            self.assertEqual(3, len(vpts))
            self.assertEqual(2, len(vedges))
            self.assertEqual(
                [{"xy": [40.0, 0.0]}], user_edit_valve_picks("__v__", td))


if __name__ == "__main__":
    unittest.main(verbosity=2)

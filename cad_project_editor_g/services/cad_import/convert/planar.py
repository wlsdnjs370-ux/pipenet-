# -*- coding: utf-8 -*-
"""시제품 최종(유저손질 포함) → K-Fire .kfp (plan · Z=0).

기본은 MF101_흰색점선범위(뒤호환). `python _tmp_export_mf101_kfp.py <KEY> [OUT]`
으로 다른 도면(예: 3F)도 같은 경로로 낸다. 이음 로직은 건드리지 않고 망만 직렬화한다.

    from services.cad_import.convert.planar import build_planar_graph

1) 급수원에서 물길로 닿는 배관·헤드만 (편집은 여러 개 합집합,
   KFP 변환은 지정한 1개 — 2개 이상이면 selected_source 필수)
2) 헤드로 물을 안 나르는 막다른 관 삭제 (헤드 끝쪽 load=0 은 1단만 남김)
3) 관말 헤드에 배관이 하나뿐이면 그 관은 그대로 둔다
4) 티 허브 겹침 정규화 — 관통 배관이 티 허브 위를 지나며 스텁과 겹치는
   확인된 패턴만 허브에서 쪼갠다(main—tee—main). convert.normalize 참조.
   [2026-08-11 오너: 2단계 이음은 불변 · 변환 시점에서만 고친다]
5) 본체 노드정리(SSOT) — 편집기 그래프 완성 후 저장 «직전»에
   PipeEditor.cleanup_collinear_intermediate_nodes 를 부른다(재구현 금지).
   반드시 티 정규화(4) 다음 — 겹침이 남아 있으면 병합 배관 검증에 걸려
   정리가 덜 된다. [2026-08-11 오너 확정]
"""
from __future__ import annotations

import heapq
import json
import math
import os
from collections import defaultdict, deque
from types import SimpleNamespace

from services.cad_import.convert.normalize import normalize_tee_overlaps
from services.cad_import.kinds import disk_kind_list
from services.cad_import.pipeline import flow as fw
from services.cad_import.pipeline.disp_cache import _disp_cache_load
from services.cad_import.pipeline.expand import _spec_path, stage1_body
from services.cad_import.pipeline.user_net import (
    apply_kind_overrides, apply_user_edits)
from domain import models
from domain.node_meta_factory import build_attribute_apply_meta
from domain.pipe_sizing import pipe_specs_for_standard
from editor_core import DEFAULT_DESIGN_SETTINGS, PipeEditor
from services.library_service import canonicalize_category_id
from services.pipenet_import import _ensure_default_libraries

SOURCE_SELECTION_REQUIRED = "source_selection_required"

KEY = "MF101_흰색점선범위"
OUT = os.path.join(os.path.expanduser("~"), "Desktop", f"{KEY}_유저정리5.kfp")

OPT_HEAD_K = 80.0
OPT_HEAD_SPEC = "80(5.6)"
OPT_PIPE_STD = "KSD3507"
OPT_PIPE_C = 120.0
# 설계기본설정 SSOT — NFPA 습식강관 ε0.100 (= 대치 C120). 옛 0.085 하드코딩 폐기.
OPT_PIPE_ROUGHNESS = float(DEFAULT_DESIGN_SETTINGS["roughness_mm"])
OPT_BASE_Z = 0.0
DEFAULT_DN = 25
GRID_M = 0.05


def default_out(key):
    """키별 기본 저장 경로 — MF101 은 기존 이름 그대로(뒤호환)."""
    if key == KEY:
        return OUT
    return os.path.join(os.path.expanduser("~"), "Desktop", f"{key}_유저정리.kfp")


def _empty_planar(error, code=None):
    return {
        "ok": False,
        "error": error or "평면 그래프를 만들지 못했습니다.",
        "code": code,
        "path": None,
        "kfp": None,
        "node_head_kinds": {},
        "head_kinds": [],
        "hcov": [],
        "sources": [],
    }


def build_planar_graph(key, out=None, write=False, **graph):
    """기존 export(main) 경로의 얇은 별칭. 변환은 기본으로 파일을 안 쓴다."""
    try:
        return main(key, out, write=write, **graph)
    except SystemExit as exc:
        return _empty_planar(str(exc))
    except OSError as exc:
        return _empty_planar(str(exc))
    except json.JSONDecodeError as exc:
        return _empty_planar(str(exc))


def _source_xy(src):
    xy = src.get("xy") if isinstance(src, dict) else src
    if not xy or len(xy) < 2:
        return None
    return xy


def _source_tag(src, i):
    if isinstance(src, dict) and src.get("tag"):
        return str(src["tag"])
    return f"Z{i + 1}"


def pick_convert_sources(sources, selected=None):
    """KFP 변환용 급수원 1개. 편집의 다중 sources[] 는 건드리지 않는다.

    1개면 그대로. 2개 이상이면 selected(태그 또는 1부터 번호)가 필요하다.
    반환: (고른 목록, None) 또는 (None, (code, message)).
    """
    srcs = [s for s in (sources or ()) if _source_xy(s) is not None]
    if len(srcs) <= 1:
        return srcs, None
    key = None if selected is None else str(selected).strip()
    if not key:
        return None, (
            SOURCE_SELECTION_REQUIRED,
            "급수원이 여러 개입니다. 변환할 급수원 하나를 지정하세요.",
        )
    hits = []
    for i, src in enumerate(srcs):
        if key == _source_tag(src, i) or key == str(i + 1):
            hits.append(src)
    if len(hits) == 1:
        return hits, None
    return None, (
        SOURCE_SELECTION_REQUIRED,
        f"급수원 '{key}'를 찾지 못했습니다.",
    )


def wet_from_sources(pts, edges, user_sources):
    """급수원마다 스냅한 뒤 BFS. 여러 급수원이면 닿는 배관 합집합."""
    g = SimpleNamespace(pts=pts, edges=edges)
    seed = []
    for i, src in enumerate(user_sources or ()):
        tag = src.get("tag") if isinstance(src, dict) else None
        xy = src.get("xy") if isinstance(src, dict) else src
        if not xy or len(xy) < 2:
            continue
        tag = tag or f"Z{i + 1}"
        sx, sy = float(xy[0]), float(xy[1])
        d, e = fw.snap(g, sx, sy)
        if e is None or d > fw.SRC_SNAP:
            print(f"  ★급수원 {tag}: 스냅 실패 ({d:.0f}mm)")
            continue
        seed += list(e)
        print(f"  급수원 {tag}: 스냅 {d:.0f}mm → 관 "
              f"({pts[e[0]][0]:.0f},{pts[e[0]][1]:.0f})–"
              f"({pts[e[1]][0]:.0f},{pts[e[1]][1]:.0f})")
    if not seed:
        raise SystemExit("급수원 스냅 실패 — 물길 필터 불가")

    adj = defaultdict(set)
    for i, j in edges:
        adj[i].add(j)
        adj[j].add(i)
    reach = set(seed)
    q = deque(seed)
    while q:
        u = q.popleft()
        for v in adj.get(u, ()):
            if v not in reach:
                reach.add(v)
                q.append(v)
    wet = {(min(i, j), max(i, j)) for (i, j) in edges
           if i in reach and j in reach}
    return reach, wet, list(dict.fromkeys(seed))


def prune_dead_pipes(pts, edges, head_vids, seed):
    """헤드로 물을 안 나르는 간선 삭제. 헤드 끝배관 1단은 보호.

    옛 water_cleanup 부하 셈과 같다: 급수원→헤드 최단경로 간선만 부하>0.
    인라인 헤드의 막다른 쪽(load=0)은 헤드에 붙은 간선 하나만 남기고
    그 너머는 삭제한다. 배관이 하나뿐인 관말 헤드는 그대로 둔다.
    """
    adj = defaultdict(set)
    for i, j in edges:
        adj[i].add(j)
        adj[j].add(i)

    dist, prev = {}, {}
    pq = []
    for s in seed:
        if s not in adj and s not in head_vids:
            continue
        dist[s] = 0.0
        heapq.heappush(pq, (0.0, s))
    while pq:
        d0, u = heapq.heappop(pq)
        if d0 > dist.get(u, 1e18) + 1e-9:
            continue
        for v in adj.get(u, ()):
            nd = d0 + math.hypot(pts[v][0] - pts[u][0],
                                 pts[v][1] - pts[u][1])
            if nd < dist.get(v, 1e18) - 1e-9:
                dist[v] = nd
                prev[v] = u
                heapq.heappush(pq, (nd, v))

    load = defaultdict(int)
    n_path = 0
    for h in head_vids:
        if h not in dist:
            continue
        n_path += 1
        n = h
        while n in prev:
            p = prev[n]
            load[(min(n, p), max(n, p))] += 1
            n = p

    deg = defaultdict(int)
    for i, j in edges:
        deg[i] += 1
        deg[j] += 1

    # 관말 = 헤드이면서 현재 배관이 하나뿐 → 그 관 보호
    terminal = {h for h in head_vids if deg.get(h, 0) == 1}
    protect = set()
    kept_sole = 0
    for h in terminal:
        for v in adj.get(h, ()):
            e = (min(h, v), max(h, v))
            if load.get(e, 0) == 0:
                protect.add(e)
                kept_sole += 1

    # 인라인 헤드: 막다른 쪽(load=0) 간선이 있으면 헤드에 붙은 것 1개만 보호
    kept_stub = 0
    for h in head_vids:
        if h in terminal:
            continue
        dead_inc = []
        for v in adj.get(h, ()):
            e = (min(h, v), max(h, v))
            if load.get(e, 0) > 0:
                continue
            ln = math.hypot(pts[v][0] - pts[h][0], pts[v][1] - pts[h][1])
            dead_inc.append((ln, e))
        if not dead_inc:
            continue
        dead_inc.sort()
        protect.add(dead_inc[0][1])
        kept_stub += 1

    # 급수원 관 보호 — 막다른이 아니라 물이 들어오는 시작이라 삭제하지 않는다
    seed_set = set(seed)
    for i, j in edges:
        if i in seed_set and j in seed_set:
            protect.add((min(i, j), max(i, j)))

    dead = set()
    for e in edges:
        if load.get(e, 0) > 0:
            continue
        if e in protect:
            continue
        dead.add(e)

    kept = {e for e in edges if e not in dead}
    return kept, dead, terminal, kept_sole, n_path, kept_stub


def _pipe_comps(graph):
    par = {}

    def find(a):
        while par[a] != a:
            par[a] = par[par[a]]
            a = par[a]
        return a

    for p in graph.pipes.values():
        par.setdefault(p.start, p.start)
        par.setdefault(p.end, p.end)
        ra, rb = find(p.start), find(p.end)
        if ra != rb:
            par[ra] = rb
    return {find(a) for a in par}


# «원래 일직선» 공선 판정 — MF101_4F 실측: 일직선 런 횡이탈 ≤3.6mm,
# 진짜 굴곡(90° 회향) ≥166mm. 그 사이 간극이 커서 10mm로 둔다.
COLLINEAR_TOL_MM = 10.0

# 사선 전용 게이트 — 스냅 전 부분사슬 양끝 방향의 축(가로/세로) 이탈이
# 이 값(°) 이상일 때만 복원한다.
# 실측 근거(8도면): 통과 최소 8.84° · 차단 최대 1.98° — 복원 대상
# MF101_4F 44.91~45.00°·MF 45.00°(스냅 전) / 제외 대상 = 축 방향 런.
DIAGONAL_GATE_DEG = 5.0


def _straight_restore_map(pts, remap, snap_edges, head_vid, user_sources,
                          used, xform):
    """«원래 일직선이던» 차수2 중간 점의 복원 좌표표 {정본 vid: (x, y)}.

    격자 스냅은 노드마다 독립이라 원래 일직선인 사선 런이 톱니가 된다.
    스냅 전 mm 좌표가 일직선인 부분사슬의 중간 점만 «스냅된 양끝을 잇는
    직선» 위 선형 보간으로 되돌린다. 사슬 양끝(차수≠2·헤드·급수원)과
    진짜 굴곡점은 스냅 그대로 둔다. 축에 평행한(<DIAGONAL_GATE_DEG)
    부분사슬은 복원하지 않는다. 표에 없는 노드는 건드리지 않는다.
    """
    adj = defaultdict(set)
    for a, b in snap_edges:
        adj[a].add(b)
        adj[b].add(a)
    anchor = {v for v in adj if len(adj[v]) != 2}
    anchor |= {remap[v] for v in head_vid}
    for src in user_sources or ():
        xy = _source_xy(src)
        if not xy:
            continue
        best = min(used, key=lambda v: (pts[v][0] - xy[0]) ** 2
                   + (pts[v][1] - xy[1]) ** 2)
        if math.hypot(pts[best][0] - xy[0],
                      pts[best][1] - xy[1]) <= fw.SRC_SNAP:
            anchor.add(remap[best])

    def lat_mm(p, a, b):
        dx, dy = b[0] - a[0], b[1] - a[1]
        L = math.hypot(dx, dy)
        if L < 1e-9:
            return float("inf")
        return abs((p[0] - a[0]) * dy - (p[1] - a[1]) * dx) / L

    out = {}
    seen = set()
    for v0 in list(anchor):
        for v1 in adj.get(v0, ()):
            if tuple(sorted((v0, v1))) in seen:
                continue
            seen.add(tuple(sorted((v0, v1))))
            chain = [v0, v1]
            while chain[-1] not in anchor:
                nxts = [w for w in adj[chain[-1]] if w != chain[-2]]
                if not nxts:
                    break
                e = tuple(sorted((chain[-1], nxts[0])))
                if e in seen:
                    break
                seen.add(e)
                chain.append(nxts[0])
            if chain[-1] not in anchor or len(chain) < 3:
                continue  # 고리·중간점 없는 사슬은 복원할 것이 없다
            # 최대 일직선 부분사슬 — 스냅 전 좌표에서 현 이탈이 큰 점
            # (진짜 굴곡)을 경계로 잘라 양쪽만 따로 본다
            segs = []
            stack = [(0, len(chain) - 1)]
            while stack:
                i0, i1 = stack.pop()
                if i1 - i0 < 2:
                    continue
                a, b = pts[chain[i0]], pts[chain[i1]]
                wi, worst = -1, COLLINEAR_TOL_MM
                for k in range(i0 + 1, i1):
                    d = lat_mm(pts[chain[k]], a, b)
                    if d > worst:
                        wi, worst = k, d
                if wi < 0:
                    segs.append((i0, i1))
                else:
                    stack.append((i0, wi))
                    stack.append((wi, i1))
            for i0, i1 in segs:
                a_mm, b_mm = pts[chain[i0]], pts[chain[i1]]
                dx, dy = b_mm[0] - a_mm[0], b_mm[1] - a_mm[1]
                L2 = dx * dx + dy * dy
                if L2 < 1e-9:
                    continue
                th = math.degrees(math.atan2(dy, dx)) % 180.0
                if min(th, 180.0 - th, abs(th - 90.0)) < DIAGONAL_GATE_DEG:
                    continue  # 축에 평행한 런은 격자 스냅 그대로 둔다
                A, B = xform(*a_mm), xform(*b_mm)  # 양끝은 스냅 유지
                tmp = {}
                for k in range(i0 + 1, i1):
                    p = pts[chain[k]]
                    t = ((p[0] - a_mm[0]) * dx
                         + (p[1] - a_mm[1]) * dy) / L2
                    if not 0.0 <= t <= 1.0:
                        break
                    # 0.1mm 자리 — 1mm 반올림이면 짧은 토막 각도가
                    # ±0.2° 흔들려 복원 의미가 없어진다
                    tmp[chain[k]] = (round(A[0] + (B[0] - A[0]) * t, 4),
                                     round(A[1] + (B[1] - A[1]) * t, 4))
                else:
                    out.update(tmp)
    return out


# «팔 전용 꺾임점» 최소 각도 — 원본에서 팔의 티 직전 마디 꺾임이 이 값(°)
# 미만이면 곧은(런 위) 중간점으로 보고 보존하지 않는다.
# 실측 근거: 보존 대상 MF3 팔 꺾임 90° / 제외 대상 MF4 곧은 팔 중간점 0°.
JOG_MIN_DEG = 5.0


def _arm_shape_map(pts, remap, snap_edges, head_vid, used, xform,
                   restore_xy):
    """차수1 헤드 팔의 «티에 붙는 마지막 토막» 방향을 원본 그대로 둔다.

    [2026-08-18 오너 승인 A″] 팔 양끝의 독립 격자 스냅이 짧은 토막을
    꺾으므로, 티 쪽 끝은 격자 스냅 그대로 두고 티 직전 마디 하나만
    «스냅된 티 + 원본 벡터»로 옮긴다. 단,
      · 직전 마디가 원본에서 꺾임점(≥JOG_MIN_DEG)일 때만 — 곧은(런 위)
        중간점을 옮기면 런 정렬과 노드정리 병합이 깨진다(MF4 시뮬 +157)
      · 티토막이 축(가로/세로) 정렬이면 건너뛴다 — 꺾임 개선이 없고
        노드정리의 축 병합만 깨진다(MF4 회귀 +151 교훈)
      · 앞 토막이 축 정렬이면 건너뛴다 — 옮기면 그 토막이 새로 꺾여
        하나를 고치려다 다른 하나를 꺾는 셈이 된다
      · 티-헤드 직결 한 토막 팔은 헤드를 옮긴다(오너 명시 승인)
      · 직선 복원 대상 마디는 복원이 우선
      · 키 충돌·미세 간선이 생기는 팔은 그 팔만 건너뛴다
    반환: ({정본 vid: (x, y)}, 보존한 팔 수)
    """
    adj = defaultdict(set)
    for a, b in snap_edges:
        adj[a].add(b)
        adj[b].add(a)

    out = {}

    def cur_fin(v):
        p = out.get(v) or restore_xy.get(v)
        return p if p is not None else xform(*pts[v])

    # make_node 키(0.1mm 자리) 기준 최종 위치표 — 팔별 충돌 검사용
    keys = {}
    for v in {remap[v] for v in used}:
        keys[(round(cur_fin(v)[0], 3), round(cur_fin(v)[1], 3),
              round(OPT_BASE_Z, 3))] = v

    n_arm = 0
    for vid in head_vid:
        hv = remap[vid]
        if len(adj.get(hv, ())) != 1:
            continue  # 관통형(인라인) 헤드는 팔이 없다
        path = [hv]
        tee = None
        while True:  # 차수1 헤드에서 차수≠2 마디(티)까지 팔 따라가기
            prev = path[-2] if len(path) > 1 else None
            nxts = [w for w in adj[path[-1]] if w != prev]
            if not nxts:
                break
            path.append(nxts[0])
            if len(adj[nxts[0]]) != 2:
                tee = nxts[0]
                break
        if tee is None:
            continue  # 막다른 팔 — 붙일 티가 없다
        last = path[-2]
        if last in restore_xy or last in out:
            continue  # 직선 복원이 우선
        if last != hv:
            # 곧은(런 위) 중간점 제외 — 직전 마디가 꺾임점일 때만 옮긴다
            pre = path[-3]
            d1x, d1y = (pts[last][0] - pts[pre][0],
                        pts[last][1] - pts[pre][1])
            d2x, d2y = (pts[tee][0] - pts[last][0],
                        pts[tee][1] - pts[last][1])
            L1, L2 = math.hypot(d1x, d1y), math.hypot(d2x, d2y)
            if L1 < 1e-9 or L2 < 1e-9:
                continue
            if abs(d1x * d2y - d1y * d2x) / (L1 * L2) < math.sin(
                    math.radians(JOG_MIN_DEG)):
                continue
            # 런 위의 점 제외 — 티토막이 축 정렬이면 옮겨도 꺾임 개선이
            # 없고 노드정리의 축 병합만 깨진다(MF4 회귀 +151 교훈).
            th_b = math.degrees(math.atan2(d2y, d2x)) % 180.0
            if min(th_b, 180.0 - th_b, abs(th_b - 90.0)) < JOG_MIN_DEG:
                continue
            # 앞 토막이 축 정렬이면 옮겼을 때 그 토막이 새로 꺾인다
            # — 티토막 꺾임을 고치려고 앞 토막을 꺾는 셈이라 건너뛴다.
            th_a = math.degrees(math.atan2(d1y, d1x)) % 180.0
            if min(th_a, 180.0 - th_a, abs(th_a - 90.0)) < JOG_MIN_DEG:
                continue
        ts = xform(*pts[tee])  # 티 쪽 끝은 격자 스냅 유지
        fx = round(ts[0] + (pts[last][0] - pts[tee][0]) / 1000.0, 4)
        fy = round(ts[1] + (pts[last][1] - pts[tee][1]) / 1000.0, 4)
        old_key = (round(cur_fin(last)[0], 3), round(cur_fin(last)[1], 3),
                   round(OPT_BASE_Z, 3))
        key = (round(fx, 3), round(fy, 3), round(OPT_BASE_Z, 3))
        if keys.get(key, last) != last:
            continue  # 키 충돌 — 이 팔만 건너뛴다
        if any(abs(fx - cur_fin(w)[0]) + abs(fy - cur_fin(w)[1])
               < GRID_M / 2 for w in adj[last]):
            continue  # 미세 간선 — 이 팔만 건너뛴다
        keys.pop(old_key, None)
        keys[key] = last
        out[last] = (fx, fy)
        n_arm += 1
    return out, n_arm


def main(key=KEY, out=None, *, write=True, pts=None, edges=None, hcov=None,
         ups=None, head_kinds=None, user_sources=None, selected_source=None,
         ho=None):
    if write and out is None:
        out = default_out(key)
    if not write:
        out = None
    kind_ovs = []
    given_ho = list(ho or ())
    ho = []
    if pts is not None:
        pts = [tuple(p) for p in pts]
        edges = {tuple(sorted(e)) for e in (edges or ())}
        hcov = [tuple(h) for h in (hcov or ())]
        ups = [tuple(u) for u in (ups or ())]
        user_sources = list(user_sources or ())
        head_kinds = fw.require_head_kinds(hcov, head_kinds or ())
        ho = given_ho
        if ho and not any(h.get("sa") is not None
                          and h.get("sweep") is not None for h in ho):
            try:
                spec = json.load(open(_spec_path(key), encoding="utf-8"))
            except Exception:  # noqa: BLE001
                spec = {}
            spec_ho = fw.ho_from_spec(spec)
            if spec_ho:
                ho = spec_ho
        print(f"현재 편집 리비전: pts={len(pts)} · 간선={len(edges)}"
              f" · 급수원={len(user_sources)}")
    else:
        cached = _disp_cache_load(key)
        if cached is not None:
            print(f"[캐시 HIT] {key}")
            pts = [tuple(p) for p in cached["pts"]]
            edges = {tuple(e) for e in cached["edges"]}
            hcov = [tuple(h) for h in cached["hcov"]]
            ups = [tuple(u) for u in (cached.get("ups") or ())]
            head_kinds = list(cached.get("head_kinds") or ())
            ho = list(cached.get("ho") or ())
            if ho and not any(h.get("sa") is not None
                              and h.get("sweep") is not None for h in ho):
                try:
                    spec = json.load(open(_spec_path(key), encoding="utf-8"))
                except Exception:  # noqa: BLE001
                    spec = {}
                spec_ho = fw.ho_from_spec(spec)
                if spec_ho:
                    ho = spec_ho
        else:
            print(f"[캐시 MISS] pipeline 재실행…")
            st = stage1_body(key)
            P = fw.pipeline(st, key=None)
            pts, edges, hcov = P["pts"], set(P["edges"]), P["hcov"]
            ups = P.get("ups") or ()
            head_kinds = list(P.get("head_kinds") or ())
            ho = fw.ho_from_spots(P.get("spots"))
            if not any(h.get("sa") is not None and h.get("sweep") is not None
                       for h in ho):
                spec_ho = fw.ho_from_spec(st.get("spec"))
                if spec_ho:
                    ho = spec_ho

        pts, edges, user_sources, kind_ovs = apply_user_edits(
            key, pts, edges, fw.default_edits_dir())
        edges = {tuple(sorted(e)) for e in edges}
        head_kinds = fw.require_head_kinds(
            hcov, apply_kind_overrides(head_kinds, kind_ovs))
    kind_n = {}
    for rec in head_kinds:
        k = fw.normalize_head_kind(rec.get("kind"))
        kind_n[k] = kind_n.get(k, 0) + 1
    print(f"손질 후: pts={len(pts)} · 간선={len(edges)}"
          f" · 급수원={len(user_sources)} · kind_ov={len(kind_ovs)}"
          f" · 종류 {kind_n}")
    user_sources, src_err = pick_convert_sources(user_sources, selected_source)
    if src_err:
        code, msg = src_err
        print(f"★{msg}")
        return _empty_planar(msg, code)

    # ★헤드 접속 완성 [2026-08-13 오너] — 편집 최종망은 헤드가 «중심 노드»로
    #   연결된 상태다(보드와 같은 SSOT·idempotent). 변환은 그 확정을 읽기만
    #   하고, 아래에서 «중심 최단» 재선정을 하지 않는다.
    pts, edges, head_centers, n_wire, multi_arm = fw.attach_heads_center(
        pts, edges, hcov)
    if n_wire or multi_arm:
        print(f"헤드 중심접속 완성: 새 연결 {n_wire}"
              f" · 팔 박빙 {len(multi_arm)}곳")
    if multi_arm:
        print(f"    ★팔 박빙 {len(multi_arm)}개 — 빨간색 테두리 헤드는"
              f" 최종 변환후 정상 연결 여부를 확인하세요")

    reach, wet, seed = wet_from_sources(pts, edges, user_sources)
    n_all, m_all = len(edges), fw.mlen(pts, edges)
    edges = wet
    print(f"물길 필터: 간선 {len(edges)}/{n_all}"
          f" · {fw.mlen(pts, edges):.1f}m / {m_all:.1f}m"
          f" · 도달노드 {len(reach)}")

    hnodes = fw.head_nodes(pts, hcov, edges=edges, upright=ups)
    kinds_aligned = disk_kind_list(hcov, head_kinds)
    ups_keys = {(round(x, 1), round(y, 1)) for (x, y, _r) in ups}
    head_vid = {}
    head_kind_by_vid = {}
    n_dry_head = 0
    n_center = 0
    wet_head_idx: list = []      # [G2] 물닿음으로 인정된 hcov 인덱스
    for _h_i, ((hx, hy, _hr), ctr, ns, kind) in enumerate(zip(
            hcov, head_centers, hnodes, kinds_aligned)):
        vid = None
        if ctr is not None and ctr in reach:
            # 편집 완성이 확정한 중심 노드 1:1 — 재선정 없음
            vid = ctr
            n_center += 1
        elif ctr is None and (round(hx, 1), round(hy, 1)) in ups_keys:
            # 상향식 원 밑 통과 — 중심 노드가 없을 때만 예전 문 그대로
            wet_ns = {n for n in ns if n in reach}
            if wet_ns:
                vid = min(wet_ns, key=lambda n: (pts[n][0] - hx) ** 2
                          + (pts[n][1] - hy) ** 2)
        if vid is None:
            n_dry_head += 1
            continue
        head_vid[vid] = (hx, hy)
        head_kind_by_vid[vid] = kind
        # [G2] 이 전개가 «물닿음» 으로 인정한 hcov 번호. 최불리 선정이 board 의
        # 도달 판정을 쓰는데 이쪽이 더 엄격해(중심 노드가 물길에 닿아야 한다),
        # 두 집합이 어긋나면 제한 전개가 통째로 가지치기된다. 밖에서 대조할 수
        # 있도록 내보낸다 — 계산에는 쓰지 않으므로 .kfp 는 그대로다.
        wet_head_idx.append(_h_i)
    print(f"헤드: 물닿음 {len(head_vid)} · 마른/미부착 {n_dry_head}"
          f" · 중심접속 {n_center}")

    before_m = fw.mlen(pts, edges)
    edges, dead, terminal, kept_sole, n_path, kept_stub = prune_dead_pipes(
        pts, edges, set(head_vid), seed)
    print(f"막다른관 삭제: {len(dead)}개 · {fw.mlen(pts, dead):.1f}m"
          f" (남김 {len(edges)} · {fw.mlen(pts, edges):.1f}m"
          f" / 직전 {before_m:.1f}m)"
          f" · 관말헤드 {len(terminal)} · 관말팔보호 {kept_sole}"
          f" · 끝배관1단보호 {kept_stub}"
          f" · 경로닿은헤드 {n_path}")

    used = set()
    for i, j in edges:
        used.add(i)
        used.add(j)
    # 간선이 남은 헤드만 (관말 팔 보호로 살아남은 것 포함)
    head_vid = {v: xy for v, xy in head_vid.items() if v in used}
    used |= set(head_vid)
    minx = min(pts[v][0] for v in used)
    miny = min(pts[v][1] for v in used)

    def xform(x_mm, y_mm):
        mx = (x_mm - minx) / 1000.0 + 1.0
        my = (y_mm - miny) / 1000.0 + 1.0
        return (round(round(mx / GRID_M) * GRID_M, 3),
                round(round(my / GRID_M) * GRID_M, 3))

    pos_vid, remap = {}, {}
    for vid in sorted(used, key=lambda v: (v not in head_vid, v)):
        p = xform(*pts[vid])
        if p in pos_vid:
            remap[vid] = pos_vid[p]
        else:
            pos_vid[p] = vid
            remap[vid] = vid

    # ---- 티 허브 겹침 정규화 (변환 시점 전용 · 2단계 이음 불변)
    #  스냅 좌표에서 «관통 배관이 티 허브 위를 지나며 스텁과 겹치는»
    #  확인된 패턴만 허브에서 쪼갠다: main—tee—main. 새 선은 만들지 않고
    #  있던 관통을 허브에서 나눌 뿐이라 총 기하·물길은 그대로다.
    snap_edges = {tuple(sorted((remap[i], remap[j])))
                  for (i, j) in edges if remap[i] != remap[j]}
    # [G2] 역참조 — 눌린 간선이 «어느 board 간선에서 왔는지». 여러 board 간선이
    # 한 자리로 눌리면 첫 것을 대표로 둔다(관경 매칭은 선분 위치만 쓰므로
    # 대표 하나로 충분하다). 이 표가 없으면 G3 의 관경이 엉뚱한 배관에 붙는다.
    _snap_origin: dict = {}
    for (i, j) in edges:
        if remap[i] == remap[j]:
            continue
        _snap_origin.setdefault(tuple(sorted((remap[i], remap[j]))), (i, j))
    snap_pos = {v: xform(*pts[v]) for e in snap_edges for v in e}
    snap_edges, n_tee_split = normalize_tee_overlaps(snap_pos, snap_edges)
    print(f"티 겹침 정규화: 관통 쪼갬 {n_tee_split}곳 · 간선 {len(snap_edges)}")

    # ---- 직선 위치 복원 [2026-08-18 오너]
    #  격자 스냅으로 톱니가 된 «원래 일직선» 런의 중간 점만 스냅된 양끝을
    #  잇는 직선 위로 되돌린다. 점 개수·연결·양끝 위치는 불변.
    restore_xy = _straight_restore_map(pts, remap, snap_edges, head_vid,
                                       user_sources, used, xform)

    # ---- 헤드팔 모양 보존 [2026-08-18 오너 승인]
    #  티에 붙는 마지막 토막의 방향을 원본 그대로 둔다(티 쪽 끝은 스냅 유지).
    #  곧은 런 위 점·복원 대상·문제 팔은 건너뛴다 — 그 팔만.
    arm_xy, n_arm = _arm_shape_map(pts, remap, snap_edges, head_vid, used,
                                   xform, restore_xy)
    n_restore = len(restore_xy)
    restore_xy.update(arm_xy)

    def final_xy(vid):
        p = restore_xy.get(vid)
        return p if p is not None else xform(*pts[vid])

    if restore_xy:
        canon = {remap[v] for v in used}
        keys = [(round(final_xy(v)[0], 3), round(final_xy(v)[1], 3),
                 round(OPT_BASE_Z, 3)) for v in canon]
        clash = len(set(keys)) != len(keys)  # make_node 키 충돌 검사
        short = any(  # 25mm 미세 간선 스킵(아래)에 걸릴 간선 검사
            sum(abs(final_xy(a)[k] - final_xy(b)[k]) for k in range(2))
            < GRID_M / 2 for a, b in snap_edges)
        if clash or short:
            print(f"★직선 위치 복원 취소 — 키충돌={clash} 미세간선={short}"
                  " (격자 스냅 그대로)")
            restore_xy = {}
        else:
            print(f"직선 위치 복원: {n_restore}노드"
                  + (f" · 팔 모양 보존 {n_arm}팔({len(arm_xy)}노드)"
                     if arm_xy else ""))

    editor = PipeEditor()
    _ensure_default_libraries(editor)
    specs = {s.nominal_mm: s
             for s in pipe_specs_for_standard(editor.pipe_library, OPT_PIPE_STD)}
    spec = specs.get(DEFAULT_DN) or next(iter(specs.values()))

    node_seq = 0
    node_id = {}
    node_cache = {}
    n_head_upgrade = 0
    head_meta = {
        "type_id": "head",
        "k_factor_si": OPT_HEAD_K,
        "head_spec_name": OPT_HEAD_SPEC,
        "required_pressure_bar": 0.0,
    }

    def make_node(x, y, z, node_type="기본", meta=None):
        nonlocal node_seq, n_head_upgrade
        key = (round(x, 3), round(y, 3), round(z, 3))
        cached_nid = node_cache.get(key)
        if cached_nid is not None:
            if node_type == "Head":
                nd = editor.graph.get_node(cached_nid)
                if str(getattr(nd, "type_id", "")) != "head":
                    nd.type = "Head"
                    nd.update_from_meta(meta or {})
                    n_head_upgrade += 1
            return cached_nid
        node_seq += 1
        nid = f"N{node_seq}"
        node = models.Node.create(nid, x, y, z, node_type=node_type)
        if meta:
            node.update_from_meta(meta)
        editor.graph.add_node(node)
        node_cache[key] = nid
        return nid

    node_ref: dict = {}          # [G2] kfp 노드 id → board 노드 인덱스
    for vid in used:
        tgt = remap[vid]
        if tgt in node_id:
            continue
        x, y = final_xy(tgt)
        node_id[tgt] = make_node(x, y, OPT_BASE_Z)
        node_ref.setdefault(node_id[tgt], vid)

    head_nids = set()
    node_head_kinds = {}
    for vid in head_vid:
        tgt = remap[vid]
        nid = node_id[tgt]
        x, y = final_xy(tgt)
        make_node(x, y, OPT_BASE_Z, "Head", dict(head_meta))
        head_nids.add(nid)
        kind = head_kind_by_vid.get(vid)
        prev = node_head_kinds.get(nid)
        if prev is not None and prev != kind:
            node_head_kinds.pop(nid, None)
        elif kind:
            node_head_kinds[nid] = kind

    pipe_seq = 0
    seen_pairs = set()
    edge_ref: dict = {}          # [G2] kfp 배관 id → (board_i, board_j)
    for ri, rj in sorted(snap_edges):
        a_id, b_id = node_id[ri], node_id[rj]
        key = tuple(sorted((a_id, b_id)))
        if key in seen_pairs:
            continue
        na = editor.graph.get_node(a_id).coords
        nb = editor.graph.get_node(b_id).coords
        if sum(abs(na[k] - nb[k]) for k in range(3)) < GRID_M / 2:
            continue
        seen_pairs.add(key)
        pipe_seq += 1
        origin = _snap_origin.get((ri, rj))
        if origin is not None:
            edge_ref[f"P{pipe_seq}"] = origin
        pipe = models.Pipe(f"P{pipe_seq}", a_id, b_id)
        pipe.type = OPT_PIPE_STD
        pipe.nominal_mm = spec.nominal_mm
        pipe.diameter = spec.inner_d_mm
        pipe.length_m = round(sum(abs(na[k] - nb[k]) for k in range(3)), 3)
        pipe.equivalent_length = 0.0
        pipe.C = OPT_PIPE_C
        pipe.roughness_mm = OPT_PIPE_ROUGHNESS
        editor.graph.add_pipe(pipe)

    n_src = 0
    for src in user_sources or ():
        tag = src.get("tag") if isinstance(src, dict) else None
        xy = src.get("xy") if isinstance(src, dict) else src
        if not xy or len(xy) < 2:
            continue
        sx, sy = float(xy[0]), float(xy[1])
        best = min(used, key=lambda v: (pts[v][0] - sx) ** 2
                   + (pts[v][1] - sy) ** 2)
        d = math.hypot(pts[best][0] - sx, pts[best][1] - sy)
        if d > fw.SRC_SNAP:
            print(f"  급수원 {tag or '?'} 스냅 실패 d={d:.0f}mm")
            continue
        nid = node_id[remap[best]]
        nd = editor.graph.get_node(nid)
        if str(getattr(nd, "type_id", "")) == "head":
            print(f"  급수원 {tag} 헤드 노드 회피 — 스킵")
            continue
        pump_meta = build_attribute_apply_meta(
            nd.to_dict(), "Pump",
            library=editor,
            canonicalize_category_id=canonicalize_category_id,
        ).meta
        nd.type = pump_meta.get("type") or "Pump"
        nd.update_from_meta(pump_meta)
        n_src += 1
        print(f"  급수원 {tag or 'Z'} → {nid} 펌프 (스냅 {d:.0f}mm)")

    editor.node_counter = {"N": node_seq}
    editor.pipe_id_counter = pipe_seq
    if hasattr(editor, "_rebind_managers"):
        editor._rebind_managers()
    if hasattr(editor, "_mark_pipe_key_index_dirty"):
        editor._mark_pipe_key_index_dirty()
    editor.update_all_pipe_lengths_in_pipe_data()
    if hasattr(editor, "_sync_all_legacy_views_from_graph"):
        editor._sync_all_legacy_views_from_graph()

    # ---- 본체 노드정리 (SSOT: 편집 메뉴 «노드정리»와 같은 함수 하나)
    #  일직선 중간 노드를 저장 전에 병합한다. 헤드(비기본 type_id)와
    #  펌프(차단 타입)는 SSOT 게이트가 후보에서 빼므로 여기서 더 할 일 없다.
    res_clean = editor.cleanup_collinear_intermediate_nodes()
    print(f"노드정리(SSOT): 삭제 {len(res_clean['removed'])}"
          f" · 실패 {len(res_clean['failed'])}")

    # [G2] 역참조 복구 — 노드정리는 일직선 배관 여럿을 하나로 **병합하고 새 id 를
    # 매긴다**. 그래서 배관 생성 때 적어 둔 edge_ref 가 최종 배관을 못 덮는다
    # (실측: 배관 53개 중 36개 미포함). 병합된 배관의 두 끝은 살아남은 원래
    # 노드이므로 node_ref 로 board 노드를 되짚어 채운다 — 병합된 조각들은
    # 일직선이라 그 두 끝이 같은 직선을 정의하고, 관경 텍스트 매칭에는 그것으로
    # 충분하다. 없는 것을 지어내지 않고 «아는 것만» 채운다.
    for _pid, _pipe in list(getattr(editor.graph, "pipes", {}).items()):
        if _pid in edge_ref:
            continue
        _bi = node_ref.get(getattr(_pipe, "start", None))
        _bj = node_ref.get(getattr(_pipe, "end", None))
        if _bi is not None and _bj is not None:
            edge_ref[_pid] = (_bi, _bj)
    # 최종 배관에 없는 낡은 키(병합으로 사라진 배관)는 버린다 — 남겨 두면
    # 「덮었다」는 착시가 생긴다.
    _final = set(getattr(editor.graph, "pipes", {}))
    for _pid in [k for k in edge_ref if k not in _final]:
        edge_ref.pop(_pid, None)

    if out:
        editor.save_json(out)
        e2 = PipeEditor()
        e2.load_json(out)
        g2 = e2.graph
        leftover = e2.collect_collinear_merge_candidates()
        tag = "재로드"
        print(f"저장: {out}")
    else:
        report = editor.validate_and_fix_integrity(strict=True)
        if report.get("errors"):
            raise SystemExit(
                "저장 차단: 무결성 오류\n" + "\n".join(report["errors"][:15]))
        g2 = editor.graph
        leftover = editor.collect_collinear_merge_candidates()
        tag = "평면 그래프 메모리"
    kfp = editor.to_dict()
    n_head = sum(1 for n in g2.nodes.values()
                 if str(getattr(n, "type_id", "")) == "head")
    n_pump = sum(1 for n in g2.nodes.values()
                 if str(getattr(n, "type_id", "")) == "pump")
    tot_m = sum(float(getattr(p, "length_m", 0) or 0)
                for p in g2.pipes.values())
    print(f"{tag}: 노드 {len(g2.nodes)} · 배관 {len(g2.pipes)}"
          f" · 헤드 {n_head} (도면 {len(hcov)} 중 물닿음)"
          f" · 펌프 {n_pump} · 연결성분 {len(_pipe_comps(g2))}"
          f" · 연장 {tot_m:.1f}m · 헤드승격 {n_head_upgrade}")
    print(f"{'재로드 정리후보' if out else '정리후보'}: {len(leftover)}"
          " (0이어야 이미 정리된 것)")
    if out:
        print(f"파일크기 {os.path.getsize(out) / 1024:.0f} KB")

    nodes = kfp.get("nodes_meta_runtime") or {}
    node_head_kinds = {
        nid: kind for nid, kind in node_head_kinds.items()
        if nid in nodes
        and str((nodes[nid] or {}).get("type_id") or "") == "head"
        and kind in ("상향식", "하향식", "상하향식")
    }
    return {
        "ok": True,
        "path": out,
        "kfp": kfp,
        "node_head_kinds": node_head_kinds,
        "head_kinds": head_kinds,
        "hcov": hcov,
        "origin_mm": (minx, miny),
        "ho": ho,
        "sources": list(user_sources or ()),
        # [G2] 역참조 — 관경(G3)이 평면 mm 좌표에서 매칭하려면 kfp 배관에서
        # 원 board 간선으로 되짚을 수 있어야 한다. 기존 키는 그대로 두고 더한다.
        "edge_ref": edge_ref,
        "node_ref": node_ref,
        "wet_head_idx": wet_head_idx,
    }


if __name__ == "__main__":
    import sys
    _r = main(sys.argv[1] if len(sys.argv) > 1 else KEY,
              sys.argv[2] if len(sys.argv) > 2 else None)
    if isinstance(_r, dict) and _r.get("ok") is False:
        raise SystemExit(_r.get("error") or 1)
